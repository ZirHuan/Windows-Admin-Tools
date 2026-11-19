#Requires -Version 5.1
#Requires -RunAsAdministrator
#Requires -Modules ActiveDirectory

<#
.SYNOPSIS
    Read-only consolidated health, inventory and security audit for a Windows Active Directory domain.

.DESCRIPTION
    Single-pass, NON-DESTRUCTIVE audit that inventories a domain and flags things that are
    commonly wrong or worth handling. Domain-agnostic; driven by parameters and an optional
    JSON config file per environment.

    NOTHING here changes state. There is no -Execute parameter. Always safe to re-run.

    Check sections (each can be skipped with -Skip):
      FOREST    Forest/domain functional level, UPN suffixes, tombstone lifetime
      AD        DC inventory, FSMO placement, replication failures, SYSVOL, DFSR/FRS,
                PDC emulator time source
      SITES     Sites, subnets, sites without subnets, DC-to-site mapping
      DNS       DC name resolution, scavenging/aging, stale A records
      CERT      LDAPS TLS handshake + cert expiry, expiring DC machine certs, CA liveness,
                autoenrollment GPOs vs dead CA, NTAuth store, unlinked GPOs,
                optional dead-CA blast-radius sweep (-CertSweep)
      USERS     Privileged accounts + risk scoring, krbtgt age, password policy,
                password-never-expires, stale users/computers, AdminSDHolder orphans,
                AS-REP roastable accounts, Protected Users
      SECURITY  Unconstrained Kerberos delegation, duplicate SPNs, LAPS coverage,
                Kerberoastable accounts
      NET       DC port reachability (LDAP/LDAPS/KRB/GC/DNS/SMB), stale IPs on 389,
                unsigned LDAP binds (2889/2887)
      LICENSE   Windows activation state on DCs (best-effort)
      EVENTS    Critical Directory Service / DFSR errors in the last N hours

    BUILDING BLOCKS. Parts that are useful on their own are separate scripts, called here.
    They are looked up next to this script first, then in the sibling folders of the repo
    (..\AD, ..\Network), so both the repo layout and a flat release folder work:
      AD\Get-FsmoRoleHolder.ps1          AD section
      AD\Get-PrivilegedAccountRisk.ps1   USERS + SECURITY sections
      AD\Get-UnsignedLdapClient.ps1      NET section
      Network\Test-TcpPort.ps1           NET section, CA liveness
      Network\Get-TlsCertificate.ps1     CERT section (LDAPS)
      DomainScanner\Get-CertBlastRadius.ps1   CERT section, only with -CertSweep
    Only the tools needed by the sections that run are required.

    Output:
      - Colour-coded console table (PASS / WARN / FAIL / INFO)
      - Structured object returned to the pipeline
      - Optional HTML report (-Html) and CSV of findings (-Csv)

.PARAMETER Domain
    FQDN of the domain to audit. Defaults to the current user's DNS domain, falling
    back to the computer's domain (USERDNSDOMAIN is empty under SYSTEM / scheduled tasks).

.PARAMETER DomainControllers
    Explicit list of DC FQDNs to probe for port/event checks. If omitted, discovered
    automatically via Get-ADDomainController.

.PARAMETER ConfigPath
    Optional path to a JSON config file describing environment-specific expectations
    (known-risk accounts, expected sites, stale IPs/DC names, expected FSMO holder,
    dead CA host/issuer, checks to skip). See sample.config.json for the format.

.PARAMETER Skip
    One or more section names to skip:
    FOREST, AD, SITES, DNS, CERT, USERS, SECURITY, NET, LICENSE, EVENTS.

.PARAMETER SkipGpoScan
    Skip the per-GPO XML report scan in the CERT section (slow on domains with many GPOs).
    Also skips the unlinked-GPO count, which comes from the same scan.

.PARAMETER CertSweep
    When a registered CA is dead (or config DeadCAIssuerMatch is set), run
    Get-CertBlastRadius.ps1 against every enabled Windows computer to find the machines
    that still hold certs from it. Needs WinRM to the estate; can take minutes.

.PARAMETER EventHoursBack
    Hours of event-log history to scan in the EVENTS and NET (2889) sections. Default 24.

.PARAMETER InactiveDays
    Threshold (days since last logon) for flagging stale users/computers. Default 90.

.PARAMETER CertExpiryWarnDays
    Warn when a DC/LDAPS certificate expires within this many days. Default 60.

.PARAMETER LogRoot
    Root folder for the log file and any reports. Default C:\Logs\DomainAudit.

.PARAMETER Html
    Also write a colour-coded HTML report suitable for hand-off / sign-off.

.PARAMETER Csv
    Also write a flat CSV of every finding.

.EXAMPLE
    .\Invoke-DomainHealthAudit.ps1 -Verbose
    Full audit of the current domain, console output only.

.EXAMPLE
    .\Invoke-DomainHealthAudit.ps1 -Domain contoso.local -ConfigPath .\contoso.config.json -Html -Csv
    Audit contoso.local using its environment config (copy of sample.config.json), write HTML + CSV.

.EXAMPLE
    .\Invoke-DomainHealthAudit.ps1 -Skip LICENSE,NET
    Skip the licensing and network-reachability sections (e.g. no remote access to DCs).

.NOTES
    Version   : 1.2.0
    Read-only : always safe to re-run

    Changelog (full history in CHANGELOG.md):
      1.2.0  2026-11-19  Merged the two diverged copies (DomainScanner 1.1.0 and
                         AD/Inventering): new SECURITY section, password policy,
                         AdminSDHolder orphans, AS-REP roastable, Protected Users,
                         unlinked GPOs, PDC time source - all rewritten to be
                         locale-independent. Reusable parts moved to separate tools
                         (see BUILDING BLOCKS). Privileged groups by SID. -CertSweep.
      1.1.0  2026-09-28  Fixes from review + lab test run (see CHANGELOG.md).
      1.0.0  2026-09     Initial consolidated seed.
#>

[CmdletBinding()]
param(
    [string]$Domain = $env:USERDNSDOMAIN,

    [string[]]$DomainControllers,

    [string]$ConfigPath = '',

    [ValidateSet('FOREST','AD','SITES','DNS','CERT','USERS','SECURITY','NET','LICENSE','EVENTS')]
    [string[]]$Skip = @(),

    [switch]$SkipGpoScan,

    [switch]$CertSweep,

    [ValidateRange(1,720)]
    [int]$EventHoursBack = 24,

    [ValidateRange(1,3650)]
    [int]$InactiveDays = 90,

    [ValidateRange(1,3650)]
    [int]$CertExpiryWarnDays = 60,

    [string]$LogRoot = 'C:\Logs\DomainAudit',

    [switch]$Html,

    [switch]$Csv
)

$ScriptVersion = '1.2.0'
$AllSections   = @('FOREST','AD','SITES','DNS','CERT','USERS','SECURITY','NET','LICENSE','EVENTS')

# ======================================================================
#  ENVIRONMENT CONFIG (defaults, overridable via -ConfigPath JSON)
# ======================================================================
$Config = [ordered]@{
    KnownRiskAccounts = @()            # sam names known to be vendor/external/orphaned
    ExpectedSites     = @()            # site names that should exist (INFO only)
    StaleDCNames      = @()            # decommissioned DC short names to hunt for
    StaleIPs          = @()            # IPs that should no longer answer LDAP
    ExpectedFsmoHolder= ''             # short name if all 5 FSMO should sit on one DC
    DeadCAHostname    = ''             # FQDN of a CA known to be decommissioned
    DeadCAIssuerMatch = ''             # issuer substring for -CertSweep (e.g. 'CN=Old-CA')
    ServiceAcctRegex  = '^(svc[-_]|sa[-_]|sql|exchange|iis|backup|print|scan|service|mssql)'
    NamedAdminRegex   = '^(adm[-_.]|admin[-_.])'   # personal admin accounts to re-confirm
}
$SkipList = @($Skip)
if ($ConfigPath) {
    if (-not (Test-Path $ConfigPath)) {
        Write-Warning "Config '$ConfigPath' not found - using defaults."
    } else {
        try {
            $loaded = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json
            foreach ($k in @($Config.Keys)) {
                if ($null -ne $loaded.$k) { $Config[$k] = $loaded.$k }
            }
            # Older AD/Inventering configs called it AdminAcctRegex.
            if ($null -eq $loaded.NamedAdminRegex -and $null -ne $loaded.AdminAcctRegex) {
                $Config.NamedAdminRegex = $loaded.AdminAcctRegex
            }
            foreach ($s in @($loaded.Skip)) {
                if (-not $s) { continue }
                if ($s -in $AllSections) { $SkipList += $s.ToUpper() }
                else { Write-Warning "Config Skip entry '$s' is not a valid section - ignored." }
            }
            $SkipList = @($SkipList | Select-Object -Unique)
        } catch {
            Write-Warning "Could not parse config '$ConfigPath': $_ - using defaults."
        }
    }
}

# USERDNSDOMAIN is empty under SYSTEM (scheduled task, guest agent) - fall back to the computer's domain.
if (-not $Domain) {
    try { $Domain = (Get-ADDomain -Current LocalComputer -ErrorAction Stop).DNSRoot } catch { }
}
if (-not $Domain) {
    throw 'Could not determine the domain. Pass -Domain <fqdn>.'
}
$Domain = $Domain.ToLower()

function Test-SectionEnabled { param([string]$Section) return ($Section -notin $SkipList) }

# ======================================================================
#  BUILDING BLOCKS (separate scripts - see .DESCRIPTION)
# ======================================================================
function Resolve-Tool {
    # Next to this script first (flat release folder), then each sibling folder of the
    # parent (repo layout: DomainScanner\..\AD, ..\Network).
    param([Parameter(Mandatory)][string]$Name)
    $candidates = @(Join-Path $PSScriptRoot $Name)
    $parent = Split-Path -Path $PSScriptRoot -Parent
    if ($parent) {
        $candidates += @(Get-ChildItem -LiteralPath $parent -Directory -ErrorAction SilentlyContinue |
                         ForEach-Object { Join-Path $_.FullName $Name })
    }
    foreach ($c in $candidates) { if (Test-Path -LiteralPath $c -PathType Leaf) { return $c } }
    return $null
}

$toolNeeds = [ordered]@{
    'Get-FsmoRoleHolder.ps1'        = @('AD')
    'Get-PrivilegedAccountRisk.ps1' = @('USERS','SECURITY')
    'Get-UnsignedLdapClient.ps1'    = @('NET')
    'Test-TcpPort.ps1'              = @('NET','CERT')
    'Get-TlsCertificate.ps1'        = @('CERT')
}
if ($CertSweep) { $toolNeeds['Get-CertBlastRadius.ps1'] = @('CERT') }
$Tool = @{}
$missing = @()
foreach ($t in $toolNeeds.GetEnumerator()) {
    if (-not @($t.Value | Where-Object { Test-SectionEnabled $_ }).Count) { continue }
    $p = Resolve-Tool $t.Key
    if ($p) { $Tool[$t.Key] = $p } else { $missing += "$($t.Key) (sections: $($t.Value -join ','))" }
}
if ($missing) {
    throw ("Building-block script(s) not found next to this script or in a sibling folder: " +
           ($missing -join '; ') + ". Copy the whole release folder, or -Skip those sections.")
}

# ======================================================================
#  SETUP
# ======================================================================
$Stamp      = Get-Date -Format 'yyyyMMdd_HHmm'
$ScriptName = 'Invoke-DomainHealthAudit'
$LogPath    = Join-Path $LogRoot ("{0}_{1}.log"  -f $ScriptName, $Stamp)
$HtmlPath   = Join-Path $LogRoot ("DomainAudit_{0}.html" -f $Stamp)
$CsvPath    = Join-Path $LogRoot ("DomainAudit_{0}.csv"  -f $Stamp)

$results  = [System.Collections.Generic.List[object]]::new()
$checkNum = 0

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','OK')][string]$Level = 'INFO'
    )
    $entry = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Verbose $entry
    try {
        if (-not (Test-Path $LogRoot)) { New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null }
        Add-Content -Path $LogPath -Value $entry -ErrorAction Stop
    } catch { Write-Warning "Could not write to log: $_" }
}

function Add-Result {
    param(
        [Parameter(Mandatory)][string]$Section,
        [Parameter(Mandatory)][string]$Name,
        [ValidateSet('PASS','WARN','FAIL','INFO')][string]$Status,
        [string]$Detail,
        [string]$Extra = ''
    )
    $script:checkNum++
    $results.Add([PSCustomObject]@{
        Num     = $script:checkNum
        Section = $Section
        Name    = $Name
        Status  = $Status
        Detail  = $Detail
        Extra   = $Extra
    })
    $color = switch ($Status) { 'PASS'{'Green'};'WARN'{'Yellow'};'FAIL'{'Red'};'INFO'{'Cyan'} }
    Write-Host (" [{0,3}] [{1,-8}] [{2,-4}] {3,-38} {4}" -f `
        $script:checkNum, $Section, $Status, $Name, $Detail) -ForegroundColor $color
    Write-Log "$Section | $Status | $Name | $Detail"
}

function ConvertTo-HtmlText {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    ($Text -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;')
}

function Test-NoEventsError {
    # Locale-independent: Get-WinEvent raises NoMatchingEventsFound when the filter is simply empty.
    param($ErrorRecord)
    return ("$($ErrorRecord.FullyQualifiedErrorId)" -match 'NoMatchingEventsFound')
}

function Test-PortOpen {
    param([string]$HostName, [int]$Port, [int]$TimeoutMs = 3000)
    return [bool](& $Tool['Test-TcpPort.ps1'] -ComputerName $HostName -Port $Port -TimeoutMs $TimeoutMs).Open
}

function Join-First {
    # First N items joined, with "(+M more)" so long lists stay readable in the report.
    param([object[]]$Items, [int]$N = 25, [string]$Sep = ', ')
    $Items = @($Items)
    $s = (@($Items | Select-Object -First $N) -join $Sep)
    if ($Items.Count -gt $N) { $s += " (+$($Items.Count - $N) more)" }
    return $s
}

# Server-side LDAP bit tests (userAccountControl)
$ldapEnabled = '(!(userAccountControl:1.2.840.113556.1.4.803:=2))'

# ======================================================================
#  HEADER
# ======================================================================
Write-Log "Audit v$ScriptVersion started. Domain=$Domain Skip=$($SkipList -join ',')"
Write-Host ""
Write-Host "=== WINDOWS DOMAIN HEALTH AUDIT v$ScriptVersion ===" -ForegroundColor Cyan
Write-Host "  Domain     : $Domain"
Write-Host "  Date       : $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
Write-Host "  Skipping   : $(if($SkipList){$SkipList -join ', '}else{'(nothing)'})"
Write-Host ""
Write-Host (" {0,-5} {1,-10} {2,-6} {3,-38} {4}" -f '#','Section','Status','Check','Detail') -ForegroundColor White
Write-Host ("-" * 110) -ForegroundColor DarkGray

# Resolve DC list once (used by several sections)
$dcObjects = @()
try {
    $dcObjects = @(Get-ADDomainController -Filter * -Server $Domain -ErrorAction Stop)
} catch {
    Write-Log "Get-ADDomainController failed: $_" -Level WARN
}
if (-not $DomainControllers) {
    $DomainControllers = @($dcObjects | ForEach-Object { $_.HostName })
}
$DomainControllers = @($DomainControllers)

# ======================================================================
#  SECTION: FOREST
# ======================================================================
if (Test-SectionEnabled 'FOREST') {
    try {
        $dom = Get-ADDomain -Server $Domain -ErrorAction Stop
        $fst = Get-ADForest -Server $Domain -ErrorAction Stop

        Add-Result 'FOREST' 'Domain Functional Level' 'INFO' "$($dom.DomainMode)"
        Add-Result 'FOREST' 'Forest Functional Level' 'INFO' "$($fst.ForestMode)"

        # Flag legacy functional levels (2008R2 and older predate 2012R2)
        if ("$($dom.DomainMode)" -match '2000|2003|2008') {
            Add-Result 'FOREST' 'DFL is legacy' 'WARN' "Domain mode $($dom.DomainMode) predates 2012R2 - consider raising"
        }
        $upn = ($fst.UPNSuffixes -join ', ')
        Add-Result 'FOREST' 'UPN Suffixes' 'INFO' $(if ($upn) { $upn } else { '(none beyond default)' })

        # Tombstone lifetime (attribute unset = 60 days)
        try {
            $configNC = (Get-ADRootDSE -Server $Domain).configurationNamingContext
            $tsl = (Get-ADObject "CN=Directory Service,CN=Windows NT,CN=Services,$configNC" `
                    -Server $Domain -Properties tombstoneLifetime -ErrorAction Stop).tombstoneLifetime
            if (-not $tsl) { $tsl = 60 }
            if ($tsl -lt 180) {
                Add-Result 'FOREST' 'Tombstone Lifetime' 'WARN' "$tsl days (default/low - 180 recommended)"
            } else {
                Add-Result 'FOREST' 'Tombstone Lifetime' 'PASS' "$tsl days"
            }
        } catch { Add-Result 'FOREST' 'Tombstone Lifetime' 'INFO' "Could not read: $_" }
    } catch {
        Add-Result 'FOREST' 'Forest/Domain query' 'FAIL' "Could not query domain '$Domain': $_"
    }
}

# ======================================================================
#  SECTION: AD  (DCs, FSMO, replication, SYSVOL, DFSR, PDC time)
# ======================================================================
if (Test-SectionEnabled 'AD') {
    # DC inventory
    if ($dcObjects) {
        $dcNames = ($dcObjects | ForEach-Object { $_.HostName }) -join ', '
        Add-Result 'AD' 'DC Inventory' 'INFO' "$(@($dcObjects).Count) DC(s): $dcNames"
    } else {
        Add-Result 'AD' 'DC Inventory' 'FAIL' 'No domain controllers discovered'
    }

    # FSMO (building block: Get-FsmoRoleHolder)
    $pdcFqdn = $null
    try {
        $fsmo = @(& $Tool['Get-FsmoRoleHolder.ps1'] -Domain $Domain -ExpectedHolder $Config.ExpectedFsmoHolder -ErrorAction Stop)
        $pdcFqdn = ($fsmo | Where-Object Role -eq 'PDCEmulator').Holder
        $fsmoLines = ($fsmo | ForEach-Object { "$($_.Role)=$($_.HolderShort)" }) -join '; '
        if ($Config.ExpectedFsmoHolder) {
            if (@($fsmo | Where-Object { -not $_.OnExpected }).Count -gt 0) {
                Add-Result 'AD' 'FSMO Placement' 'WARN' "Not all on $($Config.ExpectedFsmoHolder)" $fsmoLines
            } else {
                Add-Result 'AD' 'FSMO Placement' 'PASS' "All 5 on $($Config.ExpectedFsmoHolder)" $fsmoLines
            }
        } else {
            Add-Result 'AD' 'FSMO Placement' 'INFO' 'FSMO holders listed' $fsmoLines
        }
    } catch { Add-Result 'AD' 'FSMO Placement' 'WARN' "Could not query FSMO: $_" }

    # Replication - structured cmdlet instead of scraping repadmin text
    # (the old '\d+\s+fail' regex never matched real replsummary output, so it always passed).
    if (@($dcObjects).Count -le 1) {
        Add-Result 'AD' 'AD Replication' 'INFO' 'Single DC - no intra-domain replication partners'
    } else {
        try {
            $fails = @(Get-ADReplicationFailure -Target $Domain -Scope Domain -ErrorAction Stop |
                       Where-Object { $_.FailureCount -gt 0 })
            if ($fails.Count -gt 0) {
                $lines = $fails | Select-Object -First 5 | ForEach-Object {
                    "$(($_.Server -split '\.')[0]) <- $($_.Partner) fails=$($_.FailureCount) err=$($_.LastError) since $($_.FirstFailureTime)"
                }
                Add-Result 'AD' 'AD Replication' 'FAIL' "$($fails.Count) replication link(s) failing" ($lines -join ' | ')
            } else {
                Add-Result 'AD' 'AD Replication' 'PASS' 'No replication failures reported by any DC'
            }
        } catch { Add-Result 'AD' 'AD Replication' 'WARN' "Could not query replication failures: $_" }
    }

    # SYSVOL reachability (first DC)
    if ($DomainControllers.Count -gt 0) {
        $primary = $DomainControllers[0]
        try {
            if (Test-Path "\\$primary\SYSVOL" -ErrorAction Stop) {
                Add-Result 'AD' 'SYSVOL Share' 'PASS' "\\$primary\SYSVOL accessible"
            } else {
                Add-Result 'AD' 'SYSVOL Share' 'FAIL' "\\$primary\SYSVOL NOT accessible"
            }
        } catch { Add-Result 'AD' 'SYSVOL Share' 'WARN' "Error accessing SYSVOL: $_" }
    }

    # SYSVOL replication engine: DFSR vs legacy FRS
    try {
        $frs = Get-ADObject -Server $Domain -Filter "objectClass -eq 'nTFRSSubscriber'" -ErrorAction Stop
        if ($frs) {
            Add-Result 'AD' 'SYSVOL Engine' 'WARN' 'Legacy FRS objects present - migrate SYSVOL to DFSR'
        } else {
            Add-Result 'AD' 'SYSVOL Engine' 'PASS' 'No legacy FRS subscriber objects (DFSR expected)'
        }
    } catch { Add-Result 'AD' 'SYSVOL Engine' 'INFO' "Could not determine SYSVOL engine: $_" }

    # PDC emulator time source. The forest-root PDC is the domain's time master and must
    # sync from an external NTP source (Type=NTP), not the hierarchy (NT5DS) or nothing.
    # Read from the registry + service state - locale-independent, unlike w32tm text.
    if ($pdcFqdn) {
        try {
            $tm = Invoke-Command -ComputerName $pdcFqdn -ErrorAction Stop -ScriptBlock {
                $p = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters' -ErrorAction Stop
                $vm = $false
                try { $vm = [bool]((Get-CimInstance Win32_ComputerSystem -ErrorAction Stop).Model -match 'Virtual|VMware|KVM|QEMU|HVM') } catch { }
                [PSCustomObject]@{
                    Type      = "$($p.Type)"
                    NtpServer = "$($p.NtpServer)"
                    Service   = "$((Get-Service W32Time -ErrorAction SilentlyContinue).Status)"
                    IsVM      = $vm
                }
            }
            $isRootPdc = $true
            try { $isRootPdc = ((Get-ADForest -Server $Domain).RootDomain -ieq $Domain) } catch { }
            $desc = "Type=$($tm.Type) NtpServer=$($tm.NtpServer) Service=$($tm.Service)$(if($tm.IsVM){' (VM)'})"
            if ($tm.Service -ne 'Running') {
                Add-Result 'AD' 'PDC Time Source' 'FAIL' "W32Time is not running on the PDC ($pdcFqdn)" $desc
            } elseif ($isRootPdc -and $tm.Type -ne 'NTP' -and $tm.Type -ne 'AllSync') {
                Add-Result 'AD' 'PDC Time Source' 'WARN' "Forest-root PDC ($pdcFqdn) is not set to an external NTP source - domain time is free-running$(if($tm.IsVM){' or follows the hypervisor'})" $desc
            } else {
                Add-Result 'AD' 'PDC Time Source' 'PASS' "PDC ($pdcFqdn) time source configured" $desc
            }
        } catch { Add-Result 'AD' 'PDC Time Source' 'INFO' "Could not read W32Time config on $pdcFqdn (remoting?): $_" }
    }
}

# ======================================================================
#  SECTION: SITES  (sites, subnets, coverage gaps)
# ======================================================================
if (Test-SectionEnabled 'SITES') {
    try {
        $sites   = @(Get-ADReplicationSite -Filter * -Server $Domain -ErrorAction Stop)
        $subnets = @(Get-ADReplicationSubnet -Filter * -Server $Domain -ErrorAction Stop)

        Add-Result 'SITES' 'Site Inventory' 'INFO' "$($sites.Count) site(s): $(($sites.Name) -join ', ')"
        if ($subnets.Count -eq 0) {
            Add-Result 'SITES' 'Subnet Inventory' 'WARN' 'No subnets defined - clients cannot be mapped to a site'
        } else {
            Add-Result 'SITES' 'Subnet Inventory' 'INFO' "$($subnets.Count) subnet(s) defined"
        }

        # Sites with no subnet mapped (Get-ADReplicationSubnet returns the site DN in .Site)
        $mappedSiteDNs = @($subnets | ForEach-Object { $_.Site } | Where-Object { $_ })
        $emptySites = @($sites | Where-Object { $_.DistinguishedName -notin $mappedSiteDNs })
        if ($emptySites.Count -gt 0) {
            Add-Result 'SITES' 'Sites Without Subnets' 'WARN' "$($emptySites.Count) site(s) have no subnet: $(($emptySites.Name) -join ', ')"
        } else {
            Add-Result 'SITES' 'Sites Without Subnets' 'PASS' 'Every site has at least one subnet'
        }

        # DC-to-site mapping
        if ($dcObjects) {
            $dcSiteLines = ($dcObjects | ForEach-Object { "$($_.Name)->$($_.Site)" }) -join '; '
            Add-Result 'SITES' 'DC Site Mapping' 'INFO' 'DC-to-site placement' $dcSiteLines
            if ($Config.ExpectedSites) {
                $badSite = @($dcObjects | Where-Object { $_.Site -notin $Config.ExpectedSites })
                if ($badSite.Count -gt 0) {
                    Add-Result 'SITES' 'DC In Expected Site' 'WARN' "DC(s) outside expected sites: $(($badSite.Name) -join ', ')"
                } else {
                    Add-Result 'SITES' 'DC In Expected Site' 'PASS' 'All DCs are in expected sites'
                }
            }
        }
    } catch { Add-Result 'SITES' 'Sites & Subnets' 'WARN' "Could not query sites: $_" }
}

# ======================================================================
#  SECTION: DNS
# ======================================================================
if (Test-SectionEnabled 'DNS') {
    $dnsServer = if ($DomainControllers.Count -gt 0) { $DomainControllers[0] } else { $Domain }

    # DC name resolution
    foreach ($dc in $DomainControllers) {
        try {
            $ips = @(Resolve-DnsName -Name $dc -Type A -Server $dnsServer -ErrorAction Stop |
                     Where-Object { $_.IPAddress } | ForEach-Object { $_.IPAddress })
            if ($ips.Count -gt 0) {
                Add-Result 'DNS' "Resolve $dc" 'PASS' "$dc -> $($ips -join ', ')"
            } else {
                Add-Result 'DNS' "Resolve $dc" 'WARN' "$dc returned no A record from $dnsServer"
            }
        } catch { Add-Result 'DNS' "Resolve $dc" 'WARN' "Could not resolve ${dc}: $_" }
    }

    # Scavenging / aging on the forward zone
    try {
        $zone = Get-DnsServerZoneAging -Name $Domain -ComputerName $dnsServer -ErrorAction Stop
        if ($zone.AgingEnabled) {
            Add-Result 'DNS' 'Zone Scavenging' 'PASS' "Aging enabled on $Domain (refresh $($zone.RefreshInterval), no-refresh $($zone.NoRefreshInterval))"
        } else {
            Add-Result 'DNS' 'Zone Scavenging' 'WARN' "Aging/scavenging NOT enabled on $Domain - stale records will accumulate"
        }
    } catch { Add-Result 'DNS' 'Zone Scavenging' 'INFO' "Could not read aging: $_" }

    # Stale A records for known-decommissioned DCs / stale IPs
    $staleNames = @($Config.StaleDCNames | Where-Object { $_ })
    $staleIPs   = @($Config.StaleIPs     | Where-Object { $_ })
    if ($staleNames.Count -gt 0 -or $staleIPs.Count -gt 0) {
        try {
            $aRecords = Get-DnsServerResourceRecord -ComputerName $dnsServer -ZoneName $Domain -RRType A -ErrorAction Stop
            $hits = @($aRecords | Where-Object {
                $rrIp   = $_.RecordData.IPv4Address.IPAddressToString
                $rrHost = ($_.HostName -split '\.')[0]
                ($staleIPs -contains $rrIp) -or ($staleNames -contains $rrHost)
            })
            if ($hits.Count -gt 0) {
                Add-Result 'DNS' 'Stale A Records' 'WARN' "$($hits.Count) A record(s) match known decommissioned hosts/IPs" (($hits | ForEach-Object { "$($_.HostName)=$($_.RecordData.IPv4Address)" }) -join '; ')
            } else {
                Add-Result 'DNS' 'Stale A Records' 'PASS' 'No A records match known stale hosts/IPs'
            }
        } catch { Add-Result 'DNS' 'Stale A Records' 'INFO' "Could not enumerate A records: $_" }
    }
}

# ======================================================================
#  SECTION: CERT  (LDAPS + PKI + expiring DC certs + GPO hygiene)
# ======================================================================
if (Test-SectionEnabled 'CERT') {
    foreach ($dc in $DomainControllers) {
        # LDAPS: real TLS handshake + served cert expiry (building block: Get-TlsCertificate)
        $lc = & $Tool['Get-TlsCertificate.ps1'] -ComputerName $dc -Port 636 -WarnDays $CertExpiryWarnDays
        if ($lc.Status -eq 'ERROR') {
            Add-Result 'CERT' "LDAPS $dc" 'WARN' "No working LDAPS on ${dc}:636 (no usable cert, or blocked): $($lc.Error)"
        } else {
            $certName = if ($lc.Subject) { $lc.Subject } else { "SAN:$($lc.DnsName)" }
            $desc = "$certName expires $($lc.NotAfter.ToString('yyyy-MM-dd')) ($($lc.DaysLeft) d)"
            switch ($lc.Status) {
                'EXPIRED' { Add-Result 'CERT' "LDAPS $dc" 'FAIL' "LDAPS cert EXPIRED: $desc" "Issuer: $($lc.Issuer)" }
                'WARN'    { Add-Result 'CERT' "LDAPS $dc" 'WARN' "LDAPS cert expires soon: $desc" "Issuer: $($lc.Issuer)" }
                default   { Add-Result 'CERT' "LDAPS $dc" 'PASS' "TLS handshake OK ($($lc.Protocol)), $desc" "Issuer: $($lc.Issuer)" }
            }
        }

        # Expiring machine certs (remote store read; best-effort)
        try {
            $expiring = Invoke-Command -ComputerName $dc -ArgumentList $CertExpiryWarnDays -ScriptBlock {
                param($WarnDays)
                Get-ChildItem Cert:\LocalMachine\My -ErrorAction Stop |
                    Where-Object { $_.NotAfter -lt (Get-Date).AddDays($WarnDays) } |
                    Select-Object Subject, NotAfter, Thumbprint
            } -ErrorAction Stop
            if ($expiring) {
                Add-Result 'CERT' "Expiring Certs $dc" 'WARN' "$(@($expiring).Count) machine cert(s) expired or expiring within $CertExpiryWarnDays days" (($expiring | ForEach-Object { "$($_.Subject) -> $($_.NotAfter.ToString('yyyy-MM-dd'))" }) -join '; ')
            } else {
                Add-Result 'CERT' "Expiring Certs $dc" 'PASS' "No machine certs expire within $CertExpiryWarnDays days"
            }
        } catch { Add-Result 'CERT' "Expiring Certs $dc" 'INFO' "Could not read cert store on $dc (remoting?): $_" }
    }

    # Enterprise CA presence + LIVENESS
    # Detects the dead-CA scenario: a CA still registered in AD (Enrollment Services)
    # that no longer answers, while autoenrollment GPOs still point clients at it.
    $registeredCAs = @()
    $pkiBase = $null
    try {
        $configNC = (Get-ADRootDSE -Server $Domain).configurationNamingContext
        $pkiBase  = "CN=Public Key Services,CN=Services,$configNC"
        $registeredCAs = @(Get-ADObject -SearchBase "CN=Enrollment Services,$pkiBase" `
               -Server $Domain -Filter "objectClass -eq 'pKIEnrollmentService'" `
               -Properties dNSHostName,displayName -ErrorAction Stop)
        if ($registeredCAs.Count -gt 0) {
            Add-Result 'CERT' 'Enterprise CA Registered' 'INFO' "$($registeredCAs.Count) CA(s) in AD: $(($registeredCAs | ForEach-Object { "$($_.Name)=$($_.dNSHostName)" }) -join ', ')"
        } else {
            Add-Result 'CERT' 'Enterprise CA Registered' 'INFO' 'No enterprise CA published in AD'
        }
        if ($Config.DeadCAHostname) {
            $stillThere = @($registeredCAs | Where-Object { $_.dNSHostName -ieq $Config.DeadCAHostname })
            if ($stillThere.Count -gt 0) {
                Add-Result 'CERT' 'Known Dead CA Registered' 'WARN' "$($Config.DeadCAHostname) (config DeadCAHostname) is still registered in Enrollment Services"
            } else {
                Add-Result 'CERT' 'Known Dead CA Registered' 'PASS' "$($Config.DeadCAHostname) is no longer registered in Enrollment Services"
            }
        }
    } catch { Add-Result 'CERT' 'Enterprise CA Registered' 'INFO' "Could not query Enrollment Services: $_" }

    # Liveness: does each registered CA actually answer? Judged on certutil's exit code,
    # not its text, so it works on localized (e.g. Swedish) Windows.
    $deadCAs = [System.Collections.Generic.List[string]]::new()
    $deadCANames = [System.Collections.Generic.List[string]]::new()
    $haveCertutil = [bool](Get-Command certutil.exe -ErrorAction SilentlyContinue)
    foreach ($ca in $registeredCAs) {
        $caHost = $ca.dNSHostName
        $caName = if ($ca.displayName) { $ca.displayName } else { $ca.Name }
        if (-not $caHost) { continue }
        if (-not $haveCertutil) {
            Add-Result 'CERT' "CA Liveness $caName" 'INFO' 'certutil.exe not available - liveness not tested'
            continue
        }
        $tcp135  = Test-PortOpen -HostName $caHost -Port 135
        $pingOut = & certutil.exe -ping -config "$caHost\$caName" 2>&1 | Out-String
        $pingRc  = $LASTEXITCODE
        if ($pingRc -eq 0) {
            Add-Result 'CERT' "CA Liveness $caName" 'PASS' "$caHost answers certutil -ping"
        } else {
            $deadCAs.Add("$caName ($caHost)")
            $deadCANames.Add($caName)
            $reason = if (-not $tcp135) { 'no RPC/135 and -ping failed' } else { 'RPC reachable but -ping failed' }
            $last   = (($pingOut -split "`r?`n") | Where-Object { $_.Trim() } | Select-Object -Last 1)
            Add-Result 'CERT' "CA Liveness $caName" 'FAIL' "$caHost does NOT answer as a CA - $reason (exit $pingRc)" "$last"
        }
    }

    # Autoenrollment GPOs still in play - cross-reference against dead CAs.
    # Structured check on the Public Key policy node: EnrollCertificatesAutomatically=true,
    # in an enabled Computer/User half, in a GPO that has at least one enabled link.
    # (A plain 'AutoEnrollment' text match also hits GPOs that explicitly DISABLE it.)
    # The same pass counts GPOs with no enabled link (hygiene).
    if ($SkipGpoScan) {
        Add-Result 'CERT' 'Autoenrollment GPOs' 'INFO' 'GPO scan skipped (-SkipGpoScan)'
    } else {
        try {
            Import-Module GroupPolicy -ErrorAction Stop
            $aeGpos = [System.Collections.Generic.List[string]]::new()
            $noLinkGpos = [System.Collections.Generic.List[string]]::new()
            $aeUnlinked = 0
            foreach ($g in @(Get-GPO -All -Domain $Domain -ErrorAction Stop)) {
                try {
                    [xml]$rep = Get-GPOReport -Guid $g.Id -ReportType Xml -Domain $Domain -ErrorAction Stop
                    $links = @($rep.GPO.LinksTo | Where-Object { $_ -and $_.Enabled -eq 'true' } | ForEach-Object { $_.SOMPath })
                    if ($links.Count -eq 0) { $noLinkGpos.Add($g.DisplayName) }
                    $sides = @()
                    foreach ($n in @($rep.SelectNodes("//*[local-name()='AutoEnrollmentSettings']"))) {
                        $on = $n.SelectSingleNode("*[local-name()='EnrollCertificatesAutomatically']")
                        if (-not $on -or $on.InnerText -ne 'true') { continue }
                        $side = $n.ParentNode.ParentNode.ParentNode    # Computer|User / ExtensionData / Extension
                        if ($side -and $side.SelectSingleNode("*[local-name()='Enabled']").InnerText -eq 'true') {
                            $sides += $side.LocalName
                        }
                    }
                    if ($sides.Count -eq 0) { continue }
                    if ($links.Count -eq 0) { $aeUnlinked++; continue }
                    $aeGpos.Add("$($g.DisplayName) [$(($sides | Select-Object -Unique) -join '+'); links: $($links -join ', ')]")
                } catch { Write-Log "GPO report failed for '$($g.DisplayName)': $_" -Level WARN }
            }
            $unlinkedNote = if ($aeUnlinked -gt 0) { " ($aeUnlinked more enabled but unlinked - ignored)" } else { '' }
            if ($aeGpos.Count -gt 0) {
                if ($deadCAs.Count -gt 0) {
                    Add-Result 'CERT' 'Autoenrollment vs Dead CA' 'FAIL' `
                        "$($aeGpos.Count) linked autoenrollment GPO(s) active while $($deadCAs.Count) registered CA(s) are dead - clients enrolling against a void$unlinkedNote" `
                        ("GPOs: " + ($aeGpos -join ' | ') + "  ||  Dead CA: " + ($deadCAs -join ', '))
                } elseif ($registeredCAs.Count -eq 0) {
                    Add-Result 'CERT' 'Autoenrollment GPOs' 'WARN' "$($aeGpos.Count) linked autoenrollment GPO(s) but no enterprise CA is registered$unlinkedNote" ($aeGpos -join ' | ')
                } else {
                    Add-Result 'CERT' 'Autoenrollment GPOs' 'INFO' "$($aeGpos.Count) linked autoenrollment GPO(s), all registered CAs answer$unlinkedNote" ($aeGpos -join ' | ')
                }
            } else {
                Add-Result 'CERT' 'Autoenrollment GPOs' 'PASS' "No linked GPO enables autoenrollment$unlinkedNote"
            }
            if ($noLinkGpos.Count -gt 0) {
                Add-Result 'CERT' 'Unlinked GPOs' 'INFO' "$($noLinkGpos.Count) GPO(s) have no enabled link - candidates for cleanup" (Join-First $noLinkGpos)
            } else {
                Add-Result 'CERT' 'Unlinked GPOs' 'PASS' 'Every GPO has at least one enabled link'
            }
        } catch {
            Add-Result 'CERT' 'Autoenrollment GPOs' 'INFO' "Could not enumerate GPOs (GroupPolicy module / GPMC missing?): $_"
        }
    }

    # NTAuth store - read the AD object directly (certutil -viewstore opens a GUI dialog,
    # and its text output is localized). Flags certs whose issuer is a dead CA.
    if ($pkiBase) {
        try {
            $ntObj   = Get-ADObject "CN=NTAuthCertificates,$pkiBase" -Server $Domain -Properties cACertificate -ErrorAction Stop
            $ntCerts = @($ntObj.cACertificate | ForEach-Object {
                New-Object System.Security.Cryptography.X509Certificates.X509Certificate2(,[byte[]]$_)
            })
            if ($ntCerts.Count -eq 0) {
                Add-Result 'CERT' 'NTAuth Store' 'INFO' 'NTAuth store is empty (no enterprise-trusted CA issuer cert)'
            } else {
                $ntLines = $ntCerts | ForEach-Object { "$($_.Subject) exp $($_.NotAfter.ToString('yyyy-MM-dd'))" }
                $ntDead = @($ntCerts | Where-Object { $s = $_.Subject; @($deadCANames | Where-Object { $s -match ('CN=' + [regex]::Escape($_) + '(,|$)') }).Count -gt 0 })
                if ($ntDead.Count -gt 0) {
                    Add-Result 'CERT' 'NTAuth Store' 'WARN' "$($ntCerts.Count) cert(s) in NTAuth, $($ntDead.Count) belong to a dead CA" ($ntLines -join '; ')
                } else {
                    Add-Result 'CERT' 'NTAuth Store' 'INFO' "$($ntCerts.Count) cert(s) in NTAuth - none match a dead CA" ($ntLines -join '; ')
                }
            }
        } catch { Add-Result 'CERT' 'NTAuth Store' 'INFO' "Could not read NTAuthCertificates: $_" }
    }

    # Optional: blast radius of each dead CA (building block: Get-CertBlastRadius)
    if ($CertSweep) {
        $sweepMatches = @($deadCANames | ForEach-Object { "CN=$_" })
        if ($Config.DeadCAIssuerMatch) { $sweepMatches += $Config.DeadCAIssuerMatch }
        $sweepMatches = @($sweepMatches | Where-Object { $_ } | Select-Object -Unique)
        if ($sweepMatches.Count -eq 0) {
            Add-Result 'CERT' 'Dead CA Blast Radius' 'PASS' 'No dead CA found and no DeadCAIssuerMatch configured - nothing to sweep'
        }
        foreach ($m in $sweepMatches) {
            try {
                $sw = @(& $Tool['Get-CertBlastRadius.ps1'] -IssuerMatch $m -LogRoot $LogRoot -ErrorAction Stop)
                $hit   = @($sw | Where-Object { $_.Reachable -and $_.Subject })
                $hosts = @($hit | ForEach-Object { $_.Computer } | Sort-Object -Unique)
                $cli   = @($hit | Where-Object { $_.HasClientAuth -eq $true })
                $unr   = @($sw | Where-Object { -not $_.Reachable })
                $st    = if ($cli.Count -gt 0) { 'FAIL' } elseif ($hosts.Count -gt 0) { 'WARN' } else { 'PASS' }
                Add-Result 'CERT' "Blast Radius $m" $st `
                    "$($hosts.Count) host(s) hold $($hit.Count) cert(s) from it, $($cli.Count) client-auth capable, $($unr.Count) host(s) unreachable" `
                    (Join-First $hosts)
            } catch { Add-Result 'CERT' "Blast Radius $m" 'INFO' "Sweep failed: $_" }
        }
    }
}

# ======================================================================
#  SECTION: USERS  (privileged risk + hygiene + policy)
# ======================================================================
$privAccounts = @()      # shared with SECURITY
$privSams = @{}
if ((Test-SectionEnabled 'USERS') -or (Test-SectionEnabled 'SECURITY')) {
    # Building block: Get-PrivilegedAccountRisk (groups by well-known SID, locale-independent)
    $privErr = @()
    try {
        $privAccounts = @(& $Tool['Get-PrivilegedAccountRisk.ps1'] -Domain $Domain `
            -KnownRiskAccounts @($Config.KnownRiskAccounts) -ServiceAcctRegex $Config.ServiceAcctRegex `
            -NamedAdminRegex $Config.NamedAdminRegex -InactiveDays $InactiveDays `
            -ErrorAction SilentlyContinue -ErrorVariable privErr)
    } catch { $privErr += $_ }
    foreach ($a in $privAccounts) { $privSams[$a.SamAccountName] = $a }
}

if (Test-SectionEnabled 'USERS') {
    foreach ($e in $privErr) { Add-Result 'USERS' 'Privileged Group Expansion' 'INFO' "$e" }
    $high   = @($privAccounts | Where-Object Risk -eq 'HIGH-RISK')
    $review = @($privAccounts | Where-Object Risk -eq 'REVIEW')
    $ok     = @($privAccounts | Where-Object Risk -eq 'OK')
    $privDetail = @($privAccounts | Where-Object Risk -ne 'OK' | ForEach-Object { "$($_.SamAccountName) [$($_.Risk)]: $($_.Reason) ($($_.Groups -join '/'))" })
    $privStatus = if ($high.Count -gt 0) {'FAIL'} elseif ($review.Count -gt 0) {'WARN'} else {'PASS'}
    Add-Result 'USERS' 'Privileged Accounts' $privStatus "HIGH-RISK=$($high.Count) REVIEW=$($review.Count) OK=$($ok.Count) (total $($privAccounts.Count))" ($privDetail -join ' | ')

    # krbtgt password age
    try {
        $krbtgt = Get-ADUser krbtgt -Server $Domain -Properties PasswordLastSet -ErrorAction Stop
        $age = (New-TimeSpan -Start $krbtgt.PasswordLastSet -End (Get-Date)).Days
        if ($age -gt 180) {
            Add-Result 'USERS' 'krbtgt Password Age' 'WARN' "$age days - consider a controlled reset (twice, spaced)"
        } else {
            Add-Result 'USERS' 'krbtgt Password Age' 'PASS' "$age days"
        }
    } catch { Add-Result 'USERS' 'krbtgt Password Age' 'INFO' "Could not read krbtgt: $_" }

    # Default domain password policy
    try {
        $pol = Get-ADDefaultDomainPasswordPolicy -Server $Domain -ErrorAction Stop
        $polIssues = [System.Collections.Generic.List[string]]::new()
        if ($pol.LockoutThreshold -eq 0)            { $polIssues.Add('No account lockout threshold') }
        elseif ($pol.LockoutThreshold -gt 10)       { $polIssues.Add("Lockout threshold $($pol.LockoutThreshold) - consider 10 or lower") }
        if ($pol.MinPasswordLength -lt 12)          { $polIssues.Add("Min length $($pol.MinPasswordLength) - consider 12 or more") }
        if (-not $pol.ComplexityEnabled)            { $polIssues.Add('Complexity not required') }
        if ($pol.ReversibleEncryptionEnabled)       { $polIssues.Add('Reversible encryption ENABLED') }
        $polInfo = "MinLen=$($pol.MinPasswordLength) MaxAge=$($pol.MaxPasswordAge.Days)d Complexity=$($pol.ComplexityEnabled) Lockout=$($pol.LockoutThreshold) History=$($pol.PasswordHistoryCount)"
        if ($polIssues.Count -gt 0) {
            $st = if ($pol.ReversibleEncryptionEnabled) { 'FAIL' } else { 'WARN' }
            Add-Result 'USERS' 'Password Policy' $st ($polIssues -join '; ') $polInfo
        } else {
            Add-Result 'USERS' 'Password Policy' 'PASS' 'Default domain policy meets baseline' $polInfo
        }
        try {
            $fgpp = @(Get-ADFineGrainedPasswordPolicy -Filter * -Server $Domain -ErrorAction Stop)
            if ($fgpp.Count -gt 0) {
                Add-Result 'USERS' 'Fine-Grained Policies' 'INFO' "$($fgpp.Count) fine-grained password polic(ies) also apply" (($fgpp | ForEach-Object { "$($_.Name) (prec $($_.Precedence), minlen $($_.MinPasswordLength))" }) -join '; ')
            }
        } catch { }
    } catch { Add-Result 'USERS' 'Password Policy' 'INFO' "Could not read password policy: $_" }

    # Enabled users, password-never-expires (LDAP filter: server-side, no full pull)
    try {
        $pne = @(Get-ADUser -LDAPFilter "(&$ldapEnabled(userAccountControl:1.2.840.113556.1.4.803:=65536))" -Server $Domain -ErrorAction Stop)
        if ($pne.Count -gt 0) {
            Add-Result 'USERS' 'Password Never Expires' 'WARN' "$($pne.Count) enabled user(s) with non-expiring passwords" (Join-First ($pne | ForEach-Object { $_.SamAccountName }))
        } else {
            Add-Result 'USERS' 'Password Never Expires' 'PASS' 'No enabled users with non-expiring passwords'
        }
    } catch { Add-Result 'USERS' 'Password Never Expires' 'INFO' "Query failed: $_" }

    # Stale enabled users / computers - server-side filter on lastLogonTimestamp
    # (replicated, up to ~14 days behind by design; accounts that never logged on are not counted).
    $cutFt = (Get-Date).AddDays(-1 * $InactiveDays).ToFileTimeUtc()
    try {
        $stale = @(Get-ADUser -LDAPFilter "(&$ldapEnabled(lastLogonTimestamp<=$cutFt))" -Server $Domain -ErrorAction Stop)
        if ($stale.Count -gt 0) {
            Add-Result 'USERS' 'Stale Enabled Users' 'WARN' "$($stale.Count) enabled user(s) inactive > $InactiveDays days" (Join-First ($stale | ForEach-Object { $_.SamAccountName }))
        } else {
            Add-Result 'USERS' 'Stale Enabled Users' 'PASS' "No enabled users inactive > $InactiveDays days"
        }
    } catch { Add-Result 'USERS' 'Stale Enabled Users' 'INFO' "Query failed: $_" }

    try {
        $staleC = @(Get-ADComputer -LDAPFilter "(&$ldapEnabled(lastLogonTimestamp<=$cutFt))" -Server $Domain -ErrorAction Stop)
        if ($staleC.Count -gt 0) {
            Add-Result 'USERS' 'Stale Computer Accounts' 'WARN' "$($staleC.Count) computer(s) inactive > $InactiveDays days" (Join-First ($staleC | ForEach-Object { $_.Name }))
        } else {
            Add-Result 'USERS' 'Stale Computer Accounts' 'PASS' "No computers inactive > $InactiveDays days"
        }
    } catch { Add-Result 'USERS' 'Stale Computer Accounts' 'INFO' "Query failed: $_" }

    # AdminSDHolder orphans - enabled users with adminCount=1 that are no longer in any
    # protected group. They keep the hardened AdminSDHolder ACL (inheritance off), which
    # breaks delegated administration of those accounts.
    if ($privErr.Count -gt 0) {
        Add-Result 'USERS' 'AdminSDHolder Orphans' 'INFO' 'Skipped - privileged group expansion was incomplete (see above), orphans cannot be told apart'
    } else {
        try {
            $ac = @(Get-ADUser -LDAPFilter "(&(adminCount=1)$ldapEnabled)" -Server $Domain -ErrorAction Stop)
            $orphans = @($ac | Where-Object { -not $privSams.ContainsKey($_.SamAccountName) })
            if ($orphans.Count -gt 0) {
                Add-Result 'USERS' 'AdminSDHolder Orphans' 'WARN' "$($orphans.Count) enabled user(s) with adminCount=1 but in no protected group - ACL inheritance still blocked" (Join-First ($orphans | ForEach-Object { $_.SamAccountName }))
            } else {
                Add-Result 'USERS' 'AdminSDHolder Orphans' 'PASS' 'No adminCount=1 orphans'
            }
        } catch { Add-Result 'USERS' 'AdminSDHolder Orphans' 'INFO' "Query failed: $_" }
    }

    # AS-REP roastable - Kerberos pre-authentication disabled (UAC 0x400000)
    try {
        $asrep = @(Get-ADUser -LDAPFilter "(&$ldapEnabled(userAccountControl:1.2.840.113556.1.4.803:=4194304))" -Server $Domain -ErrorAction Stop)
        if ($asrep.Count -gt 0) {
            $asPriv = @($asrep | Where-Object { $privSams.ContainsKey($_.SamAccountName) })
            $st = if ($asPriv.Count -gt 0) { 'FAIL' } else { 'WARN' }
            Add-Result 'USERS' 'AS-REP Roastable' $st "$($asrep.Count) enabled account(s) without Kerberos pre-auth - hash crackable offline$(if($asPriv.Count){", $($asPriv.Count) privileged"})" (Join-First ($asrep | ForEach-Object { $_.SamAccountName }))
        } else {
            Add-Result 'USERS' 'AS-REP Roastable' 'PASS' 'No enabled accounts with pre-auth disabled'
        }
    } catch { Add-Result 'USERS' 'AS-REP Roastable' 'INFO' "Query failed: $_" }

    # Protected Users (well-known RID 525)
    try {
        $domSid = (Get-ADDomain -Server $Domain -ErrorAction Stop).DomainSID.Value
        $pu = @(Get-ADGroupMember -Identity "$domSid-525" -Recursive -Server $Domain -ErrorAction Stop | Where-Object objectClass -eq 'user')
        $highPriv = @($privAccounts | Where-Object { $_.InHighPriv -and $_.Enabled })
        $notIn = @($highPriv | Where-Object { $_.SamAccountName -notin @($pu.SamAccountName) })
        Add-Result 'USERS' 'Protected Users' 'INFO' "$($pu.Count) user(s) in Protected Users; $($notIn.Count) of $($highPriv.Count) enabled high-priv admin(s) are not" (Join-First ($notIn | ForEach-Object { $_.SamAccountName }))
    } catch { Add-Result 'USERS' 'Protected Users' 'INFO' "Could not read Protected Users: $_" }
}

# ======================================================================
#  SECTION: SECURITY  (delegation, SPNs, LAPS, Kerberoasting)
# ======================================================================
if (Test-SectionEnabled 'SECURITY') {
    # Unconstrained delegation (UAC 0x80000 TRUSTED_FOR_DELEGATION). DCs have it by design
    # (primaryGroupID 516 = DCs, 521 = RODCs); anything else can impersonate any user that
    # authenticates to it.
    try {
        $udC = @(Get-ADComputer -LDAPFilter "(&$ldapEnabled(userAccountControl:1.2.840.113556.1.4.803:=524288)(!(primaryGroupID=516))(!(primaryGroupID=521)))" -Server $Domain -ErrorAction Stop)
        $udU = @(Get-ADUser     -LDAPFilter "(&$ldapEnabled(userAccountControl:1.2.840.113556.1.4.803:=524288))" -Server $Domain -ErrorAction Stop)
        if (($udC.Count + $udU.Count) -gt 0) {
            $d = @()
            if ($udC.Count) { $d += 'Computers: ' + (Join-First ($udC | ForEach-Object { $_.Name })) }
            if ($udU.Count) { $d += 'Users: '     + (Join-First ($udU | ForEach-Object { $_.SamAccountName })) }
            Add-Result 'SECURITY' 'Unconstrained Delegation' 'WARN' "$($udC.Count + $udU.Count) non-DC object(s) trusted for unconstrained delegation" ($d -join ' | ')
        } else {
            Add-Result 'SECURITY' 'Unconstrained Delegation' 'PASS' 'Only DCs have unconstrained delegation'
        }
    } catch { Add-Result 'SECURITY' 'Unconstrained Delegation' 'INFO' "Query failed: $_" }

    # Duplicate SPNs (cause KDC_ERR_S_PRINCIPAL_UNKNOWN / Kerberos fallback to NTLM).
    # Grouped in PowerShell instead of parsing 'setspn -X' text (localized). Domain scope.
    try {
        $spnOwners = @{}
        foreach ($o in @(Get-ADObject -LDAPFilter '(servicePrincipalName=*)' -Server $Domain -Properties servicePrincipalName,sAMAccountName -ResultPageSize 1000 -ErrorAction Stop)) {
            foreach ($s in @($o.servicePrincipalName)) {
                $k = "$s".ToLowerInvariant()
                if (-not $spnOwners.ContainsKey($k)) { $spnOwners[$k] = [System.Collections.Generic.List[string]]::new() }
                $spnOwners[$k].Add($o.sAMAccountName)
            }
        }
        $dups = @($spnOwners.GetEnumerator() | Where-Object { @($_.Value | Select-Object -Unique).Count -gt 1 })
        if ($dups.Count -gt 0) {
            Add-Result 'SECURITY' 'Duplicate SPNs' 'WARN' "$($dups.Count) SPN(s) registered on more than one account - Kerberos fails for them" (Join-First ($dups | ForEach-Object { "$($_.Key) -> $((@($_.Value | Select-Object -Unique)) -join '+')" }) 10 ' | ')
        } else {
            Add-Result 'SECURITY' 'Duplicate SPNs' 'PASS' "No duplicate SPNs ($($spnOwners.Count) SPNs checked)"
        }
    } catch { Add-Result 'SECURITY' 'Duplicate SPNs' 'INFO' "Query failed: $_" }

    # LAPS: schema (Windows LAPS msLAPS-* or legacy ms-Mcs-AdmPwd*) + coverage on member computers.
    try {
        $schemaNC = (Get-ADRootDSE -Server $Domain).schemaNamingContext
        $hasWin    = [bool](Get-ADObject -SearchBase $schemaNC -LDAPFilter '(lDAPDisplayName=msLAPS-PasswordExpirationTime)' -Server $Domain -ErrorAction Stop)
        $hasLegacy = [bool](Get-ADObject -SearchBase $schemaNC -LDAPFilter '(lDAPDisplayName=ms-Mcs-AdmPwdExpirationTime)' -Server $Domain -ErrorAction Stop)
        if (-not $hasWin -and -not $hasLegacy) {
            Add-Result 'SECURITY' 'LAPS' 'WARN' 'No LAPS schema (Windows LAPS or legacy) - local admin passwords are not managed'
        } else {
            $attrs = @()
            if ($hasWin)    { $attrs += 'msLAPS-PasswordExpirationTime' }
            if ($hasLegacy) { $attrs += 'ms-Mcs-AdmPwdExpirationTime' }
            $members = @(Get-ADComputer -LDAPFilter "(&$ldapEnabled(!(primaryGroupID=516))(!(primaryGroupID=521)))" -Server $Domain -Properties $attrs -ErrorAction Stop)
            $covered = @($members | Where-Object { $c = $_; @($attrs | Where-Object { $c.$_ }).Count -gt 0 })
            $kind = (@(if ($hasWin) { 'Windows LAPS' }; if ($hasLegacy) { 'legacy LAPS' }) -join ' + ')
            if ($members.Count -eq 0) {
                Add-Result 'SECURITY' 'LAPS' 'INFO' "$kind schema present; no enabled member computers"
            } else {
                $pct = [int](100 * $covered.Count / $members.Count)
                $st  = if ($pct -ge 90) { 'PASS' } elseif ($pct -ge 50) { 'WARN' } else { 'FAIL' }
                $uncovered = @($members | Where-Object { $_.Name -notin @($covered.Name) } | ForEach-Object { $_.Name })
                Add-Result 'SECURITY' 'LAPS' $st "$kind schema present; $($covered.Count)/$($members.Count) member computers ($pct%) have a LAPS password" (Join-First $uncovered)
            }
        }
    } catch { Add-Result 'SECURITY' 'LAPS' 'INFO' "Could not check LAPS: $_" }

    # Kerberoastable: enabled USER accounts with an SPN (krbtgt excluded). A service ticket
    # for them can be requested by any user and cracked offline - worst when privileged.
    try {
        $kerb = @(Get-ADUser -LDAPFilter "(&$ldapEnabled(servicePrincipalName=*)(!(sAMAccountName=krbtgt)))" -Server $Domain -Properties PasswordLastSet -ErrorAction Stop)
        if ($kerb.Count -gt 0) {
            $kPriv = @($kerb | Where-Object { $privSams.ContainsKey($_.SamAccountName) })
            $st = if ($kPriv.Count -gt 0) { 'FAIL' } else { 'WARN' }
            $lines = $kerb | ForEach-Object {
                $age = if ($_.PasswordLastSet) { [int]((Get-Date) - $_.PasswordLastSet).TotalDays } else { '?' }
                "$($_.SamAccountName)$(if($privSams.ContainsKey($_.SamAccountName)){' [PRIVILEGED]'}) pwdAge=${age}d"
            }
            Add-Result 'SECURITY' 'Kerberoastable Accounts' $st "$($kerb.Count) enabled user(s) with an SPN$(if($kPriv.Count){", $($kPriv.Count) privileged"}) - use gMSA or long random passwords" (Join-First $lines)
        } else {
            Add-Result 'SECURITY' 'Kerberoastable Accounts' 'PASS' 'No enabled user accounts with SPNs'
        }
    } catch { Add-Result 'SECURITY' 'Kerberoastable Accounts' 'INFO' "Query failed: $_" }
}

# ======================================================================
#  SECTION: NET  (DC port reachability + unsigned LDAP binds)
# ======================================================================
if (Test-SectionEnabled 'NET') {
    # Building block: Test-TcpPort
    # Plain hashtable on purpose: an [ordered] dictionary indexed with an [int] key indexes by POSITION.
    $corePorts = @(389, 636, 88, 3268, 53, 445)
    $portName  = @{ 389='LDAP 389'; 636='LDAPS 636'; 88='Kerberos 88'; 3268='GC 3268'; 53='DNS 53'; 445='SMB 445' }
    foreach ($dc in $DomainControllers) {
        $probe = @(& $Tool['Test-TcpPort.ps1'] -ComputerName $dc -Port $corePorts)
        $open = @($probe | Where-Object Open      | ForEach-Object { $portName[[int]$_.Port] })
        $shut = @($probe | Where-Object { -not $_.Open } | ForEach-Object { $portName[[int]$_.Port] })
        if ($shut.Count -eq 0) {
            Add-Result 'NET' "DC Ports $dc" 'PASS' 'All core AD ports reachable' ("open: " + ($open -join ', '))
        } else {
            Add-Result 'NET' "DC Ports $dc" 'WARN' "$($shut.Count) core port(s) unreachable" ("closed: " + ($shut -join ', '))
        }
    }

    # Stale IPs still answering LDAP
    foreach ($ip in @($Config.StaleIPs | Where-Object { $_ })) {
        if (Test-PortOpen -HostName $ip -Port 389) {
            Add-Result 'NET' "Stale LDAP $ip" 'WARN' "$ip still answers on 389 - expected decommissioned"
        } else {
            Add-Result 'NET' "Stale LDAP $ip" 'PASS' "$ip does not answer on 389"
        }
    }

    # Unsigned / simple LDAP binds (building block: Get-UnsignedLdapClient)
    if ($DomainControllers.Count -gt 0) {
        foreach ($r in @(& $Tool['Get-UnsignedLdapClient.ps1'] -ComputerName $DomainControllers -HoursBack $EventHoursBack)) {
            $dc = $r.ComputerName
            switch ($r.Status) {
                'ERROR'           { Add-Result 'NET' "Unsigned LDAP $dc" 'INFO' "Could not read Directory Service log on ${dc}: $($r.Error)" }
                'UNSIGNED'        { Add-Result 'NET' "Unsigned LDAP $dc" 'WARN' "$($r.ClientCount) client(s) doing unsigned/simple binds (last ${EventHoursBack}h)" (@($r.Clients) -join ', ') }
                'UNKNOWN-CLIENTS' { Add-Result 'NET' "Unsigned LDAP $dc" 'WARN' "Event 2887 reports $($r.Count2887) unsigned bind(s), but per-client 2889 logging is off (level $($r.LoggingLevel)) - set '16 LDAP Interface Events'=2 to see who" }
                'CLEAN-2887-ONLY' { Add-Result 'NET' "Unsigned LDAP $dc" 'PASS' "No 2887 in last ${EventHoursBack}h (2887 is daily; 2889 logging off, level $($r.LoggingLevel))" }
                default           { Add-Result 'NET' "Unsigned LDAP $dc" 'PASS' "No Event 2889/2887 in last ${EventHoursBack}h" }
            }
        }
    }
}

# ======================================================================
#  SECTION: LICENSE  (activation state; best-effort, remote)
# ======================================================================
if (Test-SectionEnabled 'LICENSE') {
    foreach ($dc in $DomainControllers) {
        try {
            $lic = Invoke-Command -ComputerName $dc -ScriptBlock {
                Get-CimInstance SoftwareLicensingProduct -Filter "PartialProductKey IS NOT NULL AND Name LIKE 'Windows%'" -ErrorAction Stop |
                    Select-Object Name, LicenseStatus, @{n='KMS';e={$_.KeyManagementServiceMachine}}
            } -ErrorAction Stop | Select-Object -First 1
            if ($lic) {
                $statusText = switch ([int]$lic.LicenseStatus) { 1{'Licensed'} 0{'Unlicensed'} 2{'OOB Grace'} 3{'OOT Grace'} 4{'Non-genuine grace'} 5{'Notification'} 6{'Extended grace'} default{"State $($lic.LicenseStatus)"} }
                $st = if ([int]$lic.LicenseStatus -eq 1) { 'PASS' } else { 'WARN' }
                Add-Result 'LICENSE' "Activation $dc" $st "$statusText$(if($lic.KMS){" (KMS: $($lic.KMS))"})" "$($lic.Name)"
            } else {
                Add-Result 'LICENSE' "Activation $dc" 'INFO' 'No Windows license product returned'
            }
        } catch { Add-Result 'LICENSE' "Activation $dc" 'INFO' "Could not query activation on $dc (remoting?): $_" }
    }
}

# ======================================================================
#  SECTION: EVENTS  (critical DS / DFSR errors)
# ======================================================================
if (Test-SectionEnabled 'EVENTS') {
    $start = (Get-Date).AddHours(-1 * $EventHoursBack)
    $logs = @(
        @{ Label='DS Errors';   Log='Directory Service'; Bad='FAIL' },
        @{ Label='DFSR Errors'; Log='DFS Replication';   Bad='WARN' }
    )
    foreach ($dc in $DomainControllers) {
        foreach ($l in $logs) {
            # -ErrorAction Stop: an unreachable DC must NOT be reported as "0 errors".
            try {
                $f = @{ LogName=$l.Log; Level=@(1,2); StartTime=$start }
                $ev = @(Get-WinEvent -ComputerName $dc -FilterHashtable $f -ErrorAction Stop)
                $ids = ($ev | Group-Object Id | Sort-Object Count -Descending | Select-Object -First 5 |
                        ForEach-Object { "$($_.Name)x$($_.Count)" }) -join ', '
                Add-Result 'EVENTS' "$($l.Label) $dc" $l.Bad "$($ev.Count) error/critical '$($l.Log)' event(s) (${EventHoursBack}h)" "top IDs: $ids"
            } catch {
                if (Test-NoEventsError $_) {
                    Add-Result 'EVENTS' "$($l.Label) $dc" 'PASS' "0 error/critical '$($l.Log)' events (${EventHoursBack}h)"
                } else {
                    Add-Result 'EVENTS' "$($l.Label) $dc" 'INFO' "Query failed on ${dc}: $_"
                }
            }
        }
    }
}

# ======================================================================
#  SUMMARY
# ======================================================================
Write-Host ""
Write-Host ("-" * 110) -ForegroundColor DarkGray
$pass = @($results | Where-Object Status -eq 'PASS').Count
$warn = @($results | Where-Object Status -eq 'WARN').Count
$fail = @($results | Where-Object Status -eq 'FAIL').Count
$info = @($results | Where-Object Status -eq 'INFO').Count
$overall = if ($fail -gt 0) {'FAIL'} elseif ($warn -gt 0) {'WARN'} else {'PASS'}
Write-Host "Overall: $overall   PASS=$pass  WARN=$warn  FAIL=$fail  INFO=$info" -ForegroundColor Cyan
Write-Log "Overall=$overall PASS=$pass WARN=$warn FAIL=$fail INFO=$info"

# ======================================================================
#  CSV
# ======================================================================
if ($Csv) {
    try {
        if (-not (Test-Path $LogRoot)) { New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null }
        $results | Select-Object *, @{n='Domain';e={$Domain}}, @{n='ScriptVersion';e={$ScriptVersion}} |
            Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        Write-Host "CSV : $CsvPath" -ForegroundColor Green
        Write-Log "CSV written: $CsvPath" -Level OK
    } catch { Write-Log "CSV write failed: $_" -Level ERROR }
}

# ======================================================================
#  HTML
# ======================================================================
if ($Html) {
    $rows = [System.Text.StringBuilder]::new()
    foreach ($r in $results) {
        $cls = switch ($r.Status) { 'PASS'{'status-pass'};'WARN'{'status-warn'};'FAIL'{'status-fail'};'INFO'{'status-info'} }
        $extra = if ($r.Extra) { "<br><small style='color:#555;font-family:monospace'>$(ConvertTo-HtmlText $r.Extra)</small>" } else { '' }
        [void]$rows.Append("<tr><td style='text-align:center'>$($r.Num)</td><td>$(ConvertTo-HtmlText $r.Section)</td><td>$(ConvertTo-HtmlText $r.Name)</td><td><span class='badge $cls'>$($r.Status)</span></td><td>$(ConvertTo-HtmlText $r.Detail)$extra</td></tr>`n")
    }
    $overallCls = switch ($overall) { 'PASS'{'status-pass'};'WARN'{'status-warn'};'FAIL'{'status-fail'} }
    $runDate = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $htmlDoc = @"
<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<title>Domain Health Audit - $(ConvertTo-HtmlText $Domain)</title>
<style>
 body{font-family:Segoe UI,Arial,sans-serif;font-size:13px;color:#222;margin:30px;}
 h1{color:#1a3a5c;border-bottom:2px solid #1a3a5c;padding-bottom:6px;}
 table{border-collapse:collapse;width:100%;margin-top:10px;}
 th{background:#1a3a5c;color:#fff;padding:7px 10px;text-align:left;}
 td{border:1px solid #ddd;padding:6px 10px;vertical-align:top;}
 tr:nth-child(even){background:#f7f9fc;}
 .badge{display:inline-block;padding:2px 8px;border-radius:4px;font-weight:bold;font-size:11px;}
 .status-pass{background:#d4edda;color:#155724;}
 .status-warn{background:#fff3cd;color:#856404;}
 .status-fail{background:#f8d7da;color:#721c24;}
 .status-info{background:#d1ecf1;color:#0c5460;}
 small{word-break:break-all;}
</style></head><body>
<h1>Windows Domain Health Audit</h1>
<table style="width:auto;border:none">
 <tr><td style="border:none;font-weight:bold">Domain</td><td style="border:none">$(ConvertTo-HtmlText $Domain)</td></tr>
 <tr><td style="border:none;font-weight:bold">Report Date</td><td style="border:none">$runDate</td></tr>
 <tr><td style="border:none;font-weight:bold">Overall</td><td style="border:none"><span class='badge $overallCls'>$overall</span></td></tr>
 <tr><td style="border:none;font-weight:bold">Summary</td><td style="border:none">PASS=$pass WARN=$warn FAIL=$fail INFO=$info</td></tr>
</table>
<table>
 <tr><th style="width:40px">#</th><th style="width:70px">Section</th><th style="width:200px">Check</th><th style="width:70px">Status</th><th>Detail</th></tr>
 $($rows.ToString())
</table>
<p><small>Generated by Invoke-DomainHealthAudit.ps1 v$ScriptVersion - read-only.</small></p>
</body></html>
"@
    try {
        if (-not (Test-Path $LogRoot)) { New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null }
        $htmlDoc | Set-Content -Path $HtmlPath -Encoding UTF8
        Write-Host "HTML: $HtmlPath" -ForegroundColor Green
        Write-Log "HTML written: $HtmlPath" -Level OK
    } catch { Write-Log "HTML write failed: $_" -Level ERROR }
}

Write-Log "Audit completed"
return [PSCustomObject]@{
    Domain        = $Domain
    ScriptVersion = $ScriptVersion
    Overall       = $overall
    Pass          = $pass
    Warn          = $warn
    Fail          = $fail
    Info          = $info
    Results       = $results
    HtmlPath      = if ($Html) { $HtmlPath } else { $null }
    CsvPath       = if ($Csv)  { $CsvPath }  else { $null }
}
