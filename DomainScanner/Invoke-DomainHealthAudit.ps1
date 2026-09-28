#Requires -Version 5.1
#Requires -RunAsAdministrator
#Requires -Modules ActiveDirectory

<#
.SYNOPSIS
    Read-only consolidated health/inventory audit for a Windows Active Directory domain.

.DESCRIPTION
    Single-pass, NON-DESTRUCTIVE audit that inventories a domain and flags things
    that are commonly wrong or worth handling. Generalized from a set of AD
    remediation scripts (health report, privileged-account audit, unsigned-LDAP
    clients, DC decommission readiness) into a domain-agnostic tool driven by
    parameters and an optional config file.

    NOTHING here changes state. There is no -Execute parameter. Always safe to re-run.

    Check sections (each can be skipped with -Skip):
      FOREST   Forest/domain functional level, UPN suffixes, tombstone lifetime
      AD       DC inventory, FSMO placement, replication failures, SYSVOL, DFSR/FRS state
      SITES    Sites, subnets, sites without subnets, DC-to-site mapping
      DNS      DC name resolution, scavenging/aging, stale A records
      CERT     LDAPS TLS handshake + cert expiry, expiring DC machine certs, CA liveness,
               autoenrollment GPOs vs dead CA, NTAuth store contents
      USERS    Privileged group membership + risk scoring, krbtgt age,
               password-never-expires, stale users/computers
      NET      DC port reachability (LDAP/LDAPS/KRB/GC/DNS/SMB), unsigned LDAP binds (2889/2887)
      LICENSE  Windows activation state on DCs (best-effort)
      EVENTS   Critical Directory Service / DFSR errors in the last N hours

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
    dead CA host, checks to skip). See sample.config.json for the format.

.PARAMETER Skip
    One or more section names to skip: FOREST, AD, SITES, DNS, CERT, USERS, NET, LICENSE, EVENTS.

.PARAMETER SkipGpoScan
    Skip the per-GPO XML report scan in the CERT section (slow on domains with many GPOs).

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
    Version   : 1.1.0
    Read-only : always safe to re-run

    Changelog (full history in CHANGELOG.md):
      1.1.0  2026-09-28  Fixes from review + test run in an isolated Server 2022 lab:
                         parse errors (PS5.1 encoding + "$dc:" refs), TCP probe reported
                         refused ports as open, replication check could never fail,
                         stale-A filter matched every record, autoenrollment GPO false
                         positives, certutil -viewstore opened a GUI dialog, EVENTS
                         reported PASS when the DC was unreachable. Adds LDAPS TLS
                         handshake + cert expiry, version stamping, -SkipGpoScan.
      1.0.0  2026-09     Initial consolidated seed.
#>

[CmdletBinding()]
param(
    [string]$Domain = $env:USERDNSDOMAIN,

    [string[]]$DomainControllers,

    [string]$ConfigPath = '',

    [ValidateSet('FOREST','AD','SITES','DNS','CERT','USERS','NET','LICENSE','EVENTS')]
    [string[]]$Skip = @(),

    [switch]$SkipGpoScan,

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

$ScriptVersion = '1.1.0'
$AllSections   = @('FOREST','AD','SITES','DNS','CERT','USERS','NET','LICENSE','EVENTS')

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
    Write-Host (" [{0,3}] [{1,-7}] [{2,-4}] {3,-38} {4}" -f `
        $script:checkNum, $Section, $Status, $Name, $Detail) -ForegroundColor $color
    Write-Log "$Section | $Status | $Name | $Detail"
}

function ConvertTo-HtmlText {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    ($Text -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;')
}

function Test-TcpPort {
    param([string]$HostName, [int]$Port, [int]$TimeoutMs = 3000)
    # WaitOne() also returns true when the connect FAILED fast (RST / refused), so the
    # handle alone is not proof of an open port - EndConnect() surfaces the failure.
    $tcp = New-Object System.Net.Sockets.TcpClient
    try {
        $ar = $tcp.BeginConnect($HostName, $Port, $null, $null)
        if (-not $ar.AsyncWaitHandle.WaitOne($TimeoutMs, $false)) { return $false }
        $tcp.EndConnect($ar)
        return $tcp.Connected
    } catch { return $false } finally { $tcp.Close() }
}

function Get-LdapsCertificate {
    # Completes a real TLS handshake on 636 and returns the server certificate.
    # A DC listens on 636 even with no usable cert; only a handshake proves LDAPS works.
    param([string]$HostName, [int]$TimeoutMs = 5000)
    $tcp = New-Object System.Net.Sockets.TcpClient
    try {
        $ar = $tcp.BeginConnect($HostName, 636, $null, $null)
        if (-not $ar.AsyncWaitHandle.WaitOne($TimeoutMs, $false)) { throw "TCP 636 timed out" }
        $tcp.EndConnect($ar)
        $callback = [System.Net.Security.RemoteCertificateValidationCallback]{ param($s,$c,$ch,$e) $true }
        $ssl = New-Object System.Net.Security.SslStream($tcp.GetStream(), $false, $callback)
        try {
            $ssl.ReadTimeout = $TimeoutMs
            $ssl.AuthenticateAsClient($HostName)
            return New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($ssl.RemoteCertificate)
        } finally { $ssl.Dispose() }
    } finally { $tcp.Close() }
}

function Test-NoEventsError {
    # Locale-independent: Get-WinEvent raises NoMatchingEventsFound when the filter is simply empty.
    param($ErrorRecord)
    return ("$($ErrorRecord.FullyQualifiedErrorId)" -match 'NoMatchingEventsFound')
}

function Test-SectionEnabled { param([string]$Section) return ($Section -notin $SkipList) }

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
Write-Host (" {0,-5} {1,-9} {2,-6} {3,-38} {4}" -f '#','Section','Status','Check','Detail') -ForegroundColor White
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
#  SECTION: AD  (DCs, FSMO, replication, SYSVOL, DFSR)
# ======================================================================
if (Test-SectionEnabled 'AD') {
    # DC inventory
    if ($dcObjects) {
        $dcNames = ($dcObjects | ForEach-Object { $_.HostName }) -join ', '
        Add-Result 'AD' 'DC Inventory' 'INFO' "$(@($dcObjects).Count) DC(s): $dcNames"
    } else {
        Add-Result 'AD' 'DC Inventory' 'FAIL' 'No domain controllers discovered'
    }

    # FSMO
    try {
        $dom = Get-ADDomain -Server $Domain -ErrorAction Stop
        $fst = Get-ADForest -Server $Domain -ErrorAction Stop
        $fsmo = [ordered]@{
            PDCEmulator          = $dom.PDCEmulator
            RIDMaster            = $dom.RIDMaster
            InfrastructureMaster = $dom.InfrastructureMaster
            SchemaMaster         = $fst.SchemaMaster
            DomainNamingMaster   = $fst.DomainNamingMaster
        }
        $fsmoLines = ($fsmo.GetEnumerator() | ForEach-Object { "$($_.Key)=$(($_.Value -split '\.')[0])" }) -join '; '
        if ($Config.ExpectedFsmoHolder) {
            $offHolder = $fsmo.Values | Where-Object { (($_ -split '\.')[0]) -ine $Config.ExpectedFsmoHolder }
            if ($offHolder) {
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
#  SECTION: CERT  (LDAPS + PKI + expiring DC certs)
# ======================================================================
if (Test-SectionEnabled 'CERT') {
    foreach ($dc in $DomainControllers) {
        # LDAPS: real TLS handshake + served cert expiry
        try {
            $lc = Get-LdapsCertificate -HostName $dc
            $days = [int][math]::Floor(($lc.NotAfter - (Get-Date)).TotalDays)
            # DC certs from the Kerberos Authentication template have an empty Subject (SAN only).
            $certName = if ($lc.Subject) { $lc.Subject } else { "SAN:" + $lc.GetNameInfo([System.Security.Cryptography.X509Certificates.X509NameType]::DnsName, $false) }
            $desc = "$certName expires $($lc.NotAfter.ToString('yyyy-MM-dd')) ($days d)"
            if ($days -lt 0) {
                Add-Result 'CERT' "LDAPS $dc" 'FAIL' "LDAPS cert EXPIRED: $desc" "Issuer: $($lc.Issuer)"
            } elseif ($days -lt $CertExpiryWarnDays) {
                Add-Result 'CERT' "LDAPS $dc" 'WARN' "LDAPS cert expires soon: $desc" "Issuer: $($lc.Issuer)"
            } else {
                Add-Result 'CERT' "LDAPS $dc" 'PASS' "TLS handshake OK, $desc" "Issuer: $($lc.Issuer)"
            }
        } catch {
            Add-Result 'CERT' "LDAPS $dc" 'WARN' "No working LDAPS on ${dc}:636 (no usable cert, or blocked): $($_.Exception.InnerException.Message) $($_.Exception.Message)"
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
    $haveCertutil = [bool](Get-Command certutil.exe -ErrorAction SilentlyContinue)
    foreach ($ca in $registeredCAs) {
        $caHost = $ca.dNSHostName
        $caName = if ($ca.displayName) { $ca.displayName } else { $ca.Name }
        if (-not $caHost) { continue }
        if (-not $haveCertutil) {
            Add-Result 'CERT' "CA Liveness $caName" 'INFO' 'certutil.exe not available - liveness not tested'
            continue
        }
        $tcp135  = Test-TcpPort -HostName $caHost -Port 135 -TimeoutMs 3000
        $pingOut = & certutil.exe -ping -config "$caHost\$caName" 2>&1 | Out-String
        $pingRc  = $LASTEXITCODE
        if ($pingRc -eq 0) {
            Add-Result 'CERT' "CA Liveness $caName" 'PASS' "$caHost answers certutil -ping"
        } else {
            $deadCAs.Add("$caName ($caHost)")
            $reason = if (-not $tcp135) { 'no RPC/135 and -ping failed' } else { 'RPC reachable but -ping failed' }
            $last   = (($pingOut -split "`r?`n") | Where-Object { $_.Trim() } | Select-Object -Last 1)
            Add-Result 'CERT' "CA Liveness $caName" 'FAIL' "$caHost does NOT answer as a CA - $reason (exit $pingRc)" "$last"
        }
    }

    # Autoenrollment GPOs still in play - cross-reference against dead CAs.
    # Structured check on the Public Key policy node: EnrollCertificatesAutomatically=true,
    # in an enabled Computer/User half, in a GPO that has at least one enabled link.
    # (A plain 'AutoEnrollment' text match also hits GPOs that explicitly DISABLE it.)
    if ($SkipGpoScan) {
        Add-Result 'CERT' 'Autoenrollment GPOs' 'INFO' 'GPO scan skipped (-SkipGpoScan)'
    } else {
        try {
            Import-Module GroupPolicy -ErrorAction Stop
            $aeGpos = [System.Collections.Generic.List[string]]::new()
            $aeUnlinked = 0
            foreach ($g in @(Get-GPO -All -Domain $Domain -ErrorAction Stop)) {
                try {
                    [xml]$rep = Get-GPOReport -Guid $g.Id -ReportType Xml -Domain $Domain -ErrorAction Stop
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
                    $links = @($rep.GPO.LinksTo | Where-Object { $_ -and $_.Enabled -eq 'true' } | ForEach-Object { $_.SOMPath })
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
                $deadNames = @($deadCAs | ForEach-Object { ($_ -split ' \(')[0] })
                $ntDead = @($ntCerts | Where-Object { $s = $_.Subject; @($deadNames | Where-Object { $s -match ('CN=' + [regex]::Escape($_) + '(,|$)') }).Count -gt 0 })
                if ($ntDead.Count -gt 0) {
                    Add-Result 'CERT' 'NTAuth Store' 'WARN' "$($ntCerts.Count) cert(s) in NTAuth, $($ntDead.Count) belong to a dead CA" ($ntLines -join '; ')
                } else {
                    Add-Result 'CERT' 'NTAuth Store' 'INFO' "$($ntCerts.Count) cert(s) in NTAuth - none match a dead CA" ($ntLines -join '; ')
                }
            }
        } catch { Add-Result 'CERT' 'NTAuth Store' 'INFO' "Could not read NTAuthCertificates: $_" }
    }
}

# ======================================================================
#  SECTION: USERS  (privileged risk + hygiene)
# ======================================================================
if (Test-SectionEnabled 'USERS') {
    $auditGroups    = @('Domain Admins','Enterprise Admins','Schema Admins','Administrators','Backup Operators','Account Operators')
    $highPrivGroups = @('Domain Admins','Enterprise Admins','Schema Admins')
    $svcRegex       = $Config.ServiceAcctRegex
    $namedAdminRegex = $Config.NamedAdminRegex
    $knownRisk      = @($Config.KnownRiskAccounts)

    $membershipMap = @{}
    foreach ($g in $auditGroups) {
        try {
            $members = Get-ADGroupMember -Identity $g -Recursive -Server $Domain -ErrorAction Stop |
                       Where-Object objectClass -eq 'user'
            foreach ($m in $members) {
                if (-not $membershipMap.ContainsKey($m.SamAccountName)) {
                    $membershipMap[$m.SamAccountName] = [System.Collections.Generic.List[string]]::new()
                }
                if ($g -notin $membershipMap[$m.SamAccountName]) { $membershipMap[$m.SamAccountName].Add($g) }
            }
        } catch {
            Write-Log "Could not query group '$g': $_" -Level WARN
            Add-Result 'USERS' "Group $g" 'INFO' "Could not expand membership (foreign/orphaned members?): $_"
        }
    }

    $high = 0; $review = 0; $ok = 0
    $privDetail = [System.Collections.Generic.List[string]]::new()
    foreach ($sam in ($membershipMap.Keys | Sort-Object)) {
        try {
            $u = Get-ADUser -Identity $sam -Server $Domain -Properties Enabled,LastLogonDate,PasswordNeverExpires,Description,DisplayName -ErrorAction Stop
            $groups = $membershipMap[$sam]
            $inHigh = @($groups | Where-Object { $_ -in $highPrivGroups }).Count -gt 0
            $days   = if ($u.LastLogonDate) { (New-TimeSpan -Start $u.LastLogonDate -End (Get-Date)).Days } else { $null }

            $risk = 'OK'; $reason = 'Active account, no risk flags'
            if (-not $u.Enabled -and $inHigh)                { $risk='HIGH-RISK'; $reason='Disabled account in high-priv group' }
            elseif ($u.SamAccountName -in $knownRisk)         { $risk='HIGH-RISK'; $reason='Known vendor/external account' }
            elseif ($null -eq $u.LastLogonDate -and $inHigh)  { $risk='HIGH-RISK'; $reason='Never logged in and in high-priv group' }
            elseif ($u.SamAccountName -match $svcRegex -and $inHigh) { $risk='HIGH-RISK'; $reason='Service-account name pattern in high-priv group' }
            elseif ($null -ne $days -and $days -gt $InactiveDays -and $inHigh) { $risk='REVIEW'; $reason="Last logon $days days ago in high-priv group" }
            elseif ($u.SamAccountName -match $namedAdminRegex -and $inHigh) { $risk='REVIEW'; $reason='Named admin account in high-priv group - verify still needed' }

            switch ($risk) { 'HIGH-RISK'{$high++} 'REVIEW'{$review++} default{$ok++} }
            if ($risk -ne 'OK') { $privDetail.Add("$sam [$risk]: $reason ($($groups -join '/'))") }
        } catch { Write-Log "Could not read user '$sam': $_" -Level WARN }
    }

    $privStatus = if ($high -gt 0) {'FAIL'} elseif ($review -gt 0) {'WARN'} else {'PASS'}
    Add-Result 'USERS' 'Privileged Accounts' $privStatus "HIGH-RISK=$high REVIEW=$review OK=$ok (total $($membershipMap.Count))" ($privDetail -join ' | ')

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

    # Enabled users, password-never-expires (LDAP filter: server-side, no full pull)
    $ldapEnabled = '(!(userAccountControl:1.2.840.113556.1.4.803:=2))'
    try {
        $pne = @(Get-ADUser -LDAPFilter "(&$ldapEnabled(userAccountControl:1.2.840.113556.1.4.803:=65536))" -Server $Domain -ErrorAction Stop)
        if ($pne.Count -gt 0) {
            Add-Result 'USERS' 'Password Never Expires' 'WARN' "$($pne.Count) enabled user(s) with non-expiring passwords" (($pne | Select-Object -First 25 | ForEach-Object { $_.SamAccountName }) -join ', ')
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
            Add-Result 'USERS' 'Stale Enabled Users' 'WARN' "$($stale.Count) enabled user(s) inactive > $InactiveDays days" (($stale | Select-Object -First 25 | ForEach-Object { $_.SamAccountName }) -join ', ')
        } else {
            Add-Result 'USERS' 'Stale Enabled Users' 'PASS' "No enabled users inactive > $InactiveDays days"
        }
    } catch { Add-Result 'USERS' 'Stale Enabled Users' 'INFO' "Query failed: $_" }

    try {
        $staleC = @(Get-ADComputer -LDAPFilter "(&$ldapEnabled(lastLogonTimestamp<=$cutFt))" -Server $Domain -ErrorAction Stop)
        if ($staleC.Count -gt 0) {
            Add-Result 'USERS' 'Stale Computer Accounts' 'WARN' "$($staleC.Count) computer(s) inactive > $InactiveDays days" (($staleC | Select-Object -First 25 | ForEach-Object { $_.Name }) -join ', ')
        } else {
            Add-Result 'USERS' 'Stale Computer Accounts' 'PASS' "No computers inactive > $InactiveDays days"
        }
    } catch { Add-Result 'USERS' 'Stale Computer Accounts' 'INFO' "Query failed: $_" }
}

# ======================================================================
#  SECTION: NET  (DC port reachability + unsigned LDAP binds)
# ======================================================================
if (Test-SectionEnabled 'NET') {
    $portMap = [ordered]@{ 'LDAP 389'=389; 'LDAPS 636'=636; 'Kerberos 88'=88; 'GC 3268'=3268; 'DNS 53'=53; 'SMB 445'=445 }
    foreach ($dc in $DomainControllers) {
        $open = [System.Collections.Generic.List[string]]::new()
        $shut = [System.Collections.Generic.List[string]]::new()
        foreach ($p in $portMap.GetEnumerator()) {
            if (Test-TcpPort -HostName $dc -Port $p.Value) { $open.Add($p.Key) } else { $shut.Add($p.Key) }
        }
        if ($shut.Count -eq 0) {
            Add-Result 'NET' "DC Ports $dc" 'PASS' 'All core AD ports reachable' ("open: " + ($open -join ', '))
        } else {
            Add-Result 'NET' "DC Ports $dc" 'WARN' "$($shut.Count) core port(s) unreachable" ("closed: " + ($shut -join ', '))
        }
    }

    # Stale IPs still answering LDAP
    foreach ($ip in @($Config.StaleIPs | Where-Object { $_ })) {
        if (Test-TcpPort -HostName $ip -Port 389) {
            Add-Result 'NET' "Stale LDAP $ip" 'WARN' "$ip still answers on 389 - expected decommissioned"
        } else {
            Add-Result 'NET' "Stale LDAP $ip" 'PASS' "$ip does not answer on 389"
        }
    }

    # Unsigned / simple LDAP binds. 2889 (per-client) is only logged when
    # "16 LDAP Interface Events" >= 2; 2887 (daily count) is always logged.
    foreach ($dc in $DomainControllers) {
        try {
            $r = Invoke-Command -ComputerName $dc -ArgumentList $EventHoursBack -ErrorAction Stop -ScriptBlock {
                param($Hours)
                $start = (Get-Date).AddHours(-1 * $Hours)
                $lvl = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Services\NTDS\Diagnostics' -ErrorAction SilentlyContinue).'16 LDAP Interface Events'
                $out = [ordered]@{ Level = [int]$lvl; Clients = @(); Count2887 = 0; Error = '' }
                try {
                    $out.Clients = @(Get-WinEvent -FilterHashtable @{ LogName='Directory Service'; Id=2889; StartTime=$start } -ErrorAction Stop |
                        ForEach-Object {
                            $ep = [string]$_.Properties[0].Value           # "ip:port" or "[v6]:port"
                            $i  = $ep.LastIndexOf(':')
                            if ($i -gt 0) { $ep.Substring(0, $i).Trim('[',']') } else { $ep }
                        } | Sort-Object -Unique)
                } catch { if ("$($_.FullyQualifiedErrorId)" -notmatch 'NoMatchingEventsFound') { $out.Error = "$_" } }
                try {
                    $out.Count2887 = @(Get-WinEvent -FilterHashtable @{ LogName='Directory Service'; Id=2887; StartTime=$start } -ErrorAction Stop).Count
                } catch { if ("$($_.FullyQualifiedErrorId)" -notmatch 'NoMatchingEventsFound' -and -not $out.Error) { $out.Error = "$_" } }
                [PSCustomObject]$out
            }
            if ($r.Error) {
                Add-Result 'NET' "Unsigned LDAP $dc" 'INFO' "Could not read Directory Service log on ${dc}: $($r.Error)"
            } elseif (@($r.Clients).Count -gt 0) {
                Add-Result 'NET' "Unsigned LDAP $dc" 'WARN' "$(@($r.Clients).Count) client(s) doing unsigned/simple binds (last ${EventHoursBack}h)" (@($r.Clients) -join ', ')
            } elseif ($r.Count2887 -gt 0) {
                Add-Result 'NET' "Unsigned LDAP $dc" 'WARN' "Event 2887 reports unsigned binds, but per-client 2889 logging is off (level $($r.Level)) - set '16 LDAP Interface Events'=2 to see who"
            } elseif ($r.Level -lt 2) {
                Add-Result 'NET' "Unsigned LDAP $dc" 'PASS' "No 2887 in last ${EventHoursBack}h (2887 is daily; 2889 logging off, level $($r.Level))"
            } else {
                Add-Result 'NET' "Unsigned LDAP $dc" 'PASS' "No Event 2889/2887 in last ${EventHoursBack}h"
            }
        } catch {
            Add-Result 'NET' "Unsigned LDAP $dc" 'INFO' "Could not query ${dc} (remoting blocked?): $_"
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
