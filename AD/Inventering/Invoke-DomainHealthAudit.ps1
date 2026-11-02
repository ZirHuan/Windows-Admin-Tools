#Requires -Version 5.1
#Requires -RunAsAdministrator
#Requires -Modules ActiveDirectory

<#
.SYNOPSIS
    Read-only consolidated health and assessment audit for a Windows Active Directory domain.

.DESCRIPTION
    Single-pass, NON-DESTRUCTIVE audit that inventories a domain and surfaces common
    configuration problems, hygiene issues, and security risks in on-premises AD environments.
    Domain-agnostic; driven by parameters and an optional JSON config file per environment.

    NOTHING here changes state. Always safe to re-run.

    Check sections (each can be skipped with -Skip):
      FOREST    Forest/domain functional level, UPN suffixes, tombstone lifetime
      AD        DC inventory, FSMO, replication, SYSVOL, DFSR/FRS, PDC time sync
      SITES     Sites, subnets, coverage gaps, DC-to-site mapping
      DNS       Zone scavenging/aging, stale A records, DC name resolution
      CERT      LDAPS + expiry, CA/PKI, dead-CA detection, autoenrollment GPOs, NTAuth,
                unlinked GPOs
      USERS     Privileged group risk scoring, password policy, AdminSDHolder orphans,
                AS-REP roastable accounts, stale/inactive accounts, krbtgt age
      SECURITY  Unconstrained Kerberos delegation, duplicate SPNs, LAPS deployment,
                Kerberoastable service accounts
      NET       DC port reachability (LDAP/LDAPS/KRB/GC/DNS/SMB), unsigned LDAP binds (2889)
      LICENSE   KMS/Windows activation state on reachable servers (best-effort)
      EVENTS    Critical Directory Service / DFSR errors in the last N hours

    Output:
      - Colour-coded console table (PASS / WARN / FAIL / INFO)
      - Structured object returned to the pipeline
      - Optional HTML report (-Html) and CSV of findings (-Csv)

.PARAMETER Domain
    FQDN of the domain to audit. Defaults to the current machine's domain.

.PARAMETER DomainControllers
    Explicit list of DC FQDNs to probe. If omitted, discovered via Get-ADDomainController.

.PARAMETER ConfigPath
    Optional JSON config file with environment-specific expectations (known-risk accounts,
    expected sites, stale IPs/DC names, expected FSMO holder, dead CA info, checks to skip).
    See example.config.json for the schema.

.PARAMETER Skip
    One or more section names to skip:
    FOREST, AD, SITES, DNS, CERT, USERS, SECURITY, NET, LICENSE, EVENTS.

.PARAMETER EventHoursBack
    Hours of event-log history to scan in EVENTS and NET sections. Default 24.

.PARAMETER InactiveDays
    Threshold (days since last logon) for flagging stale users/computers. Default 90.

.PARAMETER CertExpiryWarnDays
    Warn when a DC/LDAPS certificate expires within this many days. Default 60.

.PARAMETER LogRoot
    Root folder for the log file and any reports. Default C:\Logs\DomainAudit.

.PARAMETER Html
    Write a colour-coded HTML report suitable for customer hand-off.

.PARAMETER Csv
    Write a flat CSV of every finding.

.EXAMPLE
    .\Invoke-DomainHealthAudit.ps1 -Verbose
    Full audit of the current domain, console output only.

.EXAMPLE
    .\Invoke-DomainHealthAudit.ps1 -Domain customer.local -ConfigPath .\customer.config.json -Html -Csv
    Audit with environment config, write HTML + CSV report.

.EXAMPLE
    .\Invoke-DomainHealthAudit.ps1 -Skip LICENSE,NET
    Skip licensing and network sections (e.g. no remote access to DCs).

.NOTES
    Author    : Iver / IT Consult (Christer Rosvall)
    Read-only : always safe to re-run
#>

[CmdletBinding()]
param(
    [string]$Domain = $env:USERDNSDOMAIN,

    [string[]]$DomainControllers,

    [string]$ConfigPath = '',

    [ValidateSet('FOREST','AD','SITES','DNS','CERT','USERS','SECURITY','NET','LICENSE','EVENTS')]
    [string[]]$Skip = @(),

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

# ======================================================================
#  ENVIRONMENT CONFIG (defaults, overridable via -ConfigPath JSON)
# ======================================================================
$Config = [ordered]@{
    KnownRiskAccounts = @()            # sam names known to be vendor/external/orphaned
    ExpectedSites     = @()            # site names that should exist (INFO only)
    StaleDCNames      = @()            # decommissioned DC short names to hunt for
    StaleIPs          = @()            # IPs that should no longer answer LDAP
    ExpectedFsmoHolder= ''             # short name if all 5 FSMO should sit on one DC
    ServiceAcctRegex  = '^(svc[-_]|sa[-_]|sql|exchange|iis|backup|print|scan|service|mssql)'
    AdminAcctRegex    = '^(adm\.|admin\.)'
}
if ($ConfigPath -and (Test-Path $ConfigPath)) {
    try {
        $loaded = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json
        foreach ($k in $Config.Keys.Clone()) {
            if ($null -ne $loaded.$k) { $Config[$k] = $loaded.$k }
        }
        if ($loaded.Skip) { $Skip = @($Skip + $loaded.Skip | Select-Object -Unique) }
    } catch {
        Write-Warning "Could not parse config '$ConfigPath': $_ — using defaults."
    }
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
    Write-Host (" [{0,3}] [{1,-8}] [{2,-4}] {3,-42} {4}" -f `
        $script:checkNum, $Section, $Status, $Name, $Detail) -ForegroundColor $color
    Write-Log "$Section | $Status | $Name | $Detail"
}

function Escape-Html {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    ($Text -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;')
}

function Test-TcpPort {
    param([string]$HostName, [int]$Port, [int]$TimeoutMs = 3000)
    $tcp = New-Object System.Net.Sockets.TcpClient
    try {
        $ar = $tcp.BeginConnect($HostName, $Port, $null, $null)
        return $ar.AsyncWaitHandle.WaitOne($TimeoutMs, $false)
    } catch { return $false } finally { $tcp.Close() }
}

function Should-Run { param([string]$Section) return ($Section -notin $Skip) }

# ======================================================================
#  HEADER
# ======================================================================
Write-Log "Audit started. Domain=$Domain Skip=$($Skip -join ',')"
Write-Host ""
Write-Host "=== WINDOWS DOMAIN HEALTH AUDIT ===" -ForegroundColor Cyan
Write-Host "  Domain     : $Domain"
Write-Host "  Date       : $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
Write-Host "  Skipping   : $(if($Skip){$Skip -join ', '}else{'(nothing)'})"
Write-Host ""
Write-Host (" {0,-5} {1,-10} {2,-6} {3,-44} {4}" -f '#','Section','Status','Check','Detail') -ForegroundColor White
Write-Host ("-" * 115) -ForegroundColor DarkGray

# Resolve DC list once (used by several sections)
$dcObjects = @()
try {
    $dcObjects = Get-ADDomainController -Filter * -Server $Domain -ErrorAction Stop
} catch {
    Write-Log "Get-ADDomainController failed: $_" -Level WARN
}
if (-not $DomainControllers) {
    $DomainControllers = @($dcObjects | ForEach-Object { $_.HostName })
}

# ======================================================================
#  SECTION: FOREST
# ======================================================================
if (Should-Run 'FOREST') {
    try {
        $dom = Get-ADDomain -Server $Domain -ErrorAction Stop
        $fst = Get-ADForest -Server $Domain -ErrorAction Stop

        Add-Result 'FOREST' 'Domain Functional Level' 'INFO' "$($dom.DomainMode)"
        Add-Result 'FOREST' 'Forest Functional Level' 'INFO' "$($fst.ForestMode)"

        if ("$($dom.DomainMode)" -match '2000|2003|2008') {
            Add-Result 'FOREST' 'DFL is legacy' 'WARN' "Domain mode $($dom.DomainMode) predates 2012R2 — consider raising"
        }

        $upnList = if ($fst.UPNSuffixes) { $fst.UPNSuffixes -join ', ' } else { '(none beyond default)' }
        Add-Result 'FOREST' 'UPN Suffixes' 'INFO' $upnList

        try {
            $configNC = (Get-ADRootDSE -Server $Domain).configurationNamingContext
            $tsl = (Get-ADObject "CN=Directory Service,CN=Windows NT,CN=Services,$configNC" `
                    -Server $Domain -Properties tombstoneLifetime -ErrorAction Stop).tombstoneLifetime
            if (-not $tsl) { $tsl = 60 }
            if ($tsl -lt 180) {
                Add-Result 'FOREST' 'Tombstone Lifetime' 'WARN' "$tsl days (low — 180 recommended)"
            } else {
                Add-Result 'FOREST' 'Tombstone Lifetime' 'PASS' "$tsl days"
            }
        } catch { Add-Result 'FOREST' 'Tombstone Lifetime' 'INFO' "Could not read: $_" }
    } catch {
        Add-Result 'FOREST' 'Forest/Domain query' 'FAIL' "Could not query domain '$Domain': $_"
    }
}

# ======================================================================
#  SECTION: AD  (DCs, FSMO, replication, SYSVOL, DFSR, time)
# ======================================================================
if (Should-Run 'AD') {
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

    # Replication — structured cmdlet first, repadmin fallback
    try {
        $replFailures = @(Get-ADReplicationFailure -Scope Forest -Server $Domain -ErrorAction Stop)
        if ($replFailures.Count -gt 0) {
            Add-Result 'AD' 'AD Replication' 'FAIL' "$($replFailures.Count) replication failure(s)" `
                (($replFailures | ForEach-Object { "$($_.Server) partner=$($_.Partner) err=$($_.LastError) fails=$($_.FailureCount)" }) -join ' | ')
        } else {
            Add-Result 'AD' 'AD Replication' 'PASS' '0 replication failures'
        }
    } catch {
        try {
            $replOut   = repadmin /replsummary 2>&1 | Out-String
            $failLines = ($replOut -split "`n") | Where-Object { $_ -match '\d+\s+fail' -and $_ -notmatch '^\s*0\s+fail' }
            if ($failLines) {
                Add-Result 'AD' 'AD Replication' 'FAIL' 'Failures detected (repadmin fallback)' (($failLines | Select-Object -First 5) -join ' | ')
            } else {
                Add-Result 'AD' 'AD Replication' 'PASS' '0 failures in repadmin /replsummary'
            }
        } catch { Add-Result 'AD' 'AD Replication' 'WARN' "Could not check replication: $_" }
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
        $frs = Get-ADObject -Server $Domain -Filter { objectClass -eq 'nTFRSSubscriber' } -ErrorAction SilentlyContinue
        if ($frs) {
            Add-Result 'AD' 'SYSVOL Engine' 'WARN' 'Legacy FRS objects present — migrate SYSVOL to DFSR'
        } else {
            Add-Result 'AD' 'SYSVOL Engine' 'PASS' 'No legacy FRS subscriber objects (DFSR assumed)'
        }
    } catch { Add-Result 'AD' 'SYSVOL Engine' 'INFO' "Could not determine SYSVOL engine: $_" }

    # PDC Emulator time sync
    try {
        $pdcFqdn = (Get-ADDomain -Server $Domain -ErrorAction Stop).PDCEmulator
        $w32out  = w32tm /monitor /computers:$pdcFqdn /nowarn 2>&1 | Out-String
        if ($w32out -match 'error|stratum:\s*0\b|RefID:\s*0\.0\.0\.0') {
            Add-Result 'AD' 'PDC Time Sync' 'WARN' "Time sync issue on PDC ($pdcFqdn) — verify w32tm source"
        } else {
            Add-Result 'AD' 'PDC Time Sync' 'PASS' "PDC ($pdcFqdn) time sync OK"
        }
    } catch { Add-Result 'AD' 'PDC Time Sync' 'INFO' "Could not check PDC time sync: $_" }
}

# ======================================================================
#  SECTION: SITES  (sites, subnets, coverage gaps)
# ======================================================================
if (Should-Run 'SITES') {
    try {
        $sites   = Get-ADReplicationSite   -Filter * -Server $Domain -ErrorAction Stop
        $subnets = Get-ADReplicationSubnet -Filter * -Server $Domain -Properties siteObject -ErrorAction SilentlyContinue

        Add-Result 'SITES' 'Site Inventory'   'INFO' "$(@($sites).Count) site(s): $(($sites.Name) -join ', ')"
        Add-Result 'SITES' 'Subnet Inventory' 'INFO' "$(@($subnets).Count) subnet(s) defined"

        $mappedSiteDNs = @($subnets | ForEach-Object { $_.siteObject } | Where-Object { $_ })
        $emptySites    = $sites | Where-Object { $_.DistinguishedName -notin $mappedSiteDNs }
        if ($emptySites) {
            Add-Result 'SITES' 'Sites Without Subnets' 'WARN' "$(@($emptySites).Count) site(s) have no subnet: $(($emptySites.Name) -join ', ')"
        } else {
            Add-Result 'SITES' 'Sites Without Subnets' 'PASS' 'Every site has at least one subnet'
        }

        if ($dcObjects) {
            $dcSiteLines = ($dcObjects | ForEach-Object { "$($_.Name)->$($_.Site)" }) -join '; '
            Add-Result 'SITES' 'DC Site Mapping' 'INFO' 'DC-to-site placement' $dcSiteLines
            if ($Config.ExpectedSites) {
                $badSite = $dcObjects | Where-Object { $_.Site -notin $Config.ExpectedSites }
                if ($badSite) {
                    Add-Result 'SITES' 'DC In Expected Site' 'WARN' "DC(s) outside expected sites: $(($badSite.Name) -join ', ')"
                }
            }
        }
    } catch { Add-Result 'SITES' 'Sites & Subnets' 'WARN' "Could not query sites: $_" }
}

# ======================================================================
#  SECTION: DNS
# ======================================================================
if (Should-Run 'DNS') {
    $dnsServer = if ($DomainControllers.Count -gt 0) { $DomainControllers[0] } else { $Domain }

    foreach ($dc in $DomainControllers) {
        try {
            $r = Resolve-DnsName -Name $dc -Type A -Server $dnsServer -ErrorAction Stop
            if ($r) { Add-Result 'DNS' "Resolve $dc" 'PASS' "$dc -> $(($r | Where-Object {$_.IPAddress}).IPAddress -join ', ')" }
        } catch { Add-Result 'DNS' "Resolve $dc" 'WARN' "Could not resolve $dc: $_" }
    }

    try {
        $zone = Get-DnsServerZoneAging -Name $Domain -ComputerName $dnsServer -ErrorAction Stop
        if ($zone.AgingEnabled) {
            Add-Result 'DNS' 'Zone Scavenging' 'PASS' "Aging enabled (refresh $($zone.RefreshInterval), no-refresh $($zone.NoRefreshInterval))"
        } else {
            Add-Result 'DNS' 'Zone Scavenging' 'WARN' "Aging/scavenging NOT enabled on $Domain — stale records will accumulate"
        }
    } catch { Add-Result 'DNS' 'Zone Scavenging' 'INFO' "Could not read aging: $_" }

    if ($Config.StaleIPs -or $Config.StaleDCNames) {
        try {
            $aRecords = Get-DnsServerResourceRecord -ComputerName $dnsServer -ZoneName $Domain -RRType A -ErrorAction Stop
            $hits = $aRecords | Where-Object {
                $rec = $_
                ($Config.StaleIPs -contains $rec.RecordData.IPv4Address.IPAddressToString) -or
                ($Config.StaleDCNames | Where-Object { $_ -and ($rec.HostName -ieq $_ -or $rec.HostName -ieq ($_ -split '\.')[0]) })
            }
            if ($hits) {
                Add-Result 'DNS' 'Stale A Records' 'WARN' "$(@($hits).Count) A record(s) match known decommissioned hosts/IPs" `
                    (($hits | ForEach-Object { "$($_.HostName)=$($_.RecordData.IPv4Address)" }) -join '; ')
            } else {
                Add-Result 'DNS' 'Stale A Records' 'PASS' 'No A records match known stale hosts/IPs'
            }
        } catch { Add-Result 'DNS' 'Stale A Records' 'INFO' "Could not enumerate A records: $_" }
    }
}

# ======================================================================
#  SECTION: CERT  (LDAPS + PKI + dead CA detection + GPO health)
# ======================================================================
if (Should-Run 'CERT') {
    foreach ($dc in $DomainControllers) {
        $ldaps = Test-TcpPort -HostName $dc -Port 636
        if ($ldaps) {
            Add-Result 'CERT' "LDAPS 636 $dc" 'PASS' "$dc answers on 636"
        } else {
            Add-Result 'CERT' "LDAPS 636 $dc" 'WARN' "$dc did not answer on 636 (no LDAPS cert bound, or blocked)"
        }
        try {
            $expiring = Invoke-Command -ComputerName $dc -ArgumentList $CertExpiryWarnDays -ScriptBlock {
                param($WarnDays)
                Get-ChildItem Cert:\LocalMachine\My -ErrorAction Stop |
                    Where-Object { $_.NotAfter -lt (Get-Date).AddDays($WarnDays) } |
                    Select-Object Subject, NotAfter, Thumbprint
            } -ErrorAction Stop
            if ($expiring) {
                Add-Result 'CERT' "Expiring Certs $dc" 'WARN' "$(@($expiring).Count) cert(s) expire within $CertExpiryWarnDays days" `
                    (($expiring | ForEach-Object { "$($_.Subject) -> $($_.NotAfter.ToString('yyyy-MM-dd'))" }) -join '; ')
            } else {
                Add-Result 'CERT' "Expiring Certs $dc" 'PASS' "No machine certs expire within $CertExpiryWarnDays days"
            }
        } catch { Add-Result 'CERT' "Expiring Certs $dc" 'INFO' "Could not read cert store on $dc (remoting?): $_" }
    }

    # Enterprise CA presence + liveness
    # Detects registered CAs that no longer answer while autoenrollment GPOs still point at them.
    $registeredCAs = @()
    try {
        $configNC = (Get-ADRootDSE -Server $Domain).configurationNamingContext
        $registeredCAs = Get-ADObject `
            -SearchBase "CN=Enrollment Services,CN=Public Key Services,CN=Services,$configNC" `
            -Server $Domain -Filter { objectClass -eq 'pKIEnrollmentService' } `
            -Properties dNSHostName -ErrorAction SilentlyContinue
        if ($registeredCAs) {
            Add-Result 'CERT' 'Enterprise CA Registered' 'INFO' `
                "$(@($registeredCAs).Count) CA(s) in AD: $(($registeredCAs | ForEach-Object { "$($_.Name)=$($_.dNSHostName)" }) -join ', ')"
        } else {
            Add-Result 'CERT' 'Enterprise CA Registered' 'INFO' 'No enterprise CA published in AD'
        }
    } catch { Add-Result 'CERT' 'Enterprise CA Registered' 'INFO' "Could not query Enrollment Services: $_" }

    $deadCAs = [System.Collections.Generic.List[string]]::new()
    foreach ($ca in $registeredCAs) {
        $caHost = $ca.dNSHostName
        $caName = $ca.Name
        if (-not $caHost) { continue }
        $tcp135 = Test-TcpPort -HostName $caHost -Port 135 -TimeoutMs 3000
        $pingOk = $false
        try {
            $pingOut = certutil -ping -config "$caHost\$caName" 2>&1 | Out-String
            # Note: match on exit code is more locale-robust than string matching
            $pingOk  = ($pingOut -match 'interface is alive' -or
                        $pingOut -match 'CertUtil: -ping command completed successfully' -or
                        $LASTEXITCODE -eq 0)
        } catch { $pingOk = $false }
        if ($pingOk) {
            Add-Result 'CERT' "CA Liveness $caName" 'PASS' "$caHost answers certutil -ping"
        } else {
            $deadCAs.Add("$caName ($caHost)")
            $reason = if (-not $tcp135) { 'no RPC/135 and -ping failed' } else { 'RPC reachable but -ping did not confirm CA' }
            Add-Result 'CERT' "CA Liveness $caName" 'FAIL' "$caHost does NOT answer as a CA — $reason"
        }
    }

    # Autoenrollment GPOs cross-referenced against dead CAs + unlinked GPO count
    $aeGpos       = [System.Collections.Generic.List[string]]::new()
    $unlinkedGpos = 0
    try {
        Import-Module GroupPolicy -ErrorAction Stop
        $allGpos = Get-GPO -All -Domain $Domain -ErrorAction Stop
        foreach ($g in $allGpos) {
            try {
                [xml]$rep = Get-GPOReport -Guid $g.Id -ReportType Xml -Domain $Domain -ErrorAction Stop
                $links    = @($rep.GPO.LinksTo)
                if ($links.Count -eq 0) { $unlinkedGpos++ }
                $xml = $rep.OuterXml
                if ($xml -match 'AutoEnrollment|Auto-Enrollment|AutoEnrollmentPolicy') {
                    $linkScope = ($links | ForEach-Object { $_.SOMPath }) -join ', '
                    $aeGpos.Add("$($g.DisplayName) [links: $(if($linkScope){$linkScope}else{'(unlinked)'})]")
                }
            } catch { }
        }
        if ($aeGpos.Count -gt 0) {
            if ($deadCAs.Count -gt 0) {
                Add-Result 'CERT' 'Autoenrollment vs Dead CA' 'FAIL' `
                    "$($aeGpos.Count) autoenrollment GPO(s) active while $($deadCAs.Count) CA(s) are dead — clients enrolling against a void" `
                    ("GPOs: " + ($aeGpos -join ' | ') + "  ||  Dead CAs: " + ($deadCAs -join ', '))
            } else {
                Add-Result 'CERT' 'Autoenrollment GPOs' 'INFO' "$($aeGpos.Count) autoenrollment GPO(s) present (CAs answer)" ($aeGpos -join ' | ')
            }
        } else {
            Add-Result 'CERT' 'Autoenrollment GPOs' 'PASS' 'No autoenrollment GPOs found'
        }
        if ($unlinkedGpos -gt 0) {
            Add-Result 'CERT' 'Unlinked GPOs' 'WARN' "$unlinkedGpos GPO(s) are not linked to any OU/site/domain — consider cleanup"
        } else {
            Add-Result 'CERT' 'Unlinked GPOs' 'PASS' 'All GPOs are linked'
        }
    } catch {
        Add-Result 'CERT' 'Autoenrollment GPOs' 'INFO' "Could not enumerate GPOs (GroupPolicy module?): $_"
    }

    try {
        $ntauth  = certutil -viewstore -enterprise NTAuth 2>&1 | Out-String
        $ntCount = ([regex]::Matches($ntauth, '(?im)^\s*Serial Number:')).Count
        if ($ntCount -eq 0) {
            Add-Result 'CERT' 'NTAuth Store' 'INFO' 'NTAuth store is empty (no enterprise-trusted CA issuer cert)'
        } else {
            Add-Result 'CERT' 'NTAuth Store' 'INFO' "$ntCount cert(s) in NTAuth — verify none belong to a decommissioned CA"
        }
    } catch { Add-Result 'CERT' 'NTAuth Store' 'INFO' "Could not read NTAuth: $_" }
}

# ======================================================================
#  SECTION: USERS  (privileged risk + hygiene + policy)
# ======================================================================
if (Should-Run 'USERS') {
    $auditGroups    = @('Domain Admins','Enterprise Admins','Schema Admins','Administrators','Backup Operators','Account Operators')
    $highPrivGroups = @('Domain Admins','Enterprise Admins','Schema Admins')
    $svcRegex       = $Config.ServiceAcctRegex
    $adminAcctRegex = $Config.AdminAcctRegex
    $knownRisk      = @($Config.KnownRiskAccounts)

    # Build privileged account membership map (reused by SECURITY section)
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
        } catch { Write-Log "Could not query group '$g': $_" -Level WARN }
    }

    $high = 0; $review = 0; $ok = 0
    $privDetail = [System.Collections.Generic.List[string]]::new()
    foreach ($sam in ($membershipMap.Keys | Sort-Object)) {
        try {
            $u      = Get-ADUser -Identity $sam -Server $Domain -Properties Enabled,LastLogonDate,PasswordNeverExpires,Description,DisplayName -ErrorAction Stop
            $groups = $membershipMap[$sam]
            $inHigh = ($groups | Where-Object { $_ -in $highPrivGroups }).Count -gt 0
            $days   = if ($u.LastLogonDate) { (New-TimeSpan -Start $u.LastLogonDate -End (Get-Date)).Days } else { $null }

            $risk = 'OK'; $reason = 'Active account, no risk flags'
            if (-not $u.Enabled -and $inHigh)                              { $risk='HIGH-RISK'; $reason='Disabled account in high-priv group' }
            elseif ($u.SamAccountName -in $knownRisk)                       { $risk='HIGH-RISK'; $reason='Known vendor/external account' }
            elseif ($null -eq $u.LastLogonDate -and $inHigh)                { $risk='HIGH-RISK'; $reason='Never logged in and in high-priv group' }
            elseif ($u.SamAccountName -match $svcRegex -and $inHigh)        { $risk='HIGH-RISK'; $reason='Service-account name pattern in high-priv group' }
            elseif ($null -ne $days -and $days -gt $InactiveDays -and $inHigh) { $risk='REVIEW'; $reason="Last logon $days days ago in high-priv group" }
            elseif ($u.SamAccountName -match $adminAcctRegex -and $inHigh)  { $risk='REVIEW'; $reason='Named admin account in high-priv group — verify still needed' }

            switch ($risk) { 'HIGH-RISK'{$high++} 'REVIEW'{$review++} default{$ok++} }
            if ($risk -ne 'OK') { $privDetail.Add("$sam [$risk]: $reason ($($groups -join '/'))") }
        } catch { Write-Log "Could not read user '$sam': $_" -Level WARN }
    }
    $privStatus = if ($high -gt 0) {'FAIL'} elseif ($review -gt 0) {'WARN'} else {'PASS'}
    Add-Result 'USERS' 'Privileged Accounts' $privStatus `
        "HIGH-RISK=$high REVIEW=$review OK=$ok (total $($membershipMap.Count))" `
        ($privDetail -join ' | ')

    # krbtgt password age
    try {
        $krbtgt = Get-ADUser krbtgt -Server $Domain -Properties PasswordLastSet -ErrorAction Stop
        $age    = (New-TimeSpan -Start $krbtgt.PasswordLastSet -End (Get-Date)).Days
        if ($age -gt 180) {
            Add-Result 'USERS' 'krbtgt Password Age' 'WARN' "$age days — consider a controlled double-reset"
        } else {
            Add-Result 'USERS' 'krbtgt Password Age' 'PASS' "$age days"
        }
    } catch { Add-Result 'USERS' 'krbtgt Password Age' 'INFO' "Could not read krbtgt: $_" }

    # Default domain password policy
    try {
        $pwdPol    = Get-ADDefaultDomainPasswordPolicy -Server $Domain -ErrorAction Stop
        $polIssues = [System.Collections.Generic.List[string]]::new()
        if ($pwdPol.LockoutThreshold -eq 0)          { $polIssues.Add('No account lockout threshold') }
        elseif ($pwdPol.LockoutThreshold -gt 10)     { $polIssues.Add("Lockout threshold=$($pwdPol.LockoutThreshold) — consider ≤10") }
        if ($pwdPol.MinPasswordLength -lt 12)        { $polIssues.Add("Min length=$($pwdPol.MinPasswordLength) — consider ≥12") }
        if (-not $pwdPol.ComplexityEnabled)           { $polIssues.Add('Complexity not required') }
        if ($pwdPol.MaxPasswordAge.Days -eq 0)       { $polIssues.Add('Max age=0 (passwords never expire via DDP)') }
        elseif ($pwdPol.MaxPasswordAge.Days -gt 365) { $polIssues.Add("Max age=$($pwdPol.MaxPasswordAge.Days)d (consider ≤365)") }
        $polInfo = "MinLen=$($pwdPol.MinPasswordLength) MaxAge=$($pwdPol.MaxPasswordAge.Days)d Complexity=$($pwdPol.ComplexityEnabled) Lockout=$($pwdPol.LockoutThreshold)"
        if ($polIssues.Count -gt 0) {
            Add-Result 'USERS' 'Password Policy' 'WARN' ($polIssues -join '; ') $polInfo
        } else {
            Add-Result 'USERS' 'Password Policy' 'PASS' 'Policy meets baseline' $polInfo
        }
    } catch { Add-Result 'USERS' 'Password Policy' 'INFO' "Could not read password policy: $_" }

    # Enabled users with PasswordNeverExpires
    try {
        $pne = @(Get-ADUser -Filter { Enabled -eq $true -and PasswordNeverExpires -eq $true } -Server $Domain)
        if ($pne.Count -gt 0) {
            Add-Result 'USERS' 'Password Never Expires' 'WARN' "$($pne.Count) enabled user(s) with non-expiring passwords"
        } else {
            Add-Result 'USERS' 'Password Never Expires' 'PASS' 'No enabled users with non-expiring passwords'
        }
    } catch { Add-Result 'USERS' 'Password Never Expires' 'INFO' "Query failed: $_" }

    # Stale enabled users
    try {
        $cut   = (Get-Date).AddDays(-1 * $InactiveDays)
        $stale = @(Get-ADUser -Filter { Enabled -eq $true } -Server $Domain -Properties LastLogonDate |
                   Where-Object { $_.LastLogonDate -and $_.LastLogonDate -lt $cut })
        if ($stale.Count -gt 0) {
            Add-Result 'USERS' 'Stale Enabled Users' 'WARN' "$($stale.Count) enabled user(s) inactive > $InactiveDays days"
        } else {
            Add-Result 'USERS' 'Stale Enabled Users' 'PASS' "No enabled users inactive > $InactiveDays days"
        }
    } catch { Add-Result 'USERS' 'Stale Enabled Users' 'INFO' "Query failed: $_" }

    # Stale computer accounts
    try {
        $cut    = (Get-Date).AddDays(-1 * $InactiveDays)
        $staleC = @(Get-ADComputer -Filter { Enabled -eq $true } -Server $Domain -Properties LastLogonDate |
                    Where-Object { $_.LastLogonDate -and $_.LastLogonDate -lt $cut })
        if ($staleC.Count -gt 0) {
            Add-Result 'USERS' 'Stale Computer Accounts' 'WARN' "$($staleC.Count) computer(s) inactive > $InactiveDays days"
        } else {
            Add-Result 'USERS' 'Stale Computer Accounts' 'PASS' "No computers inactive > $InactiveDays days"
        }
    } catch { Add-Result 'USERS' 'Stale Computer Accounts' 'INFO' "Query failed: $_" }

    # AdminSDHolder orphans — accounts with AdminCount=1 not in any current protected group.
    # These inherit a hardened ACL from AdminSDHolder even after removal from the group,
    # blocking normal ACL inheritance and causing subtle permission issues.
    try {
        $adminCountUsers = Get-ADUser -Filter { AdminCount -eq 1 -and Enabled -eq $true } `
                           -Server $Domain -Properties AdminCount -ErrorAction Stop
        $orphans = @($adminCountUsers | Where-Object { -not $membershipMap.ContainsKey($_.SamAccountName) })
        if ($orphans.Count -gt 0) {
            Add-Result 'USERS' 'AdminSDHolder Orphans' 'WARN' `
                "$($orphans.Count) enabled user(s) with AdminCount=1 but not in any protected group — ACL likely locked unnecessarily" `
                (($orphans | ForEach-Object { $_.SamAccountName }) -join ', ')
        } else {
            Add-Result 'USERS' 'AdminSDHolder Orphans' 'PASS' 'No AdminCount=1 orphans found'
        }
    } catch { Add-Result 'USERS' 'AdminSDHolder Orphans' 'INFO' "Could not check AdminSDHolder orphans: $_" }

    # AS-REP roastable — Kerberos pre-auth disabled, hash obtainable without credentials
    try {
        $asrep = @(Get-ADUser -Filter { DoesNotRequirePreAuth -eq $true -and Enabled -eq $true } `
                  -Server $Domain -Properties DoesNotRequirePreAuth -ErrorAction Stop)
        if ($asrep.Count -gt 0) {
            Add-Result 'USERS' 'AS-REP Roastable' 'WARN' `
                "$($asrep.Count) enabled account(s) with Kerberos pre-auth disabled — offline hash crackable" `
                (($asrep | ForEach-Object { $_.SamAccountName }) -join ', ')
        } else {
            Add-Result 'USERS' 'AS-REP Roastable' 'PASS' 'No accounts with pre-auth disabled'
        }
    } catch { Add-Result 'USERS' 'AS-REP Roastable' 'INFO' "Could not check pre-auth setting: $_" }

    # Protected Users group membership
    try {
        $pu        = @(Get-ADGroupMember -Identity 'Protected Users' -Recursive -Server $Domain -ErrorAction Stop |
                       Where-Object objectClass -eq 'user')
        $privTotal = $membershipMap.Count
        Add-Result 'USERS' 'Protected Users Group' 'INFO' `
            "$($pu.Count) user(s) in Protected Users$(if($privTotal -gt 0){" (vs $privTotal accounts in priv groups)"})"
    } catch { Add-Result 'USERS' 'Protected Users Group' 'INFO' "Could not read Protected Users: $_" }
}

# ======================================================================
#  SECTION: SECURITY  (delegation, SPNs, LAPS, Kerberoasting)
# ======================================================================
if (Should-Run 'SECURITY') {
    # Unconstrained Kerberos delegation — machines/users that can impersonate any authenticating user.
    # DCs carry this by design (PrimaryGroupID 516/521); everything else is a risk.
    try {
        $udComputers = @(Get-ADComputer -Filter { TrustedForDelegation -eq $true -and Enabled -eq $true } `
                        -Server $Domain -Properties TrustedForDelegation,PrimaryGroupID -ErrorAction Stop |
                        Where-Object { $_.PrimaryGroupID -notin @(516,521) })
        $udUsers     = @(Get-ADUser -Filter { TrustedForDelegation -eq $true -and Enabled -eq $true } `
                        -Server $Domain -Properties TrustedForDelegation -ErrorAction Stop)
        $udTotal = $udComputers.Count + $udUsers.Count
        if ($udTotal -gt 0) {
            $detail = [System.Collections.Generic.List[string]]::new()
            if ($udComputers.Count -gt 0) { $detail.Add("Computers: " + ($udComputers.Name -join ', ')) }
            if ($udUsers.Count    -gt 0)  { $detail.Add("Users: "     + ($udUsers.SamAccountName -join ', ')) }
            Add-Result 'SECURITY' 'Unconstrained Delegation' 'WARN' `
                "$udTotal non-DC object(s) with unconstrained Kerberos delegation — high-value attacker pivot" `
                ($detail -join ' | ')
        } else {
            Add-Result 'SECURITY' 'Unconstrained Delegation' 'PASS' 'No non-DC objects with unconstrained delegation'
        }
    } catch { Add-Result 'SECURITY' 'Unconstrained Delegation' 'INFO' "Could not check delegation: $_" }

    # Duplicate SPNs — root cause of KDC_ERR_S_PRINCIPAL_UNKNOWN Kerberos failures
    try {
        $setspnOut = setspn -X -F 2>&1 | Out-String
        if ($setspnOut -match 'duplicate') {
            $dupLines = ($setspnOut -split "`n") | Where-Object { $_ -match 'duplicate|Checking' } | Select-Object -First 10
            Add-Result 'SECURITY' 'Duplicate SPNs' 'WARN' 'Duplicate SPNs detected — causes Kerberos authentication failures' ($dupLines -join ' | ')
        } else {
            Add-Result 'SECURITY' 'Duplicate SPNs' 'PASS' 'No duplicate SPNs (setspn -X -F)'
        }
    } catch { Add-Result 'SECURITY' 'Duplicate SPNs' 'INFO' "Could not run setspn: $_" }

    # LAPS deployment — check schema for legacy (ms-Mcs-AdmPwd) or Windows LAPS (msLAPS-Password)
    try {
        $schemaNc   = (Get-ADRootDSE -Server $Domain).schemaNamingContext
        $legacyLaps = Get-ADObject -SearchBase $schemaNc -Filter { Name -eq 'ms-Mcs-AdmPwd' }     -Server $Domain -ErrorAction SilentlyContinue
        $winLaps    = Get-ADObject -SearchBase $schemaNc -Filter { Name -eq 'msLAPS-Password' }    -Server $Domain -ErrorAction SilentlyContinue
        if ($legacyLaps -or $winLaps) {
            $lapsType = @()
            if ($legacyLaps) { $lapsType += 'Legacy LAPS' }
            if ($winLaps)    { $lapsType += 'Windows LAPS' }
            $lapsAttr = if ($winLaps) { 'msLAPS-PasswordExpirationTime' } else { 'ms-Mcs-AdmPwdExpirationTime' }
            try {
                $allComp   = @(Get-ADComputer -Filter { Enabled -eq $true } -Server $Domain -Properties $lapsAttr -ErrorAction Stop)
                $lapsComp  = @($allComp | Where-Object { $_.$lapsAttr })
                $pct       = if ($allComp.Count -gt 0) { [int](($lapsComp.Count / $allComp.Count) * 100) } else { 0 }
                $covStatus = if ($pct -ge 80) { 'PASS' } else { 'WARN' }
                Add-Result 'SECURITY' 'LAPS Deployment' $covStatus `
                    "$($lapsType -join ' + ') schema present; $($lapsComp.Count)/$($allComp.Count) computers ($pct%) have a LAPS password set"
            } catch {
                Add-Result 'SECURITY' 'LAPS Deployment' 'INFO' "$($lapsType -join ' + ') schema present; coverage check failed: $_"
            }
        } else {
            Add-Result 'SECURITY' 'LAPS Deployment' 'WARN' 'LAPS schema NOT found — local admin passwords likely unmanaged'
        }
    } catch { Add-Result 'SECURITY' 'LAPS Deployment' 'INFO' "Could not check LAPS schema: $_" }

    # Kerberoastable accounts — enabled users with SPNs; TGS ticket obtainable and hash crackable offline
    try {
        $kerb = @(Get-ADUser -Filter { ServicePrincipalName -like '*' -and Enabled -eq $true } `
                 -Server $Domain -Properties ServicePrincipalName,PasswordLastSet -ErrorAction Stop |
                 Where-Object { $_.SamAccountName -ne 'krbtgt' })
        if ($kerb.Count -gt 0) {
            $inHigh  = @($kerb | Where-Object { $membershipMap.ContainsKey($_.SamAccountName) })
            $status  = if ($inHigh.Count -gt 0) { 'FAIL' } else { 'WARN' }
            $detail  = "$($kerb.Count) enabled user(s) with SPN$(if($inHigh.Count -gt 0){" — $($inHigh.Count) also in privileged groups (critical)"})"
            Add-Result 'SECURITY' 'Kerberoastable Accounts' $status $detail `
                (($kerb | ForEach-Object {
                    $age = if ($_.PasswordLastSet) { ([datetime]::Now - $_.PasswordLastSet).Days } else { '?' }
                    "$($_.SamAccountName) [pwdAge=${age}d]"
                }) -join ', ')
        } else {
            Add-Result 'SECURITY' 'Kerberoastable Accounts' 'PASS' 'No enabled user accounts with SPNs (beyond krbtgt)'
        }
    } catch { Add-Result 'SECURITY' 'Kerberoastable Accounts' 'INFO' "Could not check Kerberoastable accounts: $_" }
}

# ======================================================================
#  SECTION: NET  (DC port reachability + unsigned LDAP binds)
# ======================================================================
if (Should-Run 'NET') {
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

    foreach ($ip in @($Config.StaleIPs)) {
        if (-not $ip) { continue }
        if (Test-TcpPort -HostName $ip -Port 389) {
            Add-Result 'NET' "Stale LDAP $ip" 'WARN' "$ip still answers on 389 — expected decommissioned"
        } else {
            Add-Result 'NET' "Stale LDAP $ip" 'PASS' "$ip does not answer on 389"
        }
    }

    # Unsigned/simple LDAP binds — requires NTDS "LDAP Interface Events" auditing >= 2
    $start = (Get-Date).AddHours(-1 * $EventHoursBack)
    foreach ($dc in $DomainControllers) {
        try {
            $filter = @{ LogName='Directory Service'; Id=2889; StartTime=$start }
            $ev     = Invoke-Command -ComputerName $dc -ArgumentList $filter -ScriptBlock {
                param($f) Get-WinEvent -FilterHashtable $f -ErrorAction Stop | Select-Object Properties
            } -ErrorAction Stop
            $clients = $ev | ForEach-Object { ([string]$_.Properties[0].Value -split ':')[0] } | Sort-Object -Unique
            if ($clients) {
                Add-Result 'NET' "Unsigned LDAP $dc" 'WARN' "$(@($clients).Count) client(s) doing unsigned/simple binds (last ${EventHoursBack}h)" ($clients -join ', ')
            } else {
                Add-Result 'NET' "Unsigned LDAP $dc" 'PASS' "No Event 2889 in last ${EventHoursBack}h"
            }
        } catch {
            if ("$_" -match 'No events were found') {
                Add-Result 'NET' "Unsigned LDAP $dc" 'PASS' "No Event 2889 in last ${EventHoursBack}h"
            } else {
                Add-Result 'NET' "Unsigned LDAP $dc" 'INFO' "Could not read 2889 on $dc (enable NTDS LDAP Interface Events>=2, or remoting blocked): $_"
            }
        }
    }
}

# ======================================================================
#  SECTION: LICENSE  (activation state; best-effort, remote)
# ======================================================================
if (Should-Run 'LICENSE') {
    foreach ($dc in $DomainControllers) {
        try {
            $lic = Invoke-Command -ComputerName $dc -ScriptBlock {
                Get-CimInstance SoftwareLicensingProduct -ErrorAction Stop |
                    Where-Object { $_.PartialProductKey -and $_.Name -like 'Windows*' } |
                    Select-Object Name, LicenseStatus, @{n='KMS';e={$_.KeyManagementServiceMachine}}
            } -ErrorAction Stop | Select-Object -First 1
            if ($lic) {
                $statusText = switch ([int]$lic.LicenseStatus) {
                    1 {'Licensed'} 0 {'Unlicensed'} 2 {'OOB Grace'} 3 {'OOT Grace'} 5 {'Notification'} default {"State $($lic.LicenseStatus)"}
                }
                $st = if ([int]$lic.LicenseStatus -eq 1) { 'PASS' } else { 'WARN' }
                Add-Result 'LICENSE' "Activation $dc" $st "$statusText$(if($lic.KMS){" (KMS: $($lic.KMS))"})"
            } else {
                Add-Result 'LICENSE' "Activation $dc" 'INFO' 'No Windows license product returned'
            }
        } catch { Add-Result 'LICENSE' "Activation $dc" 'INFO' "Could not query activation on $dc (remoting?): $_" }
    }
}

# ======================================================================
#  SECTION: EVENTS  (critical DS / DFSR errors)
# ======================================================================
if (Should-Run 'EVENTS') {
    $start = (Get-Date).AddHours(-1 * $EventHoursBack)
    foreach ($dc in $DomainControllers) {
        try {
            $f = @{ LogName='Directory Service'; Level=@(1,2); StartTime=$start }
            $n = @(Get-WinEvent -ComputerName $dc -FilterHashtable $f -ErrorAction SilentlyContinue).Count
            if ($n -eq 0) { Add-Result 'EVENTS' "DS Errors $dc" 'PASS' "0 error/critical DS events (${EventHoursBack}h)" }
            else          { Add-Result 'EVENTS' "DS Errors $dc" 'FAIL' "$n error/critical DS event(s) (${EventHoursBack}h)" }
        } catch { Add-Result 'EVENTS' "DS Errors $dc" 'INFO' "Query failed on $dc: $_" }

        try {
            $f = @{ LogName='DFS Replication'; Level=@(1,2); StartTime=$start }
            $n = @(Get-WinEvent -ComputerName $dc -FilterHashtable $f -ErrorAction SilentlyContinue).Count
            if ($n -eq 0) { Add-Result 'EVENTS' "DFSR Errors $dc" 'PASS' "0 DFSR error events (${EventHoursBack}h)" }
            else          { Add-Result 'EVENTS' "DFSR Errors $dc" 'WARN' "$n DFSR error event(s) (${EventHoursBack}h)" }
        } catch { Add-Result 'EVENTS' "DFSR Errors $dc" 'INFO' "DFSR log query failed on $dc: $_" }
    }
}

# ======================================================================
#  SUMMARY
# ======================================================================
Write-Host ""
Write-Host ("-" * 115) -ForegroundColor DarkGray
$pass    = @($results | Where-Object Status -eq 'PASS').Count
$warn    = @($results | Where-Object Status -eq 'WARN').Count
$fail    = @($results | Where-Object Status -eq 'FAIL').Count
$info    = @($results | Where-Object Status -eq 'INFO').Count
$overall = if ($fail -gt 0) {'FAIL'} elseif ($warn -gt 0) {'WARN'} else {'PASS'}
Write-Host "Overall: $overall   PASS=$pass  WARN=$warn  FAIL=$fail  INFO=$info" -ForegroundColor Cyan
Write-Log "Overall=$overall PASS=$pass WARN=$warn FAIL=$fail INFO=$info"

# ======================================================================
#  CSV
# ======================================================================
if ($Csv) {
    try {
        if (-not (Test-Path $LogRoot)) { New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null }
        $results | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
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
        $cls   = switch ($r.Status) { 'PASS'{'status-pass'};'WARN'{'status-warn'};'FAIL'{'status-fail'};'INFO'{'status-info'} }
        $extra = if ($r.Extra) { "<br><small style='color:#555;font-family:monospace'>$(Escape-Html $r.Extra)</small>" } else { '' }
        [void]$rows.Append("<tr><td style='text-align:center'>$($r.Num)</td><td>$(Escape-Html $r.Section)</td><td>$(Escape-Html $r.Name)</td><td><span class='badge $cls'>$($r.Status)</span></td><td>$(Escape-Html $r.Detail)$extra</td></tr>`n")
    }
    $overallCls = switch ($overall) { 'PASS'{'status-pass'};'WARN'{'status-warn'};'FAIL'{'status-fail'} }
    $runDate    = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $html = @"
<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<title>Domain Health Audit - $Domain</title>
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
 <tr><td style="border:none;font-weight:bold">Domain</td><td style="border:none">$(Escape-Html $Domain)</td></tr>
 <tr><td style="border:none;font-weight:bold">Report Date</td><td style="border:none">$runDate</td></tr>
 <tr><td style="border:none;font-weight:bold">Overall</td><td style="border:none"><span class='badge $overallCls'>$overall</span></td></tr>
 <tr><td style="border:none;font-weight:bold">Summary</td><td style="border:none">PASS=$pass WARN=$warn FAIL=$fail INFO=$info</td></tr>
</table>
<table>
 <tr><th style="width:40px">#</th><th style="width:80px">Section</th><th style="width:220px">Check</th><th style="width:70px">Status</th><th>Detail</th></tr>
 $($rows.ToString())
</table>
<p><small>Generated by Invoke-DomainHealthAudit.ps1 — read-only. Iver / IT Consult.</small></p>
</body></html>
"@
    try {
        if (-not (Test-Path $LogRoot)) { New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null }
        $html | Set-Content -Path $HtmlPath -Encoding UTF8
        Write-Host "HTML: $HtmlPath" -ForegroundColor Green
        Write-Log "HTML written: $HtmlPath" -Level OK
    } catch { Write-Log "HTML write failed: $_" -Level ERROR }
}

Write-Log "Audit completed"
return [PSCustomObject]@{
    Domain   = $Domain
    Overall  = $overall
    Pass     = $pass
    Warn     = $warn
    Fail     = $fail
    Info     = $info
    Results  = $results
    HtmlPath = if ($Html) { $HtmlPath } else { $null }
    CsvPath  = if ($Csv)  { $CsvPath }  else { $null }
}
