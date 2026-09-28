#Requires -Version 5.1

<#
.SYNOPSIS
    Read-only estate sweep: finds every machine holding a certificate issued by a
    specific (typically decommissioned) CA, with expiry and EKU, to size the blast
    radius before retiring or rebuilding that CA.

.DESCRIPTION
    NON-DESTRUCTIVE. Written for the "dead CA + domain-wide autoenrollment GPO"
    situation: an enterprise CA is still registered in AD and pushed as a trusted
    root by autoenrollment GPOs linked at the domain root, but the CA no longer
    answers. GPO *scope* is the whole domain; this script finds actual *dependency* -
    which machines really hold certs from that CA and which of those can be used for
    client authentication (the 802.1x / EAP-TLS marker).

    For each reachable Windows computer it reports, per matching cert in
    Cert:\LocalMachine\My:
      Subject, Issuer, NotAfter, days-to-expiry, EKUs, HasClientAuth, Thumbprint.

    HasClientAuth is true when the cert carries the Client Authentication EKU
    (1.3.6.1.5.5.7.3.2), Any Purpose (2.5.29.37.0), or no EKU extension at all
    (which means valid for all purposes).

    Machine list comes from -ComputerName, or a text/inventory file, or (default)
    all enabled Windows computers in AD. Remote hosts are queried with a single
    fan-out Invoke-Command (-ThrottleLimit), on both Windows PowerShell 5.1 and 7.

.PARAMETER IssuerMatch
    Required. Case-insensitive SUBSTRING (not a regex) matched against each cert's
    Issuer DN, e.g. 'CN=Contoso-Old-CA' or 'CN=CA, DC=contoso, DC=local'.

.PARAMETER ComputerName
    Explicit list of computers to sweep. Overrides -InventoryPath and AD discovery.

.PARAMETER InventoryPath
    Path to a file containing one hostname per line (lines starting with # ignored),
    or the project 01-INVENTORY.md (FQDNs are extracted heuristically).

.PARAMETER OnlyWindows
    When discovering from AD, keep only computers whose OperatingSystem is Windows. Default on.

.PARAMETER ThrottleLimit
    Max concurrent remote connections. Default 16.

.PARAMETER LogRoot
    Output folder for the CSV + log. Default C:\Logs\DomainAudit.

.EXAMPLE
    .\Get-CertBlastRadius.ps1 -IssuerMatch 'CN=Contoso-Old-CA' -InventoryPath .\01-INVENTORY.md
    Sweep every host named in the inventory for certs issued by Contoso-Old-CA.

.EXAMPLE
    .\Get-CertBlastRadius.ps1 -IssuerMatch 'Contoso-Old-CA' -ComputerName srv-app01,srv-db01,srv-rds01

.NOTES
    Version   : 1.1.0
    Read-only : queries Cert:\LocalMachine\My remotely; changes nothing.
    Pairs with: Invoke-DomainHealthAudit.ps1 (CERT section flags the dead CA; this sizes it)

    Changelog (full history in CHANGELOG.md):
      1.1.0  2026-09-28  Fixes from review + test run in an isolated Server 2022 lab:
                         PS5.1 parse failure (encoding), PS7 parse error in the
                         -Parallel block (replaced by Invoke-Command fan-out that works on
                         5.1 too), ActiveDirectory module only needed for AD discovery,
                         certs without EKU now count as client-auth capable, EKUs read
                         from the raw extension, inventory parser no longer picks up
                         OIDs / version numbers, the CA's own self-signed cert is not
                         counted as a dependent, SAN shown when Subject is empty,
                         version stamping. -IssuerMatch is now mandatory (no
                         environment-specific default).
      1.0.0  2026-09     Initial version.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$IssuerMatch,
    [string[]]$ComputerName,
    [string]$InventoryPath = '',
    [bool]$OnlyWindows = $true,
    [ValidateRange(1,256)]
    [int]$ThrottleLimit = 16,
    [string]$LogRoot = 'C:\Logs\DomainAudit'
)

$ScriptVersion = '1.1.0'
$Stamp   = Get-Date -Format 'yyyyMMdd_HHmm'
$LogPath = Join-Path $LogRoot ("CertBlastRadius_{0}.log" -f $Stamp)
$CsvPath = Join-Path $LogRoot ("CertBlastRadius_{0}.csv" -f $Stamp)

function Write-Log {
    param([string]$Message, [ValidateSet('INFO','WARN','ERROR','OK')][string]$Level='INFO')
    $entry = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Verbose $entry
    try {
        if (-not (Test-Path $LogRoot)) { New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null }
        Add-Content -Path $LogPath -Value $entry -ErrorAction Stop
    } catch { Write-Warning "Log write failed: $_" }
}

# --- Build target list ---
$targets = @()
if ($ComputerName) {
    $targets = $ComputerName
} elseif ($InventoryPath) {
    if (-not (Test-Path $InventoryPath)) { throw "Inventory file not found: $InventoryPath" }
    $raw = Get-Content -Path $InventoryPath
    if ($InventoryPath -match '\.md$') {
        # Heuristic: FQDN-looking tokens whose last label is alphabetic (skips OIDs,
        # IPs and version numbers such as 1.3.6.1.5.5.7.3.2) and not a file name.
        $targets = ($raw | Select-String -Pattern '\b[a-z0-9][a-z0-9\-]*(\.[a-z0-9\-]+)*\.[a-z]{2,}\b' -AllMatches |
                    ForEach-Object { $_.Matches.Value }) |
                    Where-Object { $_ -notmatch '\.(md|ps1|psm1|psd1|csv|html|txt|json|log|zip|exe|msi)$' } |
                    ForEach-Object { $_.ToLower() } |
                    Sort-Object -Unique
    } else {
        $targets = $raw | Where-Object { $_ -and $_ -notmatch '^\s*#' } | ForEach-Object { $_.Trim() }
    }
} else {
    Write-Log "No -ComputerName/-InventoryPath - discovering from AD"
    Import-Module ActiveDirectory -ErrorAction Stop
    $comp = @(Get-ADComputer -Filter "Enabled -eq 'True'" -Properties OperatingSystem,DNSHostName)
    if ($OnlyWindows) { $comp = @($comp | Where-Object { $_.OperatingSystem -match 'Windows' }) }
    $targets = $comp | ForEach-Object { if ($_.DNSHostName) { $_.DNSHostName } else { $_.Name } }
}
$targets = @($targets | Where-Object { $_ } | Sort-Object -Unique)

Write-Host ""
Write-Host "=== CERT BLAST-RADIUS SWEEP v$ScriptVersion ===" -ForegroundColor Cyan
Write-Host "  Issuer match : $IssuerMatch"
Write-Host "  Targets      : $($targets.Count) host(s)"
Write-Host ""
Write-Log "Sweep v$ScriptVersion start. IssuerMatch='$IssuerMatch' Targets=$($targets.Count)"

if ($targets.Count -eq 0) {
    Write-Warning 'No targets to sweep.'
    return
}

# --- The remote probe: returns ONE object per host, so hosts with no match still prove reachability ---
$probe = {
    param($Match)
    $certs = @(Get-ChildItem Cert:\LocalMachine\My -ErrorAction Stop |
        Where-Object { $_.Issuer.IndexOf($Match, [StringComparison]::OrdinalIgnoreCase) -ge 0 } |
        Where-Object { $_.Subject -ne $_.Issuer } |      # skip the CA's own self-signed cert - not a dependent
        ForEach-Object {
            # Read EKUs from the raw extension (EnhancedKeyUsageList is not reliable on PS 5.1).
            $ekuExt = $_.Extensions | Where-Object { $_.Oid.Value -eq '2.5.29.37' } | Select-Object -First 1
            $ekus = @()
            if ($ekuExt) {
                $typed = New-Object System.Security.Cryptography.X509Certificates.X509EnhancedKeyUsageExtension($ekuExt, $ekuExt.Critical)
                $ekus  = @($typed.EnhancedKeyUsages | ForEach-Object { $_ })
            }
            $ekuOids = @($ekus | ForEach-Object { $_.Value })
            [PSCustomObject]@{
                Subject       = if ($_.Subject) { $_.Subject } else { 'SAN:' + $_.GetNameInfo([System.Security.Cryptography.X509Certificates.X509NameType]::DnsName, $false) }
                Issuer        = $_.Issuer
                NotAfter      = $_.NotAfter
                DaysToExpiry  = [int][math]::Floor(($_.NotAfter - (Get-Date)).TotalDays)
                EKU           = if ($ekuExt) { ($ekus | ForEach-Object { if ($_.FriendlyName) { $_.FriendlyName } else { $_.Value } }) -join ', ' } else { '(none - all purposes)' }
                HasClientAuth = (-not $ekuExt) -or ($ekuOids -contains '1.3.6.1.5.5.7.3.2') -or ($ekuOids -contains '2.5.29.37.0')
                Thumbprint    = $_.Thumbprint
            }
        })
    [PSCustomObject]@{ Host = $env:COMPUTERNAME; Certs = $certs }
}

# --- Run: one fan-out Invoke-Command (parallel on both PS 5.1 and 7) ---
$remoteErrors = @()
$answers = @(Invoke-Command -ComputerName $targets -ScriptBlock $probe -ArgumentList $IssuerMatch `
                -ThrottleLimit $ThrottleLimit -ErrorAction SilentlyContinue -ErrorVariable remoteErrors)

$byHost = @{}
foreach ($a in $answers) { $byHost[$a.PSComputerName.ToLower()] = $a }

$errByHost = @{}
foreach ($e in $remoteErrors) {
    $name = $null
    if ($e.OriginInfo -and $e.OriginInfo.PSComputerName) { $name = $e.OriginInfo.PSComputerName }
    elseif ($e.TargetObject -is [string]) { $name = $e.TargetObject }
    if ($name) { $errByHost[$name.ToLower()] = ($e.Exception.Message -split "`r?`n")[0] }
}

# --- Flatten + report ---
$flat = [System.Collections.Generic.List[object]]::new()
$unreachable = 0; $withMatch = 0; $clientAuth = 0
foreach ($t in $targets) {
    $s = $byHost[$t.ToLower()]
    if (-not $s) {
        $unreachable++
        $err = if ($errByHost.ContainsKey($t.ToLower())) { $errByHost[$t.ToLower()] } else { 'no answer (WinRM unreachable or access denied)' }
        Write-Host ("  [UNREACH] {0}  {1}" -f $t, $err) -ForegroundColor DarkGray
        $flat.Add([PSCustomObject]@{ Computer=$t; Reachable=$false; Subject='';Issuer='';NotAfter='';DaysToExpiry='';EKU='';HasClientAuth='';Thumbprint=''; Error=$err })
        continue
    }
    $certs = @($s.Certs | Where-Object { $_ })
    if ($certs.Count -eq 0) {
        Write-Host ("  [clean  ] {0}" -f $t) -ForegroundColor Green
        continue
    }
    $withMatch++
    foreach ($c in $certs) {
        if ($c.HasClientAuth) { $clientAuth++ }
        $col = if ($c.HasClientAuth) { 'Red' } elseif ($c.DaysToExpiry -lt 90) { 'Yellow' } else { 'Cyan' }
        Write-Host ("  [MATCH  ] {0,-28} exp {1}  ({2,5}d)  CliAuth={3,-5}  {4}" -f `
            $t, ([datetime]$c.NotAfter).ToString('yyyy-MM-dd'), $c.DaysToExpiry, $c.HasClientAuth, $c.Subject) -ForegroundColor $col
        $flat.Add([PSCustomObject]@{
            Computer=$t; Reachable=$true; Subject=$c.Subject; Issuer=$c.Issuer
            NotAfter=$c.NotAfter; DaysToExpiry=$c.DaysToExpiry; EKU=$c.EKU
            HasClientAuth=$c.HasClientAuth; Thumbprint=$c.Thumbprint; Error=''
        })
    }
}

Write-Host ""
Write-Host ("Summary: matched={0} host(s), clientAuth-capable certs={1}, unreachable={2}, total targets={3}" -f `
    $withMatch, $clientAuth, $unreachable, $targets.Count) -ForegroundColor Cyan
Write-Log "Summary matched=$withMatch clientAuth=$clientAuth unreachable=$unreachable total=$($targets.Count)"

if ($clientAuth -gt 0) {
    Write-Host "WARNING: $clientAuth cert(s) usable for Client Authentication depend on this CA - likely 802.1x/EAP-TLS. Do NOT retire the CA before handling these." -ForegroundColor Red
}
if ($unreachable -gt 0) {
    Write-Host "NOTE: $unreachable host(s) could not be checked - the blast radius is a MINIMUM until they are." -ForegroundColor Yellow
}

# --- CSV ---
try {
    if (-not (Test-Path $LogRoot)) { New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null }
    $flat | Select-Object *, @{n='ScriptVersion';e={$ScriptVersion}} |
        Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
    Write-Host "CSV: $CsvPath" -ForegroundColor Green
    Write-Log "CSV written: $CsvPath" -Level OK
} catch { Write-Log "CSV write failed: $_" -Level ERROR }

return $flat
