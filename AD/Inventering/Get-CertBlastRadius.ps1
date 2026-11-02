#Requires -Version 5.1
#Requires -Modules ActiveDirectory

<#
.SYNOPSIS
    Read-only estate sweep: finds every machine holding a certificate issued by a
    specific (typically decommissioned) CA, with expiry and EKU, to size the blast
    radius before retiring or rebuilding that CA.

.DESCRIPTION
    NON-DESTRUCTIVE. Designed for the "dead CA + domain-wide autoenrollment GPO"
    scenario: a CA is still registered in AD and pushed as a trusted root by
    autoenrollment GPOs, but no longer answers. GPO *scope* is the whole domain;
    this script finds actual *dependency* — which machines really hold certs from
    that CA, and which carry Client Authentication EKU (the 802.1x / EAP-TLS marker).

    For each reachable Windows computer it reports, per matching cert:
      Subject, Issuer, NotAfter, days-to-expiry, EKUs, HasClientAuth, Thumbprint.

    Machine list comes from -ComputerName, or a text/inventory file, or (default)
    all enabled Windows computers in AD.

.PARAMETER IssuerMatch
    Substring/regex matched against each cert's Issuer.
    E.g. 'CN=CA, DC=customer, DC=local', 'ca01', or 'ad-CA01-CA'.

.PARAMETER ComputerName
    Explicit list of computers to sweep. Overrides -InventoryPath and AD discovery.

.PARAMETER InventoryPath
    Path to a file containing one hostname per line (lines starting with # ignored),
    or the project 01-INVENTORY.md (FQDNs are extracted heuristically).

.PARAMETER OnlyWindows
    When discovering from AD, keep only computers whose OperatingSystem is Windows. Default on.

.PARAMETER ThrottleLimit
    Max parallel remote queries (PS7 uses ForEach-Object -Parallel; PS5 runs sequentially). Default 16.

.PARAMETER LogRoot
    Output folder for the CSV + log. Default C:\Logs\DomainAudit.

.EXAMPLE
    .\Get-CertBlastRadius.ps1 -InventoryPath .\01-INVENTORY.md
    Sweep every host named in the inventory for ca01-issued certs.

.EXAMPLE
    .\Get-CertBlastRadius.ps1 -IssuerMatch 'ca01' -ComputerName dc01.customer.local,srv01.customer.local

.NOTES
    Author    : Iver / IT Consult (Christer Rosvall)
    Read-only : queries Cert:\LocalMachine\My remotely; changes nothing.
    Pairs with: Invoke-DomainHealthAudit.ps1 (CERT section flags the dead CA; this sizes it)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$IssuerMatch,   # e.g. 'CN=CA, DC=customer, DC=local' or just 'ca01'
    [string[]]$ComputerName,
    [string]$InventoryPath = '',
    [bool]$OnlyWindows = $true,
    [int]$ThrottleLimit = 16,
    [string]$LogRoot = 'C:\Logs\DomainAudit'
)

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
} elseif ($InventoryPath -and (Test-Path $InventoryPath)) {
    $raw = Get-Content -Path $InventoryPath
    if ($InventoryPath -match '\.md$') {
        # Heuristic: pull FQDN-looking tokens out of the inventory markdown
        $targets = ($raw | Select-String -Pattern '([a-z0-9][a-z0-9\-]*\.[a-z0-9\.\-]+)' -AllMatches |
                    ForEach-Object { $_.Matches.Value }) |
                    Where-Object { $_ -notmatch '\.(md|ps1|csv|html|txt)$' } |
                    Sort-Object -Unique
    } else {
        $targets = $raw | Where-Object { $_ -and $_ -notmatch '^\s*#' } | ForEach-Object { $_.Trim() }
    }
} else {
    Write-Log "No -ComputerName/-InventoryPath — discovering from AD"
    $props = @('OperatingSystem','DNSHostName')
    $comp = Get-ADComputer -Filter { Enabled -eq $true } -Properties $props
    if ($OnlyWindows) { $comp = $comp | Where-Object { $_.OperatingSystem -match 'Windows' } }
    $targets = $comp | ForEach-Object { if ($_.DNSHostName) { $_.DNSHostName } else { $_.Name } }
}
$targets = @($targets | Where-Object { $_ } | Sort-Object -Unique)

Write-Host ""
Write-Host "=== CERT BLAST-RADIUS SWEEP ===" -ForegroundColor Cyan
Write-Host "  Issuer match : $IssuerMatch"
Write-Host "  Targets      : $($targets.Count) host(s)"
Write-Host ""
Write-Log "Sweep start. IssuerMatch='$IssuerMatch' Targets=$($targets.Count)"

# --- The remote probe (used both in PS7 parallel and PS5 sequential) ---
$probe = {
    param($HostName, $IssuerMatch)
    $result = [PSCustomObject]@{
        Computer=$HostName; Reachable=$false; MatchCount=0; Certs=@(); Error=''
    }
    try {
        $certs = Invoke-Command -ComputerName $HostName -ArgumentList $IssuerMatch -ScriptBlock {
            param($Match)
            Get-ChildItem Cert:\LocalMachine\My -ErrorAction Stop |
                Where-Object { $_.Issuer -match [regex]::Escape($Match) } |
                ForEach-Object {
                    [PSCustomObject]@{
                        Subject       = $_.Subject
                        Issuer        = $_.Issuer
                        NotAfter      = $_.NotAfter
                        DaysToExpiry  = [int]([math]::Floor((New-TimeSpan -Start (Get-Date) -End $_.NotAfter).TotalDays))
                        EKU           = ($_.EnhancedKeyUsageList.FriendlyName -join ', ')
                        HasClientAuth = [bool]($_.EnhancedKeyUsageList.ObjectId -contains '1.3.6.1.5.5.7.3.2')
                        Thumbprint    = $_.Thumbprint
                    }
                }
        } -ErrorAction Stop
        $result.Reachable  = $true
        $result.Certs      = @($certs)
        $result.MatchCount = @($certs).Count
    } catch {
        $result.Error = "$_"
    }
    return $result
}

# --- Run (parallel on PS7, sequential on PS5) ---
$sweep = [System.Collections.Generic.List[object]]::new()
if ($PSVersionTable.PSVersion.Major -ge 7) {
    $sweep = $targets | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
        $probe = [scriptblock]::Create($using:probe.ToString())
        & $probe $_ $using:IssuerMatch
    }
} else {
    foreach ($t in $targets) { $sweep.Add((& $probe $t $IssuerMatch)) }
}

# --- Flatten + report ---
$flat = [System.Collections.Generic.List[object]]::new()
$unreachable = 0; $withMatch = 0; $clientAuth = 0
foreach ($s in $sweep) {
    if (-not $s.Reachable) {
        $unreachable++
        Write-Host ("  [UNREACH] {0}  {1}" -f $s.Computer, $s.Error) -ForegroundColor DarkGray
        $flat.Add([PSCustomObject]@{ Computer=$s.Computer; Reachable=$false; Subject='';Issuer='';NotAfter='';DaysToExpiry='';EKU='';HasClientAuth='';Thumbprint=''; Error=$s.Error })
        continue
    }
    if ($s.MatchCount -eq 0) {
        Write-Host ("  [clean  ] {0}" -f $s.Computer) -ForegroundColor Green
        continue
    }
    $withMatch++
    foreach ($c in $s.Certs) {
        if ($c.HasClientAuth) { $clientAuth++ }
        $col = if ($c.HasClientAuth) { 'Red' } elseif ($c.DaysToExpiry -lt 90) { 'Yellow' } else { 'Cyan' }
        Write-Host ("  [MATCH  ] {0,-28} exp {1}  ({2,4}d)  CliAuth={3}  {4}" -f `
            $s.Computer, $c.NotAfter.ToString('yyyy-MM-dd'), $c.DaysToExpiry, $c.HasClientAuth, $c.Subject) -ForegroundColor $col
        $flat.Add([PSCustomObject]@{
            Computer=$s.Computer; Reachable=$true; Subject=$c.Subject; Issuer=$c.Issuer
            NotAfter=$c.NotAfter; DaysToExpiry=$c.DaysToExpiry; EKU=$c.EKU
            HasClientAuth=$c.HasClientAuth; Thumbprint=$c.Thumbprint; Error=''
        })
    }
}

Write-Host ""
Write-Host ("Summary: matched={0} host(s), clientAuth certs={1}, unreachable={2}, total targets={3}" -f `
    $withMatch, $clientAuth, $unreachable, $targets.Count) -ForegroundColor Cyan
Write-Log "Summary matched=$withMatch clientAuth=$clientAuth unreachable=$unreachable total=$($targets.Count)"

if ($clientAuth -gt 0) {
    Write-Host "WARNING: $clientAuth cert(s) with Client Authentication EKU depend on this CA — likely 802.1x/EAP-TLS. Do NOT retire the CA before handling these." -ForegroundColor Red
}

# --- CSV ---
try {
    if (-not (Test-Path $LogRoot)) { New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null }
    $flat | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
    Write-Host "CSV: $CsvPath" -ForegroundColor Green
    Write-Log "CSV written: $CsvPath" -Level OK
} catch { Write-Log "CSV write failed: $_" -Level ERROR }

return $flat
