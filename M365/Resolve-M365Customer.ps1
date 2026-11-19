#Requires -Version 5.1
<#
.SYNOPSIS
    Resolves which M365 customer/tenant to work against from customers.json and explicit
    parameters. Returns TenantId, AdminUPN and the matching profile. Connects to nothing.

.DESCRIPTION
    customers.json is a list of customer profiles (see customers.template.json). This
    script picks the right one and works out the effective values, the same way for every
    M365 script:

      -Customer <ShortName>   profile by ShortName (error if not found)
      -TenantId <value>       if it matches a profile's ShortName, TenantId or PrimaryDomain,
                              that profile is used and its real TenantId wins (so
                              -TenantId contoso or -TenantId contoso.com both work);
                              otherwise the value is used as-is (GUID or domain)
      neither                 with -SelectSingle and exactly one profile: ask to use it;
                              with -SelectSingle and several: error listing them;
                              without -SelectSingle: no profile

    Explicit -AdminUPN / -TenantId always win over profile values (except the -TenantId
    match described above).

    Building block for:
      M365/Get-CustomerReport.ps1
      M365/move365user-shared.ps1
      M365/Get-CADeletionAudit.ps1

.PARAMETER Customer
    ShortName of a profile in customers.json.

.PARAMETER TenantId
    Tenant GUID or domain, or a ShortName / PrimaryDomain that matches a profile.

.PARAMETER AdminUPN
    Admin account. Overrides the profile's AdminUPN.

.PARAMETER CustomersFile
    Path to customers.json. Default: customers.json next to the calling script
    (-CallerRoot), else next to this script.

.PARAMETER CallerRoot
    Folder of the calling script, used to find the default customers.json.

.PARAMETER SelectSingle
    When neither -Customer nor -TenantId is given: offer the only profile (Y/N prompt),
    or fail listing the profiles when there are several.

.OUTPUTS
    PSCustomObject: Profile (the customers.json entry, or $null), Profiles (all entries),
    CustomersFile, FileExists, ShortName, DisplayName, TenantId, AdminUPN, MatchedBy
    (Customer / TenantId / Single / None).
    Throws with a readable message when the requested customer cannot be resolved.

.EXAMPLE
    $c = & .\Resolve-M365Customer.ps1 -Customer contoso
    Connect-MgGraph -TenantId $c.TenantId -Scopes User.Read.All

.NOTES
    Version : 1.0.0
    Changelog:
      1.0.0  2026-11-19  First version. Replaces three slightly different copies of the
                         profile lookup in Get-CustomerReport, move365user-shared and
                         Get-CADeletionAudit.
#>
[CmdletBinding()]
param(
    [string]$Customer,
    [string]$TenantId,
    [string]$AdminUPN,
    [string]$CustomersFile,
    [string]$CallerRoot,
    [switch]$SelectSingle
)

$ScriptVersion = '1.0.0'
Write-Verbose "Resolve-M365Customer v$ScriptVersion"

if (-not $CustomersFile) {
    $root = if ($CallerRoot) { $CallerRoot } else { $PSScriptRoot }
    $CustomersFile = Join-Path $root 'customers.json'
}

$exists   = Test-Path -LiteralPath $CustomersFile -PathType Leaf
$profiles = @()
if ($exists) {
    try {
        # Assign first, then enumerate: on Windows PowerShell 5.1 ConvertFrom-Json writes a JSON
        # array as ONE object, so @(... | ConvertFrom-Json) would be a 1-element list.
        $json = Get-Content -LiteralPath $CustomersFile -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        $profiles = @($json | ForEach-Object { $_ })
    } catch {
        throw "Could not read $CustomersFile as JSON: $($_.Exception.Message)"
    }
}
$available = (@($profiles | ForEach-Object { $_.ShortName }) -join ', ')

$match = $null
$by    = 'None'
$effTenant = $TenantId

if ($Customer) {
    if (-not $exists) {
        throw "Customer profile file not found: $CustomersFile. Create it from customers.template.json, or pass -TenantId / -AdminUPN directly."
    }
    $match = @($profiles | Where-Object { $_.ShortName -ieq $Customer }) | Select-Object -First 1
    if (-not $match) { throw "Customer '$Customer' not found in $CustomersFile. Available: $available" }
    $by = 'Customer'
} elseif ($TenantId) {
    $match = @($profiles | Where-Object {
        $_.ShortName -ieq $TenantId -or $_.TenantId -ieq $TenantId -or $_.PrimaryDomain -ieq $TenantId
    }) | Select-Object -First 1
    if ($match) {
        $by = 'TenantId'
        $effTenant = $null          # use the profile's real TenantId, not the alias
    }
} elseif ($SelectSingle -and $profiles.Count -gt 0) {
    if ($profiles.Count -eq 1) {
        $p = $profiles[0]
        Write-Host ''
        Write-Host "[INFO] Auto-selected customer: $($p.DisplayName) [$($p.ShortName)]" -ForegroundColor Cyan
        Write-Host "       TenantId : $($p.TenantId)" -ForegroundColor Cyan
        Write-Host "       Domain   : $($p.PrimaryDomain)" -ForegroundColor Cyan
        $yn = Read-Host '  Continue with this customer? (Y/N)'
        if ($yn -notmatch '^[Yy]$') { throw "Cancelled. Re-run with -Customer <name>. Available: $available" }
        $match = $p
        $by = 'Single'
    } else {
        throw "Multiple customers in $CustomersFile - specify -Customer <name>. Available: $available"
    }
}

$prop = { param($o, $n) if ($o -and $o.PSObject.Properties[$n] -and $o.$n) { "$($o.$n)" } else { $null } }

[PSCustomObject]@{
    Profile       = $match
    Profiles      = $profiles
    CustomersFile = $CustomersFile
    FileExists    = $exists
    ShortName     = & $prop $match 'ShortName'
    DisplayName   = & $prop $match 'DisplayName'
    TenantId      = if ($effTenant) { $effTenant } else { & $prop $match 'TenantId' }
    AdminUPN      = if ($AdminUPN) { $AdminUPN } else { & $prop $match 'AdminUPN' }
    MatchedBy     = $by
}
