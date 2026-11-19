#Requires -Version 5.1
#Requires -Modules ActiveDirectory
<#
.SYNOPSIS
    Lists every user in the privileged AD groups and scores each one OK / REVIEW / HIGH-RISK.

.DESCRIPTION
    Read-only. Expands (recursively) the built-in privileged groups and classifies each
    user account found:

      HIGH-RISK  disabled account in a high-privilege group
                 account listed in -KnownRiskAccounts (vendor / external / orphaned)
                 never logged on, in a high-privilege group
                 service-account name pattern in a high-privilege group
      REVIEW     last logon more than -InactiveDays ago, in a high-privilege group
                 named admin account (adm.x / admin-x) in a high-privilege group
      OK         none of the above

    High-privilege = Domain Admins, Enterprise Admins, Schema Admins. Also expanded (and
    reported, but not scored as high-privilege): Administrators, Backup Operators,
    Account Operators, Server Operators, Print Operators, Replicator, Key Admins,
    Enterprise Key Admins - i.e. the groups AdminSDHolder protects.

    Groups are looked up by well-known SID, not by name, so this works on domains whose
    DCs were installed in another language (e.g. "Domanadministratorer" on Swedish).

    Used on its own, and as a building block by:
      DomainScanner/Invoke-DomainHealthAudit.ps1   (USERS + SECURITY sections)

.PARAMETER Domain
    Domain FQDN or DC to query. Default: the computer's domain.

.PARAMETER KnownRiskAccounts
    sAMAccountNames known to be vendor / external / orphaned accounts.

.PARAMETER ServiceAcctRegex
    Regex for service-account names. Default
    '^(svc[-_]|sa[-_]|sql|exchange|iis|backup|print|scan|service|mssql)'.

.PARAMETER NamedAdminRegex
    Regex for personal admin accounts that should be re-confirmed. Default '^(adm[-_.]|admin[-_.])'.

.PARAMETER InactiveDays
    Days since last logon before an account in a high-privilege group gets REVIEW. Default 90.

.OUTPUTS
    PSCustomObject per user: SamAccountName, DisplayName, Enabled, LastLogonDate,
    DaysSinceLogon, PasswordLastSet, PasswordNeverExpires, Groups (string[]),
    InHighPriv, Risk, Reason.
    A group that cannot be expanded (foreign / orphaned members) is reported as a
    non-terminating error, so callers can collect it with -ErrorVariable.

.EXAMPLE
    .\Get-PrivilegedAccountRisk.ps1 | Where-Object Risk -ne 'OK' | Format-Table SamAccountName,Risk,Reason

.EXAMPLE
    .\Get-PrivilegedAccountRisk.ps1 -KnownRiskAccounts vendor-support -InactiveDays 60 |
        Export-Csv .\privileged.csv -NoTypeInformation -Encoding UTF8

.NOTES
    Version : 1.0.0
    Read-only.

    Changelog:
      1.0.0  2026-11-19  First version. Extracted from Invoke-DomainHealthAudit (USERS
                         section). Groups by well-known SID instead of English name;
                         adds the other AdminSDHolder-protected groups to the expanded set.
#>

[CmdletBinding()]
param(
    [string]$Domain,

    [string[]]$KnownRiskAccounts = @(),

    [string]$ServiceAcctRegex = '^(svc[-_]|sa[-_]|sql|exchange|iis|backup|print|scan|service|mssql)',

    [string]$NamedAdminRegex = '^(adm[-_.]|admin[-_.])',

    [ValidateRange(1, 3650)]
    [int]$InactiveDays = 90
)

$ScriptVersion = '1.0.0'
Write-Verbose "Get-PrivilegedAccountRisk v$ScriptVersion"

$srv = @{}
if ($Domain) { $srv.Server = $Domain }
$dom = Get-ADDomain @srv -ErrorAction Stop
$fst = Get-ADForest @srv -ErrorAction Stop
$domSid = $dom.DomainSID.Value

# Enterprise Admins / Schema Admins live in the forest root domain.
$rootSid = $domSid
$rootSrv = $srv
if ($fst.RootDomain -and $fst.RootDomain -ine $dom.DNSRoot) {
    try {
        $rootSid = (Get-ADDomain -Identity $fst.RootDomain -Server $fst.RootDomain -ErrorAction Stop).DomainSID.Value
        $rootSrv = @{ Server = $fst.RootDomain }
    } catch {
        Write-Warning "Could not read forest root domain '$($fst.RootDomain)': $_ - Enterprise/Schema Admins skipped."
        $rootSid = $null
    }
}

# Key = stable English name used in output; Sid = lookup; High = counts as high-privilege.
$groupDefs = @(
    @{ Key = 'Domain Admins';     Sid = "$domSid-512";  High = $true;  Srv = $srv }
    @{ Key = 'Enterprise Admins'; Sid = "$rootSid-519"; High = $true;  Srv = $rootSrv; NeedRoot = $true }
    @{ Key = 'Schema Admins';     Sid = "$rootSid-518"; High = $true;  Srv = $rootSrv; NeedRoot = $true }
    @{ Key = 'Administrators';    Sid = 'S-1-5-32-544'; High = $false; Srv = $srv }
    @{ Key = 'Account Operators'; Sid = 'S-1-5-32-548'; High = $false; Srv = $srv }
    @{ Key = 'Server Operators';  Sid = 'S-1-5-32-549'; High = $false; Srv = $srv }
    @{ Key = 'Print Operators';   Sid = 'S-1-5-32-550'; High = $false; Srv = $srv }
    @{ Key = 'Backup Operators';  Sid = 'S-1-5-32-551'; High = $false; Srv = $srv }
    @{ Key = 'Replicator';        Sid = 'S-1-5-32-552'; High = $false; Srv = $srv }
    @{ Key = 'Key Admins';        Sid = "$domSid-526";  High = $false; Srv = $srv }
    @{ Key = 'Enterprise Key Admins'; Sid = "$rootSid-527"; High = $false; Srv = $rootSrv; NeedRoot = $true }
)
$highKeys = @($groupDefs | Where-Object { $_.High } | ForEach-Object { $_.Key })

$membership = @{}   # sam -> List[string] of group keys
foreach ($g in $groupDefs) {
    if ($g.NeedRoot -and -not $rootSid) { continue }
    $gSrv = $g.Srv
    try {
        $members = @(Get-ADGroupMember -Identity $g.Sid -Recursive @gSrv -ErrorAction Stop |
                     Where-Object { $_.objectClass -eq 'user' })
        foreach ($m in $members) {
            if (-not $membership.ContainsKey($m.SamAccountName)) {
                $membership[$m.SamAccountName] = [System.Collections.Generic.List[string]]::new()
            }
            if ($g.Key -notin $membership[$m.SamAccountName]) { $membership[$m.SamAccountName].Add($g.Key) }
        }
    } catch {
        Write-Error -Message "Could not expand group '$($g.Key)' ($($g.Sid)) - foreign or orphaned members?: $_" `
                    -Category ReadError -TargetObject $g.Key
    }
}

$now = Get-Date
foreach ($sam in ($membership.Keys | Sort-Object)) {
    try {
        $u = Get-ADUser -Identity $sam @srv -Properties Enabled, LastLogonDate, PasswordNeverExpires, PasswordLastSet, DisplayName -ErrorAction Stop
    } catch {
        Write-Error -Message "Could not read user '$sam': $_" -Category ReadError -TargetObject $sam
        continue
    }
    $groups = @($membership[$sam])
    $inHigh = @($groups | Where-Object { $_ -in $highKeys }).Count -gt 0
    $days   = if ($u.LastLogonDate) { [int](New-TimeSpan -Start $u.LastLogonDate -End $now).TotalDays } else { $null }

    $risk = 'OK'; $reason = 'Active account, no risk flags'
    if (-not $u.Enabled -and $inHigh)                                  { $risk = 'HIGH-RISK'; $reason = 'Disabled account in high-priv group' }
    elseif ($sam -in $KnownRiskAccounts)                               { $risk = 'HIGH-RISK'; $reason = 'Known vendor/external account' }
    elseif ($null -eq $u.LastLogonDate -and $inHigh)                   { $risk = 'HIGH-RISK'; $reason = 'Never logged in and in high-priv group' }
    elseif ($ServiceAcctRegex -and $sam -match $ServiceAcctRegex -and $inHigh) { $risk = 'HIGH-RISK'; $reason = 'Service-account name pattern in high-priv group' }
    elseif ($null -ne $days -and $days -gt $InactiveDays -and $inHigh) { $risk = 'REVIEW';    $reason = "Last logon $days days ago in high-priv group" }
    elseif ($NamedAdminRegex -and $sam -match $NamedAdminRegex -and $inHigh) { $risk = 'REVIEW'; $reason = 'Named admin account in high-priv group - verify still needed' }

    [PSCustomObject]@{
        SamAccountName       = $sam
        DisplayName          = $u.DisplayName
        Enabled              = $u.Enabled
        LastLogonDate        = $u.LastLogonDate
        DaysSinceLogon       = $days
        PasswordLastSet      = $u.PasswordLastSet
        PasswordNeverExpires = $u.PasswordNeverExpires
        Groups               = $groups
        InHighPriv           = $inHigh
        Risk                 = $risk
        Reason               = $reason
    }
}
