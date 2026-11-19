#Requires -Version 5.1
#Requires -Modules ActiveDirectory
<#
.SYNOPSIS
    Lists the five FSMO (operations master) roles and which DC holds each one.

.DESCRIPTION
    Read-only. Queries Get-ADForest (Schema Master, Domain Naming Master) and Get-ADDomain
    (PDC Emulator, RID Master, Infrastructure Master) and returns one object per role.

    Used on its own, and as a building block by:
      DomainScanner/Invoke-DomainHealthAudit.ps1   (AD section: FSMO placement)
      AD/Manage-FSMORoles.ps1                       (list before, verify after a move)

.PARAMETER Domain
    Domain FQDN or a DC to query. Default: the computer's domain.

.PARAMETER ExpectedHolder
    Optional DC name (short or FQDN). Each role gets OnExpected = true/false, compared on
    the short name, so 'dc01' matches 'dc01.contoso.local'.

.OUTPUTS
    PSCustomObject: Role, Scope (Forest/Domain), Holder (FQDN), HolderShort, OnExpected.
    The role names match Move-ADDirectoryServerOperationMasterRole -OperationMasterRole.

.EXAMPLE
    .\Get-FsmoRoleHolder.ps1 | Format-Table

.EXAMPLE
    .\Get-FsmoRoleHolder.ps1 -ExpectedHolder dc01 | Where-Object { -not $_.OnExpected }
    Roles that are NOT on dc01.

.NOTES
    Version : 1.0.0
    Read-only.

    Changelog:
      1.0.0  2026-11-19  First version. Extracted from Invoke-DomainHealthAudit and
                         Manage-FSMORoles, which both queried the holders themselves.
#>

[CmdletBinding()]
param(
    [string]$Domain,

    [string]$ExpectedHolder
)

$ScriptVersion = '1.0.0'
Write-Verbose "Get-FsmoRoleHolder v$ScriptVersion"

$srv = @{}
if ($Domain) { $srv.Server = $Domain }
$dom = Get-ADDomain @srv -ErrorAction Stop
$fst = Get-ADForest @srv -ErrorAction Stop

$expectedShort = if ($ExpectedHolder) { ($ExpectedHolder -split '\.')[0] } else { '' }

$roles = @(
    @{ Role = 'SchemaMaster';         Scope = 'Forest'; Holder = $fst.SchemaMaster }
    @{ Role = 'DomainNamingMaster';   Scope = 'Forest'; Holder = $fst.DomainNamingMaster }
    @{ Role = 'PDCEmulator';          Scope = 'Domain'; Holder = $dom.PDCEmulator }
    @{ Role = 'RIDMaster';            Scope = 'Domain'; Holder = $dom.RIDMaster }
    @{ Role = 'InfrastructureMaster'; Scope = 'Domain'; Holder = $dom.InfrastructureMaster }
)

foreach ($r in $roles) {
    $short = ("$($r.Holder)" -split '\.')[0]
    [PSCustomObject]@{
        Role        = $r.Role
        Scope       = $r.Scope
        Holder      = "$($r.Holder)"
        HolderShort = $short
        OnExpected  = if ($expectedShort) { $short -ieq $expectedShort } else { $null }
    }
}
