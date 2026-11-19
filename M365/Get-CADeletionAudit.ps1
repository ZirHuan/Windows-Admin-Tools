#Requires -Version 7.0
<#
.SYNOPSIS
    Queries Entra ID audit logs for Conditional Access policy deletion events.
.DESCRIPTION
    Searches directoryAudits for 'Delete conditionalAccessPolicy' events.
    Requires AuditLog.Read.All scope — will prompt to reconnect if missing.
    Supports -Customer to load TenantId from customers.json, or supply -TenantId directly.
.PARAMETER Customer
    Short name of a customer profile in customers.json (e.g. "contoso").
    Loads TenantId automatically. Explicit -TenantId always overrides.
.PARAMETER TenantId
    Entra ID Tenant ID or primary domain. Required if -Customer is not provided.
.EXAMPLE
    .\Get-CADeletionAudit.ps1 -Customer contoso
.EXAMPLE
    .\Get-CADeletionAudit.ps1 -TenantId "00000000-0000-0000-0000-000000000000"
.NOTES
    Version : 1.1.0
    Building block: Resolve-M365Customer.ps1 (same folder) - customers.json lookup.
    Changelog:
      1.1.0  2026-11-19  Profile lookup via Resolve-M365Customer.ps1. -TenantId may now
                         also be a profile ShortName or PrimaryDomain.
      1.0.0              First version.
#>

[CmdletBinding()]
param(
    [string]$Customer,
    [string]$TenantId
)

# ── Resolve tenant from customers.json (building block: Resolve-M365Customer) ──
$resolver = Join-Path $PSScriptRoot 'Resolve-M365Customer.ps1'
if (-not (Test-Path -LiteralPath $resolver)) { throw "Resolve-M365Customer.ps1 not found next to this script - copy the whole M365 folder." }
$cust = & $resolver -Customer $Customer -TenantId $TenantId -CallerRoot $PSScriptRoot
$TenantId = $cust.TenantId
if ($cust.Profile) {
    Write-Host "[INFO] Loaded TenantId for '$($cust.DisplayName)': $TenantId" -ForegroundColor Cyan
}

if (-not $TenantId) {
    throw "Provide -Customer <shortname> or -TenantId <guid/domain>."
}

# ── Graph connection ──────────────────────────────────────────────────────────
$context = Get-MgContext
if (-not $context) {
    Write-Host "[INFO] Not connected to Graph. Connecting..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes "AuditLog.Read.All" -TenantId $TenantId
} elseif ($context.Scopes -notcontains "AuditLog.Read.All") {
    Write-Warning "Current session is missing AuditLog.Read.All. Reconnecting..."
    Connect-MgGraph -Scopes "AuditLog.Read.All" -TenantId $TenantId
} else {
    Write-Host "[INFO] Connected as $($context.Account)" -ForegroundColor Cyan
}

# ── Query audit logs ──────────────────────────────────────────────────────────
$filter = "activityDisplayName eq 'Delete conditionalAccessPolicy'"
$select = "activityDateTime,initiatedBy,targetResources,result,correlationId"

$uri = "https://graph.microsoft.com/v1.0/auditLogs/directoryAudits" +
       "?`$filter=$([uri]::EscapeDataString($filter))" +
       "&`$select=$select" +
       "&`$top=50" +
       "&`$orderby=activityDateTime desc"

try {
    Write-Host "[INFO] Querying audit logs for CA policy deletions..." -ForegroundColor Cyan

    # Page through all results via @odata.nextLink — a single request caps at the
    # page size ($top) and would silently drop older deletion events.
    $events = [System.Collections.Generic.List[object]]::new()
    $next   = $uri
    while ($next) {
        $page = Invoke-MgGraphRequest -Method GET -Uri $next -ErrorAction Stop
        if ($page.value) { $events.AddRange([object[]]$page.value) }
        $next = $page.'@odata.nextLink'
    }

    if ($events.Count -eq 0) {
        Write-Host ""
        Write-Host "No CA policy deletion events found in audit logs." -ForegroundColor Yellow
        Write-Host "Either no policies were deleted, or log retention has expired (max 30 days without Entra P1/P2)." -ForegroundColor Yellow
    } else {
        Write-Host "[INFO] Found $($events.Count) deletion event(s):" -ForegroundColor Green
        Write-Host ""

        $results = $events | ForEach-Object {
            $actor = if ($_.initiatedBy.user) {
                "$($_.initiatedBy.user.userPrincipalName) (User)"
            } elseif ($_.initiatedBy.app) {
                "$($_.initiatedBy.app.displayName) (App)"
            } else { "Unknown" }

            $targets = ($_.targetResources | ForEach-Object { $_.displayName }) -join ", "

            [PSCustomObject]@{
                DateTime      = $_.activityDateTime
                InitiatedBy   = $actor
                PolicyDeleted = $targets
                Result        = $_.result
                CorrelationId = $_.correlationId
            }
        }

        $results | Format-Table -AutoSize
    }
} catch {
    if ($_ -match "403|Forbidden|Authorization") {
        Write-Warning "Access denied — AuditLog.Read.All scope missing or not consented."
        Write-Warning "Run: Connect-MgGraph -Scopes 'AuditLog.Read.All' -TenantId '$TenantId'"
    } elseif ($_ -match "Authentication_RequestFromNonPremiumTenantOrB2CTenant") {
        Write-Warning "Audit log access requires Entra ID P1 or P2. This tenant has no premium license."
        Write-Warning "Log history is not available for this tenant."
    } else {
        Write-Warning "Query failed: $($_.Exception.Message)"
    }
}
