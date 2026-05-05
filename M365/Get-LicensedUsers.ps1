#Requires -Version 5.1
<#
.SYNOPSIS
    Enhanced Microsoft 365 Licensed Users Report Generator
.DESCRIPTION
    Generates detailed reports showing:
    - Licensed users with last sign-in dates and activity status
    - Inactive user highlighting
    - Unused paid licenses summary at top of dashboard
    - Unlicensed users
    - License SKU breakdown
    - Interactive HTML dashboard

    FUTURE: Per-user and total license cost estimation (planned, not yet implemented)

.PARAMETER Language
    Report language (English or Swedish).
.PARAMETER SkipSignInLookup
    Skip sign-in activity lookup (faster, no AuditLog.Read.All scope needed).
.PARAMETER InactiveThresholdDays
    Days since last sign-in before a user is flagged inactive. Default: 90.
.PARAMETER EnableLogging
    Enable file logging.
.EXAMPLE
    .\Get-LicensedUsers.ps1
.EXAMPLE
    .\Get-LicensedUsers.ps1 -Language Swedish -InactiveThresholdDays 60
.EXAMPLE
    .\Get-LicensedUsers.ps1 -SkipSignInLookup
#>

[CmdletBinding()]
param(
    [string]$OutputPdfPath,
    [ValidateSet("English", "Swedish")]
    [string]$Language = "English",
    [switch]$SkipSignInLookup,
    [ValidateRange(1, 3650)]
    [int]$InactiveThresholdDays = 90,
    [switch]$EnableLogging
)

$ErrorActionPreference = 'Stop'
$script:ScriptVersion  = "3.1.0"

function Write-Log {
    param([string]$Level, [string]$Message)
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $color = switch ($Level) {
        'Success' { 'Green' }
        'Warning' { 'Yellow' }
        'Error'   { 'Red' }
        default   { 'Cyan' }
    }
    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $color
}

Write-Log -Level Info -Message "M365 Licensed Users Report v$script:ScriptVersion"
Write-Log -Level Info -Message "Installing required modules..."

$modules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Identity.DirectoryManagement'
)
foreach ($module in $modules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Log -Level Warning -Message "Installing $module..."
        Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $module -ErrorAction Stop
}

Write-Log -Level Info -Message "Connecting to Microsoft Graph..."
$scopes = @("User.Read.All", "Directory.Read.All")
if (-not $SkipSignInLookup) { $scopes += "AuditLog.Read.All" }

try {
    Connect-MgGraph -Scopes $scopes -ErrorAction Stop
}
catch {
    Write-Log -Level Error -Message "Failed to connect to Microsoft Graph: $_"
    Write-Log -Level Error -Message "Ensure the account has at least the Global Reader role in Entra ID,"
    Write-Log -Level Error -Message "or both Directory Readers + Reports Reader roles."
    throw
}

# Verify that the required scopes were actually granted after consent
$grantedScopes = (Get-MgContext).Scopes
$missingScopes  = $scopes | Where-Object { $_ -notin $grantedScopes }

if ($missingScopes) {
    foreach ($scope in $missingScopes) {
        $hint = switch ($scope) {
            'User.Read.All'        { "Assign 'Directory Readers' or 'Global Reader' role in Entra ID." }
            'Directory.Read.All'   { "Assign 'Directory Readers' or 'Global Reader' role in Entra ID." }
            'AuditLog.Read.All'    { "Assign 'Reports Reader' or 'Global Reader' role in Entra ID, or re-run with -SkipSignInLookup." }
            default                { "Check Entra ID role assignments." }
        }
        Write-Log -Level Warning -Message "Scope not granted: $scope — $hint"
    }

    # AuditLog missing — fall back to skipping sign-in rather than aborting
    if ('AuditLog.Read.All' -in $missingScopes -and -not $SkipSignInLookup) {
        Write-Log -Level Warning -Message "Falling back to -SkipSignInLookup due to missing AuditLog.Read.All."
        $SkipSignInLookup = $true
    }

    # Critical scopes missing — cannot continue
    $criticalMissing = $missingScopes | Where-Object { $_ -in @('User.Read.All', 'Directory.Read.All') }
    if ($criticalMissing) {
        Write-Log -Level Error -Message "Cannot continue — critical scopes not granted: $($criticalMissing -join ', ')"
        Write-Log -Level Error -Message "Required Entra ID role: 'Global Reader' (or 'Directory Readers' for read-only access)."
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        throw "Insufficient Microsoft Graph permissions. See log above for details."
    }
}

Write-Log -Level Info -Message "Retrieving tenant information..."
$tenantInfo   = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/organization"
$tenantDomain = ($tenantInfo.value[0].verifiedDomains | Where-Object { $_.isDefault }).name
if (-not $tenantDomain) { $tenantDomain = "unknown-tenant" }
Write-Log -Level Info -Message "Tenant: $tenantDomain"

# --- Output directory ---
$documentsFolder = [Environment]::GetFolderPath('MyDocuments')
$reportRoot      = Join-Path $documentsFolder 'User Reports'
$domainFolder    = Join-Path $reportRoot $tenantDomain
$tsStamp         = Get-Date -Format 'yyyyMMdd-HHmmss'
$outputFolder    = Join-Path $domainFolder $tsStamp
if (-not (Test-Path $outputFolder)) {
    New-Item -Path $outputFolder -ItemType Directory -Force | Out-Null
}
Write-Log -Level Info -Message "Output folder: $outputFolder"

# --- SKUs ---
Write-Log -Level Info -Message "Retrieving SKU information..."
try {
    $skus = Get-MgSubscribedSku -All -ErrorAction Stop
    Write-Log -Level Success -Message "Retrieved $($skus.Count) SKU records"
}
catch {
    if ($_.Exception.Message -match '403|Forbidden|Authorization_RequestDenied|Insufficient') {
        Write-Log -Level Error -Message "Access denied reading license SKUs. Requires Directory.Read.All."
        Write-Log -Level Error -Message "Assign 'Directory Readers' or 'Global Reader' role in Entra ID."
    }
    else {
        Write-Log -Level Error -Message "Failed to retrieve SKU information: $_"
    }
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    throw
}

$skuLookup = @{}
foreach ($sku in $skus) {
    $skuLookup[$sku.SkuId] = @{
        SkuPartNumber = $sku.SkuPartNumber
        DisplayName   = if ($sku.SkuPartNumber) { $sku.SkuPartNumber } else { $sku.SkuId }
    }
}

# --- License friendly names from Microsoft catalog ---
Write-Log -Level Info -Message "Loading license catalog..."
$licenseCatalog = @{}
try {
    $catalogUrl = "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv"
    $tempFile   = Join-Path $env:TEMP "license-catalog.csv"
    Invoke-WebRequest -Uri $catalogUrl -OutFile $tempFile -UseBasicParsing -ErrorAction Stop
    $catalog = Import-Csv -Path $tempFile
    foreach ($item in $catalog) {
        if ($item.String_Id -and $item.Product_Display_Name) {
            $licenseCatalog[$item.String_Id] = $item.Product_Display_Name
        }
    }
    Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
    Write-Log -Level Success -Message "Loaded $($licenseCatalog.Count) license names"
}
catch {
    Write-Log -Level Warning -Message "Could not load license catalog: $_"
}

# --- Retrieve users ---
Write-Log -Level Info -Message "Retrieving users..."
$userProperties = "Id,AssignedLicenses,DisplayName,UserPrincipalName,Mail,AccountEnabled"
if (-not $SkipSignInLookup) { $userProperties += ",SignInActivity" }

try {
    $users = Get-MgUser -All -Property $userProperties -ConsistencyLevel eventual -ErrorAction Stop
    Write-Log -Level Success -Message "Retrieved $($users.Count) users"
}
catch {
    if ($_.Exception.Message -match '403|Forbidden|Authorization_RequestDenied|Insufficient') {
        Write-Log -Level Error -Message "Access denied reading users. Requires User.Read.All."
        Write-Log -Level Error -Message "Assign 'Directory Readers' or 'Global Reader' role in Entra ID."
        # If SignInActivity caused the 403, retry without it
        if (-not $SkipSignInLookup -and $_.Exception.Message -match 'SignInActivity|AuditLog') {
            Write-Log -Level Warning -Message "Retrying without SignInActivity (requires AuditLog.Read.All / Reports Reader role)."
            $SkipSignInLookup = $true
            $userProperties   = "Id,AssignedLicenses,DisplayName,UserPrincipalName,Mail,AccountEnabled"
            try {
                $users = Get-MgUser -All -Property $userProperties -ConsistencyLevel eventual -ErrorAction Stop
                Write-Log -Level Success -Message "Retrieved $($users.Count) users (sign-in data skipped)"
            }
            catch {
                Write-Log -Level Error -Message "Still unable to read users: $_"
                Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
                throw
            }
        }
        else {
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            throw
        }
    }
    else {
        Write-Log -Level Error -Message "Failed to retrieve users: $_"
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        throw
    }
}

# --- Process users ---
$today           = Get-Date
$licensedUsers   = [System.Collections.Generic.List[PSCustomObject]]::new()
$unlicensedUsers = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($user in $users) {
    # Sign-in data
    $lastSignIn      = $null
    $daysSinceSignIn = $null
    $isInactive      = $false
    $neverSignedIn   = $false

    if (-not $SkipSignInLookup) {
        if ($user.SignInActivity -and $user.SignInActivity.LastSignInDateTime) {
            $lastSignIn      = $user.SignInActivity.LastSignInDateTime
            $daysSinceSignIn = [int](($today - $lastSignIn).TotalDays)
            $isInactive      = $daysSinceSignIn -ge $InactiveThresholdDays
        }
        else {
            $neverSignedIn = $true
        }
    }

    if ($user.AssignedLicenses -and $user.AssignedLicenses.Count -gt 0) {
        $userLicenses = [System.Collections.Generic.List[string]]::new()
        $licenseNames = [System.Collections.Generic.List[string]]::new()

        foreach ($license in $user.AssignedLicenses) {
            $skuId = $license.SkuId
            if ($skuLookup.ContainsKey($skuId)) {
                $skuPart     = $skuLookup[$skuId].SkuPartNumber
                $displayName = if ($licenseCatalog.ContainsKey($skuPart)) { $licenseCatalog[$skuPart] } else { $skuPart }
                $userLicenses.Add($displayName)
                $licenseNames.Add($displayName)
            }
            else {
                $userLicenses.Add("Unknown ($skuId)")
                $licenseNames.Add("Unknown")
            }
        }

        $licensedUsers.Add([PSCustomObject]@{
            DisplayName       = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
            Mail              = $user.Mail
            AccountEnabled    = $user.AccountEnabled
            LicenseCount      = $user.AssignedLicenses.Count
            Licenses          = (($userLicenses | Sort-Object -Unique) -join '; ')
            LicenseList       = $licenseNames.ToArray()
            LastSignIn        = $lastSignIn
            DaysSinceSignIn   = $daysSinceSignIn
            NeverSignedIn     = $neverSignedIn
            IsInactive        = $isInactive
        })
    }
    else {
        $unlicensedUsers.Add([PSCustomObject]@{
            DisplayName       = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
            Mail              = $user.Mail
            AccountEnabled    = $user.AccountEnabled
        })
    }
}

Write-Log -Level Success -Message "Found $($licensedUsers.Count) licensed users"
Write-Log -Level Success -Message "Found $($unlicensedUsers.Count) unlicensed users"

# --- Totals ---
$inactiveCount      = ($licensedUsers | Where-Object { $_.IsInactive }).Count
$neverSignedInCount = ($licensedUsers | Where-Object { $_.NeverSignedIn }).Count

# --- Unused paid licenses ---
$unusedLicenses = $skus | Where-Object {
    $purchased = if ($_.PrepaidUnits.Enabled) { $_.PrepaidUnits.Enabled } else { 0 }
    $consumed  = if ($_.ConsumedUnits) { $_.ConsumedUnits } else { 0 }
    $purchased -gt 0 -and $consumed -lt $purchased
} | ForEach-Object {
    $purchased   = if ($_.PrepaidUnits.Enabled) { $_.PrepaidUnits.Enabled } else { 0 }
    $consumed    = if ($_.ConsumedUnits) { $_.ConsumedUnits } else { 0 }
    $displayName = if ($licenseCatalog.ContainsKey($_.SkuPartNumber)) { $licenseCatalog[$_.SkuPartNumber] } else { $_.SkuPartNumber }
    [PSCustomObject]@{
        SkuPartNumber = $_.SkuPartNumber
        DisplayName   = $displayName
        Purchased     = $purchased
        Consumed      = $consumed
        Available     = $purchased - $consumed
    }
} | Sort-Object Available -Descending

Write-Log -Level Info -Message "Unused SKUs with available seats: $($unusedLicenses.Count)"

# ============================================================
# HTML generation — build sections as strings first
# ============================================================

# -- Unused licenses cards --
$unusedCardsHtml = ''
foreach ($ul in $unusedLicenses) {
    $unusedCardsHtml += "<div class='unused-card'>"
    $unusedCardsHtml += "<div class='unused-name'>$([System.Web.HttpUtility]::HtmlEncode($ul.DisplayName))</div>"
    $unusedCardsHtml += "<div class='unused-seats'><strong>$($ul.Available)</strong> unused of $($ul.Purchased)</div>"
    $unusedCardsHtml += "</div>"
}

$unusedSectionHtml = ''
if ($unusedLicenses.Count -gt 0) {
    $unusedSectionHtml = @"
<div class='unused-section'>
  <div class='unused-section-title'>Unused Paid Licenses &mdash; $($unusedLicenses.Count) SKU(s) with available seats</div>
  <div class='unused-grid'>$unusedCardsHtml</div>
</div>
"@
}

# -- Licensed users rows --
$licensedRowsHtml = ''
foreach ($user in ($licensedUsers | Sort-Object DisplayName)) {
    $statusClass = if ($user.AccountEnabled) { 'status-enabled' } else { 'status-disabled' }
    $statusText  = if ($user.AccountEnabled) { 'Enabled' } else { 'Disabled' }

    $signInDisplay = if ($SkipSignInLookup) {
        "<span class='na'>N/A</span>"
    }
    elseif ($user.NeverSignedIn) {
        "<span class='badge-never'>Never</span>"
    }
    else {
        "$($user.LastSignIn.ToString('yyyy-MM-dd'))<br><small>$($user.DaysSinceSignIn)d ago</small>"
    }

    $rowClass = if (-not $SkipSignInLookup -and $user.NeverSignedIn) { 'row-never-signin' }
                elseif (-not $SkipSignInLookup -and $user.IsInactive) { 'row-inactive' }
                else { '' }

    $licensePills = ($user.LicenseList | Sort-Object -Unique | ForEach-Object {
        "<span class='license-pill'>$([System.Web.HttpUtility]::HtmlEncode($_))</span>"
    }) -join ' '

    $licensedRowsHtml += "<tr class='$rowClass'>"
    $licensedRowsHtml += "<td>$([System.Web.HttpUtility]::HtmlEncode($user.DisplayName))</td>"
    $licensedRowsHtml += "<td>$([System.Web.HttpUtility]::HtmlEncode($user.UserPrincipalName))</td>"
    $licensedRowsHtml += "<td class='$statusClass'>$statusText</td>"
    $licensedRowsHtml += "<td>$signInDisplay</td>"
    $licensedRowsHtml += "<td>$licensePills</td>"
    $licensedRowsHtml += "</tr>`n"
}

# -- Unlicensed users rows --
$unlicensedRowsHtml = ''
foreach ($user in ($unlicensedUsers | Sort-Object DisplayName)) {
    $statusClass = if ($user.AccountEnabled) { 'status-enabled' } else { 'status-disabled' }
    $statusText  = if ($user.AccountEnabled) { 'Enabled' } else { 'Disabled' }
    $unlicensedRowsHtml += "<tr>"
    $unlicensedRowsHtml += "<td>$([System.Web.HttpUtility]::HtmlEncode($user.DisplayName))</td>"
    $unlicensedRowsHtml += "<td>$([System.Web.HttpUtility]::HtmlEncode($user.UserPrincipalName))</td>"
    $unlicensedRowsHtml += "<td>$([System.Web.HttpUtility]::HtmlEncode($user.Mail))</td>"
    $unlicensedRowsHtml += "<td class='$statusClass'>$statusText</td>"
    $unlicensedRowsHtml += "</tr>`n"
}

# -- SKU rows --
$skuRowsHtml = ''
foreach ($sku in ($skus | Sort-Object SkuPartNumber)) {
    $purchased   = if ($sku.PrepaidUnits.Enabled) { $sku.PrepaidUnits.Enabled } else { 0 }
    $consumed    = if ($sku.ConsumedUnits) { $sku.ConsumedUnits } else { 0 }
    $available   = $purchased - $consumed
    $displayName = if ($licenseCatalog.ContainsKey($sku.SkuPartNumber)) { $licenseCatalog[$sku.SkuPartNumber] } else { $sku.SkuPartNumber }
    $rowClass    = if ($available -gt 0) { 'sku-unused' } else { '' }
    $skuRowsHtml += "<tr class='$rowClass'>"
    $skuRowsHtml += "<td>$([System.Web.HttpUtility]::HtmlEncode($sku.SkuPartNumber))</td>"
    $skuRowsHtml += "<td>$([System.Web.HttpUtility]::HtmlEncode($displayName))</td>"
    $skuRowsHtml += "<td>$purchased</td>"
    $skuRowsHtml += "<td>$consumed</td>"
    $skuRowsHtml += "<td>$available</td>"
    $skuRowsHtml += "</tr>`n"
}

$signInNote = if ($SkipSignInLookup) { ' | Sign-in data: skipped' } else { " | Inactive threshold: $InactiveThresholdDays days" }

# ============================================================
# Full HTML
# ============================================================
Write-Log -Level Info -Message "Generating HTML report..."
$htmlPath = Join-Path $outputFolder 'M365-LicensedUsers.html'

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>M365 Users - $tenantDomain</title>
<style>
*{box-sizing:border-box}
body{font-family:'Segoe UI',sans-serif;margin:0;padding:20px;background:#f0f2f5}
.container{max-width:1500px;margin:0 auto;background:white;padding:30px;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,.1)}
h1{color:#0078d4;margin-top:0;border-bottom:3px solid #0078d4;padding-bottom:10px}
h2{color:#0078d4;margin-top:30px;padding-bottom:8px;border-bottom:2px solid #e1e4e8}

.summary{display:grid;grid-template-columns:repeat(auto-fit,minmax(175px,1fr));gap:15px;margin:20px 0}
.metric{background:#f8f9fa;padding:15px;border-radius:6px;border-left:4px solid #0078d4}
.metric-label{font-size:11px;color:#666;text-transform:uppercase;letter-spacing:.5px}
.metric-value{font-size:26px;font-weight:700;color:#0078d4;margin-top:4px}
.metric-sub{font-size:11px;color:#888;margin-top:3px}
.metric.warn{border-left-color:#d13438}.metric.warn .metric-value{color:#d13438}

.unused-section{background:#fff8e1;border:1px solid #f9a825;border-radius:8px;padding:20px;margin:20px 0}
.unused-section-title{font-size:14px;font-weight:700;color:#b45309;margin-bottom:12px}
.unused-section-title::before{content:'⚠  '}
.unused-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(210px,1fr));gap:10px}
.unused-card{background:white;border:1px solid #fcd34d;border-radius:6px;padding:12px}
.unused-name{font-size:12px;font-weight:600;color:#333;margin-bottom:5px;line-height:1.3}
.unused-seats{font-size:13px;color:#555}

.search-bar{margin:15px 0;padding:12px;background:#f8f9fa;border-radius:6px;display:flex;gap:10px;align-items:center;flex-wrap:wrap}
.search-bar input{flex:1;min-width:200px;max-width:380px;padding:9px 12px;border:2px solid #ddd;border-radius:6px;font-size:14px}
.search-bar input:focus{outline:none;border-color:#0078d4}
.filter-btns{display:flex;gap:6px;flex-wrap:wrap}
.filter-btn{padding:8px 14px;border:1px solid #ddd;background:white;cursor:pointer;border-radius:4px;font-size:12px;font-weight:600;transition:all .2s}
.filter-btn:hover{border-color:#0078d4;color:#0078d4}
.filter-btn.active{background:#0078d4;color:white;border-color:#0078d4}

table{border-collapse:collapse;width:100%;margin-top:12px;font-size:13px}
th,td{border:1px solid #ddd;padding:9px 8px;text-align:left}
th{background:#0078d4;color:white;font-weight:600;position:sticky;top:0;cursor:pointer;user-select:none;white-space:nowrap}
th:hover{background:#006cbf}
tr:nth-child(even){background:#f8f9fa}
tr:hover{background:#e7f3ff !important}

.row-inactive{background:#fff3cd !important}
.row-never-signin{background:#fce4e4 !important}
.sku-unused{background:#fff8e1 !important}

.status-enabled{color:#107c10;font-weight:600}
.status-disabled{color:#d13438;font-weight:600}
.badge-never{color:#d13438;font-weight:700}
.na{color:#bbb;font-style:italic}
.license-pill{display:inline-block;background:#e1f5ff;color:#0078d4;padding:3px 9px;border-radius:12px;margin:2px;font-size:11px;white-space:nowrap}

.legend{display:flex;gap:16px;margin:8px 0 4px;font-size:12px;flex-wrap:wrap}
.legend-item{display:flex;align-items:center;gap:6px}
.ldot{width:14px;height:14px;border-radius:2px;flex-shrink:0}
.ldot-inactive{background:#fff3cd;border:1px solid #f0c000}
.ldot-never{background:#fce4e4;border:1px solid #f5b8b8}

.tab-buttons{display:flex;gap:10px;margin:20px 0;flex-wrap:wrap}
.tab-btn{padding:11px 22px;border:2px solid #0078d4;background:white;color:#0078d4;cursor:pointer;border-radius:6px;font-size:14px;font-weight:600;transition:all .2s}
.tab-btn:hover{background:#e1f5ff}
.tab-btn.active{background:#0078d4;color:white}
.tab-content{display:none}
.tab-content.active{display:block}

.footer{margin-top:30px;padding-top:20px;border-top:1px solid #ddd;text-align:center;color:#888;font-size:12px}
</style>
<script>
function searchTable(tableId, inputId) {
    const filter = document.getElementById(inputId).value.toLowerCase();
    document.getElementById(tableId).querySelectorAll('tbody tr').forEach(row => {
        const match = Array.from(row.cells).some(c => c.textContent.toLowerCase().includes(filter));
        row.style.display = match ? '' : 'none';
    });
}
function filterByClass(tableId, cls, btnId) {
    const btn    = document.getElementById(btnId);
    const active = btn.classList.contains('active');
    document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
    document.getElementById(tableId).querySelectorAll('tbody tr').forEach(row => {
        row.style.display = (active || row.classList.contains(cls)) ? '' : 'none';
    });
    if (!active) btn.classList.add('active');
}
function showTab(tabId, btn) {
    document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.getElementById(tabId).classList.add('active');
    btn.classList.add('active');
}
function sortTable(tableId, col) {
    const table = document.getElementById(tableId);
    const asc   = table.dataset.sortCol == col && table.dataset.sortDir == 'asc';
    const rows  = Array.from(table.querySelectorAll('tbody tr'));
    rows.sort((a, b) => {
        const va = a.cells[col]?.textContent.trim() || '';
        const vb = b.cells[col]?.textContent.trim() || '';
        return asc ? vb.localeCompare(va,undefined,{numeric:true}) : va.localeCompare(vb,undefined,{numeric:true});
    });
    const tbody = table.querySelector('tbody');
    rows.forEach(r => tbody.appendChild(r));
    table.dataset.sortCol = col;
    table.dataset.sortDir = asc ? 'desc' : 'asc';
}
</script>
</head>
<body>
<div class="container">

<h1>Microsoft 365 Users Report</h1>
<p><strong>Tenant:</strong> $tenantDomain &nbsp;|&nbsp; <strong>Generated:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')$signInNote</p>

<div class="summary">
  <div class="metric"><div class="metric-label">Total Users</div><div class="metric-value">$($users.Count)</div></div>
  <div class="metric"><div class="metric-label">Licensed Users</div><div class="metric-value">$($licensedUsers.Count)</div></div>
  <div class="metric"><div class="metric-label">Unlicensed Users</div><div class="metric-value">$($unlicensedUsers.Count)</div></div>
  <div class="metric warn"><div class="metric-label">Inactive (&gt;${InactiveThresholdDays}d)</div><div class="metric-value">$inactiveCount</div><div class="metric-sub">$neverSignedInCount never signed in</div></div>
  <div class="metric"><div class="metric-label">SKU Types</div><div class="metric-value">$($skus.Count)</div></div>
</div>

$unusedSectionHtml

<div class="tab-buttons">
  <button class="tab-btn active" onclick="showTab('tab-licensed', this)">Licensed Users ($($licensedUsers.Count))</button>
  <button class="tab-btn" onclick="showTab('tab-unlicensed', this)">Unlicensed Users ($($unlicensedUsers.Count))</button>
  <button class="tab-btn" onclick="showTab('tab-skus', this)">License SKUs ($($skus.Count))</button>
</div>

<div id="tab-licensed" class="tab-content active">
  <h2>Licensed Users</h2>
  <div class="search-bar">
    <input type="text" id="lic-search" onkeyup="searchTable('lic-table','lic-search')" placeholder="Search users...">
    <div class="filter-btns">
      <button class="filter-btn" id="btn-inactive" onclick="filterByClass('lic-table','row-inactive','btn-inactive')">Show Inactive</button>
      <button class="filter-btn" id="btn-never"    onclick="filterByClass('lic-table','row-never-signin','btn-never')">Never Signed In</button>
    </div>
  </div>
  <div class="legend">
    <div class="legend-item"><div class="ldot ldot-inactive"></div><span>Inactive &gt; $InactiveThresholdDays days</span></div>
    <div class="legend-item"><div class="ldot ldot-never"></div><span>Never signed in</span></div>
  </div>
  <table id="lic-table">
    <thead><tr>
      <th onclick="sortTable('lic-table',0)">Name &#8597;</th>
      <th onclick="sortTable('lic-table',1)">UPN &#8597;</th>
      <th onclick="sortTable('lic-table',2)">Status &#8597;</th>
      <th onclick="sortTable('lic-table',3)">Last Sign-In &#8597;</th>
      <th>Licenses</th>
    </tr></thead>
    <tbody>$licensedRowsHtml</tbody>
  </table>
</div>

<div id="tab-unlicensed" class="tab-content">
  <h2>Unlicensed Users</h2>
  <div class="search-bar">
    <input type="text" id="unlic-search" onkeyup="searchTable('unlic-table','unlic-search')" placeholder="Search users...">
  </div>
  <table id="unlic-table">
    <thead><tr>
      <th onclick="sortTable('unlic-table',0)">Name &#8597;</th>
      <th onclick="sortTable('unlic-table',1)">UPN &#8597;</th>
      <th>Email</th>
      <th onclick="sortTable('unlic-table',3)">Status &#8597;</th>
    </tr></thead>
    <tbody>$unlicensedRowsHtml</tbody>
  </table>
</div>

<div id="tab-skus" class="tab-content">
  <h2>License SKUs <small style="font-size:13px;color:#888">(yellow = unused seats available)</small></h2>
  <table>
    <thead><tr>
      <th>SKU Part Number</th>
      <th>Display Name</th>
      <th>Purchased</th>
      <th>Consumed</th>
      <th>Available</th>
    </tr></thead>
    <tbody>$skuRowsHtml</tbody>
  </table>
</div>

<div class="footer">
  <p>Generated by M365 Licensed Users Report v$script:ScriptVersion &mdash; Powered by Microsoft Graph PowerShell SDK</p>
</div>
</div>
</body>
</html>
"@

Set-Content -Path $htmlPath -Value $html -Encoding UTF8
Write-Log -Level Success -Message "HTML report saved: $htmlPath"

# --- CSV exports ---
Write-Log -Level Info -Message "Exporting CSV files..."

$licensedUsers |
    Select-Object DisplayName, UserPrincipalName, Mail, AccountEnabled,
                  Licenses, LicenseCount, LastSignIn, DaysSinceSignIn, IsInactive, NeverSignedIn |
    Export-Csv -Path (Join-Path $outputFolder 'Licensed-Users.csv') -NoTypeInformation -Encoding UTF8
Write-Log -Level Success -Message "Licensed-Users.csv written"

$unlicensedUsers |
    Select-Object DisplayName, UserPrincipalName, Mail, AccountEnabled |
    Export-Csv -Path (Join-Path $outputFolder 'Unlicensed-Users.csv') -NoTypeInformation -Encoding UTF8
Write-Log -Level Success -Message "Unlicensed-Users.csv written"

$skus | Select-Object SkuPartNumber,
    @{N='DisplayName'; E={ if ($licenseCatalog.ContainsKey($_.SkuPartNumber)) { $licenseCatalog[$_.SkuPartNumber] } else { $_.SkuPartNumber } }},
    @{N='Purchased';   E={ if ($_.PrepaidUnits.Enabled) { $_.PrepaidUnits.Enabled } else { 0 } }},
    @{N='Consumed';    E={ if ($_.ConsumedUnits) { $_.ConsumedUnits } else { 0 } }},
    @{N='Available';   E={ $p = if ($_.PrepaidUnits.Enabled) { $_.PrepaidUnits.Enabled } else { 0 }
                           $c = if ($_.ConsumedUnits) { $_.ConsumedUnits } else { 0 }
                           $p - $c }} |
    Export-Csv -Path (Join-Path $outputFolder 'License-SKUs.csv') -NoTypeInformation -Encoding UTF8
Write-Log -Level Success -Message "License-SKUs.csv written"

$unusedLicenses |
    Export-Csv -Path (Join-Path $outputFolder 'Unused-Licenses.csv') -NoTypeInformation -Encoding UTF8
Write-Log -Level Success -Message "Unused-Licenses.csv written"

# --- Open report ---
Start-Process $htmlPath
Write-Log -Level Success -Message "Done!"
Write-Host ""
Write-Host "Report location: $outputFolder" -ForegroundColor Green
Write-Host ""
Write-Host "Files created:" -ForegroundColor Yellow
Write-Host "  - M365-LicensedUsers.html  (Interactive dashboard)" -ForegroundColor White
Write-Host "  - Licensed-Users.csv       (Users + sign-in data)" -ForegroundColor White
Write-Host "  - Unlicensed-Users.csv     (Unlicensed users)" -ForegroundColor White
Write-Host "  - License-SKUs.csv         (SKU information)" -ForegroundColor White
Write-Host "  - Unused-Licenses.csv      (Unused paid seats)" -ForegroundColor White
Write-Host ""

Disconnect-MgGraph | Out-Null
