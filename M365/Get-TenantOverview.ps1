#Requires -Version 7.0

<#
.SYNOPSIS
    Generates an HTML overview report of a Microsoft 365 tenant.
.DESCRIPTION
    Collects tenant info, domains, licenses, users, admin roles, groups,
    Conditional Access policies, Exchange Online services, and service health.
    Will prompt to connect to Microsoft Graph and Exchange Online if not already connected.
.PARAMETER TenantId
    Tenant ID or primary domain (e.g. "contoso.com"). Used when connecting to Graph.
.PARAMETER AdminUPN
    Admin UPN for Exchange Online connection (e.g. "admin@contoso.com").
.PARAMETER TenantName
    Friendly name used in the report title (e.g. "Contoso").
.PARAMETER OutputPath
    Folder where the HTML report is saved. Defaults to C:\Reports\.
.PARAMETER SkipExchangeOnline
    Skip Exchange Online data collection entirely.
.EXAMPLE
    .\Get-TenantOverview.ps1 -TenantId "contoso.com" -AdminUPN "admin@contoso.com" -TenantName "Contoso"
.EXAMPLE
    .\Get-TenantOverview.ps1 -TenantName "Contoso" -SkipExchangeOnline
#>

[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$AdminUPN,
    [string]$TenantName   = "M365 Tenant",
    [string]$OutputPath   = "C:\Reports",
    [switch]$SkipExchangeOnline
)

# ── REQUIRED MODULES ─────────────────────────────────────────────────────────
$requiredModules = @(
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Identity.DirectoryManagement",
    "Microsoft.Graph.Users",
    "Microsoft.Graph.Groups",
    "Microsoft.Graph.Identity.SignIns",
    "ExchangeOnlineManagement"
)

Write-Host "`n[INFO] Checking required modules..." -ForegroundColor Cyan
foreach ($mod in $requiredModules) {
    # Try importing first — module may be installed outside standard paths
    Import-Module $mod -ErrorAction SilentlyContinue
    if (-not (Get-Module -Name $mod)) {
        Write-Warning "  Module '$mod' could not be loaded."
        $install = Read-Host "  Attempt install via PSResourceGet? (Y/N)"
        if ($install -match "^[Yy]$") {
            try {
                Install-PSResource $mod -Scope CurrentUser -TrustRepository -ErrorAction Stop
                Import-Module $mod -ErrorAction Stop
                Write-Host "  Installed and loaded $mod" -ForegroundColor Green
            } catch {
                Write-Warning "  Could not install $mod — $_"
                Write-Warning "  Some report sections may be unavailable."
            }
        }
    } else {
        Write-Host "  OK — $mod" -ForegroundColor Green
    }
}

# ── GRAPH CONNECTION ──────────────────────────────────────────────────────────
$requiredScopes = @(
    "User.Read.All",
    "Directory.Read.All",
    "Policy.Read.All"
)

$graphContext = Get-MgContext
if (-not $graphContext) {
    Write-Host "`n[INFO] Not connected to Microsoft Graph." -ForegroundColor Yellow
    $connect = Read-Host "  Connect now? (Y/N)"
    if ($connect -match "^[Yy]$") {
        if (-not $TenantId) { $TenantId = Read-Host "  Enter Tenant ID or domain (e.g. contoso.com)" }
        try {
            # Try standard auth first
            Connect-MgGraph -Scopes $requiredScopes -TenantId $TenantId -ErrorAction Stop
        } catch {
            if ($_ -match "NullReferenceException|RuntimeBroker|window handle") {
                # WAM broker failure — fall back to non-WAM
                Write-Warning "  WAM authentication failed. Retrying without WAM..."
                Connect-MgGraph -Scopes $requiredScopes -TenantId $TenantId -NoWAM -ErrorAction Stop
            } else { throw }
        }
        Write-Host "  Connected to Microsoft Graph." -ForegroundColor Green
    } else {
        throw "Microsoft Graph connection is required. Exiting."
    }
} else {
    Write-Host "[INFO] Already connected to Microsoft Graph as $($graphContext.Account)" -ForegroundColor Green
}

# ── EXCHANGE ONLINE CONNECTION ────────────────────────────────────────────────
if (-not $SkipExchangeOnline) {
    $exoConnected = $false
    try {
        $null = Get-EXOMailbox -ResultSize 1 -ErrorAction Stop
        $exoConnected = $true
        Write-Host "[INFO] Already connected to Exchange Online." -ForegroundColor Green
    } catch {
        Write-Host "`n[INFO] Not connected to Exchange Online." -ForegroundColor Yellow
        $connect = Read-Host "  Connect now? (Y/N)"
        if ($connect -match "^[Yy]$") {
            try {
                # Use Device code flow — avoids WAM NullReferenceException on Windows
                Write-Host "  Opening device login flow (enter code in browser when prompted)..." -ForegroundColor Cyan
                Connect-ExchangeOnline -Device -ShowBanner:$false -ErrorAction Stop
            } catch {
                Write-Warning "  Device auth failed: $_"
                Write-Warning "  Skipping Exchange Online — EXO section will be empty in the report."
            }
            # Verify connection took effect
            try {
                $null = Get-EXOMailbox -ResultSize 1 -ErrorAction Stop
                $exoConnected = $true
                Write-Host "  Connected to Exchange Online." -ForegroundColor Green
            } catch {
                Write-Warning "  EXO session could not be verified — EXO section may be empty."
            }
        } else {
            Write-Warning "  Skipping Exchange Online — EXO section will be empty in the report."
        }
    }
}

Write-Host ""

# ── HELPER: HTML TABLE BUILDER ──────────────────────────────────────────────
function ConvertTo-HtmlTable {
    param(
        [Parameter(ValueFromPipeline)]
        [object[]]$InputObject,
        [string[]]$Properties
    )
    begin { $rows = @() }
    process { $rows += $InputObject }
    end {
        if (-not $rows) { return "<p class='empty'>No data found.</p>" }
        if (-not $Properties) { $Properties = $rows[0].PSObject.Properties.Name }

        $html = "<table><thead><tr>"
        foreach ($p in $Properties) { $html += "<th>$p</th>" }
        $html += "</tr></thead><tbody>"

        foreach ($row in $rows) {
            $html += "<tr>"
            foreach ($p in $Properties) {
                $val = $row.$p
                if ($null -eq $val) { $val = "" }
                $html += "<td>$([System.Web.HttpUtility]::HtmlEncode($val.ToString()))</td>"
            }
            $html += "</tr>"
        }
        $html += "</tbody></table>"
        return $html
    }
}

# ── HELPER: STAT CARD ────────────────────────────────────────────────────────
function New-StatCard {
    param([string]$Label, [string]$Value, [string]$Color = "#0078d4")
    return "<div class='stat-card' style='border-top:4px solid $Color'>
                <div class='stat-value'>$Value</div>
                <div class='stat-label'>$Label</div>
            </div>"
}

# ── HELPER: SECTION ──────────────────────────────────────────────────────────
function New-Section {
    param([string]$Title, [string]$Content, [string]$Icon = "")
    return "<div class='section'>
                <h2>$Icon $Title</h2>
                $Content
            </div>"
}

# ── DATA COLLECTION ──────────────────────────────────────────────────────────
Write-Host "[INFO] Collecting tenant information..." -ForegroundColor Cyan

# Org
$org = Get-MgOrganization
$tenantDisplayName = if ($TenantName -ne "M365 Tenant") { $TenantName } else { $org.DisplayName }

# Domains
Write-Host "[INFO] Collecting domains..." -ForegroundColor Cyan
$domains = Get-MgDomain | Select-Object Id, IsDefault, IsVerified, AuthenticationType

# Licenses
Write-Host "[INFO] Collecting licenses..." -ForegroundColor Cyan
$licenses = Get-MgSubscribedSku | Select-Object SkuPartNumber,
    @{N="Total";     E={ $_.PrepaidUnits.Enabled }},
    @{N="Assigned";  E={ $_.ConsumedUnits }},
    @{N="Available"; E={ $_.PrepaidUnits.Enabled - $_.ConsumedUnits }}

# Users
Write-Host "[INFO] Collecting users (this may take a moment)..." -ForegroundColor Cyan
$users = Get-MgUser -All -Property DisplayName,UserPrincipalName,AccountEnabled,AssignedLicenses,UserType,CreatedDateTime

$userStats = [PSCustomObject]@{
    Total    = $users.Count
    Enabled  = ($users | Where-Object AccountEnabled -eq $true).Count
    Disabled = ($users | Where-Object AccountEnabled -eq $false).Count
    Licensed = ($users | Where-Object { $_.AssignedLicenses.Count -gt 0 }).Count
    Guests   = ($users | Where-Object UserType -eq "Guest").Count
}

$userTable = $users | Select-Object DisplayName, UserPrincipalName,
    @{N="Enabled";E={ $_.AccountEnabled }},
    @{N="Type";E={ $_.UserType }},
    @{N="Licenses";E={ $_.AssignedLicenses.Count }},
    @{N="Created";E={ if ($_.CreatedDateTime) { ([datetime]$_.CreatedDateTime).ToString("yyyy-MM-dd") } else { "" } }} |
    Sort-Object DisplayName

# Admin roles
Write-Host "[INFO] Collecting admin role assignments..." -ForegroundColor Cyan
$roleRows = @()
$roles = Get-MgDirectoryRole -All
foreach ($role in $roles) {
    $members = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -ErrorAction SilentlyContinue
    foreach ($m in $members) {
        $u = Get-MgUser -UserId $m.Id -ErrorAction SilentlyContinue
        if ($u) {
            $roleRows += [PSCustomObject]@{
                Role              = $role.DisplayName
                DisplayName       = $u.DisplayName
                UserPrincipalName = $u.UserPrincipalName
            }
        }
    }
}

# Groups
Write-Host "[INFO] Collecting groups..." -ForegroundColor Cyan
$groups = Get-MgGroup -All -Property DisplayName,GroupTypes,SecurityEnabled,MailEnabled,MembershipRule,CreatedDateTime

$groupStats = [PSCustomObject]@{
    Total          = $groups.Count
    "M365 Groups"  = ($groups | Where-Object { $_.GroupTypes -contains "Unified" }).Count
    Security       = ($groups | Where-Object { $_.SecurityEnabled -and $_.GroupTypes -notcontains "Unified" }).Count
    Dynamic        = ($groups | Where-Object { $_.GroupTypes -contains "DynamicMembership" }).Count
}

$groupTable = $groups | Select-Object DisplayName,
    @{N="Type";E={
        if ($_.GroupTypes -contains "Unified") { "Microsoft 365" }
        elseif ($_.GroupTypes -contains "DynamicMembership") { "Dynamic Security" }
        elseif ($_.SecurityEnabled) { "Security" }
        else { "Distribution" }
    }},
    @{N="Mail Enabled";E={ $_.MailEnabled }},
    @{N="Dynamic Rule";E={ if ($_.MembershipRule) { "Yes" } else { "No" } }},
    @{N="Created";E={ if ($_.CreatedDateTime) { ([datetime]$_.CreatedDateTime).ToString("yyyy-MM-dd") } else { "" } }} |
    Sort-Object DisplayName

# Conditional Access
Write-Host "[INFO] Collecting Conditional Access policies..." -ForegroundColor Cyan
$caTable = $null
$caError = $null
try {
    $caPolicies = Get-MgIdentityConditionalAccessPolicy -All -ErrorAction Stop
    $caTable = $caPolicies | Select-Object DisplayName, State,
        @{N="Created";E={ if ($_.CreatedDateTime) { ([datetime]$_.CreatedDateTime).ToString("yyyy-MM-dd") } else { "" } }},
        @{N="Modified";E={ if ($_.ModifiedDateTime) { ([datetime]$_.ModifiedDateTime).ToString("yyyy-MM-dd") } else { "" } }} |
        Sort-Object DisplayName
} catch {
    $caError = "Could not retrieve Conditional Access policies. Reconnect with 'Policy.Read.All' scope."
}

# Exchange Online — M365 Services
Write-Host "[INFO] Collecting Exchange Online data..." -ForegroundColor Cyan
$exoError = $null
$exoSharedMailboxes = $null
$exoDistGroups = $null
$exoInbound = $null
$exoOutbound = $null
$exoTransportRules = $null
try {
    # Test if EXO session is active
    $null = Get-EXOMailbox -ResultSize 1 -ErrorAction Stop

    $exoSharedMailboxes = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -ErrorAction Stop |
        Select-Object DisplayName, PrimarySmtpAddress,
            @{N="HiddenFromGAL"; E={ $_.HiddenFromAddressListsEnabled }}

    $exoDistGroups = Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop |
        Select-Object DisplayName, PrimarySmtpAddress, GroupType

    $exoInbound = Get-InboundConnector -ErrorAction Stop |
        Select-Object Name, Enabled, ConnectorType,
            @{N="SenderDomains"; E={ $_.SenderDomains -join ", " }}

    $exoOutbound = Get-OutboundConnector -ErrorAction Stop |
        Select-Object Name, Enabled, ConnectorType,
            @{N="RecipientDomains"; E={ $_.RecipientDomains -join ", " }}

    $exoTransportRules = Get-TransportRule -ErrorAction Stop |
        Select-Object Name, State, Priority
} catch {
    $exoError = "Exchange Online data not available. Connect first: Connect-ExchangeOnline -UserPrincipalName admin@$($org.VerifiedDomains | Where-Object IsDefault | Select-Object -ExpandProperty Name)"
}

# Monitoring & Notifications
Write-Host "[INFO] Collecting notification settings..." -ForegroundColor Cyan
$techNotificationMails = $org.TechnicalNotificationMails


# Pre-compute conditional content (if/else can't be used inside here-string subexpressions)
$roleContent = if ($roleRows.Count -gt 0) { $roleRows | ConvertTo-HtmlTable } else { "<p class='empty'>No role assignments found.</p>" }

# ── BUILD HTML ───────────────────────────────────────────────────────────────
Write-Host "[INFO] Building HTML report..." -ForegroundColor Cyan

Add-Type -AssemblyName System.Web

$reportDate = Get-Date -Format "yyyy-MM-dd HH:mm"

# Tenant info section
$tenantContent = "<div class='info-grid'>
    <div><span class='label'>Tenant Name</span><span class='value'>$($org.DisplayName)</span></div>
    <div><span class='label'>Tenant ID</span><span class='value'>$($org.Id)</span></div>
    <div><span class='label'>Country</span><span class='value'>$($org.CountryLetterCode)</span></div>
    <div><span class='label'>Created</span><span class='value'>$(if ($org.CreatedDateTime) { ([datetime]$org.CreatedDateTime).ToString('yyyy-MM-dd') })</span></div>
</div>"

# Stat cards
$userCards = "<div class='stat-grid'>" +
    (New-StatCard -Label "Total Users"    -Value $userStats.Total    -Color "#0078d4") +
    (New-StatCard -Label "Enabled"        -Value $userStats.Enabled  -Color "#107c10") +
    (New-StatCard -Label "Disabled"       -Value $userStats.Disabled -Color "#797775") +
    (New-StatCard -Label "Licensed"       -Value $userStats.Licensed -Color "#0078d4") +
    (New-StatCard -Label "Guests"         -Value $userStats.Guests   -Color "#d83b01") +
    "</div>"

$groupCards = "<div class='stat-grid'>" +
    (New-StatCard -Label "Total Groups"   -Value $groupStats.Total           -Color "#0078d4") +
    (New-StatCard -Label "M365 Groups"    -Value $groupStats."M365 Groups"   -Color "#107c10") +
    (New-StatCard -Label "Security"       -Value $groupStats.Security        -Color "#8764b8") +
    (New-StatCard -Label "Dynamic"        -Value $groupStats.Dynamic         -Color "#d83b01") +
    "</div>"

$caContent = if ($caError) {
    "<p class='warning'>&#9888; $caError</p>"
} else {
    $caTable | ConvertTo-HtmlTable
}

# EXO section content
$exoContent = if ($exoError) {
    "<p class='warning'>&#9888; $exoError</p>"
} else {
    $smCount  = if ($exoSharedMailboxes) { $exoSharedMailboxes.Count } else { 0 }
    $dgCount  = if ($exoDistGroups)      { $exoDistGroups.Count }      else { 0 }
    $ibCount  = if ($exoInbound)         { $exoInbound.Count }         else { 0 }
    $obCount  = if ($exoOutbound)        { $exoOutbound.Count }        else { 0 }
    $trCount  = if ($exoTransportRules)  { $exoTransportRules.Count }  else { 0 }

    $exoCards = "<div class='stat-grid'>" +
        (New-StatCard -Label "Shared Mailboxes"      -Value $smCount -Color "#0078d4") +
        (New-StatCard -Label "Distribution Groups"   -Value $dgCount -Color "#107c10") +
        (New-StatCard -Label "Inbound Connectors"    -Value $ibCount -Color "#8764b8") +
        (New-StatCard -Label "Outbound Connectors"   -Value $obCount -Color "#8764b8") +
        (New-StatCard -Label "Transport Rules"       -Value $trCount -Color "#d83b01") +
        "</div>"

    $smHtml  = "<h3>Shared Mailboxes</h3>"  + ($exoSharedMailboxes | ConvertTo-HtmlTable)
    $dgHtml  = "<h3>Distribution Groups</h3>" + ($exoDistGroups | ConvertTo-HtmlTable)
    $ibHtml  = "<h3>Inbound Connectors</h3>"  + ($exoInbound | ConvertTo-HtmlTable)
    $obHtml  = "<h3>Outbound Connectors</h3>" + ($exoOutbound | ConvertTo-HtmlTable)
    $trHtml  = "<h3>Transport Rules</h3>"     + ($exoTransportRules | ConvertTo-HtmlTable)

    $exoCards + $smHtml + $dgHtml + $ibHtml + $obHtml + $trHtml
}

# Monitoring section content
$techMailsHtml = if ($techNotificationMails) {
    $techNotificationMails | ForEach-Object { "<li>$_</li>" } | Out-String
    "<ul style='padding-left:20px;line-height:2'>$($techNotificationMails | ForEach-Object { "<li>$_</li>" })</ul>"
} else {
    "<p class='empty'>No technical notification emails configured.</p>"
}

$monitoringContent = "<h3>Technical Notification Emails</h3>$techMailsHtml"

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>$tenantDisplayName — Tenant Overview</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Segoe UI', system-ui, sans-serif; background: #f3f2f1; color: #323130; font-size: 14px; }

        header { background: #0078d4; color: white; padding: 24px 40px; }
        header h1 { font-size: 24px; font-weight: 600; }
        header p { font-size: 13px; opacity: 0.85; margin-top: 4px; }

        main { max-width: 1400px; margin: 0 auto; padding: 24px 40px; }

        .section { background: white; border-radius: 4px; padding: 24px; margin-bottom: 20px;
                   box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
        .section h2 { font-size: 16px; font-weight: 600; color: #0078d4; margin-bottom: 16px;
                      padding-bottom: 8px; border-bottom: 1px solid #edebe9; }

        .stat-grid { display: flex; gap: 12px; flex-wrap: wrap; margin-bottom: 20px; }
        .stat-card { background: #faf9f8; border-radius: 4px; padding: 16px 20px; min-width: 130px;
                     flex: 1; border-top: 4px solid #0078d4; }
        .stat-value { font-size: 28px; font-weight: 700; color: #323130; }
        .stat-label { font-size: 12px; color: #605e5c; margin-top: 2px; }

        .info-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 12px; }
        .info-grid div { background: #faf9f8; padding: 12px 16px; border-radius: 4px; }
        .info-grid .label { display: block; font-size: 11px; color: #605e5c; text-transform: uppercase; letter-spacing: 0.5px; }
        .info-grid .value { display: block; font-size: 14px; font-weight: 600; margin-top: 2px; }

        table { width: 100%; border-collapse: collapse; font-size: 13px; }
        thead { background: #f3f2f1; }
        th { text-align: left; padding: 10px 12px; font-weight: 600; font-size: 12px;
             text-transform: uppercase; letter-spacing: 0.4px; color: #605e5c;
             border-bottom: 2px solid #edebe9; }
        td { padding: 9px 12px; border-bottom: 1px solid #f3f2f1; vertical-align: middle; }
        tr:hover td { background: #f3f2f1; }
        tr:last-child td { border-bottom: none; }

        .section h3 { font-size: 13px; font-weight: 600; color: #323130; margin: 20px 0 10px; }
        .empty { color: #605e5c; font-style: italic; padding: 12px 0; }
        .warning { background: #fff4ce; border-left: 4px solid #ffb900; padding: 12px 16px;
                   border-radius: 0 4px 4px 0; color: #323130; }

        footer { text-align: center; padding: 20px; color: #605e5c; font-size: 12px; }
    </style>
</head>
<body>
<header>
    <h1>$tenantDisplayName — Tenant Overview</h1>
    <p>Generated $reportDate &nbsp;|&nbsp; Microsoft Graph API</p>
</header>
<main>

$(New-Section -Title "Tenant Information" -Icon "&#127970;" -Content $tenantContent)

$(New-Section -Title "Registered Domains" -Icon "&#127760;" -Content ($domains | ConvertTo-HtmlTable))

$(New-Section -Title "Licenses" -Icon "&#128196;" -Content ($licenses | ConvertTo-HtmlTable))

$(New-Section -Title "Users" -Icon "&#128101;" -Content ($userCards + ($userTable | ConvertTo-HtmlTable)))

$(New-Section -Title "Admin Role Assignments" -Icon "&#128272;" -Content $roleContent)

$(New-Section -Title "Groups" -Icon "&#128101;" -Content ($groupCards + ($groupTable | ConvertTo-HtmlTable)))

$(New-Section -Title "Conditional Access Policies" -Icon "&#128274;" -Content $caContent)

$(New-Section -Title "Exchange Online — M365 Services" -Icon "&#128231;" -Content $exoContent)

$(New-Section -Title "Monitoring &amp; Notifications" -Icon "&#128276;" -Content $monitoringContent)

</main>
<footer>Get-TenantOverview.ps1 &nbsp;&bull;&nbsp; $tenantDisplayName &nbsp;&bull;&nbsp; $reportDate</footer>
</body>
</html>
"@

# ── SAVE REPORT ──────────────────────────────────────────────────────────────
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

$fileName = "$($tenantDisplayName -replace '[^a-zA-Z0-9]', '-')_TenantOverview_$(Get-Date -Format 'yyyyMMdd_HHmm').html"
$reportFile = Join-Path $OutputPath $fileName

$html | Set-Content -Path $reportFile -Encoding UTF8

Write-Host "`n[SUCCESS] Report saved to: $reportFile" -ForegroundColor Green

# Open in default browser
Start-Process $reportFile
