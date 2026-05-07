#Requires -Version 7.0

<#
.SYNOPSIS
    Comprehensive M365 customer report — Key Findings, tenant overview, licenses with
    renewal dates, users with MFA/sign-in/password age, Secure Score, and devices.
.PARAMETER CustomersFile
    Path to a customers.json file. Defaults to customers.json in the script folder.
    Point at a customer subfolder (e.g. ".\naturskog.se\customers.json") to auto-select
    that customer without needing -Customer.
.PARAMETER Customer
    Short name matching a profile in customers.json. Auto-selected if file has one entry.
.PARAMETER TenantId
    Entra ID Tenant ID or primary domain. Overrides profile value.
.PARAMETER AdminUPN
    Admin UPN for Exchange Online. Overrides profile value.
.PARAMETER TenantName
    Friendly name for the report title. Overrides profile value.
.PARAMETER OutputPath
    Folder to save the HTML report. Defaults to .\<shortname>\ next to the script.
.PARAMETER InactiveThresholdDays
    Days since last sign-in before a licensed user is flagged inactive. Default: 90.
.PARAMETER SkipExchangeOnline
    Skip Exchange Online connection and data collection.
.EXAMPLE
    .\Get-CustomerReport.ps1 -CustomersFile ".\naturskog.se\customers.json"
.EXAMPLE
    .\Get-CustomerReport.ps1 -Customer antilopgroup
.EXAMPLE
    .\Get-CustomerReport.ps1 -TenantId "contoso.com" -AdminUPN "admin@contoso.com" -TenantName "Contoso"
.NOTES
    Authors:  Rosvall & Claude
    Version:  1.3.2
    Changelog:
        1.0.0 - Initial release
        1.1.0 - Fix IsExternal to check all verified tenant domains (not only primary)
              - Fix finding detail text for external Global Admin alert
              - Add IE=edge meta tag and browser compatibility warning banner
              - Auto-update customers.json LastReportDate and LastReportPath after each run
              - Expand customers.template.json with PrimaryDomain, AllDomains, Country, etc.
        1.2.0 - Prompt to create or update customers.json with live collected data
              - Prompt to confirm or override the output folder (UseDefaultOutputPath switch to skip)
              - Prompt to disconnect Graph and Exchange Online after completion (DisconnectAfter switch to auto)
        1.3.0 - Check existing Graph/EXO sessions at startup, show who is connected, offer disconnect
              - Source folder renamed from source\ to <shortname>.source\
              - HTML report filename prefixed with shortname
              - customers.json snapshot saved into source folder after each run
        1.3.1 - Auto-select now shows customer details and asks confirmation before proceeding
              - Session check moved to single summary block after module load; warns on tenant mismatch
              - Graph/EXO connect sections simplified — no duplicate session prompts
        1.3.2 - Fix auto-select bypass: -TenantId now correctly skips single-profile auto-select
              - Replace all throw with Write-Host + exit 0 for clean non-exception exits
              - Remove all Disconnect-ExchangeOnline calls — crashes process in EXO 3.9.x
                (MSAL WAM broker NullReferenceException in ClearAllTokensAsync background thread)
              - EXO REST sessions now expire automatically; only Graph is explicitly disconnected
#>

[CmdletBinding()]
param(
    [string]$CustomersFile,
    [string]$Customer,
    [string]$TenantId,
    [string]$AdminUPN,
    [string]$TenantName,
    [string]$OutputPath,
    [ValidateRange(1, 3650)]
    [int]$InactiveThresholdDays = 90,
    [switch]$SkipExchangeOnline,
    # Skip the output path prompt and use the default .\<shortname>\ folder
    [switch]$UseDefaultOutputPath,
    # Automatically create or update customers.json with live data without prompting
    [switch]$UpdateCustomerProfile,
    # Automatically disconnect from Graph and Exchange Online after completion without prompting
    [switch]$DisconnectAfter
)

# ── CUSTOMER PROFILE RESOLUTION ───────────────────────────────────────────────
$profileData = $null
$shortName   = $null

if (-not $CustomersFile) {
    $CustomersFile = Join-Path $PSScriptRoot "customers.json"
}

if (Test-Path $CustomersFile) {
    $profiles = Get-Content $CustomersFile -Raw | ConvertFrom-Json
    $profileList = @($profiles)

    if ($Customer) {
        $profileData = $profileList | Where-Object { $_.ShortName -ieq $Customer }
        if (-not $profileData) {
            $available = ($profileList.ShortName) -join ', '
            Write-Host "Customer '$Customer' not found in $CustomersFile. Available: $available" -ForegroundColor Yellow
            exit 0
        }
    } elseif ($profileList.Count -eq 1 -and -not $TenantId) {
        $profileData = $profileList[0]
        Write-Host ""
        Write-Host "[INFO] Auto-selected customer: $($profileData.DisplayName) [$($profileData.ShortName)]" -ForegroundColor Cyan
        Write-Host "       TenantId : $($profileData.TenantId)" -ForegroundColor Cyan
        Write-Host "       Domain   : $($profileData.PrimaryDomain)" -ForegroundColor Cyan
        $yn = Read-Host "  Continue with this customer? (Y/N)"
        if ($yn -notmatch "^[Yy]$") {
            $available = ($profileList.ShortName) -join ', '
            Write-Host "Cancelled. Re-run with -Customer <name> to specify. Available: $available" -ForegroundColor Yellow
            exit 0
        }
    } elseif (-not $TenantId) {
        $available = ($profileList.ShortName) -join ', '
        Write-Host "Multiple customers in $CustomersFile — specify -Customer <name>. Available: $available" -ForegroundColor Yellow
        exit 0
    }
}

$EffectiveTenantId   = if ($TenantId)   { $TenantId }   elseif ($profileData.TenantId)    { $profileData.TenantId }    else { $null }
$EffectiveAdminUPN   = if ($AdminUPN)   { $AdminUPN }   elseif ($profileData.AdminUPN)    { $profileData.AdminUPN }    else { $null }
$EffectiveTenantName = if ($TenantName) { $TenantName } elseif ($profileData.DisplayName) { $profileData.DisplayName } else { "M365 Tenant" }
$shortName           = if ($profileData.ShortName) { $profileData.ShortName } else { $EffectiveTenantName -replace '[^a-zA-Z0-9]','' }

if (-not $OutputPath) {
    $defaultPath = Join-Path $PSScriptRoot $shortName
    if ($UseDefaultOutputPath) {
        $OutputPath = $defaultPath
    } else {
        Write-Host "[INFO] Default output folder: $defaultPath" -ForegroundColor Cyan
        $customPath = Read-Host "  Press Enter to use default, or type a new path"
        $OutputPath = if ($customPath.Trim()) { $customPath.Trim() } else { $defaultPath }
    }
}

# ── HELPERS ───────────────────────────────────────────────────────────────────
Add-Type -AssemblyName System.Web

function New-StatCard {
    param([string]$Label, $Value, [string]$Color = "#0078d4")
    return "<div class='stat-card' style='border-top:4px solid $Color'><div class='stat-value'>$Value</div><div class='stat-label'>$Label</div></div>"
}

function ConvertTo-HtmlTable {
    param([Parameter(ValueFromPipeline)][object[]]$InputObject, [string[]]$Properties)
    begin { $rows = @() }
    process { $rows += $InputObject }
    end {
        if (-not $rows) { return "<p class='empty'>No data found.</p>" }
        if (-not $Properties) { $Properties = $rows[0].PSObject.Properties.Name }
        $h = "<table><thead><tr>"
        foreach ($p in $Properties) { $h += "<th>$p</th>" }
        $h += "</tr></thead><tbody>"
        foreach ($row in $rows) {
            $h += "<tr>"
            foreach ($p in $Properties) {
                $val = $row.$p
                if ($null -eq $val) { $val = "" }
                $h += "<td>$([System.Web.HttpUtility]::HtmlEncode($val.ToString()))</td>"
            }
            $h += "</tr>"
        }
        $h += "</tbody></table>"
        return $h
    }
}

# ── MODULE CHECK ──────────────────────────────────────────────────────────────
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
    Import-Module $mod -ErrorAction SilentlyContinue
    if (-not (Get-Module -Name $mod)) {
        Write-Warning "  Module '$mod' not loaded."
        $yn = Read-Host "  Install via PSResourceGet? (Y/N)"
        if ($yn -match "^[Yy]$") {
            try {
                Install-PSResource $mod -Scope CurrentUser -TrustRepository -ErrorAction Stop
                Import-Module $mod -ErrorAction Stop
                Write-Host "  Installed $mod" -ForegroundColor Green
            } catch { Write-Warning "  Could not install $mod — $_" }
        }
    } else {
        Write-Host "  OK — $mod" -ForegroundColor Green
    }
}

# ── SESSION CHECK ─────────────────────────────────────────────────────────────
Write-Host "`n[INFO] Checking existing sessions..." -ForegroundColor Cyan
$graphContext = Get-MgContext -ErrorAction SilentlyContinue
$exoConnInfo  = Get-ConnectionInformation -ErrorAction SilentlyContinue | Select-Object -First 1

$graphStatus = if ($graphContext) {
    $tenantMatch = -not $EffectiveTenantId -or ($graphContext.TenantId -eq $EffectiveTenantId)
    $mismatch    = if (-not $tenantMatch) { "  *** DIFFERENT TENANT THAN SELECTED PROFILE ***" } else { "" }
    "  Graph : $($graphContext.Account)  |  Tenant: $($graphContext.TenantId)$mismatch"
} else { "  Graph : Not connected" }

$exoStatus = if ($exoConnInfo.UserPrincipalName) {
    "  EXO   : $($exoConnInfo.UserPrincipalName)"
} else { "  EXO   : Not connected" }

Write-Host $graphStatus -ForegroundColor $(if ($graphContext -and ($EffectiveTenantId -and $graphContext.TenantId -ne $EffectiveTenantId)) { 'Yellow' } else { 'Cyan' })
Write-Host $exoStatus   -ForegroundColor $(if ($exoConnInfo) { 'Cyan' } else { 'Gray' })

if ($graphContext -or $exoConnInfo) {
    $yn = Read-Host "`n  Disconnect existing session(s) before continuing? (Y/N)"
    if ($yn -match "^[Yy]$") {
        # EXO disconnect is intentionally skipped — Disconnect-ExchangeOnline crashes the process
        # in ExchangeOnlineManagement 3.9.x (MSAL WAM broker bug). EXO REST sessions expire automatically.
        if ($graphContext) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null; $graphContext = $null }
        Write-Host "  Graph disconnected. EXO session will expire automatically (see known EXO 3.9.x bug)." -ForegroundColor Yellow
    }
}
Write-Host ""

# ── GRAPH CONNECTION ──────────────────────────────────────────────────────────
$requiredScopes = @(
    "User.Read.All",
    "Directory.Read.All",
    "Policy.Read.All",
    "AuditLog.Read.All",
    "UserAuthenticationMethod.Read.All",
    "SecurityEvents.Read.All",
    "DeviceManagementManagedDevices.Read.All"
)

if (-not $graphContext) {
    Write-Host "[INFO] Not connected to Microsoft Graph." -ForegroundColor Yellow
    $yn = Read-Host "  Connect now? (Y/N)"
    if ($yn -match "^[Yy]$") {
        if (-not $EffectiveTenantId) { $EffectiveTenantId = Read-Host "  Enter Tenant ID or domain" }
        try {
            Connect-MgGraph -Scopes $requiredScopes -TenantId $EffectiveTenantId -ErrorAction Stop
        } catch {
            if ($_ -match "NullReferenceException|RuntimeBroker|window handle") {
                Connect-MgGraph -Scopes $requiredScopes -TenantId $EffectiveTenantId -NoWAM -ErrorAction Stop
            } else { throw }
        }
        Write-Host "  Connected." -ForegroundColor Green
    } else { throw "Microsoft Graph connection required." }
} else {
    Write-Host "[INFO] Using existing Graph session: $($graphContext.Account)" -ForegroundColor Green
}

# ── EXO CONNECTION ────────────────────────────────────────────────────────────
$exoConnected = $false
if (-not $SkipExchangeOnline) {
    $exoConnInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($exoConnInfo) {
        $exoConnected = $true
        Write-Host "[INFO] Using existing EXO session: $($exoConnInfo.UserPrincipalName)" -ForegroundColor Green
    } else {
        Write-Host "[INFO] Not connected to Exchange Online." -ForegroundColor Yellow
        $yn = Read-Host "  Connect now? (Y/N)"
        if ($yn -match "^[Yy]$") {
            try {
                Connect-ExchangeOnline -Device -ShowBanner:$false -ErrorAction Stop
                $null = Get-EXOMailbox -ResultSize 1 -ErrorAction Stop
                $exoConnected = $true
                Write-Host "  Connected." -ForegroundColor Green
            } catch { Write-Warning "  EXO connection failed — EXO section will be empty." }
        }
    }
}

Write-Host ""

# ── LICENSE CATALOG ───────────────────────────────────────────────────────────
Write-Host "[INFO] Loading license catalog..." -ForegroundColor Cyan
$licenseCatalog = @{}
try {
    $catalogUrl = "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv"
    $tempFile   = Join-Path $env:TEMP "license-catalog.csv"
    Invoke-WebRequest -Uri $catalogUrl -OutFile $tempFile -UseBasicParsing -ErrorAction Stop
    Import-Csv -Path $tempFile | ForEach-Object {
        if ($_.String_Id -and $_.Product_Display_Name -and -not $licenseCatalog.ContainsKey($_.String_Id)) {
            $licenseCatalog[$_.String_Id] = $_.Product_Display_Name
        }
    }
    Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
    Write-Host "  Loaded $($licenseCatalog.Count) license names." -ForegroundColor Green
} catch {
    Write-Warning "  Could not load license catalog — SKU part numbers will be used as display names."
}

# ── DATA COLLECTION ───────────────────────────────────────────────────────────

Write-Host "[INFO] Collecting tenant information..." -ForegroundColor Cyan
$org           = Get-MgOrganization
$primaryDomain = ($org.VerifiedDomains | Where-Object IsDefault).Name

Write-Host "[INFO] Collecting domains..." -ForegroundColor Cyan
$domains = Get-MgDomain | Select-Object Id, IsDefault, IsVerified, AuthenticationType
$tenantDomainNames = @($domains | Where-Object IsVerified | Select-Object -ExpandProperty Id)

Write-Host "[INFO] Collecting license SKUs..." -ForegroundColor Cyan
$skus = Get-MgSubscribedSku -All

# Subscription lifecycle dates (beta)
Write-Host "[INFO] Collecting subscription renewal dates..." -ForegroundColor Cyan
$subscriptions  = @()
$subLookup      = @{}
$subDatesNote   = ""
try {
    $subResp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/directory/subscriptions" -ErrorAction Stop
    $subscriptions = $subResp.value
    foreach ($sub in $subscriptions) {
        if ($sub.skuPartNumber) { $subLookup[$sub.skuPartNumber] = $sub }
    }
    Write-Host "  Retrieved $($subscriptions.Count) subscription record(s)." -ForegroundColor Green
} catch {
    $subDatesNote = "Renewal dates unavailable — $($_.Exception.Message)"
    Write-Warning "  $subDatesNote"
}

Write-Host "[INFO] Collecting users..." -ForegroundColor Cyan
$userPropsBase   = "Id,DisplayName,UserPrincipalName,AccountEnabled,AssignedLicenses,UserType,CreatedDateTime,LastPasswordChangeDateTime,Mail,JobTitle,Department"
$signInAvailable = $false
try {
    $users = Get-MgUser -All -Property ($userPropsBase + ",SignInActivity") -ErrorAction Stop
    $signInAvailable = $true
    Write-Host "  Sign-in activity data included." -ForegroundColor Green
} catch {
    if ($_ -match "Authentication_RequestFromNonPremiumTenantOrB2CTenant|Forbidden|403") {
        Write-Warning "  SignInActivity not available (requires Entra ID P1/P2). Sign-in dates will show as N/A."
        $users = Get-MgUser -All -Property $userPropsBase -ErrorAction Stop
    } else { throw }
}

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
                IsExternal        = -not ($tenantDomainNames | Where-Object { $u.UserPrincipalName -match "@$([regex]::Escape($_))$" })
            }
        }
    }
}

Write-Host "[INFO] Collecting groups..." -ForegroundColor Cyan
$groups = Get-MgGroup -All -Property DisplayName,GroupTypes,SecurityEnabled,MailEnabled,MembershipRule,CreatedDateTime

Write-Host "[INFO] Collecting Conditional Access policies..." -ForegroundColor Cyan
$caPolicies = @()
$caError    = $null
try {
    $caPolicies = Get-MgIdentityConditionalAccessPolicy -All -ErrorAction Stop
} catch {
    $caError = "Could not retrieve CA policies — $($_.Exception.Message)"
}

Write-Host "[INFO] Collecting MFA registration data..." -ForegroundColor Cyan
$mfaData  = @{}
$mfaError = $null
try {
    $mfaUri  = "https://graph.microsoft.com/beta/reports/authenticationMethods/userRegistrationDetails"
    $mfaResp = Invoke-MgGraphRequest -Method GET -Uri $mfaUri -ErrorAction Stop
    do {
        foreach ($entry in $mfaResp.value) { $mfaData[$entry.id] = $entry }
        $mfaResp = if ($mfaResp.'@odata.nextLink') {
            Invoke-MgGraphRequest -Method GET -Uri $mfaResp.'@odata.nextLink' -ErrorAction Stop
        } else { $null }
    } while ($mfaResp)
    Write-Host "  MFA data for $($mfaData.Count) users." -ForegroundColor Green
} catch {
    $mfaError = "MFA data unavailable — $($_.Exception.Message)"
    Write-Warning "  $mfaError"
}

Write-Host "[INFO] Collecting Secure Score..." -ForegroundColor Cyan
$secureScore      = $null
$secureScoreError = $null
try {
    $ssResp      = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/security/secureScores?`$top=1" -ErrorAction Stop
    $secureScore = $ssResp.value | Select-Object -First 1
} catch {
    $secureScoreError = "Secure Score unavailable — $($_.Exception.Message)"
    Write-Warning "  $secureScoreError"
}

Write-Host "[INFO] Collecting device information..." -ForegroundColor Cyan
$aadDevices     = @()
$aadDeviceError = $null
try {
    $aadDevices = Get-MgDevice -All -Property DisplayName,OperatingSystem,TrustType -ErrorAction Stop
} catch {
    $aadDeviceError = "Azure AD device data unavailable — $($_.Exception.Message)"
    Write-Warning "  $aadDeviceError"
}

$intuneDevices = @()
$intuneError   = $null
try {
    $intuneResp    = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$select=id,deviceName,operatingSystem,complianceState" -ErrorAction Stop
    $intuneDevices = $intuneResp.value
    Write-Host "  $($intuneDevices.Count) Intune device(s)." -ForegroundColor Green
} catch {
    $intuneError = "Intune data unavailable (no Intune license?) — $($_.Exception.Message)"
    Write-Warning "  $intuneError"
}

Write-Host "[INFO] Collecting Exchange Online data..." -ForegroundColor Cyan
$exoSharedMailboxes = @(); $exoDistGroups = @(); $exoTransportRules = @()
$exoInbound = @(); $exoOutbound = @(); $exoError = $null
if ($exoConnected) {
    try {
        $exoSharedMailboxes = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited |
            Select-Object DisplayName, PrimarySmtpAddress, @{N="HiddenFromGAL";E={$_.HiddenFromAddressListsEnabled}}
        $exoDistGroups      = Get-DistributionGroup -ResultSize Unlimited |
            Select-Object DisplayName, PrimarySmtpAddress, GroupType
        $exoInbound         = Get-InboundConnector  |
            Select-Object Name, Enabled, ConnectorType, @{N="SenderDomains";E={$_.SenderDomains -join ", "}}
        $exoOutbound        = Get-OutboundConnector |
            Select-Object Name, Enabled, ConnectorType, @{N="RecipientDomains";E={$_.RecipientDomains -join ", "}}
        $exoTransportRules  = Get-TransportRule | Select-Object Name, State, Priority
    } catch {
        $exoError = "Exchange Online error — $($_.Exception.Message)"
    }
} else {
    $exoError = "Exchange Online not connected. Rerun without -SkipExchangeOnline to include EXO data."
}

# ── PROCESS USERS ─────────────────────────────────────────────────────────────
Write-Host "[INFO] Processing user data..." -ForegroundColor Cyan
$today = Get-Date

$processedUsers = foreach ($user in $users) {
    $lastSignIn      = $null
    $daysSinceSignIn = $null
    $neverSignedIn   = $false

    if ($signInAvailable) {
        if ($user.SignInActivity -and $user.SignInActivity.LastSignInDateTime) {
            $lastSignIn      = [datetime]$user.SignInActivity.LastSignInDateTime
            $daysSinceSignIn = [int](($today - $lastSignIn).TotalDays)
        } else {
            $neverSignedIn = $true
        }
    }

    $isLicensed = $user.AssignedLicenses.Count -gt 0
    $isInactive = $signInAvailable -and (-not $neverSignedIn) -and ($daysSinceSignIn -ge $InactiveThresholdDays)

    $passwordAge = $null
    if ($user.LastPasswordChangeDateTime) {
        $passwordAge = [int](($today - [datetime]$user.LastPasswordChangeDateTime).TotalDays)
    }

    $userLicenseNames = @()
    foreach ($lic in $user.AssignedLicenses) {
        $matchedSku = $skus | Where-Object { $_.SkuId -eq $lic.SkuId }
        if ($matchedSku) {
            $name = if ($licenseCatalog.ContainsKey($matchedSku.SkuPartNumber)) { $licenseCatalog[$matchedSku.SkuPartNumber] } else { $matchedSku.SkuPartNumber }
            $userLicenseNames += $name
        }
    }

    $mfaEntry    = if ($mfaData.ContainsKey($user.Id)) { $mfaData[$user.Id] } else { $null }
    $mfaRegistered = if ($mfaEntry) { [bool]$mfaEntry.isMfaRegistered } else { $null }

    [PSCustomObject]@{
        Id              = $user.Id
        DisplayName     = $user.DisplayName
        UPN             = $user.UserPrincipalName
        Enabled         = $user.AccountEnabled
        UserType        = $user.UserType
        IsLicensed      = $isLicensed
        LicenseNames    = $userLicenseNames
        LicenseCount    = $user.AssignedLicenses.Count
        LastSignIn      = $lastSignIn
        DaysSinceSignIn = $daysSinceSignIn
        NeverSignedIn   = $neverSignedIn
        IsInactive      = $isInactive
        PasswordAge     = $passwordAge
        CreatedDate     = if ($user.CreatedDateTime) { ([datetime]$user.CreatedDateTime).ToString("yyyy-MM-dd") } else { "" }
        MfaRegistered   = $mfaRegistered
    }
}

$userStats = [PSCustomObject]@{
    Total         = $processedUsers.Count
    Enabled       = ($processedUsers | Where-Object { $_.Enabled }).Count
    Disabled      = ($processedUsers | Where-Object { -not $_.Enabled }).Count
    Licensed      = ($processedUsers | Where-Object { $_.IsLicensed }).Count
    Guests        = ($processedUsers | Where-Object { $_.UserType -eq "Guest" }).Count
    Inactive      = if ($signInAvailable) { ($processedUsers | Where-Object { $_.IsLicensed -and $_.IsInactive }).Count } else { $null }
    NeverSignedIn = if ($signInAvailable) { ($processedUsers | Where-Object { $_.IsLicensed -and $_.NeverSignedIn -and $_.Enabled }).Count } else { $null }
    NoMfa         = if ($mfaData.Count -gt 0) { ($processedUsers | Where-Object { $_.IsLicensed -and $_.Enabled -and $_.MfaRegistered -eq $false }).Count } else { $null }
}

# ── KEY FINDINGS ──────────────────────────────────────────────────────────────
$findings = [System.Collections.Generic.List[PSCustomObject]]::new()

# External Global Admins
$externalGAs = $roleRows | Where-Object { $_.Role -eq "Global Administrator" -and $_.IsExternal }
if ($externalGAs) {
    $names = ($externalGAs.UserPrincipalName) -join ", "
    $findings.Add([PSCustomObject]@{
        Severity = "HIGH"
        Title    = "External Global Administrator account(s) detected"
        Detail   = "Account(s) with Global Admin role whose UPN domain is not a verified tenant domain: $names"
        Action   = "Remove these accounts if they belong to a previous IT provider."
    })
}

# CA policy status
$enabledCA = @($caPolicies | Where-Object { $_.State -eq "enabled" })
if ($caPolicies.Count -eq 0 -and -not $caError) {
    $findings.Add([PSCustomObject]@{
        Severity = "HIGH"
        Title    = "No Conditional Access policies configured"
        Detail   = "This tenant has no CA policies. Access is protected only by username and password."
        Action   = "Create baseline CA policies: require MFA for all users, block legacy authentication."
    })
} elseif ($enabledCA.Count -eq 0 -and $caPolicies.Count -gt 0) {
    $roCount = ($caPolicies | Where-Object { $_.State -eq "enabledForReportingButNotEnforced" }).Count
    $findings.Add([PSCustomObject]@{
        Severity = "HIGH"
        Title    = "Conditional Access policies exist but none are enforced"
        Detail   = "$($caPolicies.Count) policy/policies found ($roCount in report-only mode) — none actively enforced."
        Action   = "Enable report-only policies after verifying they won't block legitimate access."
    })
}

# Licensed users with no MFA
if ($userStats.NoMfa -gt 0) {
    $findings.Add([PSCustomObject]@{
        Severity = "HIGH"
        Title    = "$($userStats.NoMfa) enabled licensed user(s) have no MFA registered"
        Detail   = "Users without MFA can sign in with only a password."
        Action   = "Enable a CA policy requiring MFA for all users, or enforce per-user MFA."
    })
}

# Never signed in (licensed) — only if sign-in data is available
if ($signInAvailable -and $userStats.NeverSignedIn -gt 0) {
    $findings.Add([PSCustomObject]@{
        Severity = "MEDIUM"
        Title    = "$($userStats.NeverSignedIn) licensed user(s) have never signed in"
        Detail   = "These accounts consume paid license seats but have never been used."
        Action   = "Review with customer. Remove licenses from confirmed unused accounts."
    })
}

# Inactive licensed users — only if sign-in data is available
if ($signInAvailable -and $userStats.Inactive -gt 0) {
    $findings.Add([PSCustomObject]@{
        Severity = "MEDIUM"
        Title    = "$($userStats.Inactive) licensed user(s) inactive for over $InactiveThresholdDays days"
        Detail   = "Active paid licenses assigned to accounts with no recent sign-in activity."
        Action   = "Confirm with customer whether accounts are still needed. Remove licenses from leavers."
    })
}

# Subscriptions expired or expiring within 60 days
$expiredSubs  = @()
$expiringSubs = @()
foreach ($sub in $subscriptions) {
    if ($sub.nextLifecycleDateTime -and $sub.status -ne "Deleted") {
        $expDate = [datetime]$sub.nextLifecycleDateTime
        if ($expDate -lt $today) {
            $expiredSubs  += "$($sub.skuPartNumber) — expired $($expDate.ToString('yyyy-MM-dd'))"
        } elseif ($expDate -lt $today.AddDays(60)) {
            $expiringSubs += "$($sub.skuPartNumber) — $($expDate.ToString('yyyy-MM-dd'))"
        }
    }
}
if ($expiredSubs.Count -gt 0) {
    $findings.Add([PSCustomObject]@{
        Severity = "HIGH"
        Title    = "$($expiredSubs.Count) subscription(s) have EXPIRED"
        Detail   = ($expiredSubs -join " | ")
        Action   = "Urgently verify renewal status with customer and CSP partner."
    })
}
if ($expiringSubs.Count -gt 0) {
    $findings.Add([PSCustomObject]@{
        Severity = "MEDIUM"
        Title    = "$($expiringSubs.Count) subscription(s) expiring within 60 days"
        Detail   = ($expiringSubs -join " | ")
        Action   = "Contact customer about renewal before the expiry date."
    })
}

# Unused paid seats (exclude free/viral/auto-provisioned plans with 10k+ seats)
$unusedSkus = @($skus | Where-Object {
    $_.PrepaidUnits.Enabled -gt 0 -and
    $_.PrepaidUnits.Enabled -lt 10000 -and
    $_.ConsumedUnits -lt $_.PrepaidUnits.Enabled -and
    $_.SkuPartNumber -notmatch "FREE|VIRAL|EXPLORATORY"
})
if ($unusedSkus.Count -gt 0) {
    $details = $unusedSkus | ForEach-Object {
        $n = if ($licenseCatalog.ContainsKey($_.SkuPartNumber)) { $licenseCatalog[$_.SkuPartNumber] } else { $_.SkuPartNumber }
        "$n ($($_.PrepaidUnits.Enabled - $_.ConsumedUnits) unused)"
    }
    $findings.Add([PSCustomObject]@{
        Severity = "LOW"
        Title    = "$($unusedSkus.Count) SKU(s) with unused paid license seats"
        Detail   = ($details -join " | ")
        Action   = "Consider downsizing at next renewal to match actual usage."
    })
}

# Technical notification email check
if ($org.TechnicalNotificationMails) {
    $extMails = @($org.TechnicalNotificationMails | Where-Object { $_ -notmatch "@$([regex]::Escape($primaryDomain))$" })
    if ($extMails.Count -gt 0) {
        $findings.Add([PSCustomObject]@{
            Severity = "LOW"
            Title    = "Technical notification email set to external domain"
            Detail   = "Notifications currently sent to: $($extMails -join ', ')"
            Action   = "Update technical notification email in M365 Admin Center to an Iver address."
        })
    }
}

# ── SAVE SOURCE DATA ──────────────────────────────────────────────────────────
$sourceTimestamp = Get-Date -Format "yyyyMMdd_HHmm"
$sourcePath      = Join-Path $OutputPath "$shortName.source\$sourceTimestamp"
if (-not (Test-Path $sourcePath)) {
    New-Item -ItemType Directory -Path $sourcePath -Force | Out-Null
}
Write-Host "[INFO] Saving raw source data to $sourcePath ..." -ForegroundColor Cyan

$sourceFiles = [ordered]@{
    "org.json"                     = $org
    "domains.json"                 = $domains
    "skus.json"                    = $skus
    "subscriptions.json"           = $subscriptions
    "users.json"                   = $processedUsers
    "roleAssignments.json"         = $roleRows
    "groups.json"                  = $groups
    "caPolicies.json"              = $caPolicies
    "mfaRegistration.json"         = @($mfaData.Values)
    "secureScore.json"             = $secureScore
    "aadDevices.json"              = $aadDevices
    "intuneDevices.json"           = $intuneDevices
    "exo_sharedMailboxes.json"     = $exoSharedMailboxes
    "exo_distributionGroups.json"  = $exoDistGroups
    "exo_inboundConnectors.json"   = $exoInbound
    "exo_outboundConnectors.json"  = $exoOutbound
    "exo_transportRules.json"      = $exoTransportRules
}

$savedCount = 0
foreach ($entry in $sourceFiles.GetEnumerator()) {
    try {
        $json = $entry.Value | Select-Object * |
            ConvertTo-Json -Depth 10 -EnumsAsStrings -WarningAction SilentlyContinue
        if (-not $json) { $json = "[]" }
        Set-Content -Path (Join-Path $sourcePath $entry.Key) -Encoding UTF8 -Value $json -ErrorAction Stop
        $savedCount++
    } catch {
        Write-Warning "  Could not save $($entry.Key) — $_"
    }
}
Write-Host "  Source data saved ($savedCount / $($sourceFiles.Count) files)." -ForegroundColor Green

# ── BUILD HTML SECTIONS ───────────────────────────────────────────────────────

# Key Findings
$severityOrder = @{ HIGH = 0; MEDIUM = 1; LOW = 2 }
if ($findings.Count -eq 0) {
    $findingsHtml = "<div class='finding finding-low'><div class='finding-head'><span class='badge badge-low'>ALL CLEAR</span><strong>No significant issues detected</strong></div><p>No automated issues found. Manual review is still recommended.</p></div>"
} else {
    $findingsHtml = ""
    foreach ($f in ($findings | Sort-Object { $severityOrder[$_.Severity] })) {
        $cls = "finding-$($f.Severity.ToLower())"
        $bdg = "badge-$($f.Severity.ToLower())"
        $findingsHtml += "<div class='finding $cls'>"
        $findingsHtml += "<div class='finding-head'><span class='badge $bdg'>$($f.Severity)</span><strong>$([System.Web.HttpUtility]::HtmlEncode($f.Title))</strong></div>"
        $findingsHtml += "<p>$([System.Web.HttpUtility]::HtmlEncode($f.Detail))</p>"
        $findingsHtml += "<p class='action-line'>&#128161; $([System.Web.HttpUtility]::HtmlEncode($f.Action))</p>"
        $findingsHtml += "</div>"
    }
}

# Secure Score card
$secureScoreHtml = ""
if ($secureScore) {
    $pct      = [int](($secureScore.currentScore / [Math]::Max($secureScore.maxScore, 1)) * 100)
    $barColor = if ($pct -ge 70) { "#107c10" } elseif ($pct -ge 40) { "#d83b01" } else { "#e81123" }
    $secureScoreHtml = "<div class='score-card'><div class='score-label'>Secure Score</div><div class='score-num'>$([int]$secureScore.currentScore)<span class='score-max'> / $([int]$secureScore.maxScore)</span></div><div class='score-bar-bg'><div class='score-bar' style='width:$pct%;background:$barColor'></div></div><div class='score-pct'>$pct%</div></div>"
} elseif ($secureScoreError) {
    $secureScoreHtml = "<div class='score-card'><div class='score-label'>Secure Score</div><p class='na'>Unavailable</p></div>"
}

# Tenant info
$techMails       = if ($org.TechnicalNotificationMails) { $org.TechnicalNotificationMails -join ", " } else { "Not configured" }
$tenantInfoHtml  = "<div class='info-grid'>"
$tenantInfoHtml += "<div><span class='label'>Tenant Name</span><span class='value'>$([System.Web.HttpUtility]::HtmlEncode($org.DisplayName))</span></div>"
$tenantInfoHtml += "<div><span class='label'>Tenant ID</span><span class='value'>$([System.Web.HttpUtility]::HtmlEncode($org.Id))</span></div>"
$tenantInfoHtml += "<div><span class='label'>Primary Domain</span><span class='value'>$([System.Web.HttpUtility]::HtmlEncode($primaryDomain))</span></div>"
$tenantInfoHtml += "<div><span class='label'>Country</span><span class='value'>$([System.Web.HttpUtility]::HtmlEncode($org.CountryLetterCode))</span></div>"
$tenantInfoHtml += "<div><span class='label'>Tenant Created</span><span class='value'>$(if ($org.CreatedDateTime) { ([datetime]$org.CreatedDateTime).ToString('yyyy-MM-dd') })</span></div>"
$tenantInfoHtml += "<div><span class='label'>Tech Notifications</span><span class='value'>$([System.Web.HttpUtility]::HtmlEncode($techMails))</span></div>"
$tenantInfoHtml += "</div>"

# Domains table
$domainsHtml = $domains | ConvertTo-HtmlTable

# License / Subscription table
$subRowsHtml = ""
foreach ($sku in ($skus | Sort-Object SkuPartNumber)) {
    $friendly  = if ($licenseCatalog.ContainsKey($sku.SkuPartNumber)) { $licenseCatalog[$sku.SkuPartNumber] } else { $sku.SkuPartNumber }
    $purchased = $sku.PrepaidUnits.Enabled
    $consumed  = $sku.ConsumedUnits
    $available = $purchased - $consumed
    $availStyle = if ($available -gt 0 -and $purchased -gt 0) { "style='color:#d83b01;font-weight:600'" } else { "" }

    $sub        = $subLookup[$sku.SkuPartNumber]
    $renewDate  = "—"
    $renewStyle = ""
    $subStatus  = if ($sub) { $sub.status } else { "—" }
    $isTrial    = if ($sub -and $sub.isTrial) { " <span style='background:#8764b8;color:white;font-size:10px;padding:2px 6px;border-radius:10px'>Trial</span>" } else { "" }

    if ($sub -and $sub.nextLifecycleDateTime) {
        $expDate   = [datetime]$sub.nextLifecycleDateTime
        $renewDate = $expDate.ToString("yyyy-MM-dd")
        if ($expDate -lt $today.AddDays(60))  { $renewStyle = "style='color:#e81123;font-weight:600'" }
        elseif ($expDate -lt $today.AddDays(90)) { $renewStyle = "style='color:#d83b01;font-weight:600'" }
    }

    $statusColor = switch ($subStatus) {
        "Enabled"   { "#107c10" }
        "Warning"   { "#d83b01" }
        "Suspended" { "#e81123" }
        default     { "#605e5c" }
    }

    $subRowsHtml += "<tr>"
    $subRowsHtml += "<td><small>$([System.Web.HttpUtility]::HtmlEncode($sku.SkuPartNumber))</small></td>"
    $subRowsHtml += "<td>$([System.Web.HttpUtility]::HtmlEncode($friendly))$isTrial</td>"
    $subRowsHtml += "<td><span style='color:$statusColor;font-weight:600'>$subStatus</span></td>"
    $subRowsHtml += "<td>$purchased</td>"
    $subRowsHtml += "<td>$consumed</td>"
    $subRowsHtml += "<td $availStyle>$available</td>"
    $subRowsHtml += "<td $renewStyle>$renewDate</td>"
    $subRowsHtml += "</tr>`n"
}
$subNoteHtml = if ($subDatesNote) { "<p class='warning' style='margin-top:10px'>&#9888; $([System.Web.HttpUtility]::HtmlEncode($subDatesNote))</p>" } else { "" }

# User stat cards
$inactiveVal   = if ($null -ne $userStats.Inactive)      { $userStats.Inactive }      else { "N/A" }
$neverVal      = if ($null -ne $userStats.NeverSignedIn)  { $userStats.NeverSignedIn }  else { "N/A" }
$inactiveColor = if ($null -ne $userStats.Inactive -and $userStats.Inactive -gt 0) { "#d83b01" } else { "#797775" }
$neverColor    = if ($null -ne $userStats.NeverSignedIn -and $userStats.NeverSignedIn -gt 0) { "#e81123" } else { "#797775" }
$noMfaCard     = if ($null -ne $userStats.NoMfa) { New-StatCard "No MFA" $userStats.NoMfa "#e81123" } else { "" }
$userCardsHtml = "<div class='stat-grid'>" +
    (New-StatCard "Total Users"                      $userStats.Total    "#0078d4") +
    (New-StatCard "Enabled"                          $userStats.Enabled  "#107c10") +
    (New-StatCard "Disabled"                         $userStats.Disabled "#797775") +
    (New-StatCard "Licensed"                         $userStats.Licensed "#0078d4") +
    (New-StatCard "Guests"                           $userStats.Guests   "#d83b01") +
    (New-StatCard "Inactive >$InactiveThresholdDays d" $inactiveVal      $inactiveColor) +
    (New-StatCard "Never Signed In"                  $neverVal           $neverColor) +
    $noMfaCard + "</div>"

# User table rows
$userRowsHtml = ""
foreach ($u in ($processedUsers | Sort-Object DisplayName)) {
    $rowClass = ""
    if     ($u.NeverSignedIn -and $u.IsLicensed -and $u.Enabled) { $rowClass = "row-never" }
    elseif ($u.IsInactive -and $u.IsLicensed)                    { $rowClass = "row-inactive" }
    elseif (-not $u.Enabled)                                      { $rowClass = "row-disabled" }

    # Sign-in cell
    $siHtml = if ($u.NeverSignedIn) { "<span class='badge-never'>Never</span>" }
              elseif ($u.LastSignIn) { "$($u.LastSignIn.ToString('yyyy-MM-dd'))<br><small>$($u.DaysSinceSignIn)d ago</small>" }
              else { "<span class='na'>N/A</span>" }

    # Password age cell
    $pwHtml = if ($null -eq $u.PasswordAge) { "<span class='na'>N/A</span>" }
              elseif ($u.PasswordAge -gt 365) { "<span style='color:#e81123;font-weight:600'>$($u.PasswordAge)d</span>" }
              elseif ($u.PasswordAge -gt 180) { "<span style='color:#d83b01;font-weight:600'>$($u.PasswordAge)d</span>" }
              else { "$($u.PasswordAge)d" }

    # MFA cell
    $mfaHtml = if ($null -eq $u.MfaRegistered) { "<span class='na'>N/A</span>" }
               elseif ($u.MfaRegistered) { "<span class='badge-yes'>Yes</span>" }
               else { "<span class='badge-no'>No</span>" }

    # License pills
    $pills = ($u.LicenseNames | Sort-Object -Unique | ForEach-Object { "<span class='pill'>$([System.Web.HttpUtility]::HtmlEncode($_))</span>" }) -join " "
    if (-not $pills) { $pills = "<span class='na'>—</span>" }

    $enabledHtml = if ($u.Enabled) { "<span class='ok'>Enabled</span>" } else { "<span class='bad'>Disabled</span>" }
    $typeHtml    = if ($u.UserType -eq "Guest") { "<span class='badge-guest'>Guest</span>" } else { "Member" }

    $userRowsHtml += "<tr class='$rowClass'>"
    $userRowsHtml += "<td>$([System.Web.HttpUtility]::HtmlEncode($u.DisplayName))</td>"
    $userRowsHtml += "<td><small>$([System.Web.HttpUtility]::HtmlEncode($u.UPN))</small></td>"
    $userRowsHtml += "<td>$enabledHtml</td>"
    $userRowsHtml += "<td>$typeHtml</td>"
    $userRowsHtml += "<td>$pills</td>"
    $userRowsHtml += "<td>$mfaHtml</td>"
    $userRowsHtml += "<td>$siHtml</td>"
    $userRowsHtml += "<td>$pwHtml</td>"
    $userRowsHtml += "<td><small>$($u.CreatedDate)</small></td>"
    $userRowsHtml += "</tr>`n"
}

# Admin table
$adminRowsHtml = ""
foreach ($r in ($roleRows | Sort-Object Role, DisplayName)) {
    $extBadge = if ($r.IsExternal) { " <span class='badge-ext'>EXTERNAL</span>" } else { "" }
    $rowCls   = if ($r.IsExternal) { "class='row-never'" } else { "" }
    $adminRowsHtml += "<tr $rowCls>"
    $adminRowsHtml += "<td>$([System.Web.HttpUtility]::HtmlEncode($r.Role))</td>"
    $adminRowsHtml += "<td>$([System.Web.HttpUtility]::HtmlEncode($r.DisplayName))</td>"
    $adminRowsHtml += "<td>$([System.Web.HttpUtility]::HtmlEncode($r.UserPrincipalName))$extBadge</td>"
    $adminRowsHtml += "</tr>`n"
}
$adminHtml = if ($adminRowsHtml) {
    "<table><thead><tr><th>Role</th><th>Display Name</th><th>UPN</th></tr></thead><tbody>$adminRowsHtml</tbody></table>"
} else { "<p class='empty'>No role assignments found.</p>" }

# Groups
$grpStats = [PSCustomObject]@{
    Total      = $groups.Count
    M365Groups = ($groups | Where-Object { $_.GroupTypes -contains "Unified" }).Count
    Security   = ($groups | Where-Object { $_.SecurityEnabled -and $_.GroupTypes -notcontains "Unified" }).Count
    Dynamic    = ($groups | Where-Object { $_.GroupTypes -contains "DynamicMembership" }).Count
}
$groupCardsHtml = "<div class='stat-grid'>" +
    (New-StatCard "Total Groups" $grpStats.Total      "#0078d4") +
    (New-StatCard "M365 Groups"  $grpStats.M365Groups "#107c10") +
    (New-StatCard "Security"     $grpStats.Security   "#8764b8") +
    (New-StatCard "Dynamic"      $grpStats.Dynamic    "#d83b01") + "</div>"

$groupTableHtml = $groups | Select-Object DisplayName,
    @{N="Type";E={
        if ($_.GroupTypes -contains "Unified") { "Microsoft 365" }
        elseif ($_.GroupTypes -contains "DynamicMembership") { "Dynamic Security" }
        elseif ($_.SecurityEnabled) { "Security" }
        else { "Distribution" }
    }},
    @{N="Mail Enabled";E={ $_.MailEnabled }},
    @{N="Dynamic Rule";E={ if ($_.MembershipRule) { "Yes" } else { "No" } }},
    @{N="Created";E={ if ($_.CreatedDateTime) { ([datetime]$_.CreatedDateTime).ToString("yyyy-MM-dd") } else { "" } }} |
    Sort-Object DisplayName | ConvertTo-HtmlTable

# CA policies
$caHtml = if ($caError) { "<p class='warning'>&#9888; $([System.Web.HttpUtility]::HtmlEncode($caError))</p>" }
          else {
              $caPolicies | Select-Object DisplayName, State,
                  @{N="Created";E={ if ($_.CreatedDateTime) { ([datetime]$_.CreatedDateTime).ToString("yyyy-MM-dd") } else { "" } }},
                  @{N="Modified";E={ if ($_.ModifiedDateTime) { ([datetime]$_.ModifiedDateTime).ToString("yyyy-MM-dd") } else { "" } }} |
                  Sort-Object State, DisplayName | ConvertTo-HtmlTable
          }

# Devices
$aadByOs = $aadDevices | Group-Object OperatingSystem | Sort-Object Count -Descending
$devCardsHtml = "<div class='stat-grid'>" +
    (New-StatCard "Azure AD Registered" $aadDevices.Count    "#0078d4") +
    (New-StatCard "Intune Enrolled"     $intuneDevices.Count "#107c10") + "</div>"

$devOsHtml = ""
if ($aadDeviceError) {
    $devOsHtml = "<p class='warning'>&#9888; $([System.Web.HttpUtility]::HtmlEncode($aadDeviceError))</p>"
} elseif ($aadByOs) {
    $devOsHtml = "<h3>Azure AD Device OS Breakdown</h3><table><thead><tr><th>Operating System</th><th>Count</th></tr></thead><tbody>"
    foreach ($g in $aadByOs) {
        $devOsHtml += "<tr><td>$([System.Web.HttpUtility]::HtmlEncode($g.Name))</td><td>$($g.Count)</td></tr>"
    }
    $devOsHtml += "</tbody></table>"
} else {
    $devOsHtml = "<p class='empty'>No Azure AD device records found.</p>"
}
$devIntuneNote = if ($intuneError) { "<p class='warning' style='margin-top:8px'>&#9888; $([System.Web.HttpUtility]::HtmlEncode($intuneError))</p>" } else { "" }
$devicesHtml   = $devCardsHtml + $devOsHtml + $devIntuneNote

# EXO
$exoHtml = if ($exoError) { "<p class='warning'>&#9888; $([System.Web.HttpUtility]::HtmlEncode($exoError))</p>" }
           else {
               $exoCards = "<div class='stat-grid'>" +
                   (New-StatCard "Shared Mailboxes"    $exoSharedMailboxes.Count "#0078d4") +
                   (New-StatCard "Distribution Groups" $exoDistGroups.Count      "#107c10") +
                   (New-StatCard "Transport Rules"     $exoTransportRules.Count  "#8764b8") +
                   (New-StatCard "Inbound Connectors"  $exoInbound.Count         "#797775") +
                   (New-StatCard "Outbound Connectors" $exoOutbound.Count        "#797775") + "</div>"
               $exoCards +
               "<h3>Shared Mailboxes</h3>" + ($exoSharedMailboxes | ConvertTo-HtmlTable) +
               "<h3>Distribution Groups</h3>" + ($exoDistGroups | ConvertTo-HtmlTable) +
               "<h3>Transport Rules</h3>" + ($exoTransportRules | ConvertTo-HtmlTable) +
               "<h3>Inbound Connectors</h3>" + ($exoInbound | ConvertTo-HtmlTable) +
               "<h3>Outbound Connectors</h3>" + ($exoOutbound | ConvertTo-HtmlTable)
           }

# ── ASSEMBLE HTML ─────────────────────────────────────────────────────────────
$reportDate = Get-Date -Format "yyyy-MM-dd HH:mm"
$highCount  = ($findings | Where-Object { $_.Severity -eq "HIGH" }).Count
$medCount   = ($findings | Where-Object { $_.Severity -eq "MEDIUM" }).Count
$lowCount   = ($findings | Where-Object { $_.Severity -eq "LOW" }).Count
$mfaNote     = if ($mfaError) { " | MFA data: unavailable" } else { "" }
$mfaNoteHtml = if ($mfaError) { "<div class='warning' style='margin:0 0 16px'>&#9888; MFA registration data could not be retrieved. Reconnect to Microsoft Graph with the <strong>UserAuthenticationMethod.Read.All</strong> scope included to populate the MFA column.</div>" } else { "" }

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>$EffectiveTenantName — M365 Customer Report</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',system-ui,sans-serif;background:#f3f2f1;color:#323130;font-size:14px}
header{background:#0078d4;color:white;padding:24px 40px}
header h1{font-size:22px;font-weight:600}
header p{font-size:13px;opacity:.85;margin-top:4px}
nav{background:white;border-bottom:1px solid #edebe9;padding:0 32px;display:flex;gap:0;position:sticky;top:0;z-index:100;box-shadow:0 1px 3px rgba(0,0,0,.1);overflow-x:auto}
nav a{display:block;padding:13px 14px;font-size:12px;font-weight:600;color:#605e5c;text-decoration:none;border-bottom:3px solid transparent;white-space:nowrap}
nav a:hover{color:#0078d4;border-bottom-color:#0078d4}
main{max-width:1500px;margin:0 auto;padding:24px 40px}
.section{background:white;border-radius:4px;padding:24px;margin-bottom:20px;box-shadow:0 1px 3px rgba(0,0,0,.1)}
.section h2{font-size:16px;font-weight:600;color:#0078d4;margin-bottom:16px;padding-bottom:8px;border-bottom:1px solid #edebe9}
.section h3{font-size:13px;font-weight:600;color:#323130;margin:20px 0 10px}
.stat-grid{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:20px}
.stat-card{background:#faf9f8;border-radius:4px;padding:16px 20px;min-width:130px;flex:1;border-top:4px solid #0078d4}
.stat-value{font-size:28px;font-weight:700;color:#323130}
.stat-label{font-size:12px;color:#605e5c;margin-top:2px}
.info-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:12px;margin-bottom:16px}
.info-grid div{background:#faf9f8;padding:12px 16px;border-radius:4px}
.info-grid .label{display:block;font-size:11px;color:#605e5c;text-transform:uppercase;letter-spacing:.5px}
.info-grid .value{display:block;font-size:14px;font-weight:600;margin-top:2px}
.score-card{background:#faf9f8;border-radius:4px;padding:16px 24px;min-width:180px;border:1px solid #edebe9;align-self:flex-start}
.score-label{font-size:11px;color:#605e5c;text-transform:uppercase;letter-spacing:.5px}
.score-num{font-size:32px;font-weight:700;color:#323130;margin:4px 0 2px}
.score-max{font-size:15px;color:#605e5c;font-weight:400}
.score-bar-bg{background:#edebe9;border-radius:4px;height:8px;margin:6px 0 4px}
.score-bar{height:8px;border-radius:4px}
.score-pct{font-size:12px;color:#605e5c}
.tenant-top{display:flex;gap:20px;align-items:flex-start;flex-wrap:wrap}
table{width:100%;border-collapse:collapse;font-size:13px;margin-top:6px}
thead{background:#f3f2f1}
th{text-align:left;padding:10px 12px;font-weight:600;font-size:11px;text-transform:uppercase;letter-spacing:.4px;color:#605e5c;border-bottom:2px solid #edebe9;cursor:pointer;white-space:nowrap}
th:hover{background:#edebe9}
td{padding:9px 12px;border-bottom:1px solid #f3f2f1;vertical-align:middle}
tr:hover td{background:#f3f2f1}
tr:last-child td{border-bottom:none}
.row-never td{background:#fce4e4!important}
.row-inactive td{background:#fff3cd!important}
.row-disabled td{background:#faf9f8!important;opacity:.7}
.finding{border-radius:4px;padding:16px 20px;margin-bottom:12px;border-left:5px solid #ccc}
.finding-high{background:#fdf0f0;border-left-color:#e81123}
.finding-medium{background:#fffaf0;border-left-color:#d83b01}
.finding-low{background:#f0f7ff;border-left-color:#0078d4}
.finding-head{display:flex;align-items:center;gap:10px;margin-bottom:8px}
.finding p{font-size:13px;color:#323130}
.action-line{margin-top:8px!important;font-size:12px!important;color:#605e5c!important;font-style:italic}
.badge{display:inline-block;padding:3px 10px;border-radius:12px;font-size:11px;font-weight:700;letter-spacing:.5px}
.badge-high{background:#e81123;color:white}
.badge-medium{background:#d83b01;color:white}
.badge-low{background:#0078d4;color:white}
.badge-ext{background:#e81123;color:white;font-size:10px;padding:2px 6px;border-radius:10px;font-weight:700}
.badge-guest{background:#8764b8;color:white;font-size:10px;padding:2px 6px;border-radius:10px}
.badge-never{color:#e81123;font-weight:700}
.badge-yes{background:#107c10;color:white;font-size:11px;padding:2px 8px;border-radius:10px}
.badge-no{background:#e81123;color:white;font-size:11px;padding:2px 8px;border-radius:10px}
.pill{display:inline-block;background:#e1f5ff;color:#0078d4;padding:2px 8px;border-radius:10px;margin:2px;font-size:11px;white-space:nowrap}
.ok{color:#107c10;font-weight:600}
.bad{color:#e81123;font-weight:600}
.na{color:#bbb;font-style:italic}
.empty{color:#605e5c;font-style:italic;padding:12px 0}
.warning{background:#fff4ce;border-left:4px solid #ffb900;padding:12px 16px;border-radius:0 4px 4px 0;color:#323130;margin:6px 0}
.search-bar{margin:12px 0;padding:10px;background:#f8f9fa;border-radius:4px;display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.search-bar input{flex:1;min-width:200px;max-width:360px;padding:7px 12px;border:1px solid #ddd;border-radius:4px;font-size:13px}
.search-bar input:focus{outline:none;border-color:#0078d4}
.filter-btn{padding:5px 12px;border:1px solid #ddd;background:white;cursor:pointer;border-radius:4px;font-size:12px;font-weight:600;transition:all .15s}
.filter-btn:hover{border-color:#0078d4;color:#0078d4}
.filter-btn.active{background:#0078d4;color:white;border-color:#0078d4}
.legend{display:flex;gap:14px;margin:6px 0 10px;font-size:12px;flex-wrap:wrap}
.legend-item{display:flex;align-items:center;gap:6px}
.ldot{width:12px;height:12px;border-radius:2px;flex-shrink:0}
footer{text-align:center;padding:20px;color:#605e5c;font-size:12px}
</style>
<script>
function search(tId,iId){
  const f=document.getElementById(iId).value.toLowerCase();
  document.getElementById(tId).querySelectorAll('tbody tr').forEach(r=>{
    r.style.display=Array.from(r.cells).some(c=>c.textContent.toLowerCase().includes(f))?'':'none';
  });
}
function filter(tId,cls,bId){
  const btn=document.getElementById(bId);
  const wasOn=btn.classList.contains('active');
  document.querySelectorAll('.filter-btn').forEach(b=>b.classList.remove('active'));
  document.getElementById(tId).querySelectorAll('tbody tr').forEach(r=>{
    r.style.display=(wasOn||r.classList.contains(cls))?'':'none';
  });
  if(!wasOn) btn.classList.add('active');
}
function sort(tId,col){
  const t=document.getElementById(tId);
  const asc=t.dataset.sortCol==col&&t.dataset.sortDir=='asc';
  Array.from(t.querySelectorAll('tbody tr'))
    .sort((a,b)=>{
      const va=a.cells[col]?.textContent.trim()||'';
      const vb=b.cells[col]?.textContent.trim()||'';
      return asc?vb.localeCompare(va,undefined,{numeric:true}):va.localeCompare(vb,undefined,{numeric:true});
    })
    .forEach(r=>t.querySelector('tbody').appendChild(r));
  t.dataset.sortCol=col; t.dataset.sortDir=asc?'desc':'asc';
}
</script>
</head>
<body>
<!--[if lt IE 11]>
<div style="background:#e81123;color:white;padding:10px 24px;font-family:sans-serif;font-size:13px">
  This report requires a modern browser (Edge, Chrome, or Firefox). Internet Explorer is not supported.
</div>
<![endif]-->
<header>
  <h1>$EffectiveTenantName &mdash; M365 Customer Report</h1>
  <p>Generated $reportDate &nbsp;&bull;&nbsp; Findings: <strong>$highCount HIGH</strong> &nbsp;&bull;&nbsp; $medCount MEDIUM &nbsp;&bull;&nbsp; $lowCount LOW$mfaNote</p>
</header>
<nav>
  <a href="#findings">&#128680; Key Findings</a>
  <a href="#tenant">&#127970; Tenant</a>
  <a href="#licenses">&#128196; Licenses</a>
  <a href="#users">&#128101; Users</a>
  <a href="#admins">&#128272; Admins</a>
  <a href="#groups">&#128101; Groups</a>
  <a href="#ca">&#128274; Conditional Access</a>
  <a href="#devices">&#128187; Devices</a>
  <a href="#exchange">&#128231; Exchange Online</a>
</nav>
<main>

<div class="section" id="findings">
  <h2>&#128680; Key Findings</h2>
  $findingsHtml
</div>

<div class="section" id="tenant">
  <h2>&#127970; Tenant Overview</h2>
  <div class="tenant-top">
    <div style="flex:1;min-width:300px">$tenantInfoHtml</div>
    $secureScoreHtml
  </div>
  <h3>Registered Domains</h3>
  $domainsHtml
</div>

<div class="section" id="licenses">
  <h2>&#128196; Subscriptions &amp; Licenses</h2>
  <table id="sku-table">
    <thead><tr>
      <th onclick="sort('sku-table',0)">SKU &#8597;</th>
      <th onclick="sort('sku-table',1)">Name &#8597;</th>
      <th onclick="sort('sku-table',2)">Status &#8597;</th>
      <th onclick="sort('sku-table',3)">Purchased &#8597;</th>
      <th onclick="sort('sku-table',4)">Assigned &#8597;</th>
      <th onclick="sort('sku-table',5)">Available &#8597;</th>
      <th onclick="sort('sku-table',6)">Renewal Date &#8597;</th>
    </tr></thead>
    <tbody>$subRowsHtml</tbody>
  </table>
  $subNoteHtml
</div>

<div class="section" id="users">
  <h2>&#128101; Users</h2>
  $userCardsHtml
  $mfaNoteHtml
  <div class="search-bar">
    <input type="text" id="usr-search" onkeyup="search('usr-table','usr-search')" placeholder="Search users...">
    <button class="filter-btn" id="btn-never"    onclick="filter('usr-table','row-never','btn-never')">Never Signed In</button>
    <button class="filter-btn" id="btn-inactive" onclick="filter('usr-table','row-inactive','btn-inactive')">Inactive</button>
    <button class="filter-btn" id="btn-disabled" onclick="filter('usr-table','row-disabled','btn-disabled')">Disabled</button>
  </div>
  <div class="legend">
    <div class="legend-item"><div class="ldot" style="background:#fce4e4;border:1px solid #f5b8b8"></div><span>Never signed in (licensed)</span></div>
    <div class="legend-item"><div class="ldot" style="background:#fff3cd;border:1px solid #f0c000"></div><span>Inactive &gt; $InactiveThresholdDays days</span></div>
    <div class="legend-item"><div class="ldot" style="background:#f0f0f0;border:1px solid #ddd"></div><span>Disabled account</span></div>
  </div>
  <table id="usr-table">
    <thead><tr>
      <th onclick="sort('usr-table',0)">Name &#8597;</th>
      <th onclick="sort('usr-table',1)">UPN &#8597;</th>
      <th onclick="sort('usr-table',2)">Status &#8597;</th>
      <th>Type</th>
      <th>Licenses</th>
      <th onclick="sort('usr-table',5)">MFA &#8597;</th>
      <th onclick="sort('usr-table',6)">Last Sign-In &#8597;</th>
      <th onclick="sort('usr-table',7)">Password Age &#8597;</th>
      <th onclick="sort('usr-table',8)">Created &#8597;</th>
    </tr></thead>
    <tbody>$userRowsHtml</tbody>
  </table>
</div>

<div class="section" id="admins">
  <h2>&#128272; Admin Role Assignments</h2>
  $adminHtml
</div>

<div class="section" id="groups">
  <h2>&#128101; Groups</h2>
  $groupCardsHtml
  $groupTableHtml
</div>

<div class="section" id="ca">
  <h2>&#128274; Conditional Access Policies</h2>
  $caHtml
</div>

<div class="section" id="devices">
  <h2>&#128187; Devices</h2>
  $devicesHtml
</div>

<div class="section" id="exchange">
  <h2>&#128231; Exchange Online</h2>
  $exoHtml
</div>

</main>
<footer>Get-CustomerReport.ps1 v1.3.2 &mdash; Rosvall &amp; Claude &bull; $EffectiveTenantName &bull; $reportDate &bull; Raw data: $shortName.source\$sourceTimestamp\</footer>
</body>
</html>
"@

# ── SAVE ──────────────────────────────────────────────────────────────────────
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}
$fileName   = "${shortName}_CustomerReport_$(Get-Date -Format 'yyyyMMdd_HHmm').html"
$reportFile = Join-Path $OutputPath $fileName
$html | Set-Content -Path $reportFile -Encoding UTF8

# ── UPDATE / CREATE CUSTOMERS.JSON ────────────────────────────────────────────
# Derive suggested values from live collected data
$suggestedAdminUPN = $EffectiveAdminUPN
if (-not $suggestedAdminUPN) {
    $suggestedAdminUPN = ($roleRows | Where-Object {
        $_.Role -eq "Global Administrator" -and -not $_.IsExternal
    } | Select-Object -First 1).UserPrincipalName
}
$primaryLicenseSku = ($skus | Where-Object {
    $_.PrepaidUnits.Enabled -gt 0 -and $_.PrepaidUnits.Enabled -lt 10000 -and
    $_.ConsumedUnits -gt 0 -and $_.SkuPartNumber -notmatch "FREE|VIRAL|EXPLORATORY"
} | Sort-Object ConsumedUnits -Descending | Select-Object -First 1).SkuPartNumber
$techMail = if ($org.TechnicalNotificationMails) { $org.TechnicalNotificationMails[0] } else { "" }

function Update-ProfileFields {
    param($target)
    $fields = @{
        TenantId             = $org.Id
        DisplayName          = $org.DisplayName
        PrimaryDomain        = $primaryDomain
        AllDomains           = @($tenantDomainNames)
        Country              = $org.CountryLetterCode
        TechNotificationMail = $techMail
        TenantCreated        = if ($org.CreatedDateTime) { ([datetime]$org.CreatedDateTime).ToString('yyyy-MM-dd') } else { "" }
        PrimaryLicenseSku    = $primaryLicenseSku
        LastReportDate       = (Get-Date -Format 'yyyy-MM-dd HH:mm')
        LastReportPath       = $reportFile
    }
    foreach ($key in $fields.Keys) {
        if (-not ($target | Get-Member -Name $key -ErrorAction SilentlyContinue)) {
            $target | Add-Member -NotePropertyName $key -NotePropertyValue $null -Force
        }
        $target.$key = $fields[$key]
    }
    return $target
}

$doUpdate = $false
$doCreate = $false

if ($profileData -and (Test-Path $CustomersFile)) {
    if ($UpdateCustomerProfile) {
        $doUpdate = $true
    } else {
        Write-Host ""
        $yn = Read-Host "[PROMPT] Update customers.json for '$($profileData.ShortName)' with live tenant data? (Y/N)"
        $doUpdate = $yn -match "^[Yy]$"
    }
} elseif (-not $profileData) {
    if ($UpdateCustomerProfile) {
        $doCreate = $true
    } else {
        Write-Host ""
        $yn = Read-Host "[PROMPT] No customers.json profile loaded. Create a new entry for this tenant? (Y/N)"
        $doCreate = $yn -match "^[Yy]$"
    }
}

if ($doUpdate -and (Test-Path $CustomersFile)) {
    try {
        $profiles    = Get-Content $CustomersFile -Raw | ConvertFrom-Json
        $profileList = @($profiles)
        $target      = $profileList | Where-Object { $_.ShortName -ieq $profileData.ShortName }
        if ($target) {
            $target = Update-ProfileFields $target
            $profileList | ConvertTo-Json -Depth 10 | Set-Content $CustomersFile -Encoding UTF8
            Write-Host "[INFO] customers.json updated for '$($profileData.ShortName)'." -ForegroundColor Green
        }
    } catch {
        Write-Warning "Could not update customers.json — $_"
    }
}

if ($doCreate) {
    try {
        $newShortName = Read-Host "  Enter a short name for this customer (e.g. 'contoso')"
        if ($newShortName.Trim()) {
            $newEntry = [PSCustomObject]@{
                ShortName            = $newShortName.Trim()
                DisplayName          = $org.DisplayName
                TenantId             = $org.Id
                AdminUPN             = $suggestedAdminUPN
                PrimaryDomain        = $primaryDomain
                AllDomains           = @($tenantDomainNames)
                Country              = $org.CountryLetterCode
                TechNotificationMail = $techMail
                TenantCreated        = if ($org.CreatedDateTime) { ([datetime]$org.CreatedDateTime).ToString('yyyy-MM-dd') } else { "" }
                PrimaryLicenseSku    = $primaryLicenseSku
                LastReportDate       = (Get-Date -Format 'yyyy-MM-dd HH:mm')
                LastReportPath       = $reportFile
                Notes                = ""
            }
            $existingProfiles = @()
            if (Test-Path $CustomersFile) {
                $existingProfiles = @(Get-Content $CustomersFile -Raw | ConvertFrom-Json)
            }
            ($existingProfiles + $newEntry) | ConvertTo-Json -Depth 10 | Set-Content $CustomersFile -Encoding UTF8
            Write-Host "[INFO] New customer profile '$($newShortName.Trim())' added to $CustomersFile." -ForegroundColor Green
        } else {
            Write-Warning "No short name entered — profile not created."
        }
    } catch {
        Write-Warning "Could not create customers.json entry — $_"
    }
}

# Save customers.json snapshot to source folder
if (Test-Path $CustomersFile) {
    try {
        Copy-Item $CustomersFile -Destination (Join-Path $sourcePath "customers.json") -Force -ErrorAction Stop
        Write-Host "[INFO] customers.json snapshot saved to source folder." -ForegroundColor Cyan
    } catch {
        Write-Warning "Could not save customers.json snapshot — $_"
    }
}

Write-Host "`n[SUCCESS] Report saved:      $reportFile" -ForegroundColor Green
Write-Host "[SUCCESS] Raw source data:   $sourcePath" -ForegroundColor Green
Start-Process $reportFile

# ── DISCONNECT ─────────────────────────────────────────────────────────────────
# NOTE: Disconnect-ExchangeOnline is intentionally omitted — it crashes the PowerShell
# process in ExchangeOnlineManagement 3.9.x via an unhandled MSAL WAM broker exception
# in a background thread that cannot be caught. EXO REST sessions expire automatically.
# To manually disconnect EXO safely, run: Disconnect-ExchangeOnline -Confirm:$false
# in a fresh PS window after this script completes.
if ($DisconnectAfter) {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Write-Host "[INFO] Graph disconnected. EXO session expires automatically." -ForegroundColor Cyan
} else {
    $yn = Read-Host "`n[PROMPT] Disconnect from Microsoft Graph? (Y/N)"
    if ($yn -match "^[Yy]$") {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Write-Host "[INFO] Graph disconnected. EXO session expires automatically." -ForegroundColor Cyan
    }
}
