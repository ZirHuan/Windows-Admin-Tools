#Requires -Version 7.2

<#
.SYNOPSIS
    Offboards an M365 user by converting their mailbox to Shared, blocking sign-in,
    and removing their assigned license. Supports reversal for returning employees.

.DESCRIPTION
    Performs a structured, logged, and reversible M365 user offboarding:

    OFFBOARD (default):
      1. Blocks sign-in via Entra ID (Update-MgUser)
      2. Converts the user mailbox to Shared (Set-Mailbox -Type Shared)
      3. Removes the specified license (Set-MgUserLicense)
      4. Outputs a verification object confirming final state

    REVERSE (-Reverse switch):
      1. Assigns the license back
      2. Converts the mailbox back to Regular
      3. Re-enables sign-in
      4. Outputs a verification object confirming restored state

    Supports -WhatIf for dry-run on all destructive operations.
    All actions are logged to a timestamped log file.

.PARAMETER UserPrincipalName
    UPN of the user being offboarded or restored (e.g., robin.lindberg@akerstedts.se).

.PARAMETER Customer
    Short name of a customer profile defined in customers.json (same directory as this script).
    When provided, AdminUPN, TenantId, and LicenseSkuPartNumber are loaded from the profile
    automatically. Explicit parameters always override profile values.

.PARAMETER AdminUPN
    UPN of the admin account used to connect to Exchange Online.
    Optional when -Customer is provided and the profile contains AdminUPN.

.PARAMETER LicenseSkuPartNumber
    SkuPartNumber of the license to remove/restore (e.g., "SPB").
    If not provided, the script will auto-detect all assigned licenses and prompt
    you to confirm which one to act on.

.PARAMETER Reverse
    Switch to reverse the offboarding process (restore user).
    When specified, the script assigns the license back, converts the mailbox
    to Regular, and re-enables sign-in.

.PARAMETER TenantId
    Optional. Entra ID Tenant ID or primary domain. If omitted, the Graph
    connection will use the account's home tenant.

.EXAMPLE
    .\move365user-shared.ps1 -Customer akerstedts -UserPrincipalName "robin.lindberg@akerstedts.se"

    Offboards Robin using the Åkerstedts customer profile (AdminUPN, TenantId, LicenseSku loaded automatically).

.EXAMPLE
    .\move365user-shared.ps1 -Customer akerstedts -UserPrincipalName "robin.lindberg@akerstedts.se" -WhatIf

    Dry run using customer profile — shows what would happen without making any changes.

.EXAMPLE
    .\move365user-shared.ps1 -Customer akerstedts -UserPrincipalName "robin.lindberg@akerstedts.se" `
        -LicenseSkuPartNumber "M365BP"

    Uses the Åkerstedts profile but overrides the default license SKU with M365BP.

.EXAMPLE
    .\move365user-shared.ps1 -UserPrincipalName "robin.lindberg@akerstedts.se" `
        -AdminUPN "adm.chrroses@akerstedts.se" -LicenseSkuPartNumber "SPB"

    Manual mode (no profile) — all values provided explicitly.

.EXAMPLE
    .\move365user-shared.ps1 -Customer akerstedts -UserPrincipalName "robin.lindberg@akerstedts.se" -Reverse

    Restores Robin using the Åkerstedts profile: assigns SPB license back, converts mailbox to Regular, re-enables sign-in.

.NOTES
    Modules required:
      - Microsoft.Graph.Users
      - Microsoft.Graph.Identity.DirectoryManagement
      - ExchangeOnlineManagement

    Minimum Graph scopes (least privilege):
      - User.ReadWrite.All       (block/unblock sign-in, read user)
      - Organization.Read.All   (read subscribed SKUs for license lookup)
      - Directory.Read.All      (verify user state)

    The script does NOT configure mailbox delegation (Send As, Full Access).
    That should be handled separately via a dedicated delegation script.

    Compatibility note:
      Set-MgUserLicense is part of Microsoft.Graph.Users. Ensure the module
      is version 2.x or later (Graph SDK v2). The -AddLicenses parameter
      expects [Microsoft.Graph.PowerShell.Models.IMicrosoftGraphAssignedLicense[]]
      — the hashtable cast @{ SkuId = $guid } is the supported shorthand.
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory, HelpMessage = "UPN of the target user, e.g. robin.lindberg@contoso.com")]
    [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]
    [string]$UserPrincipalName,

    [Parameter(HelpMessage = "Short name of a customer profile in customers.json (e.g. akerstedts)")]
    [string]$Customer,

    [Parameter(HelpMessage = "Admin UPN used to connect to Exchange Online. Optional when -Customer provides it.")]
    [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]
    [string]$AdminUPN,

    [Parameter(HelpMessage = "SKU part number of the license to remove/restore, e.g. SPB. If omitted, auto-detect will run.")]
    [string]$LicenseSkuPartNumber,

    [Parameter(HelpMessage = "Switch to reverse the offboarding (restore the user)")]
    [switch]$Reverse,

    [Parameter(HelpMessage = "Optional: Entra ID Tenant ID or primary domain")]
    [string]$TenantId
)

# ---------------------------------------------------------------------------
# --- Customer profile resolution ---
# ---------------------------------------------------------------------------
# If -Customer is supplied, load the matching profile from customers.json.
# Explicit parameters always win over profile values.

$profileData = $null

if ($Customer) {
    $customersFile = Join-Path $PSScriptRoot "customers.json"

    if (-not (Test-Path $customersFile)) {
        throw "Customer profile file not found: $customersFile. Create it or pass -AdminUPN directly."
    }

    $profiles = Get-Content $customersFile -Raw -ErrorAction Stop | ConvertFrom-Json
    $profileData = $profiles | Where-Object { $_.ShortName -ieq $Customer }

    if (-not $profileData) {
        $available = ($profiles.ShortName) -join ', '
        throw "Customer '$Customer' not found in $customersFile. Available profiles: $available"
    }

    Write-Host "[Profile] Loaded customer: $($profileData.DisplayName)" -ForegroundColor Cyan
}

# Resolve effective values — explicit parameter wins over profile
$EffectiveAdminUPN   = if ($AdminUPN)             { $AdminUPN }             elseif ($profileData.AdminUPN)            { $profileData.AdminUPN }            else { $null }
$EffectiveTenantId   = if ($TenantId)             { $TenantId }             elseif ($profileData.TenantId)            { $profileData.TenantId }            else { $null }
$EffectiveLicenseSku = if ($LicenseSkuPartNumber) { $LicenseSkuPartNumber } elseif ($profileData.DefaultLicenseSku)   { $profileData.DefaultLicenseSku }   else { $null }

if (-not $EffectiveAdminUPN) {
    throw "AdminUPN is required. Provide -AdminUPN or use -Customer with a profile that includes AdminUPN."
}

# ---------------------------------------------------------------------------
# --- Variables ---
# ---------------------------------------------------------------------------

# Customer/environment-specific — adapt these per engagement
$CustomerName  = if ($profileData -and $profileData.DisplayName) { $profileData.DisplayName } else { ($UserPrincipalName -split '@')[1] }
$LogDir        = "C:\Logs\M365Offboard"
$LogPath       = "$LogDir\move365user-shared_$(Get-Date -Format 'yyyyMMdd_HHmm').log"

# Graph scopes — explicit, minimum required
$RequiredGraphScopes = @(
    'User.ReadWrite.All',
    'Organization.Read.All',
    'Directory.Read.All'
)

# ---------------------------------------------------------------------------
# --- Functions ---
# ---------------------------------------------------------------------------

function Write-Log {
    <#
    .SYNOPSIS
        Writes a timestamped log entry to both the log file and the console.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [Parameter()]
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO'
    )

    $entry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"

    # Always append to log file
    Add-Content -Path $LogPath -Value $entry -ErrorAction SilentlyContinue

    # Console output varies by level
    switch ($Level) {
        'ERROR'   { Write-Warning $Message }
        'WARN'    { Write-Warning $Message }
        'SUCCESS' { Write-Host $entry -ForegroundColor Green }
        default   { Write-Verbose $entry }
    }
}

function Get-LicenseSkuId {
    <#
    .SYNOPSIS
        Resolves a license SkuId from a SkuPartNumber, or prompts if none provided.
    .OUTPUTS
        [string] The SkuId GUID for the matching license.
    #>
    param(
        [string]$SkuPartNumber,
        [string]$UserId
    )

    Write-Log "Fetching subscribed SKUs from tenant..."

    # Get all tenant-level subscribed SKUs
    $allSkus = Get-MgSubscribedSku -ErrorAction Stop

    if ($SkuPartNumber) {
        $sku = $allSkus | Where-Object { $_.SkuPartNumber -eq $SkuPartNumber }

        if (-not $sku) {
            throw "License SKU '$SkuPartNumber' not found in tenant. Available SKUs: $(($allSkus.SkuPartNumber) -join ', ')"
        }

        Write-Log "Resolved '$SkuPartNumber' to SkuId: $($sku.SkuId)"
        return $sku.SkuId
    }
    else {
        # Auto-detect: get licenses currently assigned to the target user
        Write-Log "No LicenseSkuPartNumber specified — auto-detecting licenses assigned to $UserId"

        $userLicenses = (Get-MgUser -UserId $UserId -Property AssignedLicenses -ErrorAction Stop).AssignedLicenses

        if (-not $userLicenses -or $userLicenses.Count -eq 0) {
            throw "User '$UserId' has no assigned licenses. Nothing to remove."
        }

        # Cross-reference user's assigned SkuIds with tenant SKU names for readability
        $assignedWithNames = $userLicenses | ForEach-Object {
            $skuId = $_.SkuId
            $match = $allSkus | Where-Object { $_.SkuId -eq $skuId }
            [PSCustomObject]@{
                SkuId         = $skuId
                SkuPartNumber = if ($match) { $match.SkuPartNumber } else { "Unknown" }
            }
        }

        Write-Log "Assigned licenses found:"
        $assignedWithNames | ForEach-Object {
            Write-Log "  $($_.SkuPartNumber) ($($_.SkuId))"
        }

        if ($assignedWithNames.Count -eq 1) {
            Write-Log "Single license found — auto-selecting: $($assignedWithNames[0].SkuPartNumber)"
            return $assignedWithNames[0].SkuId
        }
        else {
            # Multiple licenses assigned — cannot safely auto-select, surface as terminating error
            $skuList = ($assignedWithNames | ForEach-Object { "$($_.SkuPartNumber) ($($_.SkuId))" }) -join "`n  "
            throw "Multiple licenses assigned to '$UserId'. Specify -LicenseSkuPartNumber to target one:`n  $skuList"
        }
    }
}

function Get-VerificationState {
    <#
    .SYNOPSIS
        Retrieves current sign-in, mailbox type, and license state for a user.
    .OUTPUTS
        [PSCustomObject] with AccountEnabled, MailboxType, HasLicense properties.
    #>
    param(
        [string]$UserId,
        [string]$SkuId
    )

    Write-Log "Running verification checks for $UserId..."

    # Graph: account state and license assignments
    $mgUser = Get-MgUser -UserId $UserId -Property AccountEnabled, AssignedLicenses, DisplayName -ErrorAction Stop

    $hasLicense = $false
    if ($SkuId) {
        $hasLicense = ($mgUser.AssignedLicenses | Where-Object { $_.SkuId -eq $SkuId }).Count -gt 0
    }

    # EXO: mailbox recipient type
    $mailbox = Get-Mailbox -Identity $UserId -ErrorAction Stop
    $mailboxType = $mailbox.RecipientTypeDetails

    $result = [PSCustomObject]@{
        DisplayName    = $mgUser.DisplayName
        UPN            = $UserId
        AccountEnabled = $mgUser.AccountEnabled
        MailboxType    = $mailboxType
        LicenseSkuId   = if ($SkuId) { $SkuId } else { 'N/A' }
        HasLicense     = $hasLicense
        Timestamp      = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    }

    return $result
}

# ---------------------------------------------------------------------------
# --- Main ---
# ---------------------------------------------------------------------------

try {
    # Ensure log directory exists before first Write-Log call
    if (-not (Test-Path $LogDir)) {
        New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
    }

    # -------------------------------------------------------------------
    # Module bootstrap — install and import required modules if missing
    # #Requires -Modules is intentionally omitted so this block can run
    # before PowerShell aborts on a missing module.
    # -------------------------------------------------------------------
    $requiredModules = @(
        'ExchangeOnlineManagement',
        'Microsoft.Graph.Users',
        'Microsoft.Graph.Users.Actions',
        'Microsoft.Graph.Identity.DirectoryManagement'
    )

    foreach ($mod in $requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $mod)) {
            Write-Host "[BOOTSTRAP] Module '$mod' not found — installing from PSGallery (Scope: CurrentUser)..." -ForegroundColor Yellow
            Install-Module -Name $mod -Scope CurrentUser -Repository PSGallery -Force -ErrorAction Stop
            Write-Host "[BOOTSTRAP] '$mod' installed successfully." -ForegroundColor Green
        }
        Import-Module -Name $mod -ErrorAction Stop
    }

    $mode = if ($Reverse) { "REVERSAL" } else { "OFFBOARD" }
    Write-Log "=== move365user-shared.ps1 | Mode: $mode | User: $UserPrincipalName | Customer: $CustomerName ==="

    # -------------------------------------------------------------------
    # Connect to Exchange Online FIRST
    # EXO must load its MSAL assemblies before Microsoft.Graph loads its
    # own version — loading Graph first causes a WithBroker() signature
    # conflict (MissingMethodException) in EXO 3.9.x on Windows PS 5.1.
    # -------------------------------------------------------------------
    Write-Log "Connecting to Exchange Online as $EffectiveAdminUPN"

    Connect-ExchangeOnline -UserPrincipalName $EffectiveAdminUPN -ShowBanner:$false -ErrorAction Stop
    Write-Log "Exchange Online connected successfully." -Level 'SUCCESS'

    # Verify mailbox exists before proceeding
    $mailboxCheck = Get-Mailbox -Identity $UserPrincipalName -ErrorAction Stop
    Write-Log "Target mailbox confirmed: $($mailboxCheck.DisplayName) | Current type: $($mailboxCheck.RecipientTypeDetails)"

    # -------------------------------------------------------------------
    # Connect to Microsoft Graph (after EXO to avoid MSAL DLL conflict)
    # -------------------------------------------------------------------
    Write-Log "Connecting to Microsoft Graph (scopes: $($RequiredGraphScopes -join ', '))"

    $graphParams = @{
        Scopes    = $RequiredGraphScopes
        NoWelcome = $true
        ErrorAction = 'Stop'
    }
    if ($EffectiveTenantId) { $graphParams['TenantId'] = $EffectiveTenantId }

    Connect-MgGraph @graphParams
    Write-Log "Microsoft Graph connected successfully." -Level 'SUCCESS'

    # Verify the target user actually exists before proceeding — fail fast
    $targetUser = Get-MgUser -UserId $UserPrincipalName -Property DisplayName, AccountEnabled, AssignedLicenses -ErrorAction Stop
    Write-Log "Target user confirmed: $($targetUser.DisplayName) ($UserPrincipalName)"

    # -------------------------------------------------------------------
    # Resolve license SKU ID (used in both offboard and reversal paths)
    # -------------------------------------------------------------------
    $resolvedSkuId = Get-LicenseSkuId -SkuPartNumber $EffectiveLicenseSku -UserId $UserPrincipalName

    # ===================================================================
    # OFFBOARD PATH (default)
    # ===================================================================
    if (-not $Reverse) {

        Write-Log "--- Beginning OFFBOARD sequence ---"

        # ---------------------------------------------------------------
        # Step 1: Block sign-in
        # ---------------------------------------------------------------
        # Blocking sign-in immediately prevents new sessions while the
        # mailbox conversion and license removal complete. Existing tokens
        # will expire on their own schedule (typically up to 1 hour for
        # access tokens) unless you also revoke sessions via
        # Revoke-MgUserSignInSession — consider adding that for high-risk departures.

        if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Block Entra ID sign-in (AccountEnabled = false)")) {
            Write-Log "Step 1: Blocking sign-in for $UserPrincipalName"
            Update-MgUser -UserId $UserPrincipalName -AccountEnabled:$false -ErrorAction Stop
            Write-Log "Sign-in blocked successfully." -Level 'SUCCESS'
        }

        # ---------------------------------------------------------------
        # Step 2: Convert mailbox to Shared
        # ---------------------------------------------------------------
        # Converting to Shared removes the requirement for a user license
        # to retain mail data. The mailbox remains accessible to delegated
        # accounts. License removal (Step 3) should follow — not precede —
        # this conversion to avoid a race condition where the mailbox is
        # briefly unlicensed but still typed as UserMailbox.

        if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Convert mailbox to Shared")) {
            Write-Log "Step 2: Converting mailbox to Shared for $UserPrincipalName"
            Set-Mailbox -Identity $UserPrincipalName -Type Shared -ErrorAction Stop
            Write-Log "Mailbox converted to Shared." -Level 'SUCCESS'
        }

        # ---------------------------------------------------------------
        # Step 3: Remove license
        # ---------------------------------------------------------------
        # -RemoveLicenses expects an array of SkuId GUIDs (strings).
        # -AddLicenses must be passed as an empty array (not omitted)
        # when only removing — the cmdlet requires both parameters.

        if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Remove license SkuId '$resolvedSkuId'")) {
            Write-Log "Step 3: Removing license (SkuId: $resolvedSkuId) from $UserPrincipalName"
            Set-MgUserLicense -UserId $UserPrincipalName `
                -RemoveLicenses @($resolvedSkuId) `
                -AddLicenses @() `
                -ErrorAction Stop
            Write-Log "License removed successfully." -Level 'SUCCESS'
        }

        Write-Log "--- OFFBOARD sequence complete ---" -Level 'SUCCESS'
    }

    # ===================================================================
    # REVERSAL PATH (-Reverse switch)
    # ===================================================================
    else {

        Write-Log "--- Beginning REVERSAL sequence ---"

        # ---------------------------------------------------------------
        # Step 1: Assign license back
        # ---------------------------------------------------------------
        # -AddLicenses expects an array of objects with a SkuId property.
        # The hashtable cast @{ SkuId = $guid } is the supported shorthand
        # for IMicrosoftGraphAssignedLicense in the Graph PS SDK.

        if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Assign license SkuId '$resolvedSkuId'")) {
            Write-Log "Step 1: Assigning license (SkuId: $resolvedSkuId) back to $UserPrincipalName"
            Set-MgUserLicense -UserId $UserPrincipalName `
                -AddLicenses @(@{ SkuId = $resolvedSkuId }) `
                -RemoveLicenses @() `
                -ErrorAction Stop
            Write-Log "License assigned successfully." -Level 'SUCCESS'
        }

        # ---------------------------------------------------------------
        # Step 2: Convert mailbox back to Regular
        # ---------------------------------------------------------------
        # Converting back to Regular re-enables the full user mailbox
        # experience. Note: this does not automatically re-assign any
        # previously configured mail policies or retention tags.

        if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Convert mailbox back to Regular")) {
            Write-Log "Step 2: Converting mailbox back to Regular for $UserPrincipalName"
            Set-Mailbox -Identity $UserPrincipalName -Type Regular -ErrorAction Stop
            Write-Log "Mailbox converted to Regular." -Level 'SUCCESS'
        }

        # ---------------------------------------------------------------
        # Step 3: Re-enable sign-in
        # ---------------------------------------------------------------
        # Re-enabling sign-in alone is not sufficient for the user to log in.
        # A password reset is required separately — the previous password is
        # not retained after offboarding. Use the Entra ID portal or:
        #   Reset-MgUserPassword -UserId $UserPrincipalName
        #   (requires Authentication.ReadWrite.All scope — not included here
        #    intentionally to keep this script scoped to offboard/restore only)

        if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Re-enable Entra ID sign-in (AccountEnabled = true)")) {
            Write-Log "Step 3: Re-enabling sign-in for $UserPrincipalName"
            Update-MgUser -UserId $UserPrincipalName -AccountEnabled $true -ErrorAction Stop
            Write-Log "Sign-in re-enabled." -Level 'SUCCESS'
            Write-Log "REMINDER: A password reset is required before the user can sign in. Run Reset-MgUserPassword or reset via Entra ID portal." -Level 'WARN'
        }

        Write-Log "--- REVERSAL sequence complete ---" -Level 'SUCCESS'
    }

    # ===================================================================
    # Verification
    # ===================================================================
    # Only run if not in WhatIf mode — state won't have changed
    if (-not $WhatIfPreference) {

        Write-Log "--- Running post-action verification ---"

        $verificationResult = Get-VerificationState -UserId $UserPrincipalName -SkuId $resolvedSkuId

        # Evaluate expected state based on mode
        $expectedAccountEnabled = $Reverse.IsPresent   # true after reversal, false after offboard
        $expectedMailboxType    = if ($Reverse) { 'UserMailbox' } else { 'SharedMailbox' }
        $expectedHasLicense     = $Reverse.IsPresent   # true after reversal, false after offboard

        $verificationResult | Add-Member -MemberType NoteProperty -Name 'ExpectedAccountEnabled' -Value $expectedAccountEnabled
        $verificationResult | Add-Member -MemberType NoteProperty -Name 'ExpectedMailboxType'    -Value $expectedMailboxType
        $verificationResult | Add-Member -MemberType NoteProperty -Name 'ExpectedHasLicense'     -Value $expectedHasLicense

        # Determine overall pass/fail
        $accountCheck = $verificationResult.AccountEnabled -eq $expectedAccountEnabled
        $mailboxCheck2 = $verificationResult.MailboxType -eq $expectedMailboxType
        $licenseCheck = $verificationResult.HasLicense -eq $expectedHasLicense

        $allPassed = $accountCheck -and $mailboxCheck2 -and $licenseCheck
        $verificationResult | Add-Member -MemberType NoteProperty -Name 'VerificationPassed' -Value $allPassed

        # Log individual check results
        Write-Log "Verification | AccountEnabled : Expected=$expectedAccountEnabled | Actual=$($verificationResult.AccountEnabled) | $(if ($accountCheck) {'PASS'} else {'FAIL'})"
        Write-Log "Verification | MailboxType    : Expected=$expectedMailboxType | Actual=$($verificationResult.MailboxType) | $(if ($mailboxCheck2) {'PASS'} else {'FAIL'})"
        Write-Log "Verification | HasLicense     : Expected=$expectedHasLicense | Actual=$($verificationResult.HasLicense) | $(if ($licenseCheck) {'PASS'} else {'FAIL'})"

        if ($allPassed) {
            Write-Log "All verification checks PASSED." -Level 'SUCCESS'
        }
        else {
            Write-Log "One or more verification checks FAILED. Review the output object and log." -Level 'WARN'
        }

        # Output the full result object — caller can pipe this to Export-Csv, ConvertTo-Json, etc.
        Write-Output $verificationResult

        Write-Log "Full log written to: $LogPath"
    }
    else {
        Write-Log "WhatIf mode — verification skipped (no changes were made)." -Level 'WARN'
    }
}
catch {
    Write-Log "FATAL ERROR: $($_.Exception.Message)" -Level 'ERROR'
    Write-Log "Stack Trace: $($_.ScriptStackTrace)" -Level 'ERROR'
    # Re-throw so the caller (or pipeline) gets the terminating error
    throw
}
finally {
    # Always disconnect cleanly — suppress errors if connections were never established
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    Write-Log "=== Session ended ==="
}
