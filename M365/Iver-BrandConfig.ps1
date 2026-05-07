########################################
# Iver Brand Configuration Module
# Dark-themed reporting with yellow accents
########################################

# Brand Color Palette
$IverBrandConfig = @{
    # Primary Colors
    Yellow          = '#FCDE06'    # Iver Yellow - hero accent color
    Black           = '#000000'    # Pure black
    White           = '#FFFFFF'    # Pure white
    
    # Greys (dark background theme)
    Grey80          = '#505050'    # Dark grey for main background
    Grey50          = '#323232'    # Very dark grey for deep sections
    Grey130         = '#828282'    # Medium grey for borders/dividers
    Grey160         = '#A0A0A0'    # Light-medium grey for accents
    Grey210         = '#D2D2D2'    # Light grey for hover states
    Grey240         = '#f0f0f0'    # Very light grey for alt backgrounds
    
    # Semantic Colors
    Success         = '#4CAF50'
    Warning         = '#FF9800'
    Error           = '#F44336'
    Info            = '#2196F3'
    
    # Typography
    FontHeaderFamily = 'Palatino, "Palatino Linotype", serif'
    FontBodyFamily   = 'Arial Nova, Arial, sans-serif'
    
    # Spacing
    PaddingSmall    = '8px'
    PaddingMedium   = '12px'
    PaddingLarge    = '16px'
    PaddingXL       = '24px'
    
    # Border Radius
    BorderRadius    = '4px'
    
    # Icon Map (for report sections)
    Icons = @{
        'Users'              = 'user.png'
        'Licenses'           = 'box.png'
        'MFA'                = 'key.png'
        'Security'           = 'shield.png'
        'Devices'            = 'laptop.png'
        'Exchange'           = 'discussion.png'
        'Cloud'              = 'cloud-server.png'
        'Infrastructure'     = 'server.png'
        'Database'           = 'database.png'
        'Network'            = 'globe.png'
        'Secure Score'       = 'shield.png'
        'Email'              = 'discussion.png'
        'Compliance'         = 'document.png'
        'Settings'           = 'cogs.png'
        'Alerts'             = 'warning.png'
        'Group'              = 'group.png'
        'Dashboard'          = 'graph.png'
        'Configuration'      = 'tools.png'
    }
}

# Function: Get complete CSS stylesheet with Iver branding
function Get-IverCSSStyles {
    param(
        [string]$LogoPath = './iver-complete-neg.png',
        [string]$BackgroundColor = $IverBrandConfig.Grey80
    )
    
    $accentYellow = $IverBrandConfig.Yellow
    $textLight = $IverBrandConfig.White
    $darkBg = $IverBrandConfig.Grey80
    $darkBgDarker = $IverBrandConfig.Grey50
    $borderGrey = $IverBrandConfig.Grey130
    $hoverGrey = $IverBrandConfig.Grey160
    
    @"
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        html, body {
            font-family: $($IverBrandConfig.FontBodyFamily);
            color: $textLight;
            background-color: $darkBg;
            line-height: 1.6;
        }
        
        body {
            padding: 0;
            background: linear-gradient(135deg, $darkBg 0%, $darkBgDarker 100%);
            min-height: 100vh;
        }
        
        /* Header with Logo and Title */
        .report-header {
            background: linear-gradient(135deg, $darkBgDarker 0%, $darkBg 100%);
            padding: $($IverBrandConfig.PaddingXL);
            border-bottom: 4px solid $accentYellow;
            position: relative;
            overflow: hidden;
            margin-bottom: $($IverBrandConfig.PaddingXL);
        }
        
        .report-header::before {
            content: '';
            position: absolute;
            right: -50px;
            top: -50px;
            width: 200px;
            height: 200px;
            background: rgba(252, 222, 6, 0.05);
            border-radius: 50%;
            transform: rotate(45deg);
        }
        
        .header-content {
            position: relative;
            z-index: 1;
            display: flex;
            align-items: center;
            justify-content: space-between;
            flex-wrap: wrap;
            gap: $($IverBrandConfig.PaddingXL);
        }
        
        .logo-container {
            flex: 0 0 auto;
        }
        
        .logo-container img {
            height: 50px;
            filter: drop-shadow(0 2px 4px rgba(0, 0, 0, 0.3));
        }
        
        .header-title {
            flex: 1 1 auto;
            min-width: 200px;
        }
        
        .header-title h1 {
            font-family: $($IverBrandConfig.FontHeaderFamily);
            font-size: 32px;
            font-weight: bold;
            color: $textLight;
            margin-bottom: $($IverBrandConfig.PaddingSmall);
        }
        
        .header-title p {
            color: $hoverGrey;
            font-size: 14px;
        }
        
        .header-meta {
            flex: 0 1 auto;
            text-align: right;
            font-size: 12px;
            color: $hoverGrey;
        }
        
        /* Main Container */
        .report-container {
            max-width: 1400px;
            margin: 0 auto;
            padding: 0 $($IverBrandConfig.PaddingXL);
            padding-bottom: $($IverBrandConfig.PaddingXL);
        }
        
        /* Section Headers */
        .section-header {
            background: linear-gradient(90deg, $darkBgDarker 0%, $darkBg 100%);
            border-left: 6px solid $accentYellow;
            padding: $($IverBrandConfig.PaddingLarge);
            margin-top: $($IverBrandConfig.PaddingXL);
            margin-bottom: $($IverBrandConfig.PaddingLarge);
            border-radius: 4px 0 0 4px;
            display: flex;
            align-items: center;
            gap: $($IverBrandConfig.PaddingLarge);
        }
        
        .section-icon {
            width: 32px;
            height: 32px;
            flex-shrink: 0;
            opacity: 0.9;
            filter: brightness(1.2);
        }
        
        .section-header h2 {
            font-family: $($IverBrandConfig.FontHeaderFamily);
            font-size: 22px;
            font-weight: bold;
            color: $accentYellow;
            margin: 0;
        }
        
        /* Stat Cards */
        .stat-card {
            background: linear-gradient(135deg, rgba(255, 255, 255, 0.05) 0%, rgba(255, 255, 255, 0.02) 100%);
            border: 1px solid $borderGrey;
            border-top: 4px solid $accentYellow;
            border-radius: $($IverBrandConfig.BorderRadius);
            padding: $($IverBrandConfig.PaddingLarge);
            margin-bottom: $($IverBrandConfig.PaddingMedium);
            transition: all 0.3s ease;
        }
        
        .stat-card:hover {
            border-color: $accentYellow;
            background: linear-gradient(135deg, rgba(252, 222, 6, 0.08) 0%, rgba(255, 255, 255, 0.03) 100%);
            transform: translateY(-2px);
            box-shadow: 0 8px 24px rgba(252, 222, 6, 0.1);
        }
        
        .stat-card.success { border-top-color: $($IverBrandConfig.Success); }
        .stat-card.warning { border-top-color: $($IverBrandConfig.Warning); }
        .stat-card.error { border-top-color: $($IverBrandConfig.Error); }
        .stat-card.info { border-top-color: $($IverBrandConfig.Info); }
        
        .stat-label {
            font-size: 12px;
            color: $hoverGrey;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: $($IverBrandConfig.PaddingSmall);
        }
        
        .stat-value {
            font-family: monospace;
            font-size: 28px;
            font-weight: bold;
            color: $accentYellow;
            margin-bottom: $($IverBrandConfig.PaddingSmall);
        }
        
        .stat-value.success { color: $($IverBrandConfig.Success); }
        .stat-value.warning { color: $($IverBrandConfig.Warning); }
        .stat-value.error { color: $($IverBrandConfig.Error); }
        .stat-value.info { color: $($IverBrandConfig.Info); }
        
        .stat-description {
            font-size: 13px;
            color: $hoverGrey;
        }
        
        /* Alert Boxes */
        .alert {
            border-left: 4px solid $borderGrey;
            border-radius: $($IverBrandConfig.BorderRadius);
            padding: $($IverBrandConfig.PaddingLarge);
            margin-bottom: $($IverBrandConfig.PaddingMedium);
            background: rgba(255, 255, 255, 0.03);
        }
        
        .alert.success { 
            border-left-color: $($IverBrandConfig.Success);
            background: rgba(76, 175, 80, 0.08);
        }
        
        .alert.warning {
            border-left-color: $($IverBrandConfig.Warning);
            background: rgba(255, 152, 0, 0.08);
        }
        
        .alert.error {
            border-left-color: $($IverBrandConfig.Error);
            background: rgba(244, 67, 54, 0.08);
        }
        
        .alert.info {
            border-left-color: $($IverBrandConfig.Info);
            background: rgba(33, 150, 243, 0.08);
        }
        
        .alert-title {
            font-weight: bold;
            margin-bottom: $($IverBrandConfig.PaddingSmall);
        }
        
        .alert-message {
            font-size: 13px;
            color: $hoverGrey;
        }
        
        /* Tables */
        table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: $($IverBrandConfig.PaddingLarge);
            background: rgba(0, 0, 0, 0.2);
            border-radius: $($IverBrandConfig.BorderRadius);
            overflow: hidden;
        }
        
        thead {
            background: linear-gradient(90deg, $darkBgDarker 0%, $darkBg 100%);
            border-bottom: 2px solid $accentYellow;
        }
        
        th {
            padding: $($IverBrandConfig.PaddingMedium);
            text-align: left;
            font-weight: bold;
            color: $accentYellow;
            font-size: 13px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        
        td {
            padding: $($IverBrandConfig.PaddingMedium);
            border-bottom: 1px solid rgba(130, 130, 130, 0.2);
        }
        
        tbody tr:hover {
            background: rgba(252, 222, 6, 0.05);
        }
        
        /* Grid Layout */
        .grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
            gap: $($IverBrandConfig.PaddingLarge);
            margin-bottom: $($IverBrandConfig.PaddingXL);
        }
        
        /* Lists */
        ul, ol {
            margin-left: $($IverBrandConfig.PaddingLarge);
            color: $hoverGrey;
        }
        
        li {
            margin-bottom: $($IverBrandConfig.PaddingSmall);
            line-height: 1.8;
        }
        
        /* Footer */
        .report-footer {
            text-align: center;
            padding: $($IverBrandConfig.PaddingXL);
            border-top: 1px solid $borderGrey;
            margin-top: $($IverBrandConfig.PaddingXL);
            color: $hoverGrey;
            font-size: 12px;
        }
        
        .footer-logo {
            height: 30px;
            margin-bottom: $($IverBrandConfig.PaddingMedium);
            opacity: 0.7;
        }
        
        /* Responsive */
        @media (max-width: 768px) {
            .report-header { padding: $($IverBrandConfig.PaddingLarge); }
            .header-content { flex-direction: column; align-items: flex-start; }
            .header-meta { text-align: left; }
            .grid { grid-template-columns: 1fr; }
            .report-container { padding: 0 $($IverBrandConfig.PaddingMedium); }
            h1 { font-size: 24px; }
            h2 { font-size: 18px; }
            .stat-value { font-size: 24px; }
        }
        
        /* Print Styles */
        @media print {
            body { background: white; color: black; }
            .report-header { border-bottom: 2px solid black; }
            .section-header { border-left: 3px solid black; color: black; }
            h2 { color: black; }
            .stat-value { color: black; }
            table { border: 1px solid black; }
            th { background: #f0f0f0; color: black; border-bottom: 1px solid black; }
        }
    </style>
"@
}

# Function: Create a branded stat card
function New-IverStatCard {
    param(
        [string]$Label,
        [string]$Value,
        [string]$Description,
        [ValidateSet('default', 'success', 'warning', 'error', 'info')]
        [string]$Status = 'default',
        [string]$IconPath = $null
    )
    
    $iconHtml = if ($IconPath -and (Test-Path $IconPath)) {
        "<img src='$IconPath' class='stat-icon' alt='$Label' />"
    } else {
        ""
    }
    
    @"
    <div class="stat-card $Status">
        $iconHtml
        <div class="stat-label">$Label</div>
        <div class="stat-value $Status">$Value</div>
        <div class="stat-description">$Description</div>
    </div>
"@
}

# Function: Create a branded alert box
function New-IverAlertBox {
    param(
        [string]$Title,
        [string]$Message,
        [ValidateSet('info', 'success', 'warning', 'error')]
        [string]$Type = 'info'
    )
    
    @"
    <div class="alert $Type">
        <div class="alert-title">$Title</div>
        <div class="alert-message">$Message</div>
    </div>
"@
}

# Function: Create a branded section wrapper
function New-IverSection {
    param(
        [string]$Title,
        [string]$Content,
        [string]$IconPath = $null
    )
    
    $iconHtml = if ($IconPath -and (Test-Path $IconPath)) {
        "<img src='$IconPath' class='section-icon' alt='$Title' />"
    } else {
        ""
    }
    
    @"
    <div class="section">
        <div class="section-header">
            $iconHtml
            <h2>$Title</h2>
        </div>
        <div class="section-content">
            $Content
        </div>
    </div>
"@
}

# Function: Get icon path from standard location
function Get-IverIcon {
    param(
        [string]$IconName,
        [string]$IconDirectory = './'
    )
    
    $iconFile = $IverBrandConfig.Icons[$IconName]
    if ($iconFile) {
        $iconPath = Join-Path $IconDirectory $iconFile
        if (Test-Path $iconPath) {
            return $iconPath
        }
    }
    
    return $null
}

# Function: Export brand config for reference
function Export-IverBrandConfig {
    param(
        [string]$OutputPath = './iver-brand-config.json'
    )
    
    $IverBrandConfig | ConvertTo-Json | Out-File -FilePath $OutputPath -Encoding UTF8
    Write-Host "Brand config exported to: $OutputPath"
}

# Function: Create HTML report header with logo
function New-IverReportHeader {
    param(
        [string]$Title,
        [string]$TenantName,
        [string]$LogoPath = './iver-complete-neg.png',
        [string]$ReportDate = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    )
    
    $logoHtml = if ($LogoPath -and (Test-Path $LogoPath)) {
        "<div class='logo-container'><img src='$LogoPath' alt='Iver Logo' /></div>"
    } else {
        ""
    }
    
    @"
    <div class="report-header">
        <div class="header-content">
            $logoHtml
            <div class="header-title">
                <h1>$Title</h1>
                <p>$TenantName</p>
            </div>
            <div class="header-meta">
                <div>Generated: $ReportDate</div>
            </div>
        </div>
    </div>
"@
}

# Function: Create HTML report footer with logo
function New-IverReportFooter {
    param(
        [string]$LogoPath = './iver-complete-neg.png'
    )
    
    $logoHtml = if ($LogoPath -and (Test-Path $LogoPath)) {
        "<img src='$LogoPath' class='footer-logo' alt='Iver Logo' />"
    } else {
        ""
    }
    
    @"
    <div class="report-footer">
        $logoHtml
        <p>&copy; $(Get-Date -Format 'yyyy') Iver Managed Services. All rights reserved.</p>
    </div>
"@
}

# Export-ModuleMember is only needed if used as a .psm1 module file
# When sourced with dot-sourcing (. "Iver-BrandConfig.ps1"), functions are automatically available
# Uncomment below only if importing as a module (Import-Module Iver-BrandConfig.ps1)
# Export-ModuleMember -Function @(
#     'Get-IverCSSStyles',
#     'New-IverStatCard',
#     'New-IverAlertBox',
#     'New-IverSection',
#     'Get-IverIcon',
#     'Export-IverBrandConfig',
#     'New-IverReportHeader',
#     'New-IverReportFooter'
# ) -Variable 'IverBrandConfig'
