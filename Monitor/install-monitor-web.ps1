#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Installs the ServiceMonitor web admin UI as a Windows service.

.DESCRIPTION
    1. Verifies Python 3 and pip are available.
    2. Installs fastapi, uvicorn, and python-multipart via pip.
    3. Locates or downloads NSSM (Non-Sucking Service Manager).
    4. Creates C:\ServiceMonitor if it does not exist and copies files there.
    5. Registers the web app as a Windows service (ServiceMonitorWeb).
    6. Optionally migrates an existing services.txt + recipients.txt into
       monitor-config.json (recipients go to the Dev group by default).
    7. Places a desktop shortcut for all users.

.PARAMETER InstallDir
    Where to install the monitor files. Default: C:\ServiceMonitor

.PARAMETER Port
    Port the web UI listens on (127.0.0.1 only). Default: 8080

.PARAMETER SourceDir
    Folder containing the script files. Default: the folder of this script.

.PARAMETER NssmPath
    Path to nssm.exe. If not provided, the script looks in SourceDir and PATH,
    then falls back to downloading from the official site.

.PARAMETER SkipMigration
    Do not migrate existing services.txt / recipients.txt even if they exist.

.EXAMPLE
    .\install-monitor-web.ps1

.EXAMPLE
    .\install-monitor-web.ps1 -Port 9090 -SkipMigration

.NOTES
    Version: 1.2.3
    To remove: nssm remove ServiceMonitorWeb confirm
#>

[CmdletBinding()]
param(
    [string] $InstallDir     = 'C:\ServiceMonitor',
    [int]    $Port           = 8080,
    [string] $SourceDir      = $PSScriptRoot,
    [string] $NssmPath       = '',
    [switch] $SkipMigration,

    # Optional shared access token for the web UI (defense-in-depth on multi-user
    # RDP hosts). When set, it is stored in the service environment as SM_TOKEN
    # and users must enter it once. Leave empty to keep the UI open (RDP = auth).
    [string] $AccessToken    = ''
)

$ErrorActionPreference = 'Stop'
$ServiceName           = 'ServiceMonitorWeb'

function Write-Step { param([string]$msg) Write-Host "  $msg" -ForegroundColor Cyan }
function Write-Ok   { param([string]$msg) Write-Host "  OK: $msg" -ForegroundColor Green }
function Write-Warn { param([string]$msg) Write-Host "  WARN: $msg" -ForegroundColor Yellow }

# Write JSON without a UTF-8 BOM. Set-Content -Encoding UTF8 emits a BOM on
# Windows PowerShell 5.1, which Python's json.load chokes on. monitor_web.py
# reads with utf-8-sig as a belt-and-braces measure, but we still write clean.
function Write-JsonFile {
    param([string] $Path, $InputObject)
    $json = $InputObject | ConvertTo-Json -Depth 5
    [System.IO.File]::WriteAllText($Path, $json, (New-Object System.Text.UTF8Encoding($false)))
}

Write-Host ''
Write-Host 'ServiceMonitor Web Admin - installer' -ForegroundColor White
Write-Host '=====================================' -ForegroundColor White
Write-Host ''

# ---------------------------------------------------------------------------
# 1. Python
# ---------------------------------------------------------------------------
Write-Step 'Checking Python...'
$pythonExe = $null
# Try 'py' launcher first (Windows official launcher resolves system installs correctly)
foreach ($candidate in @('py', 'python', 'python3')) {
    try {
        $ver = & $candidate --version 2>&1
        if ($ver -notmatch '3\.\d+') { continue }

        $src = (Get-Command $candidate -ErrorAction SilentlyContinue).Source
        if (-not $src) { continue }

        # Skip the Windows Store stub - it is a redirect that does not work under SYSTEM
        if ($src -match 'WindowsApps') {
            Write-Warn "Skipping Windows Store Python stub: $src"
            continue
        }

        # Resolve the real interpreter path (handles 'py.exe' launcher indirection)
        $real = (& $src -c 'import sys; print(sys.executable)' 2>&1 | Select-Object -Last 1).Trim()
        if ($real -and (Test-Path -LiteralPath $real) -and $real -notmatch 'WindowsApps') {
            $pythonExe = $real
        } else {
            $pythonExe = $src
        }

        Write-Ok "Found $ver at $pythonExe"
        if ($pythonExe -match '\\Users\\') {
            Write-Warn 'Python is a per-user install. The SYSTEM service account may not be able to run it.'
            Write-Warn 'For reliable service operation, install Python system-wide (https://python.org - choose "Install for all users").'
        }
        break
    } catch { continue }
}
# PATH only yielded Store stubs (or nothing). Fall back to scanning standard
# install locations and the registry - covers Python installed but not on PATH.
if (-not $pythonExe) {
    Write-Warn 'No usable Python on PATH; scanning standard install locations...'
    $candidatePaths = New-Object System.Collections.Generic.List[string]

    # Registry (PythonCore) - HKLM + HKCU, native + WOW6432
    $regRoots = @(
        'HKLM:\SOFTWARE\Python\PythonCore',
        'HKLM:\SOFTWARE\WOW6432Node\Python\PythonCore',
        'HKCU:\SOFTWARE\Python\PythonCore'
    )
    foreach ($root in $regRoots) {
        if (-not (Test-Path $root)) { continue }
        foreach ($verKey in Get-ChildItem $root -ErrorAction SilentlyContinue) {
            $ipKey = Join-Path $verKey.PSPath 'InstallPath'
            $ip = (Get-ItemProperty -Path $ipKey -ErrorAction SilentlyContinue).'(default)'
            if ($ip) { $candidatePaths.Add((Join-Path $ip 'python.exe')) }
        }
    }

    # Common filesystem locations
    $globs = @(
        "$env:ProgramFiles\Python3*\python.exe",
        "${env:ProgramFiles(x86)}\Python3*\python.exe",
        "$env:LOCALAPPDATA\Programs\Python\Python3*\python.exe",
        "$env:SystemDrive\Python3*\python.exe"
    )
    foreach ($g in $globs) {
        Get-ChildItem -Path $g -ErrorAction SilentlyContinue |
            ForEach-Object { $candidatePaths.Add($_.FullName) }
    }

    foreach ($p in $candidatePaths) {
        if (-not (Test-Path -LiteralPath $p)) { continue }
        if ($p -match 'WindowsApps') { continue }
        try {
            $ver = & $p --version 2>&1
            if ($ver -match '3\.\d+') {
                # Verify pip is functional - a Python with broken/missing pip
                # would fail hard in step 2; skip it and keep searching instead.
                & $p -m pip --version > $null 2>&1
                if ($LASTEXITCODE -ne 0) {
                    Write-Warn "Skipping $p ($ver found but pip is not functional)"
                    continue
                }
                $pythonExe = $p
                Write-Ok "Found $ver at $pythonExe"
                if ($pythonExe -match '\\Users\\') {
                    Write-Warn 'Python is a per-user install. The SYSTEM service account may not be able to run it.'
                    Write-Warn 'For reliable service operation, install Python system-wide (https://python.org - choose "Install for all users").'
                }
                break
            }
        } catch { continue }
    }
}
if (-not $pythonExe) {
    throw @'
Python 3 not found.

PATH only exposed the Windows Store stubs (App execution aliases), which do not
work under a service account, and no real install was found in the registry or
standard locations. To fix:

  1) Install Python for all users:  winget install Python.Python.3
     (or download from https://python.org and tick "Install for all users")
  2) Recommended: disable the Store aliases at
     Settings > Apps > Advanced app settings > App execution aliases
     (turn OFF python.exe / python3.exe), then re-run this installer.
'@
}

# ---------------------------------------------------------------------------
# 2. pip packages
# ---------------------------------------------------------------------------
# python-multipart is required by FastAPI for Form() handling (the /auth token
# route) - without it uvicorn raises RuntimeError at startup and never binds.
Write-Step 'Installing fastapi, uvicorn, and python-multipart...'
$pipOut = & $pythonExe -m pip install --quiet --upgrade fastapi uvicorn python-multipart 2>&1
if ($LASTEXITCODE -ne 0) {
    throw "pip install failed (exit $LASTEXITCODE): $($pipOut -join "`n")"
}
Write-Ok 'fastapi + uvicorn + python-multipart installed'

# ---------------------------------------------------------------------------
# 3. NSSM
# ---------------------------------------------------------------------------
Write-Step 'Locating nssm.exe...'
$nssmExe = $null
if ($NssmPath -and (Test-Path -LiteralPath $NssmPath)) {
    $nssmExe = $NssmPath
} else {
    foreach ($candidate in @(
        (Join-Path $SourceDir 'nssm.exe'),
        (Join-Path $SourceDir 'nssm\win64\nssm.exe'),
        (Join-Path $InstallDir 'nssm.exe')
    )) {
        if (Test-Path -LiteralPath $candidate) { $nssmExe = $candidate; break }
    }
    if (-not $nssmExe) {
        $inPath = Get-Command 'nssm.exe' -ErrorAction SilentlyContinue
        if ($inPath) { $nssmExe = $inPath.Source }
    }
}
if (-not $nssmExe) {
    Write-Warn 'nssm.exe not found locally. Trying to obtain it...'
    # PS 5.1 defaults to SSL3/TLS1.0; nssm.cc (and most sites) require TLS 1.2+.
    [Net.ServicePointManager]::SecurityProtocol = `
        [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $nssmZip  = Join-Path $env:TEMP 'nssm.zip'
    $nssmTemp = Join-Path $env:TEMP 'nssm'

    # Attempt 1: download from nssm.cc (stable release first, then CI build).
    # The site sometimes returns 503, so retry each URL a couple of times.
    $urls = @(
        'https://nssm.cc/release/nssm-2.24.zip',
        'https://nssm.cc/ci/nssm-2.24-101-g897c7ad.zip'
    )
    foreach ($u in $urls) {
        for ($try = 1; $try -le 2 -and -not $nssmExe; $try++) {
            try {
                # Clear any partial/corrupt artifacts from a prior attempt or run.
                Remove-Item $nssmZip  -Force -ErrorAction SilentlyContinue
                Remove-Item $nssmTemp -Recurse -Force -ErrorAction SilentlyContinue
                Write-Step "Downloading nssm from $u (attempt $try)..."
                Invoke-WebRequest -Uri $u -OutFile $nssmZip -UseBasicParsing -TimeoutSec 30
                Expand-Archive -Path $nssmZip -DestinationPath $nssmTemp -Force
                $nssmExe = Get-ChildItem -Recurse -Filter 'nssm.exe' $nssmTemp |
                           Where-Object { $_.FullName -match 'win64' } |
                           Select-Object -First 1 -ExpandProperty FullName
                if ($nssmExe) { Write-Ok "Downloaded nssm: $nssmExe" }
            } catch {
                Write-Warn "Download failed: $($_.Exception.Message)"
                Start-Sleep -Seconds 3
            }
        }
        if ($nssmExe) { break }
    }

    # Attempt 2: winget (uses its own CDN, so works even when nssm.cc is down).
    if (-not $nssmExe -and (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-Step 'nssm.cc unreachable; trying winget install NSSM.NSSM ...'
        try {
            & winget install --id NSSM.NSSM --silent --accept-source-agreements --accept-package-agreements 2>&1 | Out-Null
        } catch { Write-Warn "winget failed: $($_.Exception.Message)" }
        $wingetGlobs = @(
            "$env:LOCALAPPDATA\Microsoft\WinGet\Links\nssm.exe",
            "$env:LOCALAPPDATA\Microsoft\WinGet\Packages\NSSM.NSSM*\*\win64\nssm.exe",
            "$env:LOCALAPPDATA\Microsoft\WinGet\Packages\NSSM.NSSM*\*\nssm.exe"
        )
        # Prefer the resolved command (robust to future WinGet layout changes);
        # fall back to globbing the package store only if it is not on PATH yet.
        $cmd = Get-Command 'nssm.exe' -ErrorAction SilentlyContinue
        if ($cmd) { $nssmExe = $cmd.Source }
        if (-not $nssmExe) {
            foreach ($g in $wingetGlobs) {
                $hit = Get-ChildItem -Path $g -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($hit) { $nssmExe = $hit.FullName; break }
            }
        }
        if ($nssmExe) { Write-Ok "Obtained nssm via winget: $nssmExe" }
    }

    if (-not $nssmExe) {
        throw @'
Could not obtain nssm.exe automatically (nssm.cc may be down and winget did not yield it).
Fix it one of these ways, then re-run the installer:
  * winget install NSSM.NSSM
  * Download nssm.exe from https://nssm.cc, then either drop it next to this script
    or pass it explicitly:  .\install-monitor-web.ps1 -NssmPath C:\path\to\nssm.exe
'@
    }
}
Write-Ok "nssm: $nssmExe"

# ---------------------------------------------------------------------------
# 4. Copy files to install dir
# ---------------------------------------------------------------------------
Write-Step "Creating install dir: $InstallDir ..."
New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null

$filesToCopy = @(
    'monitor_web.py',
    'ServiceMonitor.ps1',
    'Install-ScheduledTask.ps1',
    'New-CredStore.ps1'
)
foreach ($f in $filesToCopy) {
    $src = Join-Path $SourceDir $f
    if (Test-Path -LiteralPath $src) {
        Copy-Item -LiteralPath $src -Destination $InstallDir -Force
        Write-Ok "Copied $f"
    } else {
        Write-Warn "$f not found in source dir -- skipping"
    }
}

# Copy nssm - but skip if it is already the install-dir copy (e.g. on a re-run,
# where the detection step found nssm.exe already inside $InstallDir).
$nssmDest = Join-Path $InstallDir 'nssm.exe'
if ([System.IO.Path]::GetFullPath($nssmExe) -ieq [System.IO.Path]::GetFullPath($nssmDest)) {
    Write-Ok "nssm already in place: $nssmDest"
} else {
    Copy-Item -LiteralPath $nssmExe -Destination $nssmDest -Force
    Write-Ok 'Copied nssm.exe'
}

# Copy or create monitor-config.json
$configDest = Join-Path $InstallDir 'monitor-config.json'
if (-not (Test-Path -LiteralPath $configDest)) {
    $configSrc = Join-Path $SourceDir 'monitor-config.json'
    if (Test-Path -LiteralPath $configSrc) {
        Copy-Item -LiteralPath $configSrc -Destination $configDest
        Write-Ok 'Copied monitor-config.json'
    } else {
        # Create default
        $default = @{
            version  = '1.2'
            groups   = @{
                dev  = @{ label = 'Dev Team';     recipients = @() }
                iver = @{ label = 'Iver Support'; recipients = @() }
            }
            services = @()
        }
        Write-JsonFile -Path $configDest -InputObject $default
        Write-Ok 'Created default monitor-config.json'
    }
}

# ---------------------------------------------------------------------------
# 5. Migrate services.txt + recipients.txt (optional)
# ---------------------------------------------------------------------------
if (-not $SkipMigration) {
    $oldSvc  = Join-Path $SourceDir 'services.txt'
    $oldRcpt = Join-Path $SourceDir 'recipients.txt'

    if ((Test-Path $oldSvc) -or (Test-Path $oldRcpt)) {
        Write-Step 'Migrating services.txt / recipients.txt into monitor-config.json...'
        $cfg = Get-Content -Path $configDest -Raw | ConvertFrom-Json

        if (Test-Path $oldSvc) {
            $lines = [System.IO.File]::ReadAllLines($oldSvc) |
                     Where-Object { $_ -match '\S' -and $_ -notmatch '^\s*#' }
            foreach ($line in $lines) {
                $trimmed = $line.Trim()
                $paused  = $false
                if ($trimmed -match '^[\*\-]\s*(.+)') {
                    $trimmed = $Matches[1].Trim()
                    $paused  = $true
                }
                $exists = $cfg.services | Where-Object { $_.name -eq $trimmed }
                if (-not $exists -and $trimmed) {
                    $entry = [pscustomobject]@{ name = $trimmed; paused = $paused; alerts = 'both' }
                    $cfg.services += $entry
                }
            }
            Write-Ok "Migrated $($cfg.services.Count) services from services.txt"
        }

        if (Test-Path $oldRcpt) {
            $emails = [System.IO.File]::ReadAllLines($oldRcpt) |
                      Where-Object { $_ -match '@' -and $_ -notmatch '^\s*#' } |
                      ForEach-Object { $_.Trim() }
            foreach ($email in $emails) {
                if ($email -notin $cfg.groups.dev.recipients) {
                    $cfg.groups.dev.recipients += $email
                }
            }
            Write-Ok "Migrated $($emails.Count) recipients to Dev group from recipients.txt"
        }

        Write-JsonFile -Path $configDest -InputObject $cfg
    }
}

# ---------------------------------------------------------------------------
# 6. Register Windows service via NSSM
# ---------------------------------------------------------------------------
Write-Step "Registering Windows service: $ServiceName ..."

$nssmFull    = Join-Path $InstallDir 'nssm.exe'
$appPath     = $pythonExe
$appArgs     = "-m uvicorn monitor_web:app --host 127.0.0.1 --port $Port --workers 1"
$appDir      = $InstallDir
$logFile     = Join-Path $InstallDir 'monitor-web.log'

# Service environment: always SM_CONFIG; add SM_TOKEN only when a token is set.
$envPairs = @("SM_CONFIG=$configDest")
if ($AccessToken) { $envPairs += "SM_TOKEN=$AccessToken" }

# Remove existing service if present
$existing = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($existing) {
    & $nssmFull stop $ServiceName confirm 2>&1 | Out-Null
    & $nssmFull remove $ServiceName confirm 2>&1 | Out-Null
    Write-Warn "Removed existing service $ServiceName"
}

& $nssmFull install $ServiceName $appPath $appArgs
& $nssmFull set     $ServiceName AppDirectory      $appDir
& $nssmFull set     $ServiceName AppEnvironmentExtra @envPairs
if ($AccessToken) { Write-Ok 'Access token gate enabled (SM_TOKEN set in service env)' }
& $nssmFull set     $ServiceName AppStdout         $logFile
& $nssmFull set     $ServiceName AppStderr         $logFile
& $nssmFull set     $ServiceName AppStdoutCreationDisposition 4   # append
& $nssmFull set     $ServiceName AppStderrCreationDisposition 4
& $nssmFull set     $ServiceName Start             SERVICE_AUTO_START
& $nssmFull set     $ServiceName DisplayName       'ServiceMonitor Web Admin'
& $nssmFull set     $ServiceName Description       "Web UI for managing monitored services and alert recipients. http://localhost:$Port"

& $nssmFull start $ServiceName
Write-Ok "Service $ServiceName started"

# ---------------------------------------------------------------------------
# 7. Desktop shortcut for all users
# ---------------------------------------------------------------------------
Write-Step 'Creating desktop shortcut...'
$publicDesktop = [Environment]::GetFolderPath('CommonDesktopDirectory')
# A .url internet shortcut is the correct type for a URL target (a .lnk via
# WScript.Shell expects a filesystem path and is unreliable for http targets).
$shortcutPath  = Join-Path $publicDesktop 'ServiceMonitor Admin.url'
$urlBody = "[InternetShortcut]`r`nURL=http://localhost:$Port`r`n"
[System.IO.File]::WriteAllText($shortcutPath, $urlBody, (New-Object System.Text.ASCIIEncoding))
Write-Ok "Shortcut created: $shortcutPath"

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
Write-Host ''
Write-Host 'Installation complete!' -ForegroundColor Green
Write-Host ''
Write-Host "  Web UI : http://localhost:$Port"
Write-Host "  Config : $configDest"
Write-Host "  Log    : $logFile"
Write-Host "  Service: $ServiceName (auto-start, runs as LocalSystem)"
Write-Host ''
Write-Host 'Next steps:' -ForegroundColor Cyan
Write-Host "  1. Open http://localhost:$Port in a browser"
Write-Host '  2. Add mail recipients to each group (Dev Team / Iver Support)'
Write-Host '  3. Add services to monitor from the right panel'
Write-Host '  4. Run Install-ScheduledTask.ps1 to register ServiceMonitor.ps1 as a task'
Write-Host '  5. Optionally run New-CredStore.ps1 if your SMTP relay requires auth'
Write-Host ''
