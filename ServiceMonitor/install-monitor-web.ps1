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

.PARAMETER NoPythonAutoInstall
    Do not auto-install Python. If no system-wide Python is found the script
    will throw with manual-install instructions instead of installing one.

.EXAMPLE
    .\install-monitor-web.ps1

.EXAMPLE
    .\install-monitor-web.ps1 -Port 9090 -SkipMigration

.NOTES
    Version: 1.3.0
    To remove: nssm remove ServiceMonitorWeb confirm
#>

[CmdletBinding()]
param(
    [string] $InstallDir     = 'C:\ServiceMonitor',
    [int]    $Port           = 8080,
    [string] $SourceDir      = '',
    [string] $NssmPath       = '',
    [switch] $SkipMigration,

    # By default, if no system-wide Python is found the installer installs one
    # for all users (winget machine scope, or the official python.org silent
    # installer). Pass this to disable that and fail with instructions instead.
    [switch] $NoPythonAutoInstall,

    # Optional shared access token for the web UI (defense-in-depth on multi-user
    # RDP hosts). When set, it is stored in the service environment as SM_TOKEN
    # and users must enter it once. Leave empty to keep the UI open (RDP = auth).
    [string] $AccessToken    = ''
)

$ErrorActionPreference = 'Stop'
$ServiceName           = 'ServiceMonitorWeb'

# $PSScriptRoot is empty inside param() when launched via 'powershell.exe -File'
# (Explorer's "Run with PowerShell") - resolve at body scope instead.
if (-not $SourceDir) {
    $SourceDir = if     ($PSScriptRoot)  { $PSScriptRoot }
                 elseif ($PSCommandPath) { Split-Path -Parent $PSCommandPath }
                 else { throw 'Cannot resolve the source folder - pass -SourceDir explicitly.' }
}

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
# Locate a usable Python 3 (with a working pip). With -SystemOnly, any per-user
# install (under \Users\) is skipped - used after an all-users install so we do
# not re-select a per-user copy that still happens to be first on PATH.
function Find-Python {
    param([switch] $SystemOnly)

    # Function-local EAP: the native probes below redirect stderr (2>&1), which
    # under the script's EAP=Stop becomes a terminating NativeCommandError on the
    # first stderr line (e.g. a pip deprecation warning) - silently skipping a
    # perfectly usable interpreter. Function scope reverts automatically on return.
    $ErrorActionPreference = 'Continue'

    # Try 'py' launcher first (resolves system installs correctly), then python/python3.
    foreach ($candidate in @('py', 'python', 'python3')) {
        try {
            # Out-String: the capture can be an array (stdout + stderr lines),
            # and -notmatch on an array FILTERS instead of testing - flatten first.
            $ver = (& $candidate --version 2>&1 | Out-String).Trim()
            if ($ver -notmatch '3\.\d+') { continue }

            $src = (Get-Command $candidate -ErrorAction SilentlyContinue).Source
            if (-not $src) { continue }

            # Skip the Windows Store stub - a redirect that does not work under SYSTEM.
            if ($src -match 'WindowsApps') {
                Write-Warn "Skipping Windows Store Python stub: $src"
                continue
            }

            # Resolve the real interpreter path (handles 'py.exe' launcher indirection).
            $real  = (& $src -c 'import sys; print(sys.executable)' 2>&1 | Select-Object -Last 1).Trim()
            $found = if ($real -and (Test-Path -LiteralPath $real) -and $real -notmatch 'WindowsApps') { $real } else { $src }

            if ($SystemOnly -and $found -match '\\Users\\') { continue }

            # Verify pip here too (the fallback scan below already does): a PATH
            # Python without pip would pass this loop and then kill the install
            # at the dependency step.
            & $found -m pip --version > $null 2>&1
            if ($LASTEXITCODE -ne 0) {
                Write-Warn "Skipping $found ($ver found but pip is not functional)"
                continue
            }

            Write-Ok "Found $ver at $found"
            if ($found -match '\\Users\\') {
                Write-Warn 'Python is a per-user install. The SYSTEM service account may not be able to run it.'
                Write-Warn 'For reliable service operation, install Python system-wide (https://python.org - choose "Install for all users").'
            }
            return $found
        } catch { continue }
    }

    # PATH only yielded Store stubs (or nothing). Scan the registry + standard
    # install locations - covers Python installed but not on PATH.
    Write-Warn 'No usable Python on PATH; scanning standard install locations...'
    $candidatePaths = New-Object System.Collections.Generic.List[string]

    # Registry (PythonCore) - HKLM + HKCU, native + WOW6432. HKCU is skipped in
    # SystemOnly mode (per-user hive of whoever is running the installer).
    $regRoots = @(
        'HKLM:\SOFTWARE\Python\PythonCore',
        'HKLM:\SOFTWARE\WOW6432Node\Python\PythonCore'
    )
    if (-not $SystemOnly) { $regRoots += 'HKCU:\SOFTWARE\Python\PythonCore' }
    foreach ($root in $regRoots) {
        if (-not (Test-Path $root)) { continue }
        foreach ($verKey in Get-ChildItem $root -ErrorAction SilentlyContinue) {
            $ipKey = Join-Path $verKey.PSPath 'InstallPath'
            $ip = (Get-ItemProperty -Path $ipKey -ErrorAction SilentlyContinue).'(default)'
            if ($ip) { $candidatePaths.Add((Join-Path $ip 'python.exe')) }
        }
    }

    # Common filesystem locations.
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
        if ($SystemOnly -and $p -match '\\Users\\') { continue }
        try {
            # Out-String for the same array-vs-string reason as the PATH loop above.
            $ver = (& $p --version 2>&1 | Out-String).Trim()
            if ($ver -match '3\.\d+') {
                # Verify pip is functional - a Python with broken/missing pip
                # would fail hard in step 2; skip it and keep searching instead.
                & $p -m pip --version > $null 2>&1
                if ($LASTEXITCODE -ne 0) {
                    Write-Warn "Skipping $p ($ver found but pip is not functional)"
                    continue
                }
                Write-Ok "Found $ver at $p"
                if ($p -match '\\Users\\') {
                    Write-Warn 'Python is a per-user install. The SYSTEM service account may not be able to run it.'
                    Write-Warn 'For reliable service operation, install Python system-wide (https://python.org - choose "Install for all users").'
                }
                return $p
            }
        } catch { continue }
    }
    return $null
}

# Install Python 3 for ALL users (so the SYSTEM service account can run it).
# Prefers winget machine scope; falls back to the official python.org silent
# installer. Returns $true if an install appears to have succeeded.
function Install-PythonAllUsers {
    # --- winget (machine scope) --------------------------------------------
    # Drop to 'Continue' so benign winget stderr does not become a terminating
    # NativeCommandError under $ErrorActionPreference='Stop' (which would skip the
    # python.org fallback below). Success is decided from the exit code.
    $winget = (Get-Command winget -ErrorAction SilentlyContinue).Source
    if ($winget) {
        $prevEAP = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        try {
            foreach ($id in @('Python.Python.3.13', 'Python.Python.3.12', 'Python.Python.3.14')) {
                Write-Step "winget install $id (machine scope)..."
                & $winget install --id $id --scope machine --silent `
                    --accept-package-agreements --accept-source-agreements --disable-interactivity 2>&1 | Out-Null
                if ($LASTEXITCODE -eq 0) {
                    Write-Ok "winget installed $id for all users"
                    return $true
                }
                Write-Warn "winget could not install $id (exit $LASTEXITCODE); trying next candidate..."
            }
        } finally {
            $ErrorActionPreference = $prevEAP
        }
    } else {
        Write-Warn 'winget not available; falling back to python.org installer download.'
    }

    # --- python.org silent all-users installer -----------------------------
    try {
        # -bor, not '=': plain assignment would CLOBBER already-enabled protocols
        # (e.g. Tls13 on newer stacks) for the rest of this PowerShell session.
        [Net.ServicePointManager]::SecurityProtocol = `
            [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        $ftp   = 'https://www.python.org/ftp/python/'
        $index = (Invoke-WebRequest -Uri $ftp -UseBasicParsing).Content
        # Stable 3.x.y version folders only (pre-releases like 3.14.0a1 are files
        # inside a folder, not folder names, so this naturally excludes them).
        $vers = [regex]::Matches($index, 'href="(3\.\d+\.\d+)/"') |
            ForEach-Object { [version] $_.Groups[1].Value } |
            Sort-Object -Descending -Unique
        if (-not $vers) {
            Write-Warn 'Could not determine a Python version from python.org.'
            return $false
        }
        $suffix = if ([Environment]::Is64BitOperatingSystem) { '-amd64' } else { '' }
        foreach ($v in $vers) {
            $url = "$ftp$v/python-$v$suffix.exe"
            # Verify the exact installer exists before downloading (skips .0 finals
            # that are not published yet, falling back to the previous release).
            try { Invoke-WebRequest -Uri $url -Method Head -UseBasicParsing -ErrorAction Stop | Out-Null }
            catch { continue }

            $tmp = Join-Path $env:TEMP "python-$v$suffix.exe"
            Write-Step "Downloading Python $v ..."
            Invoke-WebRequest -Uri $url -OutFile $tmp -UseBasicParsing
            Write-Step 'Running silent all-users install...'
            $proc = Start-Process -FilePath $tmp -Wait -PassThru -ArgumentList @(
                '/quiet', 'InstallAllUsers=1', 'PrependPath=1', 'Include_pip=1', 'Include_launcher=1'
            )
            Remove-Item -LiteralPath $tmp -ErrorAction SilentlyContinue
            if ($proc.ExitCode -eq 0) {
                Write-Ok "Installed Python $v for all users"
                return $true
            }
            Write-Warn "Installer for $v exited $($proc.ExitCode); trying an older release..."
        }
    } catch {
        Write-Warn "Auto-download install failed: $($_.Exception.Message)"
    }
    return $false
}

Write-Step 'Checking Python...'
$pythonExe = Find-Python

# A per-user Python (under \Users\) is invisible to other accounts and to the
# SYSTEM service NSSM registers. If we found only one of those - or none - and
# auto-install is allowed, install Python system-wide and prefer that copy.
if ((-not $pythonExe -or $pythonExe -match '\\Users\\') -and -not $NoPythonAutoInstall) {
    if ($pythonExe) {
        Write-Warn 'Only a per-user Python was found; installing system-wide for the service account...'
    } else {
        Write-Warn 'No usable Python found; installing system-wide (all users)...'
    }
    if (Install-PythonAllUsers) {
        $sys = Find-Python -SystemOnly
        if ($sys) { $pythonExe = $sys }
    }
}
if (-not $pythonExe) {
    throw @'
Python 3 not found (and auto-install did not produce a usable interpreter).

PATH only exposed the Windows Store stubs (App execution aliases), which do not
work under a service account, and no all-users install was found. To fix:

  1) Install Python for all users:  winget install Python.Python.3 --scope machine
     (or download from https://python.org and tick "Install for all users")
  2) Recommended: disable the Store aliases at
     Settings > Apps > Advanced app settings > App execution aliases
     (turn OFF python.exe / python3.exe), then re-run this installer.

(Run with -NoPythonAutoInstall to skip the automatic install attempt.)
'@
}

# ---------------------------------------------------------------------------
# 2. pip packages
# ---------------------------------------------------------------------------
# python-multipart is required by FastAPI for Form() handling (the /auth token
# route) - without it uvicorn raises RuntimeError at startup and never binds.
Write-Step 'Installing fastapi, uvicorn, and python-multipart...'
# pip prints benign warnings to stderr (e.g. "script idna.exe ... not on PATH").
# With $ErrorActionPreference='Stop', 2>&1 turns any native-command stderr into a
# *terminating* NativeCommandError - so a harmless warning would kill the install
# before we ever check the exit code. Drop to 'Continue' for the capture, then
# decide success from the real exit code.
$prevEAP = $ErrorActionPreference
$ErrorActionPreference = 'Continue'
try {
    $pipOut  = & $pythonExe -m pip install --quiet --upgrade fastapi uvicorn python-multipart 2>&1
    $pipExit = $LASTEXITCODE
} finally {
    $ErrorActionPreference = $prevEAP
}
if ($pipExit -ne 0) {
    throw "pip install failed (exit $pipExit): $($pipOut -join "`n")"
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
    # 'confirm' is only valid for 'nssm remove' - on 'stop' it is a junk
    # argument that makes nssm error out instead of stopping the service.
    & $nssmFull stop $ServiceName 2>&1 | Out-Null
    & $nssmFull remove $ServiceName confirm 2>&1 | Out-Null
    Write-Warn "Removed existing service $ServiceName"
}

# Check the exit code of the critical nssm calls: a failed 'install' (e.g. SCM
# still holds the old service 'marked for deletion') would otherwise cascade
# into ten confusing follow-up errors and a false 'Service started' message.
& $nssmFull install $ServiceName $appPath $appArgs
if ($LASTEXITCODE -ne 0) {
    throw "nssm install failed (exit $LASTEXITCODE). If the service is 'marked for deletion', close services.msc / Task Manager and re-run."
}
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
if ($LASTEXITCODE -ne 0) {
    throw "nssm start failed (exit $LASTEXITCODE). Check $logFile for the service's own error output."
}
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
Write-Host ''
Write-Host 'Best practice: run the two setup scripts from the INSTALLED folder' -ForegroundColor Yellow
Write-Host "(not the extraction/download folder), so the scheduled task points at" -ForegroundColor Yellow
Write-Host "the permanent copies and cred files under $InstallDir :" -ForegroundColor Yellow
Write-Host ''
Write-Host "  cd `"$InstallDir`""
Write-Host '  # a) If your SMTP relay needs auth, create the cred store FIRST:'
Write-Host '  .\New-CredStore.ps1 -SmtpUser <user> -SmtpServer <host> -SmtpPort <port> -FromAddress <from>'
Write-Host '  # b) Then register the monitor task (point it at the cred files just made):'
Write-Host "  .\Install-ScheduledTask.ps1 -SmtpCredFile `"$InstallDir\smtp.cred`" -SmtpKeyFile `"$InstallDir\smtp.key`""
Write-Host ''
Write-Host "  Reason: both scripts default their paths to their own folder, and the" -ForegroundColor DarkGray
Write-Host "  task's action + cred paths are baked in at registration time. Running" -ForegroundColor DarkGray
Write-Host "  them from a temp/download folder registers a task that breaks once that" -ForegroundColor DarkGray
Write-Host "  folder is deleted." -ForegroundColor DarkGray
Write-Host ''
