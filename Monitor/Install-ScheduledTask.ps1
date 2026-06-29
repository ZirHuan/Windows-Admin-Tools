#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Registers ServiceMonitor.ps1 as a repeating Windows Scheduled Task.

.DESCRIPTION
    Creates a task that runs ServiceMonitor.ps1 every IntervalMinutes minutes
    under the SYSTEM account (highest privilege). The task repeats indefinitely
    and ignores new triggers while a previous instance is still running.

    ExecutionTimeLimit is set to a fixed 30 minutes to safely accommodate the
    worst-case restart loop (3 attempts x 60 s + settle time + SMTP overhead)
    regardless of the polling interval.

.PARAMETER ScriptFolder
    Folder containing ServiceMonitor.ps1 and the config files.
    Default: the folder containing this script.

.PARAMETER IntervalMinutes
    How often to check services, in minutes. Minimum 1.
    Default: 5.

.PARAMETER TaskName
    The Scheduled Task name. Default: ServiceMonitor.

.PARAMETER TaskFolder
    Task Scheduler folder path (e.g. '\MyOrg\'). Default: '\'  (root).

.PARAMETER PowerShell7
    Use pwsh.exe (PowerShell 7+) instead of powershell.exe (Windows PowerShell 5.1).
    pwsh.exe must be installed and on the system PATH or at its default location.

.EXAMPLE
    .\Install-ScheduledTask.ps1

.EXAMPLE
    .\Install-ScheduledTask.ps1 -IntervalMinutes 10 -TaskName MyServiceCheck

.EXAMPLE
    .\Install-ScheduledTask.ps1 -PowerShell7

.EXAMPLE
    # Remove the task:
    Unregister-ScheduledTask -TaskName ServiceMonitor -Confirm:$false
#>

[CmdletBinding()]
param(
    [string] $ScriptFolder    = $PSScriptRoot,
    [int]    $IntervalMinutes = 5,
    [string] $TaskName        = 'ServiceMonitor',
    [string] $TaskFolder      = '\',
    [switch] $PowerShell7,

    # SMTP authentication pass-through. When supplied, these are forwarded to
    # ServiceMonitor.ps1 in the task action so authenticated relay actually works
    # under the SYSTEM account. Omit all for anonymous relay.
    [string] $SmtpUser        = '',
    [string] $SmtpCredFile    = '',
    [string] $SmtpKeyFile     = '',
    [switch] $SmtpUseSsl
)

$ErrorActionPreference = 'Stop'

if ($IntervalMinutes -lt 1) {
    throw "IntervalMinutes must be >= 1 (got $IntervalMinutes)."
}

$scriptPath = Join-Path $ScriptFolder 'ServiceMonitor.ps1'
if (-not (Test-Path -LiteralPath $scriptPath)) {
    throw "ServiceMonitor.ps1 not found at: $scriptPath"
}

# Read version from the script rather than hardcoding it
$versionLine = Select-String -LiteralPath $scriptPath -Pattern "^\`$ScriptVersion\s*=\s*'([^']+)'"
$scriptVersion = if ($versionLine) { $versionLine.Matches[0].Groups[1].Value } else { 'unknown' }

# Resolve PowerShell executable
if ($PowerShell7) {
    $defaultPwsh = 'C:\Program Files\PowerShell\7\pwsh.exe'
    $exe = if (Test-Path -LiteralPath $defaultPwsh) { $defaultPwsh } else { 'pwsh.exe' }
    $psLabel = 'PowerShell 7 (pwsh.exe)'
} else {
    $exe = 'powershell.exe'
    $psLabel = 'Windows PowerShell 5.1 (powershell.exe)'
}

Write-Host "Installing '$TaskName' v$scriptVersion (every $IntervalMinutes min, SYSTEM) ..." -ForegroundColor Cyan
Write-Host "  PowerShell: $psLabel"

# Build the script arguments, appending SMTP auth flags only when provided so
# authenticated relay is actually wired into the SYSTEM task (not dead config).
$scriptArgs = "-NonInteractive -ExecutionPolicy Bypass -File `"$scriptPath`""
if ($SmtpUser)     { $scriptArgs += " -SmtpUser `"$SmtpUser`"" }
if ($SmtpCredFile) { $scriptArgs += " -SmtpCredFile `"$SmtpCredFile`"" }
if ($SmtpKeyFile)  { $scriptArgs += " -SmtpKeyFile `"$SmtpKeyFile`"" }
if ($SmtpUseSsl)   { $scriptArgs += " -SmtpUseSsl" }

if ($SmtpUser) { Write-Host "  SMTP auth : enabled (user $SmtpUser)" }

$action = New-ScheduledTaskAction `
    -Execute $exe `
    -Argument $scriptArgs `
    -WorkingDirectory $ScriptFolder

# RepetitionDuration: [TimeSpan]::MaxValue is rejected by Task Scheduler on some
# Windows builds. Try it first (shows 'Indefinitely'), then fall back to a long
# finite span (~10 years) which every build accepts.
try {
    $trigger = New-ScheduledTaskTrigger `
        -Once `
        -At (Get-Date) `
        -RepetitionInterval  (New-TimeSpan -Minutes $IntervalMinutes) `
        -RepetitionDuration  ([TimeSpan]::MaxValue)
} catch {
    Write-Host "  Note: [TimeSpan]::MaxValue rejected ($($_.Exception.Message)); using 10-year duration." -ForegroundColor Yellow
    $trigger = New-ScheduledTaskTrigger `
        -Once `
        -At (Get-Date) `
        -RepetitionInterval  (New-TimeSpan -Minutes $IntervalMinutes) `
        -RepetitionDuration  (New-TimeSpan -Days 3650)
}

$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit  (New-TimeSpan -Minutes 30) `
    -MultipleInstances   IgnoreNew `
    -StartWhenAvailable

$principal = New-ScheduledTaskPrincipal `
    -UserId    'SYSTEM' `
    -LogonType ServiceAccount `
    -RunLevel  Highest

if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host "Removed existing task '$TaskName'." -ForegroundColor Yellow
}

Register-ScheduledTask `
    -TaskPath   $TaskFolder `
    -TaskName   $TaskName `
    -Action     $action `
    -Trigger    $trigger `
    -Settings   $settings `
    -Principal  $principal `
    -Description "Monitors services defined in monitor-config.json and alerts via email. ServiceMonitor.ps1 v$scriptVersion" |
    Out-Null

Write-Host ''
Write-Host "Task '$TaskName' registered successfully." -ForegroundColor Green
Write-Host "  Version  : $scriptVersion"
Write-Host "  Interval : every $IntervalMinutes min"
Write-Host "  Runs as  : SYSTEM (Highest)"
Write-Host "  Exe limit: 30 minutes"
Write-Host "  Repeats  : indefinitely"
Write-Host ''
Write-Host "Verify in Task Scheduler UI that 'Duration' shows 'Indefinitely'."
Write-Host "Run manually: Start-ScheduledTask -TaskName '$TaskName'"
