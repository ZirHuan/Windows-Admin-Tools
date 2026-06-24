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
    worst-case restart loop (3 attempts × 60 s + settle time + SMTP overhead)
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

.EXAMPLE
    .\Install-ScheduledTask.ps1

.EXAMPLE
    .\Install-ScheduledTask.ps1 -IntervalMinutes 10 -TaskName MyServiceCheck

.EXAMPLE
    # Remove the task:
    Unregister-ScheduledTask -TaskName ServiceMonitor -Confirm:$false
#>

[CmdletBinding()]
param(
    [string] $ScriptFolder    = $PSScriptRoot,
    [int]    $IntervalMinutes = 5,
    [string] $TaskName        = 'ServiceMonitor',
    [string] $TaskFolder      = '\'
)

$ErrorActionPreference = 'Stop'

if ($IntervalMinutes -lt 1) {
    throw "IntervalMinutes must be >= 1 (got $IntervalMinutes)."
}

$scriptPath = Join-Path $ScriptFolder 'ServiceMonitor.ps1'
if (-not (Test-Path -LiteralPath $scriptPath)) {
    throw "ServiceMonitor.ps1 not found at: $scriptPath"
}

Write-Host "Installing '$TaskName' (every $IntervalMinutes min, SYSTEM) ..." -ForegroundColor Cyan

# Action: run PowerShell non-interactively
$action = New-ScheduledTaskAction `
    -Execute 'powershell.exe' `
    -Argument "-NonInteractive -ExecutionPolicy Bypass -File `"$scriptPath`"" `
    -WorkingDirectory $ScriptFolder

# Trigger: repeat indefinitely starting immediately.
# -RepetitionDuration ([TimeSpan]::MaxValue) ensures the task repeats forever -
# without it some Windows builds silently stop repeating after a finite period.
$trigger = New-ScheduledTaskTrigger `
    -Once `
    -At (Get-Date) `
    -RepetitionInterval  (New-TimeSpan -Minutes $IntervalMinutes) `
    -RepetitionDuration  ([TimeSpan]::MaxValue)

# Settings: fixed 30-minute execution limit (safe for max restart loop),
# ignore new triggers while running, and start if missed (e.g. after reboot).
$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit  (New-TimeSpan -Minutes 30) `
    -MultipleInstances   IgnoreNew `
    -StartWhenAvailable

# Run as SYSTEM with highest privilege (required for service restart)
$principal = New-ScheduledTaskPrincipal `
    -UserId    'SYSTEM' `
    -LogonType ServiceAccount `
    -RunLevel  Highest

# Remove existing task with the same name before re-registering
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
    -Description "Monitors services listed in services.txt and alerts via email. ServiceMonitor.ps1 v1.0.1" |
    Out-Null

Write-Host ''
Write-Host "Task '$TaskName' registered successfully." -ForegroundColor Green
Write-Host "  Interval : every $IntervalMinutes min"
Write-Host "  Runs as  : SYSTEM (Highest)"
Write-Host "  Exe limit: 30 minutes"
Write-Host "  Repeats  : indefinitely"
Write-Host ''
Write-Host "Verify in Task Scheduler UI that 'Duration' shows 'Indefinitely'."
Write-Host "Run manually: Start-ScheduledTask -TaskName '$TaskName'"
