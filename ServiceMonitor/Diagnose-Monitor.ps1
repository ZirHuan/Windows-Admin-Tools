#Requires -Version 5.1
<#
.SYNOPSIS
    One-shot diagnostic for "service stopped but ServiceMonitor did nothing".
    Run in an ELEVATED PowerShell on the monitored host. Read-only except the
    final manual engine test, which WILL restart stopped monitored services
    (it runs -NoEmail, so no mail is sent).
#>

$ErrorActionPreference = 'Continue'
$dir = 'C:\ServiceMonitor'

Write-Host "===== 1. Scheduled task =====" -ForegroundColor Cyan
$task = Get-ScheduledTask -TaskName 'ServiceMonitor' -ErrorAction SilentlyContinue
if (-not $task) {
    Write-Host "  ServiceMonitor task is NOT REGISTERED. <-- this is almost certainly the problem." -ForegroundColor Red
    Write-Host "  Fix: run  .\Install-ScheduledTask.ps1  from $dir (elevated)."
} else {
    Write-Host "  State : $($task.State)"
    $task | Get-ScheduledTaskInfo |
        Format-List TaskName, LastRunTime, LastTaskResult, NextRunTime, NumberOfMissedRuns
    # What the task actually runs (path + args) - catches a stale/wrong path.
    $act = $task.Actions | Select-Object -First 1
    Write-Host "  Execute  : $($act.Execute)"
    Write-Host "  Argument : $($act.Arguments)"
}

Write-Host "`n===== 2. Deployed script version =====" -ForegroundColor Cyan
$sp = Join-Path $dir 'ServiceMonitor.ps1'
if (Test-Path $sp) {
    (Select-String -Path $sp -Pattern "^\`$ScriptVersion\s*=\s*'([^']+)'").Matches.Groups[1].Value |
        ForEach-Object { Write-Host "  ServiceMonitor.ps1 version: $_" }
} else {
    Write-Host "  ServiceMonitor.ps1 NOT FOUND at $sp" -ForegroundColor Red
}

Write-Host "`n===== 3. Log tail (did the engine ever run?) =====" -ForegroundColor Cyan
$log = Join-Path $dir 'ServiceMonitor.log'
if (Test-Path $log) {
    Write-Host "  Last write: $((Get-Item $log).LastWriteTime)"
    Get-Content $log -Tail 25
} else {
    Write-Host "  No log file at $log - the engine has never successfully started." -ForegroundColor Yellow
}

Write-Host "`n===== 4. Cooldown state file =====" -ForegroundColor Cyan
$st = Join-Path $dir 'ServiceMonitor.state.json'
if (Test-Path $st) { Get-Content $st } else { Write-Host "  (none yet - no alert has been committed)" }

Write-Host "`n===== 5. Manual engine test (-NoEmail; WILL restart stopped services) =====" -ForegroundColor Cyan
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $sp -NoEmail
Write-Host "  ENGINE EXIT CODE = $LASTEXITCODE" -ForegroundColor Cyan
