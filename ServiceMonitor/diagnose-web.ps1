#Requires -Version 5.1
<#
.SYNOPSIS
    Diagnoses why the ServiceMonitorWeb service / web UI is not reachable.
.DESCRIPTION
    Reports the service state, the exact command NSSM runs (and as which account),
    whether anything is listening on the port, the tail of the service log, and -
    most usefully - tries to launch uvicorn in the foreground so any import/bind
    error is printed directly. Run in an elevated PowerShell on the RDP host.
.NOTES
    Version: 1.0.0
#>
param(
    [string] $InstallDir = 'C:\ServiceMonitor',
    [int]    $Port       = 8080
)

$svcName = 'ServiceMonitorWeb'

Write-Host "=== 1. Service status ===" -ForegroundColor Cyan
$svc = Get-Service -Name $svcName -ErrorAction SilentlyContinue
if ($svc) { $svc | Format-List Name, Status, StartType }
else      { Write-Host "Service $svcName NOT FOUND" -ForegroundColor Red }

Write-Host "=== 2. NSSM service config (command + account) ===" -ForegroundColor Cyan
$nssm = Join-Path $InstallDir 'nssm.exe'
if (Test-Path -LiteralPath $nssm) {
    foreach ($k in 'Application', 'AppParameters', 'AppDirectory', 'ObjectName') {
        "{0,-14}: {1}" -f $k, (& $nssm get $svcName $k 2>&1)
    }
} else { Write-Host "nssm.exe not found at $nssm" -ForegroundColor Red }

Write-Host "=== 3. Listening on port $Port? ===" -ForegroundColor Cyan
$conn = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue
if ($conn) { $conn | Select-Object LocalAddress, LocalPort, State, OwningProcess | Format-Table -AutoSize }
else       { Write-Host "Nothing is listening on $Port." -ForegroundColor Yellow }

Write-Host "=== 4. Last 40 lines of monitor-web.log ===" -ForegroundColor Cyan
$log = Join-Path $InstallDir 'monitor-web.log'
if (Test-Path -LiteralPath $log) { Get-Content -LiteralPath $log -Tail 40 }
else { Write-Host "No log at $log" -ForegroundColor Red }

Write-Host "=== 5. Foreground launch test (Ctrl+C to stop) ===" -ForegroundColor Cyan
$appPath = (& $nssm get $svcName Application 2>&1) -join ''
Write-Host "Running: $appPath -m uvicorn monitor_web:app --host 127.0.0.1 --port $Port"
Write-Host "(If this works here but the service does not, it is an account/per-user-Python issue.)" -ForegroundColor Yellow
Push-Location $InstallDir
& $appPath -m uvicorn monitor_web:app --host 127.0.0.1 --port $Port
Pop-Location
