#Requires -Version 5.1
<#
.SYNOPSIS
    Real-time memory usage monitor with configurable warning/critical thresholds.
.DESCRIPTION
    Polls physical memory usage every 30 seconds and displays colour-coded output.
    Warning and critical thresholds are set interactively at startup.
.EXAMPLE
    .\Watch-Memory.ps1
#>

$warning  = Read-Host "Warning threshold  % (default 75)"
$critical = Read-Host "Critical threshold % (default 90)"

$warning  = if ($warning)  { [int]$warning  } else { 75 }
$critical = if ($critical) { [int]$critical } else { 90 }

Write-Host "Monitoring memory — Warning: $warning%  Critical: $critical%  (Ctrl+C to stop)`n"

while ($true) {
    $os      = Get-CimInstance Win32_OperatingSystem
    $used    = $os.TotalVisibleMemorySize - $os.FreePhysicalMemory
    $pct     = [math]::Round(($used / $os.TotalVisibleMemorySize) * 100, 1)
    $totalGB = [math]::Round($os.TotalVisibleMemorySize / 1MB, 1)
    $usedGB  = [math]::Round($used / 1MB, 1)

    $color = if ($pct -ge $critical) { "Red" } elseif ($pct -ge $warning) { "Yellow" } else { "Green" }

    Write-Host "[$((Get-Date).ToString('HH:mm:ss'))]  RAM: " -NoNewline
    Write-Host "$pct%" -ForegroundColor $color -NoNewline
    Write-Host "  ($usedGB GB / $totalGB GB)"

    Start-Sleep -Seconds 30
}
