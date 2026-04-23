<#
.SYNOPSIS
    Quick TCP port scanner for a single target host.
.DESCRIPTION
    Tests a predefined list of common service ports against a target IP/hostname
    and reports which ports are open.
.PARAMETER Target
    IP address or hostname to scan (default: 10.120.1.141).
.EXAMPLE
    .\Test-OpenPorts.ps1
    .\Test-OpenPorts.ps1 -Target 192.168.1.1
#>

param (
    [string]$Target = "10.120.1.141"
)

$ports = 21, 22, 23, 25, 80, 135, 139, 443, 445, 1433, 3306, 3389, 5985, 5986, 8080, 8443

foreach ($port in $ports) {
    $r = Test-NetConnection -ComputerName $Target -Port $port -WarningAction SilentlyContinue
    if ($r.TcpTestSucceeded) {
        Write-Host "OPEN: $port"
    }
}
