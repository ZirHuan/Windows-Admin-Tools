#Requires -Version 5.1
<#
.SYNOPSIS
    Tests whether TCP ports are open on one or more hosts. Returns one object per host+port.

.DESCRIPTION
    Read-only TCP connect test. Completes the TCP handshake and closes the socket; nothing
    is sent. Much faster than Test-NetConnection, which pings first and waits for that too.

    A port counts as Open only when the connect really succeeded. WaitOne() on the async
    handle also returns true when the connect FAILED fast (RST / connection refused), so
    the result is confirmed with EndConnect() before it is reported as open.

    Host names are resolved once per host before connecting, so a name that does not
    resolve is reported as a DNS error (Error column) instead of looking like a closed
    port, and -TimeoutMs applies to the TCP connect only.

    Used on its own, and as a building block by:
      DomainScanner/Invoke-DomainHealthAudit.ps1   (NET section, CA liveness, stale IPs)
      AD/Manage-FSMORoles.ps1                       (DC health score)

.PARAMETER ComputerName
    One or more host names or IP addresses. Accepts pipeline input.

.PARAMETER Port
    One or more TCP ports. Default: a common service list
    (21,22,23,25,53,80,88,135,139,389,443,445,636,1433,3268,3306,3389,5985,5986,8080,8443).

.PARAMETER TimeoutMs
    TCP connect timeout per port in milliseconds (name resolution not included). Default 3000.

.PARAMETER OpenOnly
    Return only the ports that are open.

.OUTPUTS
    PSCustomObject: ComputerName, Port, Open (bool), Address (IP used), Error (string:
    'DNS: ...' when the name does not resolve; empty when the port answered, refused or
    timed out).

.EXAMPLE
    .\Test-TcpPort.ps1 -ComputerName 192.168.1.1
    Common service ports on one host.

.EXAMPLE
    .\Test-TcpPort.ps1 -ComputerName dc01,dc02 -Port 389,636,88,3268 | Format-Table
    AD ports on two DCs.

.EXAMPLE
    'srv01','srv02' | .\Test-TcpPort.ps1 -Port 3389 -OpenOnly

.NOTES
    Version : 1.0.0
    Read-only.
    Replaces Test-OpenPorts.ps1 (one host, fixed list, Test-NetConnection, text output).

    Changelog:
      1.0.0  2026-11-19  First version. Extracted from Invoke-DomainHealthAudit (Test-TcpPort,
                         with the EndConnect fix) and replaces Test-OpenPorts.ps1.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, Position = 0)]
    [Alias('HostName', 'Target', 'IPAddress')]
    [string[]]$ComputerName,

    [ValidateRange(1, 65535)]
    [int[]]$Port = @(21, 22, 23, 25, 53, 80, 88, 135, 139, 389, 443, 445, 636, 1433, 3268, 3306, 3389, 5985, 5986, 8080, 8443),

    [ValidateRange(100, 60000)]
    [int]$TimeoutMs = 3000,

    [switch]$OpenOnly
)

begin {
    $ScriptVersion = '1.0.0'
    Write-Verbose "Test-TcpPort v$ScriptVersion"
}

process {
    foreach ($h in $ComputerName) {
        if ([string]::IsNullOrWhiteSpace($h)) { continue }

        # Resolve once per host. Prefer IPv4 (what the rest of the tooling assumes).
        $addr = $null
        $err  = ''
        $ip   = $null
        if ([System.Net.IPAddress]::TryParse($h, [ref]$ip)) {
            $addr = $ip
        } else {
            try {
                $all  = @([System.Net.Dns]::GetHostAddresses($h))
                $addr = @($all | Where-Object { $_.AddressFamily -eq 'InterNetwork' }) + @($all) | Select-Object -First 1
                if (-not $addr) { $err = 'DNS: no address returned' }
            } catch {
                $err = "DNS: $($_.Exception.GetBaseException().Message)"
            }
        }

        foreach ($p in $Port) {
            $open = $false
            if ($addr) {
                $tcp = New-Object System.Net.Sockets.TcpClient($addr.AddressFamily)
                try {
                    $ar = $tcp.BeginConnect($addr, $p, $null, $null)
                    if ($ar.AsyncWaitHandle.WaitOne($TimeoutMs, $false)) {
                        # Throws on refused/reset - only a completed connect is "open".
                        $tcp.EndConnect($ar)
                        $open = $tcp.Connected
                    }
                } catch {
                    # Refused / unreachable are normal "closed" outcomes.
                } finally {
                    $tcp.Close()
                }
            }
            if ($OpenOnly -and -not $open) { continue }
            [PSCustomObject]@{
                ComputerName = $h
                Port         = $p
                Open         = $open
                Address      = if ($addr) { "$addr" } else { '' }
                Error        = $err
            }
        }
    }
}
