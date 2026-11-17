#Requires -Version 3.0
<#
.SYNOPSIS
    Summarise failed logons (4625) and recover the source IP that 4625 leaves blank.
.DESCRIPTION
    Built for Windows Server 2012 / 2012 R2 (PowerShell 3.0/4.0 compatible).

    Event 4625 with Logon Type 3 + NtLmSsp + "Source Network Address: -" is the
    classic RDP-with-NLA (CredSSP) or SMB brute-force pattern: the NTLM layer does
    not know the client IP, so Security log cannot show it. This script:

      1. Parses all 4625 events in the window (user, workstation, IP, status, type).
      2. Reads the RDP transport log (RdpCoreTS/Operational) event 131
         "server accepted a new TCP connection from client IP:port" and event 140
         "connection from IP failed because user name or password is not correct".
      3. Correlates each 4625 that has no IP with RDP connections accepted within
         -CorrelateSeconds before it. This is TIME-BASED and marked "probable".
      4. Shows live TCP sessions on 3389/445 and whether firewall logging is on.
      5. Writes CSVs + a text summary to -OutDir.

    Read-only unless -EnableFirewallLog is given (then it turns on Windows Firewall
    logging of allowed + dropped connections for all profiles, 32 MB log).
.PARAMETER Hours
    How far back to look. Default 24.
.PARAMETER CorrelateSeconds
    Window before a 4625 in which an RDP 131 connection counts as a candidate. Default 5.
.PARAMETER OutDir
    Output folder. Default C:\Temp\FailedLogonReport-<timestamp>.
.PARAMETER EnableFirewallLog
    Turn on Windows Firewall connection logging (changes configuration).
.EXAMPLE
    .\Get-FailedLogonReport.ps1
.EXAMPLE
    .\Get-FailedLogonReport.ps1 -Hours 72 -EnableFirewallLog
.NOTES
    Version : 1.0.0
    1.0.0  Initial release: 4625 parsing, RdpCoreTS 131/140 IP recovery,
           time-based correlation, live 3389/445 sessions, optional firewall logging.
#>
[CmdletBinding()]
param(
    [int]$Hours = 24,
    [int]$CorrelateSeconds = 5,
    [string]$OutDir = ("C:\Temp\FailedLogonReport-{0}" -f (Get-Date -Format 'yyyyMMdd-HHmmss')),
    [switch]$EnableFirewallLog
)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) { throw 'Run this from an elevated PowerShell (Security log needs admin).' }

if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }
$start = (Get-Date).AddHours(-$Hours)

$statusText = @{
    '0XC000006A' = 'Bad password (user exists)'
    '0XC0000064' = 'User does not exist'
    '0XC000006D' = 'Bad user name or password'
    '0XC0000234' = 'Account locked out'
    '0XC0000072' = 'Account disabled'
    '0XC000006F' = 'Outside allowed logon hours'
    '0XC0000070' = 'Workstation restriction'
    '0XC0000071' = 'Password expired'
    '0XC0000193' = 'Account expired'
    '0XC0000224' = 'Must change password'
    '0XC000015B' = 'Logon type not granted'
    '0XC0000133' = 'Clock skew (Kerberos)'
}

function Get-EventDataMap {
    param($Event)
    $map = @{}
    $xml = [xml]$Event.ToXml()
    foreach ($d in $xml.Event.EventData.Data) { $map[[string]$d.Name] = [string]$d.'#text' }
    return $map
}

# ---------------------------------------------------------------- 1. 4625
Write-Host "[1/4] Reading Security 4625 since $start ..." -ForegroundColor Cyan
$failed = @()
try {
    $ev = Get-WinEvent -FilterHashtable @{ LogName = 'Security'; Id = 4625; StartTime = $start } -ErrorAction Stop
} catch {
    if ($_.Exception.Message -match 'No events were found') { $ev = @() } else { throw }
}
foreach ($e in $ev) {
    $m = Get-EventDataMap $e
    $sub = ([string]$m['SubStatus']).ToUpper()
    $st  = ([string]$m['Status']).ToUpper()
    $reasonKey = $sub
    if (-not $statusText.ContainsKey($reasonKey)) { $reasonKey = $st }
    $reason = 'Other'
    if ($statusText.ContainsKey($reasonKey)) { $reason = $statusText[$reasonKey] }
    $failed += New-Object PSObject -Property ([ordered]@{
        Time        = $e.TimeCreated
        User        = $m['TargetUserName']
        Domain      = $m['TargetDomainName']
        LogonType   = $m['LogonType']
        Workstation = $m['WorkstationName']
        IpAddress   = $m['IpAddress']
        IpPort      = $m['IpPort']
        Status      = $st
        SubStatus   = $sub
        Reason      = $reason
        LogonProc   = ([string]$m['LogonProcessName']).Trim()
        AuthPkg     = $m['AuthenticationPackageName']
        ProcessName = $m['ProcessName']
        ProbableIP  = ''
    })
}
$failed = @($failed | Sort-Object Time)
Write-Host ("      {0} failed logons" -f $failed.Count)

# ---------------------------------------------------------------- 2. RDP transport log
Write-Host '[2/4] Reading RdpCoreTS/Operational (131 accepted, 140 bad creds) ...' -ForegroundColor Cyan
$rdpLog = 'Microsoft-Windows-RemoteDesktopServices-RdpCoreTS/Operational'
$rdpConn = @(); $rdpFail = @(); $rdpNote = ''
try {
    $rev = Get-WinEvent -FilterHashtable @{ LogName = $rdpLog; Id = 131, 140; StartTime = $start } -ErrorAction Stop
    foreach ($e in $rev) {
        $m = Get-EventDataMap $e
        $ip = ''
        foreach ($k in @('ClientIP', 'IPString')) { if ($m.ContainsKey($k) -and $m[$k]) { $ip = $m[$k] } }
        if (-not $ip) {
            $rx = [regex]::Match([string]$e.Message, '(\d{1,3}(\.\d{1,3}){3}|\[?[0-9a-fA-F:]{3,}\]?)(:\d+)?')
            if ($rx.Success) { $ip = $rx.Value }
        }
        $ipOnly = $ip -replace '^\[?([^\]]+?)\]?(:\d+)?$', '$1'
        if ($ip -match '^(\d{1,3}(\.\d{1,3}){3}):\d+$') { $ipOnly = $matches[1] }
        $o = New-Object PSObject -Property ([ordered]@{ Time = $e.TimeCreated; EventId = $e.Id; ClientIP = $ipOnly; Raw = $ip })
        if ($e.Id -eq 131) { $rdpConn += $o } else { $rdpFail += $o }
    }
} catch {
    if ($_.Exception.Message -match 'No events were found') { $rdpNote = 'RdpCoreTS log present, no 131/140 in window.' }
    else { $rdpNote = "RdpCoreTS log not readable: $($_.Exception.Message)" }
}
Write-Host ("      {0} RDP connections accepted, {1} RDP bad-credential events. {2}" -f $rdpConn.Count, $rdpFail.Count, $rdpNote)

# ---------------------------------------------------------------- 3. correlate
Write-Host "[3/4] Correlating blank-IP 4625 with RDP connects (<= $CorrelateSeconds s before) ..." -ForegroundColor Cyan
$rdpConn = @($rdpConn | Sort-Object Time)
foreach ($f in $failed) {
    if ($f.IpAddress -and $f.IpAddress -ne '-') { continue }
    $lo = $f.Time.AddSeconds(-$CorrelateSeconds)
    $cand = @($rdpConn | Where-Object { $_.Time -ge $lo -and $_.Time -le $f.Time.AddSeconds(1) } | Select-Object -ExpandProperty ClientIP -Unique)
    if ($cand.Count -gt 0) { $f.ProbableIP = ($cand -join ';') }
}

# ---------------------------------------------------------------- 4. live state
Write-Host '[4/4] Live sessions on 3389/445 + firewall logging state ...' -ForegroundColor Cyan
$live = @()
try {
    $live = @(Get-NetTCPConnection -State Established -ErrorAction Stop | Where-Object { $_.LocalPort -eq 3389 -or $_.LocalPort -eq 445 } |
        Select-Object LocalPort, RemoteAddress, RemotePort, OwningProcess)
} catch { $live = @() }

$fwProfiles = @()
try { $fwProfiles = @(Get-NetFirewallProfile -ErrorAction Stop | Select-Object Name, Enabled, LogAllowed, LogBlocked, LogFileName, LogMaxSizeKilobytes) } catch { }

$nla = $null
try { $nla = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp' -Name UserAuthentication -ErrorAction Stop).UserAuthentication } catch { }
$rdpDeny = $null
try { $rdpDeny = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server' -Name fDenyTSConnections -ErrorAction Stop).fDenyTSConnections } catch { }

if ($EnableFirewallLog) {
    Write-Host '      Enabling firewall logging (allowed + blocked, 32 MB) on all profiles ...' -ForegroundColor Yellow
    Set-NetFirewallProfile -All -LogAllowed True -LogBlocked True -LogMaxSizeKilobytes 32767 -LogFileName '%systemroot%\system32\LogFiles\Firewall\pfirewall.log'
    Write-Host '      Done. Log: C:\Windows\System32\LogFiles\Firewall\pfirewall.log'
}

# ---------------------------------------------------------------- output
$failed  | Export-Csv (Join-Path $OutDir 'failed-4625.csv')      -NoTypeInformation -Encoding UTF8
$rdpConn | Export-Csv (Join-Path $OutDir 'rdp-131-connects.csv') -NoTypeInformation -Encoding UTF8
$rdpFail | Export-Csv (Join-Path $OutDir 'rdp-140-badcreds.csv') -NoTypeInformation -Encoding UTF8

$sb = New-Object System.Text.StringBuilder
function Add-Section { param([string]$Title, $Data)
    [void]$sb.AppendLine(''); [void]$sb.AppendLine("=== $Title ===")
    if ($Data) { [void]$sb.AppendLine(($Data | Format-Table -AutoSize | Out-String -Width 200).TrimEnd()) } else { [void]$sb.AppendLine('(none)') }
}
[void]$sb.AppendLine("Failed logon report - $env:COMPUTERNAME - window $start .. $(Get-Date)")
[void]$sb.AppendLine("4625 total: $($failed.Count)   RDP 131: $($rdpConn.Count)   RDP 140: $($rdpFail.Count)   $rdpNote")
[void]$sb.AppendLine("RDP enabled: $(if ($rdpDeny -eq 0) {'yes'} elseif ($rdpDeny -eq 1) {'no'} else {'unknown'})   NLA required: $(if ($nla -eq 1) {'yes'} elseif ($nla -eq 0) {'NO'} else {'unknown'})")

Add-Section 'Top target accounts'        ($failed | Group-Object User        | Sort-Object Count -Descending | Select-Object -First 15 Count, Name)
Add-Section 'Top workstation names'      ($failed | Group-Object Workstation | Sort-Object Count -Descending | Select-Object -First 15 Count, Name)
Add-Section 'Source IP from 4625'        ($failed | Group-Object IpAddress   | Sort-Object Count -Descending | Select-Object -First 15 Count, Name)
Add-Section 'PROBABLE source IP (time-correlated with RDP 131 - not proof)' ($failed | Where-Object { $_.ProbableIP } | Group-Object ProbableIP | Sort-Object Count -Descending | Select-Object -First 20 Count, Name)
Add-Section 'RDP 140 bad-credential IPs (logged by RDP itself - authoritative)' ($rdpFail | Group-Object ClientIP | Sort-Object Count -Descending | Select-Object -First 20 Count, Name)
Add-Section 'RDP 131 connecting IPs (all accepted TCP connects)' ($rdpConn | Group-Object ClientIP | Sort-Object Count -Descending | Select-Object -First 20 Count, Name)
Add-Section 'Logon type / process / package' ($failed | Group-Object LogonType, LogonProc, AuthPkg | Sort-Object Count -Descending | Select-Object Count, Name)
Add-Section 'Failure reason'             ($failed | Group-Object Reason | Sort-Object Count -Descending | Select-Object Count, Name)
Add-Section 'Per hour'                   ($failed | Group-Object { $_.Time.ToString('yyyy-MM-dd HH:00') } | Sort-Object Name | Select-Object Name, Count)
Add-Section 'Live sessions on 3389/445 right now' $live
Add-Section 'Firewall profiles / logging' $fwProfiles

$summary = $sb.ToString()
$summary | Out-File (Join-Path $OutDir 'summary.txt') -Encoding UTF8
Write-Host $summary
Write-Host "`nFiles written to $OutDir" -ForegroundColor Green
