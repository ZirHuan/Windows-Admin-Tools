#Requires -Version 5.1
<#
.SYNOPSIS
    Finds clients doing unsigned / simple-bind LDAP against domain controllers
    (Directory Service events 2889 and 2887). Read-only.

.DESCRIPTION
    Before enforcing LDAP signing / channel binding you need to know who would break.
    For each DC this reads, over PowerShell remoting:

      2889  one event per unsigned or simple bind, with the client IP. Logged ONLY when
            HKLM\SYSTEM\CurrentControlSet\Services\NTDS\Diagnostics
            '16 LDAP Interface Events' is 2 or higher.
      2887  a once-a-day count of unsigned binds. Always logged.

    So: clients listed = who to fix. 2887 > 0 with no clients = logging level too low to
    see who; raise it (see -ShowEnableCommand) and run again after a day.

    Used on its own, and as a building block by:
      DomainScanner/Invoke-DomainHealthAudit.ps1   (NET section)

.PARAMETER ComputerName
    DCs to query. Default: every DC in the computer's domain (needs the ActiveDirectory
    module for discovery only).

.PARAMETER HoursBack
    Hours of Directory Service log to read. Default 24. 2887 is written once a day, so
    use at least 24.

.PARAMETER ShowEnableCommand
    Print the command that turns per-client logging on (does not run it).

.OUTPUTS
    PSCustomObject per DC: ComputerName, LoggingLevel, Clients (string[] of IPs),
    ClientCount, Count2887, Status (UNSIGNED/UNKNOWN-CLIENTS/CLEAN/CLEAN-2887-ONLY/ERROR),
    Error.

.EXAMPLE
    .\Get-UnsignedLdapClient.ps1 | Format-Table ComputerName,Status,ClientCount,Count2887

.EXAMPLE
    (.\Get-UnsignedLdapClient.ps1 -HoursBack 168).Clients | Sort-Object -Unique
    Every client IP seen doing unsigned binds in the last week, all DCs.

.NOTES
    Version : 1.0.0
    Read-only. Needs WinRM to the DCs and rights to read the Directory Service log.

    Changelog:
      1.0.0  2026-11-19  First version. Extracted from Invoke-DomainHealthAudit (NET section).
#>

[CmdletBinding()]
param(
    [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName, Position = 0)]
    [Alias('HostName', 'DomainController')]
    [string[]]$ComputerName,

    [ValidateRange(1, 720)]
    [int]$HoursBack = 24,

    [switch]$ShowEnableCommand
)

begin {
    $ScriptVersion = '1.0.0'
    Write-Verbose "Get-UnsignedLdapClient v$ScriptVersion"
    if ($ShowEnableCommand) {
        Write-Host 'Turn on per-client 2889 logging on a DC (no restart needed):' -ForegroundColor Cyan
        Write-Host "  Set-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Services\NTDS\Diagnostics' -Name '16 LDAP Interface Events' -Value 2 -Type DWord"
        Write-Host 'Turn it back off (0) when done - level 2 logs one event per bind.' -ForegroundColor DarkGray
    }
    $all = [System.Collections.Generic.List[string]]::new()

    $remote = {
        param($Hours)
        $start = (Get-Date).AddHours(-1 * $Hours)
        $lvl = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Services\NTDS\Diagnostics' -ErrorAction SilentlyContinue).'16 LDAP Interface Events'
        $out = [ordered]@{ Level = [int]$lvl; Clients = @(); Count2887 = 0; Error = '' }
        try {
            $out.Clients = @(Get-WinEvent -FilterHashtable @{ LogName = 'Directory Service'; Id = 2889; StartTime = $start } -ErrorAction Stop |
                ForEach-Object {
                    $ep = [string]$_.Properties[0].Value           # "ip:port" or "[v6]:port"
                    $i  = $ep.LastIndexOf(':')
                    if ($i -gt 0) { $ep.Substring(0, $i).Trim('[', ']') } else { $ep }
                } | Sort-Object -Unique)
        } catch { if ("$($_.FullyQualifiedErrorId)" -notmatch 'NoMatchingEventsFound') { $out.Error = "$_" } }
        try {
            # 2887 = daily summary. Properties[0] = simple binds without SSL/TLS,
            # Properties[1] = Negotiate/Kerberos/NTLM/Digest binds without signing.
            $n = 0
            foreach ($e in @(Get-WinEvent -FilterHashtable @{ LogName = 'Directory Service'; Id = 2887; StartTime = $start } -ErrorAction Stop)) {
                foreach ($i in 0, 1) {
                    $v = 0
                    if ($e.Properties.Count -gt $i -and [int]::TryParse("$($e.Properties[$i].Value)", [ref]$v)) { $n += $v }
                }
            }
            $out.Count2887 = $n
        } catch { if ("$($_.FullyQualifiedErrorId)" -notmatch 'NoMatchingEventsFound' -and -not $out.Error) { $out.Error = "$_" } }
        [PSCustomObject]$out
    }
}

process {
    foreach ($c in $ComputerName) { if ($c) { $all.Add($c) } }
}

end {
    if ($all.Count -eq 0) {
        try {
            Import-Module ActiveDirectory -ErrorAction Stop
            foreach ($dc in @(Get-ADDomainController -Filter * -ErrorAction Stop)) { $all.Add($dc.HostName) }
        } catch {
            throw "No -ComputerName given and DC discovery failed (ActiveDirectory module?): $_"
        }
    }

    foreach ($dc in $all) {
        $res = [ordered]@{
            ComputerName = $dc; LoggingLevel = $null; Clients = @(); ClientCount = 0
            Count2887 = 0; Status = 'ERROR'; Error = ''
        }
        try {
            $r = Invoke-Command -ComputerName $dc -ArgumentList $HoursBack -ScriptBlock $remote -ErrorAction Stop
            $res.LoggingLevel = $r.Level
            $res.Clients      = @($r.Clients)
            $res.ClientCount  = @($r.Clients).Count
            $res.Count2887    = [int]$r.Count2887
            $res.Error        = "$($r.Error)"
            $res.Status = if ($r.Error)                   { 'ERROR' }
                          elseif ($res.ClientCount -gt 0) { 'UNSIGNED' }
                          elseif ($res.Count2887 -gt 0)   { 'UNKNOWN-CLIENTS' }   # binds happen, logging too low to see who
                          elseif ($r.Level -lt 2)         { 'CLEAN-2887-ONLY' }   # only the daily counter to go on
                          else                            { 'CLEAN' }
        } catch {
            $res.Error = "Could not query $dc (remoting blocked?): $_"
        }
        [PSCustomObject]$res
    }
}
