#Requires -Version 5.1

<#
.SYNOPSIS
    Monitors Windows services, restarts failed ones, and alerts via email.

.DESCRIPTION
    PRIMARY MODE (v1.2+): Reads monitor-config.json produced by the web admin UI.
    Each service entry carries an 'alerts' field: 'both' | 'dev' | 'iver' | 'none'.
    Alerts are routed to the matching mail group(s) configured in the JSON file.

    LEGACY FALLBACK: If -ConfigFile is omitted and the file does not exist, falls
    back to services.txt + recipients.txt (all alerts go to all recipients, no
    per-service routing). This keeps existing Task Scheduler jobs working without
    any changes.

    Designed for Windows Server 2016+ and compatible with PowerShell 5.1 and 7+.
    Run as Administrator - required for Start-Service on system services.

.PARAMETER ConfigFile
    Path to monitor-config.json managed by the web admin UI.
    Default: .\monitor-config.json
    If OMITTED and the default file does not exist, the script falls back to
    -ServicesFile and -RecipientsFile for backward compatibility. An explicitly
    passed -ConfigFile that does not exist is a fatal error (no silent fallback).

.PARAMETER ServicesFile
    Legacy: Path to text file listing service names. Used only when ConfigFile
    does not exist. Default: .\services.txt

.PARAMETER RecipientsFile
    Legacy: Path to text file listing recipient email addresses. Used only when
    ConfigFile does not exist. Default: .\recipients.txt

.PARAMETER LogFile
    Path to the append-only log file.
    Default: .\ServiceMonitor.log

.PARAMETER LogMaxSizeMB
    Maximum log file size in MB before rotation (renamed to .1 / .2 / .3).
    Default: 10

.PARAMETER SmtpServer
    SMTP relay hostname or IP.
    Default: 192.168.2.15

.PARAMETER SmtpPort
    SMTP relay port.
    Default: 25

.PARAMETER FromAddress
    Sender address used in alert emails.
    Default: servicemonitor@rosvalls.com

.PARAMETER SmtpUser
    SMTP username for authenticated relay. Omit for anonymous relay.

.PARAMETER SmtpCredFile
    Path to an encrypted SMTP password file created by New-CredStore.ps1.
    Must be combined with -SmtpKeyFile.

.PARAMETER SmtpKeyFile
    Path to the AES-256 key file used to decrypt -SmtpCredFile.

.PARAMETER SmtpUseSsl
    If specified, enables SSL/TLS for the SMTP connection.

.PARAMETER SmtpConfigFile
    Path to an smtp.json (written by New-CredStore.ps1) holding the per-deployment
    SMTP settings (server, port, from, user, useSsl). If omitted, looks for
    smtp.json next to -SmtpCredFile, then beside this script. Explicit -Smtp*
    parameters always override the file.

.PARAMETER MaxAttempts
    Number of restart attempts before marking the service as failed.
    Default: 3

.PARAMETER AttemptDelaySeconds
    Seconds to wait between restart attempts.
    Default: 60

.PARAMETER StartSettleSeconds
    Seconds to wait after issuing a start command before re-checking status.
    Default: 8

.PARAMETER AlertCooldownMinutes
    Minimum minutes between repeat alert emails for the same service and
    condition (down / restart failed / not found / disabled / invalid name),
    so a service that stays broken does not flood the inbox every run.
    RECOVERED alerts always send and reset the cooldown for that service.
    0 disables the cooldown (alert on every run). Default: 60

.PARAMETER RunBudgetMinutes
    Soft time budget for a single run. Once exceeded, remaining stopped
    services are still detected and alerted, but restart attempts are skipped
    until the next run - so many simultaneously failed services cannot push
    the run past the scheduled task's execution time limit.
    0 disables the budget (never skip restarts). Default: 20

.PARAMETER NoEmail
    Suppresses all email sending (useful for local testing).

.PARAMETER TestEmail
    Sends a test alert to all configured recipients and exits without checking
    any services. Uses the same routing as a real alert.

.EXAMPLE
    .\ServiceMonitor.ps1

.EXAMPLE
    .\ServiceMonitor.ps1 -ConfigFile C:\ServiceMonitor\monitor-config.json

.EXAMPLE
    .\ServiceMonitor.ps1 -NoEmail

.EXAMPLE
    .\ServiceMonitor.ps1 -TestEmail

.NOTES
    Designed for Windows Server 2016+, PowerShell 5.1 and 7+.
    Run as Administrator for full restart capability.
    Schedule via Install-ScheduledTask.ps1 for continuous monitoring.
    Config file managed by monitor_web.py web admin UI.
    Exit codes: 0 = all monitored services OK (or recovered / nothing to check)
                1 = fatal script error (bad config, unwritable log, unhandled)
                2 = run completed, but at least one service failed
                    (not found / disabled / restart failed)
    Version: 1.3.0
#>

[CmdletBinding()]
param(
    [string] $ConfigFile           = '',
    [string] $ServicesFile         = '',
    [string] $RecipientsFile       = '',
    [string] $LogFile              = '',
    [int]    $LogMaxSizeMB         = 10,
    [string] $SmtpServer           = '192.168.2.15',
    [int]    $SmtpPort             = 25,
    [string] $FromAddress          = 'servicemonitor@rosvalls.com',
    [string] $SmtpUser             = '',
    [string] $SmtpCredFile         = '',
    [string] $SmtpKeyFile          = '',
    [switch] $SmtpUseSsl,
    [string] $SmtpConfigFile        = '',
    [int]    $MaxAttempts          = 3,
    [int]    $AttemptDelaySeconds  = 60,
    [int]    $StartSettleSeconds   = 8,
    [int]    $AlertCooldownMinutes = 60,
    [int]    $RunBudgetMinutes     = 20,
    [switch] $NoEmail,
    [switch] $TestEmail
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Remember whether -ConfigFile was passed explicitly: an explicit path that does
# not exist must be a fatal error, never a silent fallback to legacy files.
$script:ExplicitConfigFile = $PSBoundParameters.ContainsKey('ConfigFile')

# Resolve the script's own directory robustly. $PSScriptRoot is EMPTY inside the
# param() block when the script is launched via 'powershell.exe -File ...' (e.g. the
# SYSTEM scheduled task), which made the old Join-Path defaults throw before any
# logging could start. Resolve here, after param binding, with layered fallbacks.
# If nothing resolves (script text piped / Invoke-Expression), FAIL HARD rather
# than guess a directory - a monitor silently running from the wrong folder
# (e.g. C:\Windows\system32 under SYSTEM) is worse than one that stops loudly.
$ScriptDir = $null
if     ($PSScriptRoot)  { $ScriptDir = $PSScriptRoot }
elseif ($PSCommandPath) { $ScriptDir = Split-Path -Parent $PSCommandPath }
else {
    # StrictMode-safe: MyCommand may be a type without a Path property here.
    $mc = $MyInvocation.MyCommand
    if ($mc -and $mc.PSObject.Properties['Path'] -and $mc.Path) {
        $ScriptDir = Split-Path -Parent $mc.Path
    }
}
if (-not $ScriptDir) {
    throw 'Cannot resolve the script directory (script not run from a file). Pass -ConfigFile and -LogFile explicitly.'
}

if (-not $ConfigFile)     { $ConfigFile     = Join-Path $ScriptDir 'monitor-config.json' }
if (-not $ServicesFile)   { $ServicesFile   = Join-Path $ScriptDir 'services.txt' }
if (-not $RecipientsFile) { $RecipientsFile = Join-Path $ScriptDir 'recipients.txt' }
if (-not $LogFile)        { $LogFile        = Join-Path $ScriptDir 'ServiceMonitor.log' }

# Normalize to absolute paths. Several code paths use .NET file APIs, which
# resolve relative paths against the PROCESS working directory, not the
# PowerShell location - a relative -LogFile would otherwise land elsewhere.
$ConfigFile     = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ConfigFile)
$ServicesFile   = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ServicesFile)
$RecipientsFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($RecipientsFile)
$LogFile        = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($LogFile)
# The SMTP file params default to '' (no cred store) - normalizing an empty
# string throws, so only normalize when set.
if ($SmtpCredFile)   { $SmtpCredFile   = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($SmtpCredFile) }
if ($SmtpKeyFile)    { $SmtpKeyFile    = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($SmtpKeyFile) }
if ($SmtpConfigFile) { $SmtpConfigFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($SmtpConfigFile) }

# Alert cooldown state lives next to the script (see Test-AlertDue below).
$script:StateFile = Join-Path $ScriptDir 'ServiceMonitor.state.json'

# Last-resort error visibility. Any terminating error that escapes to script
# scope is appended to the log and mirrored to the Application event log before
# exiting 1. Without this, a SYSTEM scheduled-task run dies with only stderr
# output that nobody can see (the exact failure mode debugged in v1.2.9).
trap {
    $fatalMsg = "FATAL: unhandled error - $_"
    try {
        $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        Add-Content -LiteralPath $LogFile -Value "$stamp [ERROR  ] $fatalMsg" -Encoding UTF8
    } catch { }
    try {
        if ([System.Diagnostics.EventLog]::SourceExists('ServiceMonitor')) {
            [System.Diagnostics.EventLog]::WriteEntry('ServiceMonitor', $fatalMsg,
                [System.Diagnostics.EventLogEntryType]::Error, 3001)
        }
    } catch { }
    Write-Error -Message $fatalMsg -ErrorAction Continue
    exit 1
}

$ScriptVersion = '1.3.0'
$Hostname      = $env:COMPUTERNAME

$ServiceNamePattern = '^[A-Za-z0-9_.$ -]+$'

$script:SmtpNetworkCredential = $null

# Windows Application event log integration. Source is created on first run (needs
# admin / SYSTEM - the scheduled task account qualifies). If creation fails, file
# logging still works and event writes are skipped.
$script:EventSource   = 'ServiceMonitor'
$script:EventLogName  = 'Application'
$script:EventLogReady = $false

# ---------------------------------------------------------------------------
# Log helpers
# ---------------------------------------------------------------------------

function Initialize-EventLog {
    try {
        if (-not [System.Diagnostics.EventLog]::SourceExists($script:EventSource)) {
            [System.Diagnostics.EventLog]::CreateEventSource($script:EventSource, $script:EventLogName)
        }
        $script:EventLogReady = $true
    } catch {
        $script:EventLogReady = $false
        Write-Warning "Event Log source unavailable (run as admin to enable): $_"
    }
}

function Write-AppEvent {
    param(
        [ValidateSet('Information', 'Warning', 'Error')] [string] $EntryType,
        [string] $Message,
        [int]    $EventId = 1000
    )
    if (-not $script:EventLogReady) { return }
    try {
        [System.Diagnostics.EventLog]::WriteEntry(
            $script:EventSource, $Message,
            [System.Diagnostics.EventLogEntryType]::$EntryType, $EventId)
    } catch { }
}

function Initialize-Log {
    $logDir = Split-Path -LiteralPath $LogFile
    if ($logDir -and -not (Test-Path -LiteralPath $logDir)) {
        try { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
        catch { Write-Warning "Cannot create log directory '$logDir': $_" }
    }
    if (Test-Path -LiteralPath $LogFile) {
        $sizeBytes = (Get-Item -LiteralPath $LogFile).Length
        if ($sizeBytes -gt ($LogMaxSizeMB * 1MB)) {
            $maxBackups = 3
            for ($i = $maxBackups; $i -ge 1; $i--) {
                $src  = if ($i -eq 1) { $LogFile } else { "$LogFile.$($i-1)" }
                $dest = "$LogFile.$i"
                if (Test-Path -LiteralPath $src) { Move-Item -LiteralPath $src -Destination $dest -Force }
            }
        }
    }
    # Zero-length append: verifies writability (and creates the file if missing)
    # without adding a blank line to the log on every run. $LogFile is already
    # normalized to an absolute path, so the .NET call is safe here.
    try { [System.IO.File]::AppendAllText($LogFile, '') }
    catch { throw "Log file '$LogFile' is not writable: $_" }
}

function Write-Log {
    param(
        [ValidateSet('INFO','WARNING','ERROR')]
        [string] $Level,
        [string] $Message
    )
    $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line  = "$stamp [$($Level.PadRight(7))] $Message"
    try { Add-Content -LiteralPath $LogFile -Value $line -Encoding UTF8 }
    catch { Write-Warning "Log write failed: $_" }
    switch ($Level) {
        'INFO'    { Write-Host $line -ForegroundColor Cyan }
        'WARNING' { Write-Host $line -ForegroundColor Yellow }
        'ERROR'   { Write-Host $line -ForegroundColor Red }
    }
    # Mirror actionable levels to the Windows Application event log (INFO is
    # file-only to keep the event log readable). RECOVERED is logged explicitly.
    switch ($Level) {
        'WARNING' { Write-AppEvent -EntryType Warning -Message $Message -EventId 2000 }
        'ERROR'   { Write-AppEvent -EntryType Error   -Message $Message -EventId 3000 }
    }
}

# ---------------------------------------------------------------------------
# Config loading - JSON (v1.2) or legacy flat files
# ---------------------------------------------------------------------------

# Represents a service entry internally
class ServiceEntry {
    [string] $Name
    [bool]   $Paused
    [string] $Alerts   # 'both' | 'dev' | 'iver' | 'none'
}

# Represents loaded config
$script:DevRecipients  = [string[]]@()
$script:IverRecipients = [string[]]@()
$script:AllRecipients  = [string[]]@()
$script:ServiceEntries = [System.Collections.Generic.List[ServiceEntry]]::new()
$script:UsingJsonConfig = $false

# Safe nested-property read. Under Set-StrictMode -Version Latest, touching a
# property that does not exist on a PSCustomObject throws; this returns $null
# instead so a partial/old config file cannot crash the monitor.
function Get-Prop {
    param($Object, [string] $Name)
    if ($null -eq $Object) { return $null }
    if ($Object.PSObject.Properties[$Name]) { return $Object.$Name }
    return $null
}

function Import-MonitorConfig {
    # An explicitly passed -ConfigFile that does not exist is an operator error
    # (typo in the task action). Falling back to a possibly-stale services.txt
    # would silently monitor the wrong set forever - fail loudly instead.
    if ($script:ExplicitConfigFile -and -not (Test-Path -LiteralPath $ConfigFile)) {
        throw "Config file explicitly specified but not found: $ConfigFile"
    }
    if (Test-Path -LiteralPath $ConfigFile) {
        # --- JSON mode ---
        Write-Log -Level INFO -Message "Loading config from: $ConfigFile"
        try {
            $json = Get-Content -LiteralPath $ConfigFile -Raw | ConvertFrom-Json

            $groups    = Get-Prop $json   'groups'
            $devRcpts  = Get-Prop (Get-Prop $groups 'dev')  'recipients'
            $iverRcpts = Get-Prop (Get-Prop $groups 'iver') 'recipients'
            $services  = Get-Prop $json   'services'

            $script:DevRecipients  = @($devRcpts  | Where-Object { $_ -match '@' })
            $script:IverRecipients = @($iverRcpts | Where-Object { $_ -match '@' })
            $script:AllRecipients  = @($script:DevRecipients + $script:IverRecipients | Select-Object -Unique)

            foreach ($s in $services) {
                $name = Get-Prop $s 'name'
                if (-not $name) { continue }
                $entry         = [ServiceEntry]::new()
                $entry.Name    = $name
                $entry.Paused  = [bool](Get-Prop $s 'paused')
                $sAlerts       = Get-Prop $s 'alerts'
                $entry.Alerts  = if ($sAlerts -in @('both','dev','iver','none')) { $sAlerts } else { 'both' }
                $script:ServiceEntries.Add($entry)
            }
            $script:UsingJsonConfig = $true

            Write-Log -Level INFO -Message ("JSON config: {0} services, dev={1} iver={2} recipients" -f `
                $script:ServiceEntries.Count, $script:DevRecipients.Count, $script:IverRecipients.Count)
        } catch {
            throw "Failed to parse $ConfigFile : $_"
        }
    } else {
        # --- Legacy flat-file mode ---
        Write-Log -Level WARNING -Message "monitor-config.json not found - falling back to services.txt / recipients.txt"

        # Load services
        $lines  = Read-ListFile -Path $ServicesFile -Label 'Services'
        $parsed = ConvertFrom-LegacyServices -Lines $lines
        $active = $parsed.Active
        $paused = $parsed.Paused
        foreach ($n in $active) {
            $e = [ServiceEntry]::new(); $e.Name = $n; $e.Paused = $false; $e.Alerts = 'both'
            $script:ServiceEntries.Add($e)
        }
        foreach ($n in $paused) {
            $e = [ServiceEntry]::new(); $e.Name = $n; $e.Paused = $true; $e.Alerts = 'both'
            $script:ServiceEntries.Add($e)
        }

        # Load recipients (all go to AllRecipients; dev/iver not used in legacy mode)
        $rcptLines = @(Read-ListFile -Path $RecipientsFile -Label 'Recipients')
        $script:AllRecipients  = @($rcptLines | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '@' })
        $script:DevRecipients  = $script:AllRecipients
        $script:IverRecipients = $script:AllRecipients
    }
}

function Get-AlertRecipients {
    param([string] $AlertsValue)
    switch ($AlertsValue) {
        'dev'   { return $script:DevRecipients  }
        'iver'  { return $script:IverRecipients }
        'none'  { return [string[]]@()          }
        default { return $script:AllRecipients  }  # 'both' + unknown values
    }
}

# ---------------------------------------------------------------------------
# SMTP credential loader
# ---------------------------------------------------------------------------

function Initialize-SmtpCredential {
    if (-not $SmtpCredFile -and -not $SmtpKeyFile) { return }
    if (-not $SmtpCredFile -or -not $SmtpKeyFile) {
        Write-Log -Level WARNING -Message '-SmtpCredFile and -SmtpKeyFile must both be specified - SMTP auth disabled.'
        return
    }
    if (-not (Test-Path -LiteralPath $SmtpCredFile)) {
        Write-Log -Level WARNING -Message "SMTP cred file not found: $SmtpCredFile - SMTP auth disabled."
        return
    }
    if (-not (Test-Path -LiteralPath $SmtpKeyFile)) {
        Write-Log -Level WARNING -Message "SMTP key file not found: $SmtpKeyFile - SMTP auth disabled."
        return
    }
    try {
        $keyBytes = [System.IO.File]::ReadAllBytes($SmtpKeyFile)
        $encStr   = [System.IO.File]::ReadAllText($SmtpCredFile).Trim()
        $secPwd   = ConvertTo-SecureString -String $encStr -Key $keyBytes
        $script:SmtpNetworkCredential = New-Object System.Net.NetworkCredential($SmtpUser, $secPwd)
        Write-Log -Level INFO -Message "SMTP credentials loaded (user: $SmtpUser, SSL: $($SmtpUseSsl.IsPresent))"
    } catch {
        Write-Log -Level WARNING -Message "Failed to load SMTP credentials: $_ - SMTP auth disabled."
    }
}

# ---------------------------------------------------------------------------
# Email
# ---------------------------------------------------------------------------

function Send-Alert {
    param(
        [string[]] $Recipients,
        [string]   $Subject,
        [string]   $Body
    )
    if ($NoEmail) {
        Write-Log -Level INFO -Message "[NoEmail] Would send: $Subject"
        return $true
    }
    # Case-insensitive dedup: 'both' merges the dev+iver lists, and
    # User@x.com / user@x.com must not receive the alert twice.
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $rcptArray = @($Recipients | Where-Object { $_ -match '@' -and $seen.Add($_) })
    if ($rcptArray.Count -eq 0) {
        Write-Log -Level WARNING -Message 'No recipients for this alert - skipping email.'
        return $false
    }
    $smtp = $null
    $msg  = $null
    try {
        $smtp = New-Object System.Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
        $smtp.EnableSsl             = $SmtpUseSsl.IsPresent
        $smtp.UseDefaultCredentials = $false
        $smtp.Timeout               = 15000
        if ($null -ne $script:SmtpNetworkCredential) {
            $smtp.Credentials = $script:SmtpNetworkCredential
        }
        $msg = New-Object System.Net.Mail.MailMessage
        $msg.From       = New-Object System.Net.Mail.MailAddress($FromAddress, "Service Monitor [$Hostname]")
        $msg.Subject    = $Subject
        $msg.Body       = $Body
        $msg.IsBodyHtml = $false
        foreach ($addr in $rcptArray) { $msg.To.Add($addr) }
        $smtp.Send($msg)
        Write-Log -Level INFO -Message "Alert sent to: $($rcptArray -join ', ')"
        return $true
    } catch {
        Write-Log -Level ERROR -Message "Failed to send email alert: $_"
        return $false
    } finally {
        # Dispose even if Send threw, so sockets/handles are not leaked
        if ($null -ne $msg)  { $msg.Dispose() }
        if ($null -ne $smtp) { $smtp.Dispose() }
    }
}

# ---------------------------------------------------------------------------
# Alert cooldown state - prevents an alert storm when a service stays broken.
# State is a small JSON map of 'service|condition' -> last-alerted timestamp.
# A repeat alert for the same key is suppressed until AlertCooldownMinutes has
# passed. RECOVERED / OK clears the service's keys so a NEW failure after a
# recovery always alerts immediately.
# ---------------------------------------------------------------------------

function Get-AlertState {
    $state = @{}
    if (Test-Path -LiteralPath $script:StateFile) {
        try {
            $raw = Get-Content -LiteralPath $script:StateFile -Raw | ConvertFrom-Json
            foreach ($p in $raw.PSObject.Properties) { $state[$p.Name] = [string]$p.Value }
        } catch {
            Write-Log -Level WARNING -Message "Alert state file unreadable - starting fresh: $_"
        }
    }
    return $state
}

function Save-AlertState {
    param([hashtable] $State)
    try {
        if ($State.Count -eq 0) {
            if (Test-Path -LiteralPath $script:StateFile) {
                Remove-Item -LiteralPath $script:StateFile -Force
            }
        } else {
            # UTF8, not ASCII: a service name containing non-ASCII letters (for
            # example a Swedish A-ring) written as ASCII is mangled to '?', the
            # key never matches on the next run,
            # and the cooldown never suppresses - a permanent alert storm.
            $State | ConvertTo-Json | Set-Content -LiteralPath $script:StateFile -Encoding UTF8
        }
    } catch {
        Write-Log -Level WARNING -Message "Alert state file write failed: $_"
    }
}

function Test-AlertDue {
    # Check only - does NOT stamp. The timestamp is committed by Set-AlertSent
    # after Send-Alert reports success; stamping here would suppress the next
    # alert for a full cooldown window even when the send FAILED (SMTP blip),
    # silently losing the first alert of an outage.
    param([string] $Key)
    if ($AlertCooldownMinutes -le 0) { return $true }
    $state = Get-AlertState
    if ($state.ContainsKey($Key)) {
        # Timestamps are saved in round-trip ('o') format; parse invariant so
        # the system locale (e.g. sv-SE under SYSTEM) cannot break comparison.
        $last = [datetime]::MinValue
        $parsed = [datetime]::TryParse($state[$Key],
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::RoundtripKind, [ref]$last)
        if ($parsed -and (Get-Date) -lt $last.AddMinutes($AlertCooldownMinutes)) {
            Write-Log -Level INFO -Message "Alert '$Key' suppressed (cooldown $AlertCooldownMinutes min, last sent $($state[$Key]))."
            return $false
        }
    }
    return $true
}

function Set-AlertSent {
    # Commit the cooldown timestamp - call ONLY after a successful send.
    # Never under -NoEmail: a simulated send must not suppress the scheduled
    # task's next real alert.
    param([string] $Key)
    if ($AlertCooldownMinutes -le 0 -or $NoEmail) { return }
    $state = Get-AlertState
    $state[$Key] = (Get-Date).ToString('o')
    Save-AlertState -State $state
}

function Clear-AlertState {
    param([string] $ServiceName)
    $state = Get-AlertState
    $keys = @($state.Keys | Where-Object { $_ -like "$ServiceName|*" })
    if ($keys.Count -gt 0) {
        foreach ($k in $keys) { $state.Remove($k) }
        Save-AlertState -State $state
        Write-Log -Level INFO -Message "Alert cooldown cleared for '$ServiceName'."
    }
}

function New-AlertBody {
    param(
        [string] $ServiceName,
        [string] $Status,
        [int]    $Attempt,
        [string] $ErrorDetail,
        [string] $AlertGroup = 'both'
    )
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $groupLabel = switch ($AlertGroup) {
        'dev'  { 'Dev Team only' }
        'iver' { 'Iver Support only' }
        'none' { 'none (should not be reached)' }
        default { 'Both groups' }
    }
    $lines = @(
        '=== Service Monitor Alert ==='
        "Host      : $Hostname"
        "Time      : $ts"
        "Service   : $ServiceName"
        "Status    : $Status"
        "Sent to   : $groupLabel"
    )
    if ($Attempt -gt 0) { $lines += "Attempt   : $Attempt / $MaxAttempts" }
    if ($ErrorDetail) {
        $lines += ''
        $lines += '--- Error Detail ---'
        $lines += $ErrorDetail
    }
    if (Test-Path -LiteralPath $LogFile) {
        try {
            # Filter 'Alert sent to:' lines: a dev-only alert body must not leak
            # the iver group's addresses (and vice versa) via the log tail.
            $tail = @([System.IO.File]::ReadAllLines($LogFile) |
                      Where-Object { $_ -notmatch 'Alert sent to' } |
                      Select-Object -Last 20)
            if ($tail.Count -gt 0) {
                $lines += ''
                $lines += '--- Recent Log (last 20 lines) ---'
                $lines += $tail
            }
        } catch { <# non-fatal #> }
    }
    $lines += ''
    $lines += "--- Log File ---`n$LogFile"
    $lines += ''
    $lines += "ServiceMonitor.ps1 v$ScriptVersion"
    return $lines -join "`r`n"
}

# ---------------------------------------------------------------------------
# Legacy file loaders (used when ConfigFile is absent)
# ---------------------------------------------------------------------------

function Read-ListFile {
    param([string] $Path, [string] $Label)
    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Log -Level ERROR -Message "$Label file not found: $Path"
        return @()
    }
    $lines = [System.IO.File]::ReadAllLines($Path) |
             Where-Object { $_ -match '\S' -and $_ -notmatch '^\s*#' }
    return @($lines)
}

function ConvertFrom-LegacyServices {
    param([string[]] $Lines)
    $active = [System.Collections.Generic.List[string]]::new()
    $paused = [System.Collections.Generic.List[string]]::new()
    foreach ($line in $Lines) {
        $trimmed = $line.Trim()
        if (-not $trimmed -or $trimmed -match '^#') { continue }
        if ($trimmed -match '^[\*\-]\s*(.+)') {
            $name = $Matches[1].Trim()
            if ($name) { $paused.Add($name) }
        } else {
            $active.Add($trimmed)
        }
    }
    # Return a single object. A bare 'return $a, $b' drops an empty array in the
    # pipeline, so '$active, $paused = ...' would misalign when one list is empty.
    return [pscustomobject]@{
        Active = [string[]]$active.ToArray()
        Paused = [string[]]$paused.ToArray()
    }
}

# ---------------------------------------------------------------------------
# Service helpers
# ---------------------------------------------------------------------------

function Get-ServiceStartupType {
    param([string] $ServiceName)
    # Safe by construction: the name is interpolated into a WQL filter below, so
    # guard here too instead of relying on the caller having validated it.
    if ($ServiceName -notmatch $ServiceNamePattern) { return 'Unknown' }
    $wmi = Get-CimInstance -ClassName Win32_Service -Filter "Name='$ServiceName'" -ErrorAction SilentlyContinue
    if ($null -eq $wmi) { return 'Unknown' }
    return $wmi.StartMode
}

function Get-ServiceRunningStatus {
    param([string] $ServiceName)
    $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($null -eq $svc) { return $null }
    return $svc.Status
}

function Wait-ServiceState {
    param(
        [string] $ServiceName,
        [System.ServiceProcess.ServiceControllerStatus] $TargetState,
        [int]    $TimeoutSeconds = 15
    )
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        $status = Get-ServiceRunningStatus -ServiceName $ServiceName
        if ($status -eq $TargetState) { return $true }
        Start-Sleep -Seconds 1
    }
    return $false
}

function Invoke-ServiceRestart {
    param([string] $ServiceName)
    $result = @{ Success = $false; ErrorText = '' }
    # Safe by construction: the name reaches sc.exe below, so guard here too.
    if ($ServiceName -notmatch $ServiceNamePattern) {
        $result.ErrorText = "Invalid service name rejected: '$ServiceName'"
        return $result
    }
    try {
        $status = Get-ServiceRunningStatus -ServiceName $ServiceName
        # A StopPending service is already on its way down (e.g. an admin's slow
        # manual restart). Issuing a start now fails instantly and produces false
        # RESTART FAILED alerts - wait for Stopped first, then proceed.
        if ($status -eq [System.ServiceProcess.ServiceControllerStatus]::StopPending) {
            Write-Log -Level INFO -Message "'$ServiceName' is StopPending - waiting up to 60 s for it to stop..."
            $null = Wait-ServiceState -ServiceName $ServiceName `
                -TargetState ([System.ServiceProcess.ServiceControllerStatus]::Stopped) `
                -TimeoutSeconds 60
            $status = Get-ServiceRunningStatus -ServiceName $ServiceName
        }
        if ($null -ne $status -and $status -notin @(
            [System.ServiceProcess.ServiceControllerStatus]::Stopped,
            [System.ServiceProcess.ServiceControllerStatus]::StopPending
        )) {
            Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue
            $stopped = Wait-ServiceState -ServiceName $ServiceName `
                -TargetState ([System.ServiceProcess.ServiceControllerStatus]::Stopped) `
                -TimeoutSeconds 20
            if (-not $stopped) {
                $result.ErrorText = 'Service did not reach Stopped state within 20 s; attempting start anyway.'
                Write-Log -Level WARNING -Message $result.ErrorText
            }
        }
        # sc.exe can write to stderr; with $ErrorActionPreference='Stop' a 2>&1
        # capture would turn that into a terminating NativeCommandError before
        # $LASTEXITCODE is read. Capture under 'Continue', judge by exit code.
        $prevEAP = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        try {
            $scOutput = @(& sc.exe start $ServiceName 2>&1)
            $exitCode = $LASTEXITCODE
        } finally {
            $ErrorActionPreference = $prevEAP
        }
        if ($exitCode -ne 0) {
            $result.ErrorText = "sc.exe exit $exitCode - $($scOutput -join ' ')"
        } else {
            $result.Success = $true
        }
    } catch {
        $result.ErrorText = $_.ToString()
    }
    return $result
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

Initialize-EventLog
Initialize-Log
Write-Log -Level INFO -Message "=== ServiceMonitor v$ScriptVersion starting on $Hostname ==="

# Load per-deployment SMTP settings (server/port/from/user/ssl) written by
# New-CredStore.ps1, so each server uses its own relay without editing the script.
# Explicit -Smtp* / -FromAddress parameters always win; the file fills the rest.
$smtpCfgPath = if ($SmtpConfigFile)    { $SmtpConfigFile }
               elseif ($SmtpCredFile)  { Join-Path ([System.IO.Path]::GetDirectoryName([System.IO.Path]::GetFullPath($SmtpCredFile))) 'smtp.json' }
               else                    { Join-Path $ScriptDir 'smtp.json' }
if (Test-Path -LiteralPath $smtpCfgPath) {
    try {
        $smtpCfg = Get-Content -LiteralPath $smtpCfgPath -Raw | ConvertFrom-Json
        if ((Get-Prop $smtpCfg 'server') -and -not $PSBoundParameters.ContainsKey('SmtpServer'))  { $SmtpServer  = [string]$smtpCfg.server }
        if ((Get-Prop $smtpCfg 'port')   -and -not $PSBoundParameters.ContainsKey('SmtpPort'))    { $SmtpPort    = [int]   $smtpCfg.port }
        if ((Get-Prop $smtpCfg 'from')   -and -not $PSBoundParameters.ContainsKey('FromAddress')) { $FromAddress = [string]$smtpCfg.from }
        if ((Get-Prop $smtpCfg 'user')   -and -not $PSBoundParameters.ContainsKey('SmtpUser'))    { $SmtpUser    = [string]$smtpCfg.user }
        if ((Get-Prop $smtpCfg 'useSsl') -and -not $PSBoundParameters.ContainsKey('SmtpUseSsl'))  { $SmtpUseSsl  = [switch]([bool]$smtpCfg.useSsl) }
        Write-Log -Level INFO -Message "Loaded SMTP settings from ${smtpCfgPath} (server ${SmtpServer}:${SmtpPort}, ssl: $($SmtpUseSsl.IsPresent))"
    } catch {
        Write-Log -Level WARNING -Message "Failed to read SMTP config ${smtpCfgPath}: $_"
    }
}

Initialize-SmtpCredential
Import-MonitorConfig

$activeEntries = @($script:ServiceEntries | Where-Object { -not $_.Paused })
$pausedEntries = @($script:ServiceEntries | Where-Object { $_.Paused })

# TestEmail mode - short-circuit before any service logging or checking.
if ($TestEmail) {
    Write-Log -Level INFO -Message '=== TestEmail mode: sending connectivity test ==='
    $body = New-AlertBody -ServiceName '(test)' -Status 'TEST ALERT' -Attempt 0 `
                            -ErrorDetail 'This is a connectivity test sent by -TestEmail. No services were checked.' `
                            -AlertGroup 'both'
    $sent = Send-Alert -Recipients $script:AllRecipients `
               -Subject "[ServiceMonitor] TEST: connectivity check from $Hostname" `
               -Body $body
    if ($sent) {
        Write-Log -Level INFO -Message '=== TestEmail complete ==='
        exit 0
    }
    # A test that could not send must not report success - operators script
    # against this exit code to verify SMTP before scheduling.
    Write-Log -Level ERROR -Message '=== TestEmail FAILED - see the error above ==='
    exit 1
}

# Note: build the name lists with ForEach-Object (not .Name member access) so an
# empty collection does not trip Set-StrictMode -Version Latest, and wrap counts
# in @() so a scalar/$null never throws 'property Count cannot be found'.
if (@($pausedEntries).Count -gt 0) {
    $pausedNames = @($pausedEntries | ForEach-Object { $_.Name }) -join ', '
    Write-Log -Level INFO -Message "Paused (skipped): $pausedNames"
}
$activeNames = @($activeEntries | ForEach-Object { $_.Name }) -join ', '
Write-Log -Level INFO -Message "Active services to check ($(@($activeEntries).Count)): $activeNames"

if (@($activeEntries).Count -eq 0) {
    # Not an error: a fresh install has zero services until the operator adds
    # them in the web UI, and pausing everything during maintenance is a
    # deliberate choice. Exiting 1 here spammed Error events every 5 minutes.
    if (@($script:ServiceEntries).Count -gt 0) {
        Write-Log -Level INFO -Message 'All services are paused - nothing to check this run.'
    } else {
        Write-Log -Level WARNING -Message 'No services configured yet - add services in the web admin UI.'
    }
    exit 0
}

# Check each service
$failureCount = 0

# Soft time budget: with many simultaneously failed services the serial restart
# loops (attempts + delays + SMTP timeouts) could exceed the scheduled task's
# 30-minute execution limit and get killed mid-run. Past the budget, remaining
# down services are still detected and alerted but not restarted this run.
# 0 (or negative) disables the budget. AddMinutes would make the deadline
# 'now' and skip every restart - MaxValue means 'never'.
$runDeadline = if ($RunBudgetMinutes -le 0) { [datetime]::MaxValue }
               else { (Get-Date).AddMinutes($RunBudgetMinutes) }

foreach ($entry in $activeEntries) {
    $svcName   = $entry.Name
    # Wrap in @() so $alertsTo is always an array: Get-AlertRecipients can return
    # an unrolled empty array (-> $null) or a scalar, and the .Count checks below
    # would otherwise throw under Set-StrictMode -Version Latest.
    $alertsTo  = @(Get-AlertRecipients -AlertsValue $entry.Alerts)

    if ($svcName -notmatch $ServiceNamePattern) {
        $msg = "Rejected invalid service name: '$svcName' - fix the entry in the web admin."
        Write-Log -Level ERROR -Message $msg
        if ($alertsTo.Count -gt 0 -and (Test-AlertDue -Key "$svcName|invalid")) {
            $body = New-AlertBody -ServiceName $svcName -Status 'INVALID NAME rejected' -Attempt 0 `
                                    -ErrorDetail $msg -AlertGroup $entry.Alerts
            if (Send-Alert -Recipients $alertsTo `
                    -Subject "[ServiceMonitor] ERROR: invalid service name in config on $Hostname" -Body $body) {
                Set-AlertSent -Key "$svcName|invalid"
            }
        }
        $failureCount++
        continue
    }

    Write-Log -Level INFO -Message "Checking: $svcName  [alerts->$($entry.Alerts)]"

    $svcObj = Get-Service -Name $svcName -ErrorAction SilentlyContinue
    if ($null -eq $svcObj) {
        $msg  = "Service '$svcName' not found on $Hostname - verify the name in the web admin."
        Write-Log -Level ERROR -Message $msg
        if ($alertsTo.Count -gt 0 -and (Test-AlertDue -Key "$svcName|notfound")) {
            $body = New-AlertBody -ServiceName $svcName -Status 'NOT FOUND on host' -Attempt 0 `
                                    -ErrorDetail $msg -AlertGroup $entry.Alerts
            if (Send-Alert -Recipients $alertsTo `
                    -Subject "[ServiceMonitor] ERROR: $svcName not found on $Hostname" -Body $body) {
                Set-AlertSent -Key "$svcName|notfound"
            }
        }
        $failureCount++
        continue
    }

    if ($svcObj.Status -eq [System.ServiceProcess.ServiceControllerStatus]::StartPending) {
        Write-Log -Level INFO -Message "'$svcName' is StartPending - waiting up to 30 s..."
        $started = Wait-ServiceState -ServiceName $svcName `
            -TargetState ([System.ServiceProcess.ServiceControllerStatus]::Running) -TimeoutSeconds 30
        if ($started) {
            Write-Log -Level INFO -Message "OK: '$svcName' reached Running (was StartPending)."
            Clear-AlertState -ServiceName $svcName
            continue
        }
        Write-Log -Level WARNING -Message "'$svcName' still not Running after 30 s - proceeding with restart logic."
    }

    $currentStatus = Get-ServiceRunningStatus -ServiceName $svcName
    if ($currentStatus -eq [System.ServiceProcess.ServiceControllerStatus]::Running) {
        Write-Log -Level INFO -Message "OK: '$svcName' is running."
        Clear-AlertState -ServiceName $svcName
        continue
    }

    $startupType = Get-ServiceStartupType -ServiceName $svcName
    if ($startupType -eq 'Disabled') {
        $msg = "Service '$svcName' is DISABLED. Manual intervention required - skipping restart."
        Write-Log -Level ERROR -Message $msg
        if ($alertsTo.Count -gt 0 -and (Test-AlertDue -Key "$svcName|disabled")) {
            $body = New-AlertBody -ServiceName $svcName -Status 'DOWN - StartupType=Disabled' `
                                    -Attempt 0 -ErrorDetail $msg -AlertGroup $entry.Alerts
            if (Send-Alert -Recipients $alertsTo `
                    -Subject "[ServiceMonitor] ERROR: $svcName is DISABLED on $Hostname" -Body $body) {
                Set-AlertSent -Key "$svcName|disabled"
            }
        }
        $failureCount++
        continue
    }

    Write-Log -Level WARNING -Message "Service '$svcName' is NOT running (status: $currentStatus, startup: $startupType)"
    if ($alertsTo.Count -gt 0) {
        if (Test-AlertDue -Key "$svcName|down") {
            $warnBody = New-AlertBody -ServiceName $svcName -Status "DOWN ($currentStatus)" `
                                        -Attempt 0 -ErrorDetail '' -AlertGroup $entry.Alerts
            if (Send-Alert -Recipients $alertsTo `
                    -Subject "[ServiceMonitor] WARNING: $svcName is down on $Hostname" -Body $warnBody) {
                Set-AlertSent -Key "$svcName|down"
            }
        }
    } else {
        # Distinguish 'alerts=none' (deliberate) from a group with no recipients
        # configured (probably a mistake) - the old message conflated the two.
        Write-Log -Level INFO -Message "No alert email for '$svcName' (alerts=$($entry.Alerts), resolved recipients: $($alertsTo.Count))."
    }

    # Past the run budget: alert-only mode for the rest of this run.
    if ((Get-Date) -gt $runDeadline) {
        Write-Log -Level WARNING -Message "Run time budget ($RunBudgetMinutes min) exceeded - skipping restart attempts for '$svcName' until next run."
        $failureCount++
        continue
    }

    $restarted = $false
    $lastError  = ''

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        Write-Log -Level WARNING -Message "Restart attempt $attempt / $MaxAttempts for '$svcName'..."
        $restartResult = Invoke-ServiceRestart -ServiceName $svcName
        if (-not $restartResult.Success) {
            $lastError = $restartResult.ErrorText
            Write-Log -Level WARNING -Message "sc.exe reported failure on attempt ${attempt}: $lastError"
        }
        Write-Log -Level INFO -Message "Settling $StartSettleSeconds s after start command..."
        Start-Sleep -Seconds $StartSettleSeconds
        $nowStatus = Get-ServiceRunningStatus -ServiceName $svcName
        if ($nowStatus -eq [System.ServiceProcess.ServiceControllerStatus]::Running) {
            Write-Log -Level INFO -Message "Service '$svcName' recovered on attempt $attempt."
            Write-AppEvent -EntryType Information -EventId 1001 `
                -Message "Service '$svcName' recovered on attempt $attempt/$MaxAttempts on $Hostname."
            Clear-AlertState -ServiceName $svcName
            if ($alertsTo.Count -gt 0) {
                # RECOVERED always sends - it closes the loop the DOWN alert opened.
                $body = New-AlertBody -ServiceName $svcName -Status 'RECOVERED' `
                                        -Attempt $attempt -ErrorDetail '' -AlertGroup $entry.Alerts
                $null = Send-Alert -Recipients $alertsTo `
                           -Subject "[ServiceMonitor] RECOVERED: $svcName on $Hostname (attempt $attempt/$MaxAttempts)" `
                           -Body $body
            }
            $restarted = $true
            break
        }
        if ($attempt -lt $MaxAttempts) {
            Write-Log -Level INFO -Message "Still not running. Waiting $AttemptDelaySeconds s before attempt $($attempt+1)..."
            Start-Sleep -Seconds $AttemptDelaySeconds
        }
    }

    if (-not $restarted) {
        $errMsg = "Service '$svcName' failed to restart after $MaxAttempts attempts. Last error: $lastError"
        Write-Log -Level ERROR -Message $errMsg
        if ($alertsTo.Count -gt 0 -and (Test-AlertDue -Key "$svcName|failed")) {
            $body = New-AlertBody -ServiceName $svcName `
                                    -Status "RESTART FAILED after $MaxAttempts attempts" `
                                    -Attempt $MaxAttempts -ErrorDetail $lastError -AlertGroup $entry.Alerts
            if (Send-Alert -Recipients $alertsTo `
                    -Subject "[ServiceMonitor] ERROR: $svcName restart FAILED on $Hostname" -Body $body) {
                Set-AlertSent -Key "$svcName|failed"
            }
        }
        $failureCount++
    }
}

Write-Log -Level INFO -Message "=== Check complete. Failures: $failureCount / $($activeEntries.Count) active services ==="

# Exit 2 (not 1) for service failures: LastTaskResult must distinguish 'the
# script crashed' (1) from 'the script worked but a service is broken' (2).
if ($failureCount -gt 0) { exit 2 }
exit 0
