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
    If this file does not exist, the script falls back to -ServicesFile and
    -RecipientsFile for backward compatibility.

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
    Version: 1.2.3
#>

[CmdletBinding()]
param(
    [string] $ConfigFile           = (Join-Path $PSScriptRoot 'monitor-config.json'),
    [string] $ServicesFile         = (Join-Path $PSScriptRoot 'services.txt'),
    [string] $RecipientsFile       = (Join-Path $PSScriptRoot 'recipients.txt'),
    [string] $LogFile              = (Join-Path $PSScriptRoot 'ServiceMonitor.log'),
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
    [switch] $NoEmail,
    [switch] $TestEmail
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ScriptVersion = '1.2.3'
$Hostname      = $env:COMPUTERNAME

$ServiceNamePattern = '^[A-Za-z0-9_.$ -]+$'

$script:SmtpNetworkCredential = $null

# ---------------------------------------------------------------------------
# Log helpers
# ---------------------------------------------------------------------------

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
    try { Add-Content -LiteralPath $LogFile -Value '' -Encoding UTF8 }
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
        return
    }
    $rcptArray = @($Recipients | Where-Object { $_ -match '@' })
    if ($rcptArray.Count -eq 0) {
        Write-Log -Level WARNING -Message 'No recipients for this alert - skipping email.'
        return
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
    } catch {
        Write-Log -Level ERROR -Message "Failed to send email alert: $_"
    } finally {
        # Dispose even if Send threw, so sockets/handles are not leaked
        if ($null -ne $msg)  { $msg.Dispose() }
        if ($null -ne $smtp) { $smtp.Dispose() }
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
            $tail = @([System.IO.File]::ReadAllLines($LogFile) | Select-Object -Last 20)
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
    try {
        $status = Get-ServiceRunningStatus -ServiceName $ServiceName
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
        $scOutput = @(& sc.exe start $ServiceName 2>&1)
        $exitCode = $LASTEXITCODE
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

Initialize-Log
Write-Log -Level INFO -Message "=== ServiceMonitor v$ScriptVersion starting on $Hostname ==="

# Load per-deployment SMTP settings (server/port/from/user/ssl) written by
# New-CredStore.ps1, so each server uses its own relay without editing the script.
# Explicit -Smtp* / -FromAddress parameters always win; the file fills the rest.
$smtpCfgPath = if ($SmtpConfigFile)    { $SmtpConfigFile }
               elseif ($SmtpCredFile)  { Join-Path ([System.IO.Path]::GetDirectoryName([System.IO.Path]::GetFullPath($SmtpCredFile))) 'smtp.json' }
               else                    { Join-Path $PSScriptRoot 'smtp.json' }
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
    Send-Alert -Recipients $script:AllRecipients `
               -Subject "[ServiceMonitor] TEST: connectivity check from $Hostname" `
               -Body $body
    Write-Log -Level INFO -Message '=== TestEmail complete ==='
    exit 0
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
    Write-Log -Level ERROR -Message 'No active services to check - exiting.'
    exit 1
}

# Check each service
$failureCount = 0

foreach ($entry in $activeEntries) {
    $svcName   = $entry.Name
    # Wrap in @() so $alertsTo is always an array: Get-AlertRecipients can return
    # an unrolled empty array (-> $null) or a scalar, and the .Count checks below
    # would otherwise throw under Set-StrictMode -Version Latest.
    $alertsTo  = @(Get-AlertRecipients -AlertsValue $entry.Alerts)

    if ($svcName -notmatch $ServiceNamePattern) {
        Write-Log -Level ERROR -Message "Rejected invalid service name: '$svcName'."
        $failureCount++
        continue
    }

    Write-Log -Level INFO -Message "Checking: $svcName  [alerts->$($entry.Alerts)]"

    $svcObj = Get-Service -Name $svcName -ErrorAction SilentlyContinue
    if ($null -eq $svcObj) {
        $msg  = "Service '$svcName' not found on $Hostname - verify the name in the web admin."
        Write-Log -Level ERROR -Message $msg
        if ($alertsTo.Count -gt 0) {
            $body = New-AlertBody -ServiceName $svcName -Status 'NOT FOUND on host' -Attempt 0 `
                                    -ErrorDetail $msg -AlertGroup $entry.Alerts
            Send-Alert -Recipients $alertsTo `
                       -Subject "[ServiceMonitor] ERROR: $svcName not found on $Hostname" -Body $body
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
            continue
        }
        Write-Log -Level WARNING -Message "'$svcName' still not Running after 30 s - proceeding with restart logic."
    }

    $currentStatus = Get-ServiceRunningStatus -ServiceName $svcName
    if ($currentStatus -eq [System.ServiceProcess.ServiceControllerStatus]::Running) {
        Write-Log -Level INFO -Message "OK: '$svcName' is running."
        continue
    }

    $startupType = Get-ServiceStartupType -ServiceName $svcName
    if ($startupType -eq 'Disabled') {
        $msg = "Service '$svcName' is DISABLED. Manual intervention required - skipping restart."
        Write-Log -Level ERROR -Message $msg
        if ($alertsTo.Count -gt 0) {
            $body = New-AlertBody -ServiceName $svcName -Status 'DOWN - StartupType=Disabled' `
                                    -Attempt 0 -ErrorDetail $msg -AlertGroup $entry.Alerts
            Send-Alert -Recipients $alertsTo `
                       -Subject "[ServiceMonitor] ERROR: $svcName is DISABLED on $Hostname" -Body $body
        }
        $failureCount++
        continue
    }

    Write-Log -Level WARNING -Message "Service '$svcName' is NOT running (status: $currentStatus, startup: $startupType)"
    if ($alertsTo.Count -gt 0) {
        $warnBody = New-AlertBody -ServiceName $svcName -Status "DOWN ($currentStatus)" `
                                    -Attempt 0 -ErrorDetail '' -AlertGroup $entry.Alerts
        Send-Alert -Recipients $alertsTo `
                   -Subject "[ServiceMonitor] WARNING: $svcName is down on $Hostname" -Body $warnBody
    } else {
        Write-Log -Level INFO -Message "Alerts suppressed for '$svcName' (alerts=none)."
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
            if ($alertsTo.Count -gt 0) {
                $body = New-AlertBody -ServiceName $svcName -Status 'RECOVERED' `
                                        -Attempt $attempt -ErrorDetail '' -AlertGroup $entry.Alerts
                Send-Alert -Recipients $alertsTo `
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
        if ($alertsTo.Count -gt 0) {
            $body = New-AlertBody -ServiceName $svcName `
                                    -Status "RESTART FAILED after $MaxAttempts attempts" `
                                    -Attempt $MaxAttempts -ErrorDetail $lastError -AlertGroup $entry.Alerts
            Send-Alert -Recipients $alertsTo `
                       -Subject "[ServiceMonitor] ERROR: $svcName restart FAILED on $Hostname" -Body $body
        }
        $failureCount++
    }
}

Write-Log -Level INFO -Message "=== Check complete. Failures: $failureCount / $($activeEntries.Count) active services ==="

exit $(if ($failureCount -gt 0) { 1 } else { 0 })
