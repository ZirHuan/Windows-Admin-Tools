#Requires -Version 5.1

<#
.SYNOPSIS
    Monitors Windows services, restarts failed ones, and alerts via email.

.DESCRIPTION
    Reads a list of Windows services from a text file (one per line).
    For each active service (not prefixed with * or -), checks if it is running.
    If not: logs a WARNING, sends an alert email, and attempts up to MaxAttempts
    restarts with AttemptDelaySeconds between each. If all restarts fail, logs an
    ERROR with the captured error output and alerts all configured recipients.

    Designed for Windows Server 2016+ and compatible with PowerShell 5.1 and 7+.
    Run as Administrator - required for Start-Service on system services.

.PARAMETER ServicesFile
    Path to text file listing service names, one per line.
    Lines starting with * or - (optionally followed by spaces) are skipped (paused).
    Blank lines and lines starting with # are also ignored.
    Default: .\services.txt

.PARAMETER RecipientsFile
    Path to text file listing recipient email addresses, one per line.
    Blank lines and # comments ignored.
    Default: .\recipients.txt

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

.PARAMETER MaxAttempts
    Number of restart attempts before marking the service as failed.
    Default: 3

.PARAMETER AttemptDelaySeconds
    Seconds to wait between restart attempts.
    Default: 60

.PARAMETER StartSettleSeconds
    Seconds to wait after issuing a start command before re-checking status.
    Needed because sc.exe returns immediately (async); the service may still be
    in StartPending for a few seconds.
    Default: 8

.PARAMETER NoEmail
    If specified, suppresses all email sending (useful for local testing).

.EXAMPLE
    .\ServiceMonitor.ps1

.EXAMPLE
    .\ServiceMonitor.ps1 -ServicesFile C:\Monitoring\services.txt `
                         -RecipientsFile C:\Monitoring\recipients.txt `
                         -LogFile C:\Logs\svc-monitor.log

.EXAMPLE
    .\ServiceMonitor.ps1 -NoEmail   # check only, no mail

.NOTES
    Designed for Windows Server 2016+, PowerShell 5.1 and 7+.
    Run as Administrator for full restart capability.
    Schedule via Task Scheduler for continuous monitoring.
    Version: 1.0.1
#>

[CmdletBinding()]
param(
    [string] $ServicesFile        = (Join-Path $PSScriptRoot 'services.txt'),
    [string] $RecipientsFile      = (Join-Path $PSScriptRoot 'recipients.txt'),
    [string] $LogFile             = (Join-Path $PSScriptRoot 'ServiceMonitor.log'),
    [int]    $LogMaxSizeMB        = 10,
    [string] $SmtpServer          = '192.168.2.15',
    [int]    $SmtpPort            = 25,
    [string] $FromAddress         = 'servicemonitor@rosvalls.com',
    [int]    $MaxAttempts         = 3,
    [int]    $AttemptDelaySeconds = 60,
    [int]    $StartSettleSeconds  = 8,
    [switch] $NoEmail
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ScriptVersion = '1.0.1'
$Hostname      = $env:COMPUTERNAME

# Valid service-name characters: letters, digits, underscore, dot, hyphen, $, space
# This prevents argument injection into sc.exe.
$ServiceNamePattern = '^[A-Za-z0-9_.$ -]+$'

# ---------------------------------------------------------------------------
# Log helpers
# ---------------------------------------------------------------------------

function Initialize-Log {
    # Ensure log directory exists
    $logDir = Split-Path -LiteralPath $LogFile -Parent
    if ($logDir -and -not (Test-Path -LiteralPath $logDir)) {
        try {
            New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        } catch {
            Write-Warning "Cannot create log directory '$logDir': $_"
        }
    }

    # Rotate if over size limit
    if (Test-Path -LiteralPath $LogFile) {
        $sizeBytes = (Get-Item -LiteralPath $LogFile).Length
        if ($sizeBytes -gt ($LogMaxSizeMB * 1MB)) {
            $maxBackups = 3
            for ($i = $maxBackups; $i -ge 1; $i--) {
                $src  = if ($i -eq 1) { $LogFile } else { "$LogFile.$($i-1)" }
                $dest = "$LogFile.$i"
                if (Test-Path -LiteralPath $src) {
                    Move-Item -LiteralPath $src -Destination $dest -Force
                }
            }
        }
    }

    # Verify we can write (fail loud rather than silently losing logs)
    try {
        Add-Content -LiteralPath $LogFile -Value '' -Encoding UTF8
    } catch {
        throw "Log file '$LogFile' is not writable: $_"
    }
}

function Write-Log {
    param(
        [ValidateSet('INFO','WARNING','ERROR')]
        [string] $Level,
        [string] $Message
    )
    $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line  = "$stamp [$($Level.PadRight(7))] $Message"

    try {
        Add-Content -LiteralPath $LogFile -Value $line -Encoding UTF8
    } catch {
        Write-Warning "Log write failed: $_"
    }

    switch ($Level) {
        'INFO'    { Write-Host $line -ForegroundColor Cyan }
        'WARNING' { Write-Host $line -ForegroundColor Yellow }
        'ERROR'   { Write-Host $line -ForegroundColor Red }
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

    # Force array so .Count is always valid (fixes StrictMode scalar/$null issue)
    $rcptArray = @($Recipients | Where-Object { $_ -match '@' })
    if ($rcptArray.Count -eq 0) {
        Write-Log -Level WARNING -Message 'No recipients - skipping email.'
        return
    }

    try {
        $smtp = New-Object System.Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
        $smtp.EnableSsl             = $false
        $smtp.UseDefaultCredentials = $false
        $smtp.Timeout               = 15000   # 15 s

        $msg = New-Object System.Net.Mail.MailMessage
        $msg.From    = New-Object System.Net.Mail.MailAddress($FromAddress, "Service Monitor [$Hostname]")
        $msg.Subject = $Subject
        $msg.Body    = $Body
        $msg.IsBodyHtml = $false

        foreach ($addr in $rcptArray) {
            $msg.To.Add($addr)
        }

        $smtp.Send($msg)
        $smtp.Dispose()
        $msg.Dispose()

        Write-Log -Level INFO -Message "Alert sent to: $($rcptArray -join ', ')"
    } catch {
        Write-Log -Level ERROR -Message "Failed to send email alert: $_"
    }
}

function Build-AlertBody {
    param(
        [string] $ServiceName,
        [string] $Status,
        [int]    $Attempt,
        [string] $ErrorDetail
    )

    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $lines = @(
        '=== Service Monitor Alert ==='
        "Host    : $Hostname"
        "Time    : $ts"
        "Service : $ServiceName"
        "Status  : $Status"
    )

    if ($Attempt -gt 0) {
        $lines += "Attempt : $Attempt / $MaxAttempts"
    }

    if ($ErrorDetail) {
        $lines += ''
        $lines += '--- Error Detail ---'
        $lines += $ErrorDetail
    }

    $lines += ''
    $lines += '--- Log File ---'
    $lines += $LogFile
    $lines += ''
    $lines += "ServiceMonitor.ps1 v$ScriptVersion"

    return $lines -join "`r`n"
}

# ---------------------------------------------------------------------------
# File loaders
# ---------------------------------------------------------------------------

function Read-ListFile {
    param([string] $Path, [string] $Label)

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Log -Level ERROR -Message "$Label file not found: $Path"
        return @()
    }

    # ReadAllLines handles BOM transparently on both PS 5.1 and 7
    $lines = [System.IO.File]::ReadAllLines($Path) |
             Where-Object { $_ -match '\S' -and $_ -notmatch '^\s*#' }

    # Force array so callers always get a collection
    return @($lines)
}

function Parse-Services {
    param([string[]] $Lines)

    $active = [System.Collections.Generic.List[string]]::new()
    $paused = [System.Collections.Generic.List[string]]::new()

    foreach ($line in $Lines) {
        $trimmed = $line.Trim()
        if ($trimmed -match '^[\*\-]\s*(.+)') {
            $name = $Matches[1].Trim()
            if ($name) { $paused.Add($name) }
        } elseif ($trimmed) {
            $active.Add($trimmed)
        }
    }

    # Return as plain arrays (avoids PS 5.1 multiple-return unrolling ambiguity
    # when directly assigning: $a, $b = Parse-Services $x)
    return [string[]]$active.ToArray(), [string[]]$paused.ToArray()
}

# ---------------------------------------------------------------------------
# Service helpers
# ---------------------------------------------------------------------------

function Get-ServiceStartupType {
    param([string] $ServiceName)
    # Win32_Service.StartMode: 'Auto','Manual','Disabled','Boot','System'
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
    # Wait up to $TimeoutSeconds for service to reach $TargetState
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

        # Stop if not already stopped - wait for clean stop before issuing start
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

        # Use sc.exe so we capture richer Windows error text.
        # The name is already validated before this function is called.
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

# Load and parse services
$serviceLines = Read-ListFile -Path $ServicesFile -Label 'Services'
if ($serviceLines.Count -eq 0) {
    Write-Log -Level ERROR -Message 'No services loaded - exiting.'
    exit 1
}

$activeServices, $pausedServices = Parse-Services -Lines $serviceLines

if ($pausedServices.Count -gt 0) {
    Write-Log -Level INFO -Message "Paused (skipped): $($pausedServices -join ', ')"
}
Write-Log -Level INFO -Message "Active services to check ($($activeServices.Count)): $($activeServices -join ', ')"

# Load recipients - force array throughout
$recipientLines = @(Read-ListFile -Path $RecipientsFile -Label 'Recipients')
$recipients     = @($recipientLines | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '@' })
if ($recipients.Count -eq 0) {
    Write-Log -Level WARNING -Message 'No valid recipient addresses found - email alerts disabled.'
} else {
    Write-Log -Level INFO -Message "Alert recipients ($($recipients.Count)): $($recipients -join ', ')"
}

# Check each service
$failureCount = 0

foreach ($svcName in $activeServices) {

    # -- Security: validate service name before passing to sc.exe ---------------
    if ($svcName -notmatch $ServiceNamePattern) {
        Write-Log -Level ERROR -Message "Rejected invalid service name: '$svcName' (contains disallowed characters)."
        $failureCount++
        continue
    }

    Write-Log -Level INFO -Message "Checking: $svcName"

    # -- Service must exist on this host ----------------------------------------
    $svcObj = Get-Service -Name $svcName -ErrorAction SilentlyContinue
    if ($null -eq $svcObj) {
        $msg = "Service '$svcName' not found on $Hostname - verify the name in services.txt."
        Write-Log -Level ERROR -Message $msg
        $body = Build-AlertBody -ServiceName $svcName -Status 'NOT FOUND on host' -Attempt 0 -ErrorDetail $msg
        Send-Alert -Recipients $recipients -Subject "[ServiceMonitor] ERROR: $svcName not found on $Hostname" -Body $body
        $failureCount++
        continue
    }

    # -- Handle transient StartPending: wait briefly before deciding it is DOWN --
    if ($svcObj.Status -eq [System.ServiceProcess.ServiceControllerStatus]::StartPending) {
        Write-Log -Level INFO -Message "'$svcName' is StartPending - waiting up to 30 s for it to reach Running."
        $started = Wait-ServiceState -ServiceName $svcName `
            -TargetState ([System.ServiceProcess.ServiceControllerStatus]::Running) `
            -TimeoutSeconds 30
        if ($started) {
            Write-Log -Level INFO -Message "OK: '$svcName' reached Running (was StartPending)."
            continue
        }
        Write-Log -Level WARNING -Message "'$svcName' still not Running after 30 s StartPending wait - proceeding with restart logic."
    }

    $currentStatus = Get-ServiceRunningStatus -ServiceName $svcName
    if ($currentStatus -eq [System.ServiceProcess.ServiceControllerStatus]::Running) {
        Write-Log -Level INFO -Message "OK: '$svcName' is running."
        continue
    }

    # -- Service is DOWN --------------------------------------------------------

    # Check startup type: Disabled services cannot be started by sc.exe
    $startupType = Get-ServiceStartupType -ServiceName $svcName
    if ($startupType -eq 'Disabled') {
        $msg = "Service '$svcName' is DISABLED (StartupType=Disabled). Manual intervention required - skipping restart."
        Write-Log -Level ERROR -Message $msg
        $body = Build-AlertBody -ServiceName $svcName -Status "DOWN - StartupType=Disabled" -Attempt 0 -ErrorDetail $msg
        Send-Alert -Recipients $recipients -Subject "[ServiceMonitor] ERROR: $svcName is DISABLED on $Hostname" -Body $body
        $failureCount++
        continue
    }

    Write-Log -Level WARNING -Message "Service '$svcName' is NOT running (status: $currentStatus, startup: $startupType) - alerting."
    $warnBody = Build-AlertBody -ServiceName $svcName -Status "DOWN ($currentStatus)" -Attempt 0 -ErrorDetail ''
    Send-Alert -Recipients $recipients `
               -Subject "[ServiceMonitor] WARNING: $svcName is down on $Hostname" `
               -Body $warnBody

    # -- Restart loop -----------------------------------------------------------
    $restarted = $false
    $lastError  = ''

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        Write-Log -Level WARNING -Message "Restart attempt $attempt / $MaxAttempts for '$svcName'..."

        $restartResult = Invoke-ServiceRestart -ServiceName $svcName

        if (-not $restartResult.Success) {
            $lastError = $restartResult.ErrorText
            Write-Log -Level WARNING -Message "sc.exe reported failure on attempt $attempt for '$svcName': $lastError"
        }

        # Always wait the settle period so async StartPending resolves before
        # we re-check status (sc.exe returns before the service reaches Running).
        Write-Log -Level INFO -Message "Settling $StartSettleSeconds s after start command..."
        Start-Sleep -Seconds $StartSettleSeconds

        $nowStatus = Get-ServiceRunningStatus -ServiceName $svcName
        if ($nowStatus -eq [System.ServiceProcess.ServiceControllerStatus]::Running) {
            Write-Log -Level INFO -Message "Service '$svcName' recovered on attempt $attempt."
            $body = Build-AlertBody -ServiceName $svcName -Status 'RECOVERED' -Attempt $attempt -ErrorDetail ''
            Send-Alert -Recipients $recipients `
                       -Subject "[ServiceMonitor] RECOVERED: $svcName on $Hostname (attempt $attempt/$MaxAttempts)" `
                       -Body $body
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
        $body = Build-AlertBody -ServiceName $svcName `
                                -Status "RESTART FAILED after $MaxAttempts attempts" `
                                -Attempt $MaxAttempts `
                                -ErrorDetail $lastError
        Send-Alert -Recipients $recipients `
                   -Subject "[ServiceMonitor] ERROR: $svcName restart FAILED on $Hostname" `
                   -Body $body
        $failureCount++
    }
}

Write-Log -Level INFO -Message "=== Check complete. Failures: $failureCount / $($activeServices.Count) active services ==="

# Exit 0 on clean run (even if services failed and were handled).
# Exit 1 only for unrecoverable issues (logged separately above).
exit 0
