<#
.SYNOPSIS
    Register (or remove) the scheduled task that runs Export-SftpFileLog.ps1
    every few minutes, so <Folder>\SFTPfiletransfers stays current.
    <Folder> = the folder this script is in, unless -Folder is given.

.DESCRIPTION
    Creates the task "<TaskName>" that runs as SYSTEM:
        - starts now, then repeats every -IntervalMinutes (default 5), indefinitely
        - never starts a second copy while one is still running
        - is stopped if a run takes longer than 10 minutes
        - runs PowerShell 7 (pwsh.exe) if installed, else Windows PowerShell 5.1

    Re-running the script replaces the task with the new settings.
    Run from an elevated PowerShell on the SFTP server.

.PARAMETER Folder
    Base folder holding Export-SftpFileLog.ps1, SFTP_Log and SFTPfiletransfers.
    Default: the folder this script is in. Passed on to the export as -Folder.

.PARAMETER ScriptPath
    Full path of Export-SftpFileLog.ps1. Default: <Folder>\Export-SftpFileLog.ps1
    transfers.html must sit next to it (the export copies it to the output folder).

.PARAMETER IntervalMinutes
    How often to run. Default: 5

.PARAMETER TaskName
    Default: FileZilla - SFTP file transfer log

.PARAMETER ExportArguments
    Extra arguments for Export-SftpFileLog.ps1, e.g. "-WebDays 60".

.PARAMETER RunNow
    Start the task once right after registering it and show the result.

.PARAMETER Remove
    Unregister the task instead.

.EXAMPLE
    .\Install-SftpFileLogTask.ps1 -RunNow

.EXAMPLE
    .\Install-SftpFileLogTask.ps1 -Folder D:\SftpReport -WhatIf

.EXAMPLE
    .\Install-SftpFileLogTask.ps1 -Remove

.NOTES
    Version 1.0.0 - first public release (FileZilla Server 1.x SFTP logs).
    See CHANGELOG.md in this folder.

    UNSIGNED SCRIPT. If PowerShell stops with "... is not digitally signed. You
    cannot run this script on the current system", the files still carry the
    downloaded-from-the-internet mark (zip fetched over SFTP or a browser).
    Fix once, in the folder holding the scripts (elevated PowerShell):
        Get-ChildItem <folder with the scripts> -Recurse | Unblock-File
    Or run it once without changing anything:
        powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-SftpFileLogTask.ps1 -RunNow
    Running from a network share (\\server\share\...)? Unblock-File is not
    enough there - RemoteSigned treats every UNC script as remote. Use
        Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
    (this window only), or copy the scripts to a local disk first.
    The scheduled task already starts the export with -ExecutionPolicy Bypass.
    If Group Policy enforces AllSigned/RemoteSigned, Bypass is ignored and the
    scripts must be code-signed instead.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$Folder          = '',
    [string]$ScriptPath      = '',
    [ValidateRange(1, 1440)]
    [int]$IntervalMinutes    = 5,
    [string]$TaskName        = 'FileZilla - SFTP file transfer log',
    [string]$ExportArguments = '',
    [switch]$RunNow,
    [switch]$Remove
)

$ErrorActionPreference = 'Stop'

# Paths default to the folder the script is in (or -Folder)
if (-not $Folder)     { $Folder = $PSScriptRoot }
if (-not $ScriptPath) { $ScriptPath = Join-Path $Folder 'Export-SftpFileLog.ps1' }

$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw 'Run this from an elevated PowerShell (Run as administrator).'
}

$existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue

if ($Remove) {
    if (-not $existing) { Write-Host "Task '$TaskName' does not exist - nothing to remove."; return }
    if ($PSCmdlet.ShouldProcess($TaskName, 'Unregister scheduled task')) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        Write-Host "Task '$TaskName' removed."
    }
    return
}

if (-not (Test-Path $ScriptPath -PathType Leaf)) { throw "Export script not found: $ScriptPath" }
$ScriptPath = (Resolve-Path $ScriptPath).Path
$htmlPath   = Join-Path (Split-Path $ScriptPath) 'transfers.html'
if (-not (Test-Path $htmlPath)) { Write-Warning "transfers.html not found next to the script ($htmlPath) - the web page will not be copied." }

# Prefer PowerShell 7, fall back to Windows PowerShell 5.1
$pwsh = Get-Command pwsh.exe -ErrorAction SilentlyContinue
if ($pwsh) { $exe = $pwsh.Source }
else       { $exe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe' }

$Folder = (Resolve-Path $Folder).Path.TrimEnd('\')
$arg = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{0}" -Folder "{1}"' -f $ScriptPath, $Folder
if ($ExportArguments) { $arg += ' ' + $ExportArguments }

$action    = New-ScheduledTaskAction -Execute $exe -Argument $arg -WorkingDirectory (Split-Path $ScriptPath)
$trigger   = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) `
                 -RepetitionInterval (New-TimeSpan -Minutes $IntervalMinutes)
$settings  = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -StartWhenAvailable `
                 -ExecutionTimeLimit (New-TimeSpan -Minutes 10) `
                 -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
$runAs     = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest

Write-Host "Task      : $TaskName"
Write-Host "Runs      : $exe $arg"
Write-Host "Logs from : $(Join-Path $Folder 'SFTP_Log')"
Write-Host "Output to : $(Join-Path $Folder 'SFTPfiletransfers')"
Write-Host "Every     : $IntervalMinutes min, as SYSTEM"
if ($existing) { Write-Host "Existing task will be replaced." }

if ($PSCmdlet.ShouldProcess($TaskName, 'Register scheduled task')) {
    Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings `
        -Principal $runAs -Description "Builds $(Join-Path $Folder 'SFTPfiletransfers')\transfers.*.csv + transfers.html from the FileZilla SFTP logs. Installed by Install-SftpFileLogTask.ps1." `
        -Force | Out-Null
    Write-Host "Task registered. First run within a minute."

    if ($RunNow) {
        Start-ScheduledTask -TaskName $TaskName
        $deadline = (Get-Date).AddMinutes(5)
        do { Start-Sleep -Seconds 2 } while ((Get-ScheduledTask -TaskName $TaskName).State -eq 'Running' -and (Get-Date) -lt $deadline)
        $info = Get-ScheduledTaskInfo -TaskName $TaskName
        Write-Host ("Last run  : {0}  result 0x{1:X}" -f $info.LastRunTime, $info.LastTaskResult)
        if ($info.LastTaskResult -ne 0) { Write-Warning 'The run did not succeed - see the run log (Export-SftpFileLog.log in the output folder).' }
    }
}
