<#
.SYNOPSIS
    Turn the FileZilla Server SFTP logs into one CSV of file transfers per day:
        <Folder>\SFTPfiletransfers\transfers.YYYY-MM-DD.csv
    <Folder> = the folder this script is in, unless -Folder is given.
    Columns: Date;Time;ID;User;IP;Folder;File;Action;OldName;Error

.DESCRIPTION
    Reads the FileZilla Server logs in -LogDir and keeps only the lines where an
    authenticated user touched a FILE:

        Opening file '...' [mode=reading]       -> Download
        Opening file '...' [mode=writing]       -> Upload
        Failed to open file [file=...]          -> turns that Upload/Download row into
                                                   UploadFailed/DownloadFailed (+ Error)
        Removing file [file=...]                -> Delete
        Removing file failed [file=...]         -> turns that Delete row into DeleteFailed
        Renaming entity [old=..., new=...]      -> Rename (Folder/File = NEW name,
                                                   OldName = old name or old full path)
        Creating directory [dir=...]            -> CreateDir

    FileZilla logs "Opening file" BEFORE it knows whether the open works, so a
    failed upload first looks like a normal one. The script pairs the failure
    with its open (same session + same path) and rewrites that row.

    Logins, directory listings and brute-force noise are skipped.
    ID is the FileZilla session number - all rows from one connection share an ID.
    FileZilla reuses session numbers over the day, so ID + Date + Time identify a
    connection, not ID alone.

    Days are LOCAL days (the server's time zone). FileZilla rotates its log on the UTC day, so
    local day X is built from the logs of UTC day X-1 and X.

    INCREMENTAL, built to run every few minutes from a scheduled task:
      - local today and yesterday are always rebuilt (the live sftp_debug.log is
        still being written, and NTFS does not reliably update its write time);
      - older days are only rebuilt when the CSV is missing or a source log is
        newer than it. Use -Rebuild to regenerate everything.
    Logs are opened with shared read access, so the live log can be read while
    FileZilla writes to it. A CSV that is open in Excel (locked) is skipped with
    a warning and retried on the next run.

    After writing, the script regenerates transfers-data.js (the last -WebDays
    days) next to the CSVs and copies transfers.html there if it sits next to
    this script. Double-click transfers.html to search the transfers in a browser.

    One line per run is appended to -RunLog (rotated at 1 MB).

    CSV: ';' separated, every field quoted, UTF-8 with BOM -> opens directly in
    Excel in locales that use ';' as the list separator.

.PARAMETER Folder
    Base folder. Default: the folder this script is in. LogDir and OutDir are
    taken from it unless they are given themselves.

.PARAMETER LogDir
    Folder with the FileZilla logs. Default: <Folder>\SFTP_Log

.PARAMETER LogFilter
    Which files in LogDir to read. Default: *.log

.PARAMETER OutDir
    Where the daily CSVs go. Default: <Folder>\SFTPfiletransfers

.PARAMETER WebDays
    How many days (newest first) go into transfers-data.js for the web page.
    0 = all days. Default: 90. The daily CSVs are always kept.

.PARAMETER RunLog
    Run log, one line per run. Default: <OutDir>\Export-SftpFileLog.log

.PARAMETER Rebuild
    Regenerate every day, even if the CSV is up to date.

.EXAMPLE
    .\Export-SftpFileLog.ps1
    Logs from .\SFTP_Log, output to .\SFTPfiletransfers (next to the script).

.EXAMPLE
    .\Export-SftpFileLog.ps1 -Folder D:\SftpReport -Rebuild -Verbose

.NOTES
    Version 1.0.0 - first public release (FileZilla Server 1.x SFTP logs).
    See CHANGELOG.md in this folder.

    UNSIGNED SCRIPT. If PowerShell stops with "... is not digitally signed. You
    cannot run this script on the current system", the files still carry the
    downloaded-from-the-internet mark (zip fetched over SFTP or a browser).
    Fix once, in the folder holding the scripts (elevated PowerShell):
        Get-ChildItem <folder with the scripts> -Recurse | Unblock-File
    Or run it once without changing anything:
        powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Export-SftpFileLog.ps1 -Verbose
    Running from a network share (\\server\share\...)? Unblock-File is not
    enough there - RemoteSigned treats every UNC script as remote. Use
        Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
    (this window only), or copy the scripts to a local disk first.
    The scheduled task already starts the export with -ExecutionPolicy Bypass.
    If Group Policy enforces AllSigned/RemoteSigned, Bypass is ignored and the
    scripts must be code-signed instead.
#>
[CmdletBinding()]
param(
    [string]$Folder    = '',
    [string]$LogDir    = '',
    [string]$LogFilter = '*.log',
    [string]$OutDir    = '',
    [ValidateRange(0, 100000)]
    [int]$WebDays      = 90,
    [string]$RunLog    = '',
    [switch]$Rebuild
)

$ErrorActionPreference = 'Stop'
$started = Get-Date

$Columns = @('Date', 'Time', 'ID', 'User', 'IP', 'Folder', 'File', 'Action', 'OldName', 'Error')
$Utf8Bom = New-Object System.Text.UTF8Encoding($true)
$Inv     = [Globalization.CultureInfo]::InvariantCulture
# Paths default to the folder the script is in (or -Folder)
if (-not $Folder) { $Folder = $PSScriptRoot }
if (-not $LogDir) { $LogDir = Join-Path $Folder 'SFTP_Log' }
if (-not $OutDir) { $OutDir = Join-Path $Folder 'SFTPfiletransfers' }
if (-not $RunLog) { $RunLog = Join-Path $OutDir 'Export-SftpFileLog.log' }

# 2026-10-01T07:25:04.241Z >> [SSH SFTP Session 246466 192.0.2.10 alice] <message>
# Only lines WITH a username (4th token inside the brackets) = authenticated session.
$lineRx = [regex]'^(?<ts>\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?Z)\s+\S+\s+\[SSH SFTP Session (?<id>\d+) (?<ip>\S+) (?<user>[^\]]+)\]\s+(?<msg>.*)$'
$tsRx   = [regex]'^(?<d>\d{4}-\d{2}-\d{2})T\d{2}:\d{2}:\d{2}'

# Order matters: the "failed" forms must be tested before the plain forms.
$opRx = @(
    @{ Rx = [regex]"^Opening file '(?<path>[^']+)' \[mode=reading";                          Kind = 'Download'   }
    @{ Rx = [regex]"^Opening file '(?<path>[^']+)' \[mode=writing";                          Kind = 'Upload'     }
    @{ Rx = [regex]'^Failed to open file \[file=(?<path>[^\]]+?), mode=.*?\(raw = \d+: (?<err>[^)]*)\)'; Kind = 'OpenFailed' }
    @{ Rx = [regex]'^Failed to open file \[file=(?<path>[^\]]+?), mode=';                    Kind = 'OpenFailed' }
    @{ Rx = [regex]'^Removing file failed \[file=(?<path>[^\],]+).*?(?:\(raw = \d+: (?<err>[^)]*)\))?$'; Kind = 'DeleteFailed' }
    @{ Rx = [regex]'^Removing file \[file=(?<path>[^\]]+)\]';                                Kind = 'Delete'     }
    @{ Rx = [regex]'^Renaming entity \[old=(?<old>[^,]+), new=(?<path>[^\]]+)\]';            Kind = 'Rename'     }
    @{ Rx = [regex]'^Creating directory \[dir=(?<path>[^\]]+)\]';                            Kind = 'CreateDir'  }
)
# Cheap pre-filter so the regexes only run on the few interesting lines
$keywords = @('Opening file', 'Failed to open file', 'Removing file', 'Renaming entity', 'Creating directory')

function Format-CsvField([string]$s) { '"' + $s.Replace('"', '""') + '"' }

function Split-SftpPath([string]$p) {
    $p = $p.Trim()
    $slash = $p.LastIndexOf('/')
    if ($slash -lt 0) { return @('/', $p) }
    $folder = if ($slash -gt 0) { $p.Substring(0, $slash) } else { '/' }
    return @($folder, $p.Substring($slash + 1))
}

# Open for reading while FileZilla may still be writing to (or rotating) the file
function Open-SharedReader([string]$path) {
    $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
    $fs = New-Object System.IO.FileStream($path, [System.IO.FileMode]::Open,
                                          [System.IO.FileAccess]::Read, $share)
    return New-Object System.IO.StreamReader($fs, $true)
}

# UTC day a log covers = date of its first timestamped line
function Get-LogUtcDay([string]$path) {
    $sr = Open-SharedReader $path
    try {
        for ($i = 0; $i -lt 200 -and -not $sr.EndOfStream; $i++) {
            $m = $tsRx.Match($sr.ReadLine())
            if ($m.Success) { return $m.Groups['d'].Value }
        }
    } finally { $sr.Dispose() }
    return $null
}

# Write via a temp file + replace, so readers never see a half-written file.
# Returns $false (and warns) if the target is locked, e.g. open in Excel.
function Write-FileSafely([string]$path, [string]$text) {
    $tmp = "$path.tmp"
    try {
        [System.IO.File]::WriteAllText($tmp, $text, $Utf8Bom)
        if (Test-Path $path) { [System.IO.File]::Replace($tmp, $path, [NullString]::Value) }
        else                 { [System.IO.File]::Move($tmp, $path) }
        return $true
    } catch {
        Write-Warning "Could not write $path (open in another program?): $($_.Exception.Message)"
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        return $false
    }
}

function Write-RunLog([string]$text) {
    try {
        if ((Test-Path $RunLog) -and (Get-Item $RunLog).Length -gt 1MB) {
            Move-Item $RunLog "$RunLog.1" -Force
        }
        $line = '{0}  {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $text
        [System.IO.File]::AppendAllText($RunLog, $line + [Environment]::NewLine, $Utf8Bom)
    } catch { Write-Warning "Could not write run log $RunLog : $($_.Exception.Message)" }
}

try {
    if (-not (Test-Path $LogDir)) { throw "Log folder not found: $LogDir" }
    if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }

    # --- 1. Map UTC day -> log files ---------------------------------------
    $logsByDay = @{}
    foreach ($f in Get-ChildItem -Path $LogDir -Filter $LogFilter -File) {
        $d = Get-LogUtcDay $f.FullName
        if (-not $d) { Write-Verbose "Skipping $($f.Name): no timestamped lines"; continue }
        if (-not $logsByDay.ContainsKey($d)) { $logsByDay[$d] = @() }
        $logsByDay[$d] += $f
    }
    if ($logsByDay.Count -eq 0) { throw "No FileZilla logs found in $LogDir ($LogFilter)" }

    # --- 2. Which local days need (re)building? ----------------------------
    $candidates = @{}
    foreach ($d in @($logsByDay.Keys)) {
        $dt = [datetime]::ParseExact($d, 'yyyy-MM-dd', $Inv)
        $candidates[$d] = $true
        $candidates[$dt.AddDays(1).ToString('yyyy-MM-dd')] = $true
    }

    $today     = $started.ToString('yyyy-MM-dd')
    $yesterday = $started.AddDays(-1).ToString('yyyy-MM-dd')
    $needDays  = @{}
    $needLogs  = @{}
    foreach ($day in @($candidates.Keys)) {
        if ($day -gt $today) { continue }   # no files for days that have not started yet
        $prev    = [datetime]::ParseExact($day, 'yyyy-MM-dd', $Inv).AddDays(-1).ToString('yyyy-MM-dd')
        $sources = @(@($logsByDay[$day]) + @($logsByDay[$prev]) | Where-Object { $_ })
        $csv     = Join-Path $OutDir "transfers.$day.csv"

        $stale = $Rebuild -or ($day -ge $yesterday) -or -not (Test-Path $csv)
        if (-not $stale) {
            $csvTime = (Get-Item $csv).LastWriteTime
            $stale = @($sources | Where-Object { $_.LastWriteTime -gt $csvTime }).Count -gt 0
        }
        if ($stale) {
            $needDays[$day] = New-Object System.Collections.Generic.List[object]
            foreach ($s in $sources) { $needLogs[$s.FullName] = $true }
        }
    }

    # --- 3. Parse the needed logs, bucket rows by local day ----------------
    foreach ($path in ($needLogs.Keys | Sort-Object)) {
        Write-Verbose "Reading $path"
        $lastOpen   = @{}   # "id|path" -> row of the latest Upload/Download in this log
        $lastDelete = @{}   # "id|path" -> row of the latest Delete in this log
        $sr = Open-SharedReader $path
        try {
            while ($null -ne ($line = $sr.ReadLine())) {
                if ($line.IndexOf('[SSH SFTP Session ') -lt 0) { continue }
                $hit = $false
                foreach ($k in $keywords) { if ($line.IndexOf($k) -ge 0) { $hit = $true; break } }
                if (-not $hit) { continue }

                $m = $lineRx.Match($line)
                if (-not $m.Success) { continue }
                $msg = $m.Groups['msg'].Value

                foreach ($op in $opRx) {
                    $o = $op.Rx.Match($msg)
                    if (-not $o.Success) { continue }

                    $id   = $m.Groups['id'].Value
                    $p    = $o.Groups['path'].Value.Trim()
                    $key  = "$id|$p"
                    $err  = if ($o.Groups['err'].Success) { $o.Groups['err'].Value.Trim() } else { '' }

                    # A failure rewrites the row it belongs to, if we have it
                    if ($op.Kind -eq 'OpenFailed' -and $lastOpen.ContainsKey($key)) {
                        $r = $lastOpen[$key]
                        $r.Action = $r.Action + 'Failed'
                        $r.Error  = $err
                        $lastOpen.Remove($key)
                        break
                    }
                    if ($op.Kind -eq 'DeleteFailed' -and $lastDelete.ContainsKey($key)) {
                        $r = $lastDelete[$key]
                        $r.Action = 'DeleteFailed'
                        $r.Error  = $err
                        $lastDelete.Remove($key)
                        break
                    }

                    $ts  = [datetime]::Parse($m.Groups['ts'].Value, $Inv,
                                             [Globalization.DateTimeStyles]::AdjustToUniversal).ToLocalTime()
                    $day = $ts.ToString('yyyy-MM-dd')
                    $parts = Split-SftpPath $p

                    $old = ''
                    if ($op.Kind -eq 'Rename') {
                        $oldParts = Split-SftpPath $o.Groups['old'].Value
                        # Same folder -> just the old file name, else the full old path
                        $old = if ($oldParts[0] -eq $parts[0]) { $oldParts[1] } else { $o.Groups['old'].Value.Trim() }
                    }

                    $row = [pscustomobject]@{
                        Date    = $day
                        Time    = $ts.ToString('HH:mm:ss')
                        ID      = $id
                        User    = $m.Groups['user'].Value
                        IP      = $m.Groups['ip'].Value
                        Folder  = $parts[0]
                        File    = $parts[1]
                        Action  = $op.Kind
                        OldName = $old
                        Error   = $err
                        Sort    = $ts.ToString('HH:mm:ss.fff')
                    }
                    # Keep the pairing even when the row's day is not being rebuilt,
                    # so a failure is never added as a separate row by mistake.
                    if ($op.Kind -eq 'Upload' -or $op.Kind -eq 'Download') { $lastOpen[$key] = $row }
                    if ($op.Kind -eq 'Delete') { $lastDelete[$key] = $row }

                    if ($needDays.ContainsKey($day)) { $needDays[$day].Add($row) }
                    break
                }
            }
        } finally { $sr.Dispose() }
    }

    # --- 4. Write the daily CSVs -------------------------------------------
    $dayText = @{}      # day -> CSV text, used for the web data even if the file was locked
    $written = 0; $locked = 0; $failedOps = 0; $totalRows = 0
    $header = ($Columns | ForEach-Object { Format-CsvField $_ }) -join ';'
    foreach ($day in ($needDays.Keys | Sort-Object)) {
        $csv  = Join-Path $OutDir "transfers.$day.csv"
        # A local day spans two logs -> sort so the rows are in time order
        $rows = @($needDays[$day] | Sort-Object Sort, ID)
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.Append($header).Append("`r`n")
        foreach ($r in $rows) {
            $vals = foreach ($c in $Columns) { Format-CsvField ([string]$r.$c) }
            [void]$sb.Append($vals -join ';').Append("`r`n")
            if ($r.Action -like '*Failed') { $failedOps++ }
        }
        $dayText[$day] = $sb.ToString()
        $totalRows += $rows.Count
        if (Write-FileSafely $csv $dayText[$day]) {
            $written++
            Write-Verbose ("{0}  {1,6} transfers  -> {2}" -f $day, $rows.Count, $csv)
        } else { $locked++ }
    }

    # --- 5. Web page data + page -------------------------------------------
    $dataJs  = Join-Path $OutDir 'transfers-data.js'
    $webOk   = $true
    if ($needDays.Count -gt 0 -or -not (Test-Path $dataJs)) {
        $files = @(Get-ChildItem -Path $OutDir -Filter 'transfers.*.csv' -File | Sort-Object Name -Descending)
        $days  = @($files | ForEach-Object { $_.Name.Substring(10, 10) })
        foreach ($d in $dayText.Keys) { if ($days -notcontains $d) { $days += $d } }   # locked + brand new
        $days = @($days | Sort-Object -Descending)
        if ($WebDays -gt 0) { $days = @($days | Select-Object -First $WebDays) }

        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine('// Generated by Export-SftpFileLog.ps1 - do not edit')
        [void]$sb.AppendLine('window.SFTP_DAYS = [];')
        foreach ($d in ($days | Sort-Object)) {
            if ($dayText.ContainsKey($d)) { $text = $dayText[$d] }
            else {
                $sr = Open-SharedReader (Join-Path $OutDir "transfers.$d.csv")
                try { $text = $sr.ReadToEnd() } finally { $sr.Dispose() }
            }
            [void]$sb.AppendLine(('window.SFTP_DAYS.push({0});' -f (ConvertTo-Json -InputObject $text -Compress)))
        }
        [void]$sb.AppendLine(('window.SFTP_GENERATED = {0};' -f (ConvertTo-Json -InputObject (Get-Date -Format 'yyyy-MM-dd HH:mm'))))
        $webOk = Write-FileSafely $dataJs $sb.ToString()
    }

    $htmlSrc = Join-Path $PSScriptRoot 'transfers.html'
    $htmlDst = Join-Path $OutDir 'transfers.html'
    if ((Test-Path $htmlSrc) -and (-not (Test-Path $htmlDst) -or
        (Get-Item $htmlSrc).LastWriteTime -gt (Get-Item $htmlDst).LastWriteTime)) {
        Copy-Item $htmlSrc $htmlDst -Force
        Write-Verbose "Web page copied -> $htmlDst"
    }

    $secs = ((Get-Date) - $started).TotalSeconds
    $summary = 'OK  days={0} rows={1} failed-ops={2} locked={3} web={4} logs={5} {6:N1}s' -f `
        $written, $totalRows, $failedOps, $locked, $(if ($webOk) { 'ok' } else { 'LOCKED' }), $needLogs.Count, $secs
    Write-Host $summary
    Write-RunLog $summary
} catch {
    Write-RunLog ("ERROR  " + $_.Exception.Message)
    throw
}
