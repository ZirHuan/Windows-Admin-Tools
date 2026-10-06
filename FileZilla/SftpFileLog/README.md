# FileZilla Server - SFTP file-transfer log

Turns **FileZilla Server 1.x** SFTP logs (`sftp_debug*.log`) into one CSV of file
transfers per local day, plus a searchable web page, refreshed every few minutes
by a scheduled task. Answers "did user X upload / download / delete file Y, and
did it work?" without grepping through debug logs.

| File | Purpose |
|---|---|
| `Export-SftpFileLog.ps1` | `<Folder>\SFTP_Log\*.log` -> `<Folder>\SFTPfiletransfers\transfers.YYYY-MM-DD.csv` + `transfers-data.js` |
| `transfers.html` | Double-click to search: free-text (AND), user/action/date filters, failures-only, session drill-down, CSV export of the result, auto-refresh keeping filters, stale-data warning, **Open folder... / Open file...** to view another server's output |
| `Install-SftpFileLogTask.ps1` | Registers a SYSTEM scheduled task (default every 5 min, no overlap). `-RunNow`, `-Remove`, `-WhatIf` |

CSV columns: `Date;Time;ID;User;IP;Folder;File;Action;OldName;Error`
(`;`-separated, all fields quoted, UTF-8 with BOM).

Actions: `Upload`, `Download`, `Delete`, `Rename`, `CreateDir`,
`UploadFailed`, `DownloadFailed`, `DeleteFailed`.

## Requirements
- FileZilla Server 1.x (tested with FileZilla Pro Enterprise Server 1.9.4) with
  SFTP logging at a level that records file operations (`Opening file`,
  `Removing file`, `Renaming entity`, `Creating directory`).
- Windows PowerShell 5.1 or PowerShell 7. No modules.
- The old FileZilla Server 0.9.x (plain FTP, `fzs-*.log`) uses a different log
  format and is **not** supported.

## Folder layout
Everything works from the folder the scripts are in (or `-Folder <path>` on either script):

```
<Folder>\Export-SftpFileLog.ps1, Install-SftpFileLogTask.ps1, transfers.html
<Folder>\SFTP_Log\              FileZilla logs (input)
<Folder>\SFTPfiletransfers\     daily CSVs, transfers-data.js, transfers.html, run log (output)
```

If FileZilla writes its logs somewhere else, point at them instead of moving them:
`-LogDir` / `-OutDir` on the export override single paths, and the installer
passes extra arguments on with `-ExportArguments`:

```powershell
.\Install-SftpFileLogTask.ps1 -RunNow -ExportArguments '-LogDir "D:\FileZilla\logs"'
```

## Deploy (elevated PowerShell on the SFTP server)
1. Copy the three files into a folder, e.g. `C:\Tools\SftpFileLog\`.
   - **Unsigned scripts:** if PowerShell says "is not digitally signed", remove the
     downloaded-from-internet mark once: `Get-ChildItem <folder> -Recurse | Unblock-File`
     (or run with `powershell.exe -ExecutionPolicy Bypass -File <script>`). The scheduled
     task already runs with `-ExecutionPolicy Bypass`. If GPO enforces
     AllSigned/RemoteSigned, Bypass is ignored and the scripts must be signed.
2. Test: `.\Export-SftpFileLog.ps1 -Verbose`
3. Schedule: `.\Install-SftpFileLogTask.ps1 -RunNow`
4. Open `<Folder>\SFTPfiletransfers\transfers.html` in a browser.

## How it reads the log (traps handled)
- **FileZilla logs `Opening file` *before* the open succeeds.** A failed upload
  first looks like a normal one; the script pairs the following
  `Failed to open file` (same session + path) with it and rewrites the row to
  `UploadFailed` / `DownloadFailed`. Same for `Removing file failed` -> `DeleteFailed`.
- **Session IDs are reused** over the day - `ID` + `Date` + `Time` identify a
  connection, not `ID` alone.
- **The live log is open for writing** - it is read with shared access, and since
  NTFS does not reliably update its write time, local today + yesterday are always
  rebuilt. Older days are rebuilt only when a source log is newer than the CSV
  (`-Rebuild` forces all).
- **Logs rotate on the UTC day**, so local day X is built from UTC logs X-1 and X.
- A CSV locked by Excel is skipped with a warning and retried on the next run.
- Logins, directory listings and brute-force noise are ignored - only
  authenticated sessions that touch files end up in the CSV.

## Web page
`transfers.html` reads `transfers-data.js` (the last `-WebDays` days, default 90)
from its own folder, so it works from a double-click with no web server.
`transfers-data.js` is parsed as data, never executed, when loaded via
**Open file...**.
