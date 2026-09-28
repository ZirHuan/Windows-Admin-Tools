# ServiceMonitor v1.3.0 — bench test runbook

Everything below was installed and verified on the test bench on **2026-08-27**.
Follow it top to bottom; each **Verify** step tells you exactly what you should see.

---

## 0. Connect

| | |
|---|---|
| Address | `192.0.2.50` |
| Host name | `WIN-BENCH01` |
| OS | Windows Server 2022 Standard **Evaluation** |
| User | `.\benchadmin` |
| Password | *not stored in this repo — see the local bench notes* |
| Groups | Administrators, Remote Desktop Users |

From Windows: `mstsc /v:192.0.2.50`
From Linux: `xfreerdp /v:192.0.2.50 /u:benchadmin /p:'<password>'`

> **Note:** Network Level Authentication is **disabled** on this bench (Windows
> left it at 0). That is fine here, but turn it back on before a customer server:
>
> ```powershell
> $o = Get-WmiObject -Class Win32_TSGeneralSetting -Namespace root\cimv2\terminalservices |
>      Where-Object { $_.TerminalName -eq 'RDP-Tcp' }
> $o.SetUserAuthenticationRequired(1)
> ```

---

## 1. What is already installed

| Item | Value |
|---|---|
| Install folder | `C:\ServiceMonitor` |
| Web UI | `http://localhost:8080` — **bound to 127.0.0.1 only** |
| Windows service | `ServiceMonitorWeb` (NSSM, LocalSystem, auto-start) |
| Scheduled task | `ServiceMonitor` — every 5 min, SYSTEM, Highest |
| Python | 3.12 (`C:\Program Files\Python312`) |
| Config | `C:\ServiceMonitor\monitor-config.json` |
| Audit log | `C:\ServiceMonitor\changes.log` |
| Monitor log | `C:\ServiceMonitor\ServiceMonitor.log` |
| Web service log | `C:\ServiceMonitor\monitor-web.log` |
| Desktop shortcut | `ServiceMonitor Admin` (all users' desktop) |
| SMTP relay | `192.0.2.25:25`, from `servicemonitor@example.com` — anonymous, no TLS |
| Relay config | `C:\ServiceMonitor\smtp.json` (written by the **Relay settings** view) |
| Recipients | `support.lead@example.com`, `dev.lead@example.com` (Dev Team) |

The web UI listens on loopback only. It is **not** reachable from your laptop —
that is deliberate. You must be inside the RDP session to use it.

---

## 2. Open the admin UI

1. Log in over RDP.
2. Double-click **ServiceMonitor Admin** on the desktop (or browse to
   `http://localhost:8080`).
3. Top right, type your name next to **You are** and click **Set**.

**Verify:** the page is dark, with a left sidebar (Services / Mail groups /
Recent changes / Relay settings) and a green dot at the bottom reading
**Scheduled task active — ServiceMonitor · every 5 min · SYSTEM**.
Until you set your name, a yellow banner asks you to.

---

## 3. Services view

The landing page. Left column is what is monitored, right column adds more.

- **Status** — green dot Running, red Stopped, grey Paused, amber Unknown.
- **Alert routing** dropdown — Both / Dev only / Support only / None. Changes save
  immediately, no Save button.
- **Pause** (⏸) — keeps the service in the list but skips it on every run. Use
  this for planned maintenance instead of removing it.
- **Remove** (✕) — drops it from monitoring, with a confirm prompt.
- **Add a service** — type in the filter box to narrow the list, set
  *Alerts for newly added*, then click **+**. Standard Windows services are
  hidden from this list on purpose.

**Verify:** pause a service, then resume it. Both actions appear in
**Recent changes** attributed to the name you set.

---

## 4. Mail groups view

Two groups, `dev` (Dev Team) and `support` (Support Team). Add an address and press
**Add**; remove with ✕.

Routing is per service: a service set to *Dev only* mails just the Dev Team list.

**Verify:** add a throwaway address, confirm it appears in
`C:\ServiceMonitor\monitor-config.json`, then remove it again.

---

## 5. Recent changes view

The last 20 lines of `changes.log` — timestamp, who, what. This is the audit
trail; every config change through the UI lands here.

---

## 6. Relay settings view

Everything about the outgoing mail path, in one place. Before this existed the
relay lived only in `smtp.json` and `New-CredStore.ps1` on the command line.

The left column is what gets saved; the right column explains the selected
preset, previews the JSON live, and shows the file as it is on disk right now.
Nothing is hidden — what the preview shows is exactly what the Save button
writes.

### a) Relay server

| Field | Notes |
|---|---|
| Preset | Prefills the four fields below. Changing it never saves anything on its own. |
| Server | Hostname or IPv4 **without a port**. It goes straight into `SmtpClient(host, port)`, so `host:port` fails DNS resolution — the form rejects it. |
| Port | 25 anonymous, 587 STARTTLS submission. |
| STARTTLS | Sets `EnableSsl`. This is STARTTLS on the port above, *not* implicit TLS on 465. |
| From address | Must be on a domain the relay accepts, or the relay takes the message and silently drops it. |

Bad input is refused outright and `smtp.json` is left untouched — a scheme in
the server field, a port in the server field, a port outside 1–65535, a
malformed From. Questionable-but-legal input is saved *with a yellow warning*:
port 587 with TLS off, port 465 at all, a preset that needs a username when none
is set.

**Verify:** put `192.0.2.25:25` in **Server** and save. You get a red banner
reading *Nothing was saved — fix these first*, naming the port problem, and
`C:\ServiceMonitor\smtp.json` is unchanged.

**Verify:** set Server `192.0.2.25`, Port `25`, STARTTLS off, From
`servicemonitor@example.com`, preset **Custom / LAN relay**, and save. Green
banner, and `smtp.json` reads exactly:

```json
{
  "server": "192.0.2.25",
  "port": 25,
  "from": "servicemonitor@example.com",
  "user": "",
  "useSsl": false,
  "preset": "custom"
}
```

`preset` is an extra key this UI writes so it can re-select your choice later.
`ServiceMonitor.ps1` reads named keys only and ignores it.

### b) The Microsoft 365 presets

Four M365 paths are offered. They are **not** interchangeable — read the right
column before picking one.

| Preset | Endpoint | Port | TLS | Auth | External recipients |
|---|---|---|---|---|---|
| SMTP AUTH (client submission) | `smtp.office365.com` | 587 | required | mailbox user + password | yes |
| High Volume Email (HVE) | `smtp.hve.mx.microsoft` | 587 | required | HVE account + password | **no** |
| SMTP relay (connector) | tenant MX host | 25 | on | none — connector matches your IP | yes |
| Direct Send | tenant MX host | 25 | optional | none | **no** |

For the two MX-host presets the server field is **derived from the From
domain** (`contoso.com` → `contoso-com.mail.protection.outlook.com`) and is a
guess. Confirm it against the domain's real MX record in the Microsoft 365 admin
center before you trust it.

Three things that will bite on a customer tenant:

- **Basic auth only.** ServiceMonitor sends through .NET
  `System.Net.Mail.SmtpClient`, which has no way to present an OAuth token. The
  SMTP AUTH and HVE presets therefore work only with Basic auth.
- **Microsoft is retiring Basic auth for SMTP AUTH.** Per Message Center
  MC786329 as revised 2026-01-27: unchanged until December 2026, disabled by
  default for existing tenants at the end of December 2026 (admins can
  re-enable), unavailable by default for tenants created after that, final
  removal date to be announced in the second half of 2027.
- **Internal-only presets fail loudly.** HVE and Direct Send reject anything
  outside the tenant.

**Verify:** pick **Microsoft 365 — Direct Send** with From `alerts@example.com`
and save. The server auto-fills to
`example-com.mail.protection.outlook.com`, and you get a yellow warning naming
every recipient in **Mail groups** that is not on `example.com` — on this bench,
`support.lead@example.com`. Then put the LAN relay values back.

> **Do not point the bench at a real Microsoft 365 tenant to test this.** The
> presets are verified by what they write to `smtp.json`, not by sending
> through a tenant.

### c) Authentication

Optional — the anonymous LAN relay needs none of it. A status row shows
**Stored** or **Not stored**, checked by file existence only: the UI never
decrypts or displays a stored password.

**Store credentials** shells out to `New-CredStore.ps1 -PasswordFromStdin`. The
password goes down the child process's stdin pipe and nowhere else — not the
command line (any user could read that with `Get-CimInstance Win32_Process`),
not a temp file, not a log, not the redirect URL. `changes.log` records the
username and never the password.

`New-CredStore.ps1` also rewrites `smtp.json`, so **save the relay settings
first** — the current values are passed back in and round-trip unchanged.

**Clear** deletes `smtp.key`, `smtp.cred` and `smtp.user` and blanks `user` in
`smtp.json`. ServiceMonitor then connects anonymously again.

**Verify:** store a throwaway credential, then check:

```powershell
Get-ChildItem C:\ServiceMonitor\smtp.* | Select-Object Name,Length
(Get-Acl C:\ServiceMonitor\smtp.key).Access | Select-Object IdentityReference,FileSystemRights
Select-String -Path C:\ServiceMonitor\changes.log -Pattern 'Stored SMTP'
```

You should see `smtp.key` at **32 bytes**, an ACL of exactly
`NT AUTHORITY\SYSTEM` and `BUILTIN\Administrators` with FullControl, and a
`changes.log` line naming the user with no password in it. Then run:

```powershell
cd C:\ServiceMonitor
.\ServiceMonitor.ps1 -TestEmail -NoEmail
Get-Content C:\ServiceMonitor\ServiceMonitor.log -Tail 5
```

**Verify:** the log shows three lines — `Loaded SMTP settings from
C:\ServiceMonitor\smtp.json (server 192.0.2.25:25, ssl: False)`,
`Auto-discovered SMTP credential store in C:\ServiceMonitor (smtp.cred +
smtp.key)`, and `SMTP credentials loaded (user: ..., SSL: False)`. `-NoEmail`
keeps it off the network, so a throwaway credential is never used to
authenticate anywhere.

Then click **Clear** and confirm all three `smtp.*` credential files are gone.

> The auto-discovery line is new in this build. Previously the credential store
> was only read when the scheduled task was registered with `-SmtpCredFile` and
> `-SmtpKeyFile`; the task on this bench has neither, so credentials stored from
> the UI would have been silently ignored. `ServiceMonitor.ps1` now looks for
> `smtp.cred` + `smtp.key` beside `smtp.json` when **neither** parameter was
> given and a username is set. Explicit parameters still win, and a
> half-specified pair still produces the old warning.

### d) Send test email

Runs `ServiceMonitor.ps1 -TestEmail` with **no** SMTP overrides, so it resolves
the relay exactly the way the scheduled task does. It mails everyone in both
groups, and the result — success or the actual error text — appears in a banner
at the top of the page.

**Verify:** with the LAN relay config restored, click **Send test**. Green
banner reading *The relay accepted the test message*, with the log tail
underneath showing `Alert sent to: support.lead@example.com,
dev.lead@example.com`. The mail arrives.

A green banner means the relay **accepted** the message. It does not prove the
message reached an inbox — check the mailbox too.

---

## 7. Prove the alerting actually works

### a) Connectivity only

```powershell
cd C:\ServiceMonitor
.\ServiceMonitor.ps1 -TestEmail
```

**Verify:** log line `Alert sent to: support.lead@example.com, dev.lead@example.com`,
and the mail arrives.

### b) A real failure, with recovery

```powershell
Stop-Service Spooler -Force
cd C:\ServiceMonitor
.\ServiceMonitor.ps1 -AttemptDelaySeconds 5
```

**Verify:** the run reports Spooler not running, sends an alert, restarts it,
settles 8 s, confirms `recovered on attempt 1`, then sends a recovery mail.

### c) A failure it cannot fix

```powershell
Set-Service Spooler -StartupType Disabled
Stop-Service Spooler -Force
.\ServiceMonitor.ps1
```

**Verify:** `Service 'Spooler' is DISABLED. Manual intervention required -
skipping restart.` plus one alert. Then put it back:

```powershell
Set-Service Spooler -StartupType Automatic
Start-Service Spooler
```

### d) The scheduled task

```powershell
Start-ScheduledTask -TaskName ServiceMonitor
Get-ScheduledTaskInfo -TaskName ServiceMonitor | Select-Object LastRunTime,LastTaskResult
Get-Content C:\ServiceMonitor\ServiceMonitor.log -Tail 20
```

**Verify:** `LastTaskResult` is `0`.

---

## 8. Service management

```powershell
Get-Service ServiceMonitorWeb
Restart-Service ServiceMonitorWeb
Get-Content C:\ServiceMonitor\monitor-web.log -Tail 30

Get-ScheduledTask -TaskName ServiceMonitor
Disable-ScheduledTask -TaskName ServiceMonitor   # pause all monitoring
Enable-ScheduledTask  -TaskName ServiceMonitor
```

Built-in health check:

```powershell
cd C:\ServiceMonitor
.\Diagnose-Monitor.ps1
.\diagnose-web.ps1
```

---

## 9. Known behaviour worth knowing before the customer

- **`wuauserv` will flap.** Windows Update is `Manual` start by design — it stops
  itself when idle, so ServiceMonitor alerts and restarts it on most runs. It
  came across from `services.txt`. Remove it from monitoring unless the customer
  genuinely wants it, or you will train them to ignore alerts.
- **Alert cooldown is 60 min per service** (`-AlertCooldownMinutes`), so a
  service that stays down does not mail every 5 minutes. Recovery clears it.
- **The clock.** This bench shipped as *Pacific Standard Time*; I set it to
  *W. Europe Standard Time*. **Check `Get-TimeZone` on any new server first** —
  UTC syncs over NTP regardless, but the log and mail timestamps use local time.
- **UI is loopback-only.** If the customer wants it off-box, that is a real
  design decision (reverse proxy + auth), not a config flag.

---

## 10. Installing on the customer test server

```powershell
# 1. Copy the whole folder to the server, then from an elevated prompt:
cd <copied folder>
.\install-monitor-web.ps1                       # add -AccessToken '<secret>' on multi-user RDP hosts

# 2. From the INSTALLED folder, not the copy:
cd C:\ServiceMonitor
.\New-CredStore.ps1 -SmtpUser <user> -SmtpServer <host> -SmtpPort <port> -FromAddress <from>
.\Install-ScheduledTask.ps1 -SmtpCredFile C:\ServiceMonitor\smtp.cred -SmtpKeyFile C:\ServiceMonitor\smtp.key
```

Step 2 must run from `C:\ServiceMonitor` — both scripts bake their own folder
into the task action and cred paths at registration time. Registering from a
download folder gives you a task that breaks when that folder is deleted.

The `New-CredStore.ps1` line is now optional: the **Relay settings** view does
the same job from the UI, and `ServiceMonitor.ps1` finds `smtp.cred` + `smtp.key`
beside `smtp.json` on its own. Registering the task with the cred paths anyway
does no harm and keeps the intent explicit.

Pre-flight for a customer box:

- [ ] `Get-TimeZone` is correct
- [ ] NLA enabled on RDP
- [ ] SMTP relay reachable: `Test-NetConnection <relay> -Port 25`
- [ ] Relay configured in the **Relay settings** view, and **Send test email**
      comes back green
- [ ] If Microsoft 365: the chosen preset's caveats read and accepted — Basic
      auth only, MC786329 retirement timeline, internal-only limits on HVE and
      Direct Send
- [ ] `-AccessToken` set if more than one person logs in
- [ ] `wuauserv` and other Manual-start services removed from the list
- [ ] Recipients set in both groups before enabling the task

### Uninstall

```powershell
nssm remove ServiceMonitorWeb confirm
Unregister-ScheduledTask -TaskName ServiceMonitor -Confirm:$false
Remove-Item C:\ServiceMonitor -Recurse -Force
```

---

## 11. Rolling the bench back

From the Linux box:

```bash
~/scripts/testbench.sh list
~/scripts/testbench.sh rollback pre-servicemonitor        # before this install
~/scripts/testbench.sh rollback servicemonitor-v130-installed  # before the Relay settings view
~/scripts/testbench.sh rollback smtp-settings-installed        # back to this working state
```

Snapshots taken, oldest first: `pre-servicemonitor` (clean, pre-install),
`servicemonitor-v130-installed` (v1.3.0 working, no Settings view),
`smtp-settings-pre` (immediately before the Settings view was deployed) and
`smtp-settings-installed` (this verified state).
