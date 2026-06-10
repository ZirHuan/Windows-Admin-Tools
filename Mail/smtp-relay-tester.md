# smtp-relay-tester.ps1

Interactive SMTP relay diagnostic and test tool.
Compatible with **Windows PowerShell 5.1** and **PowerShell 7+** from the same file.

---

## Requirements

| Requirement | Detail |
|-------------|--------|
| PowerShell  | 5.1 (Windows PowerShell) or 7.0+ (`pwsh`) |
| .NET        | .NET Framework 4.5+ (PS 5.1) or .NET 6+ (PS7) |
| Network     | Outbound TCP to the target server/port |
| OS          | Windows (PS 5.1 or PS7); Linux/macOS (PS7 only) |

No external modules are required. All SMTP communication is implemented
directly over `System.Net.Sockets.TcpClient` + `System.Net.Security.SslStream`.

---

## Usage

### Windows PowerShell 5.1
```powershell
powershell -File smtp-relay-tester.ps1
```

### PowerShell 7
```powershell
pwsh -File smtp-relay-tester.ps1
```

All parameters are gathered interactively. Press **Enter** at any prompt to
accept the bracketed default from the previous run.

### Optional: custom settings file

```powershell
pwsh -File smtp-relay-tester.ps1 -SettingsFile C:\Temp\relay-staging.json
```

Useful when testing multiple relays with separate saved credentials.

---

## Interactive Prompts

The script walks through the following prompts in order:

1. **SMTP server hostname** - e.g. `mail.example.com`
2. **Port picker** - choose from a numbered menu or enter a raw port:
   - `[1]` 25   - Plain SMTP (no encryption)
   - `[2]` 587  - SMTP + STARTTLS (submission, recommended)
   - `[3]` 465  - SMTP over implicit SSL/TLS
   - `[4]` 2525 - Common alternate submission port
   - `[5]` custom
3. **Verbose output** - `y` shows every SMTP command/response line plus TLS details
4. **Authenticate** - `y/n`
5. **Username** (if auth enabled)
6. **Password** (if auth enabled) - masked input; saved password shown as `********`
7. **From address** - defaults to `username@server`
8. **To address** - leave blank to skip the test send

---

## TLS Behavior by Port

| Port | TLS Mode |
|------|----------|
| 465  | Implicit SSL - socket is wrapped in TLS before the server banner |
| 587  | Mandatory STARTTLS - script issues STARTTLS after EHLO and re-issues EHLO over TLS. Aborts if server does not advertise STARTTLS. |
| other | Opportunistic STARTTLS - upgrades if advertised, otherwise continues plain |

---

## EHLO Capability Display

After connecting (and after TLS negotiation where applicable) the script
prints the full EHLO capability list in a formatted table, e.g.:

```
  --- EHLO capabilities ---
  SIZE                 10240000
  AUTH                 LOGIN PLAIN NTLM
  STARTTLS             (present)
  8BITMIME             (present)
  -------------------------
```

---

## AUTH Method Analysis

After printing the capability table the script checks the `AUTH` line:

- **NTLM advertised**: prints a `[WARN]` with the message:
  > NTLM is advertised - this breaks Veeam 13 / .NET 8 SmtpClient.
  > Ask the relay provider to remove NTLM, leaving only LOGIN and/or PLAIN.

- **NTLM not advertised**: prints `[OK] NTLM not advertised - AUTH methods look fine.`

- **No AUTH in EHLO**: prints `[INFO] No AUTH methods advertised in EHLO.`

### Supported AUTH methods

The script implements both `AUTH LOGIN` and `AUTH PLAIN` manually over the
socket using base64 encoding.  It prefers LOGIN if both are advertised.
If the server advertises neither, it attempts LOGIN anyway (some servers
require auth but omit it from EHLO).  If the server advertises only NTLM or
another unsupported mechanism, the script fails with a clear error.

---

## Verbose Mode

When verbose mode is enabled every line of the SMTP protocol conversation is
printed with directional arrows:

```
  C>>> EHLO myhost
  S<<< 250-mail.example.com Hello
  S<<< 250-SIZE 10240000
  S<<< 250-AUTH LOGIN PLAIN
  S<<< 250 8BITMIME
  [TLS] Protocol : Tls13
  [TLS] Cipher   : Aes256
  [TLS] Strength : 256 bits
  [TLS] Cert CN  : CN=mail.example.com
  [TLS] Cert exp : 2025-12-31 00:00:00
```

Passwords are redacted in the protocol trace:

```
  C>>> ******** (saved, Enter to keep)
```

In non-verbose mode only the `[OK]`, `[WARN]`, `[FAIL]`, `[INFO]` status
lines are shown - matching the concise output style of the Python reference.

---

## Settings Persistence

Settings are saved to (and loaded from) a JSON file:

- **Default path**: `$HOME/.smtp-relay-tester.json`
- **Custom path**: use the `-SettingsFile` parameter

Saved fields: `server`, `port`, `username`, `doAuth`, `fromAddress`,
`toAddress`, `verbose`, `encryptedPassword`.

### Password Security

The password is **never stored in plaintext**.

| Platform | Encryption mechanism |
|----------|---------------------|
| Windows (PS 5.1 or PS7) | `ConvertFrom-SecureString` with no key - uses Windows DPAPI (Data Protection API), encrypted per-user per-machine. The ciphertext is unreadable on any other Windows account or machine. |
| Linux / macOS (PS7 only) | `ConvertFrom-SecureString` with a randomly-generated 256-bit AES key stored in `$HOME/.smtp-relay-tester.json.key`. The key file is created with `chmod 600` on first run. Both the JSON and the key file must be present for decryption to succeed. |

**Limitations and assumptions:**

- Protection relies on filesystem permissions on Linux/macOS. An attacker with
  read access to your home directory can decrypt the password if they also have
  the key file.
- On Windows, the DPAPI ciphertext is tied to your Windows user account.
  Copying `.smtp-relay-tester.json` to another Windows machine will cause
  decryption to fail gracefully - you will be prompted to re-enter the password.
- If decryption fails for any reason the script falls back to prompting for
  the password fresh, without crashing.

To **clear the saved password**, delete `$HOME/.smtp-relay-tester.json`
(and `$HOME/.smtp-relay-tester.json.key` on Linux/macOS).

---

## Exit Codes

| Code | Meaning |
|------|---------|
| 0    | All requested operations succeeded |
| 1    | Any error: connection failure, authentication failure, SMTP error, send failure |

---

## Examples

### Test relay connectivity only (no auth, no send)

```
SMTP server hostname []: mail.example.com
Select SMTP port: [2] (Enter for 587)
Verbose output: n
Authenticate [Y/n]: n
From address [test@mail.example.com]: (Enter)
Send test mail to (Enter to skip): (Enter)
```

Expected output:
```
[OK]   TCP connection established to mail.example.com:587
[OK]   Server banner: 220 mail.example.com ESMTP
[OK]   STARTTLS negotiating TLS ...
[OK]   TLS channel established
  --- EHLO capabilities ---
  AUTH                 LOGIN PLAIN
  SIZE                 10240000
  -------------------------
[OK]   NTLM not advertised - AUTH methods look fine.
[INFO] Skipped authentication (not requested).
[INFO] No To address supplied - skipping test send.
[OK]   Session closed (server: 221 Bye)
[OK]   Done - relay is working.
```

### Full test with auth and send

```
SMTP server hostname [mail.example.com]: (Enter)
Select SMTP port: (Enter for 587)
Verbose output: y
Authenticate [Y/n]: y
Username [relay-user]: (Enter)
Password [******** (saved, Enter to keep)]: (Enter)
From address [relay-user@mail.example.com]: relay@example.com
Send test mail to: admin@example.com
```

---

## Relationship to smtp-relay-tester.py

| Feature | Python version | PowerShell version |
|---------|---------------|-------------------|
| Raw socket EHLO | No (smtplib) | Yes (TcpClient) |
| Full protocol trace | No | Yes (verbose mode) |
| TLS details | No | Yes |
| AUTH LOGIN/PLAIN | Via smtplib | Manual base64 |
| NTLM warning | Yes | Yes (same wording) |
| Settings persistence | No | Yes (DPAPI/AES) |
| Port picker menu | No | Yes |
| PS 5.1 compatible | N/A | Yes |
| PS 7+ compatible | N/A | Yes |

---

## File Locations

| File | Purpose |
|------|---------|
| `smtp-relay-tester.ps1` | Main script |
| `smtp-relay-tester.ps1.md` | This README |
| `$HOME/.smtp-relay-tester.json` | Saved settings (auto-created) |
| `$HOME/.smtp-relay-tester.json.key` | AES key for password encryption on Linux/macOS (auto-created, PS7 only) |

---

## GitHub / Commit Suggestion

Suggested commit message:
```
feat: Make smtp-relay-tester.ps1 compatible with Windows PowerShell 5.1 and PS7
```

Suggested repo location: `M365-Tools/` or `Exchange/` folder.
