# Changelog - smtp-relay-tester.ps1

All notable changes to the `smtp-relay-tester.ps1` tool are documented here.
Format based on [Keep a Changelog](https://keepachangelog.com/);
this tool follows [Semantic Versioning](https://semver.org/).

Git tags for this tool are prefixed `smtp-relay-tester-ps-v*` (this repo holds
multiple independent scripts).

## [1.0.0] - 2026-06-10

First versioned, feature-complete release. PowerShell 5.1 / 7 port of
[smtp-relay-tester.py](https://github.com/ZirHuan/Linux-commands-and-Scripts/blob/main/smtp-relay-tester.py).

### Added
- Dual-version support: runs on Windows PowerShell 5.1 and PowerShell 7+.
- Raw-socket SMTP via `TcpClient` + `SslStream` (surfaces the real EHLO / AUTH
  exchange, unlike `Send-MailMessage` / `SmtpClient`).
- Standard SMTP port picker (25 / 587 / 465 / 2525 / custom).
- Verbose mode showing the full SMTP protocol conversation, including TLS details.
- STARTTLS (587), implicit SSL/TLS (465), and opportunistic STARTTLS (25).
- NTLM detection warning (NTLM breaks Veeam 13 / .NET 8 `SmtpClient`).
- Remembered settings saved to `$HOME/.smtp-relay-tester.json`, offered as
  Enter-to-accept defaults.
- Optional password persistence, encrypted at rest: DPAPI on Windows,
  AES-256 keyfile on Linux/macOS. Never stored in plaintext.
- In-script `$ScriptVersion` constant shown in the startup banner.

### Notes
- No external modules required; standard library / .NET only.
