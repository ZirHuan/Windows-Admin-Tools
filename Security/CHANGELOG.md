# Changelog - Security tools

## Get-FailedLogonReport

### 1.0.2 - 2026-11-17
- Fix: crashed with "The property '#text' cannot be found" on 4625 events that
  contain empty `<Data/>` fields (StrictMode). EventData now read via
  namespace-agnostic XPath + InnerText.

### 1.0.1 - 2026-11-17
- Performance: 1.0.0 crawled on servers with many 4625 events (array `+=` and a
  per-event pipeline scan in correlation were both quadratic). Now List[object] +
  single-pass two-pointer correlation, progress output every 2000 events.
- New `-MaxEvents` to cap how many of the newest 4625 events are read.

### 1.0.0 - 2026-11-17
- Initial release. Summarises Security 4625 failed logons and recovers the source IP
  that 4625 leaves blank for NTLM/NLA logons, using RdpCoreTS/Operational events
  131 (connection accepted) and 140 (bad credentials).
- Time-based correlation of blank-IP 4625 to RDP connects (marked "probable").
- Live 3389/445 sessions, NLA state, firewall logging state; `-EnableFirewallLog` switch.
- PowerShell 3.0+ (Windows Server 2012 / 2012 R2).
