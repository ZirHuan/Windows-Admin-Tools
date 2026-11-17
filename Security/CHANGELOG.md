# Changelog - Security tools

## Get-FailedLogonReport

### 1.0.0 - 2026-11-17
- Initial release. Summarises Security 4625 failed logons and recovers the source IP
  that 4625 leaves blank for NTLM/NLA logons, using RdpCoreTS/Operational events
  131 (connection accepted) and 140 (bad credentials).
- Time-based correlation of blank-IP 4625 to RDP connects (marked "probable").
- Live 3389/445 sessions, NLA state, firewall logging state; `-EnableFirewallLog` switch.
- PowerShell 3.0+ (Windows Server 2012 / 2012 R2).
