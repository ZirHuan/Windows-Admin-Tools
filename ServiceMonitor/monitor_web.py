#!/usr/bin/env python3
"""
monitor_web.py — ServiceMonitor web admin UI
Bind: 127.0.0.1:8080  (localhost-only; RDP login is the primary auth layer)
Start: uvicorn monitor_web:app --host 127.0.0.1 --port 8080 --workers 1

Optional defense-in-depth: set the SM_TOKEN environment variable (via the NSSM
service config) to require a shared access token before the UI is usable. This
matters on multi-user RDP hosts where any logged-in user can reach 127.0.0.1.
If SM_TOKEN is unset (default) the UI is open and RDP login is the only gate.
"""

import ctypes
import hmac
import html
import json
import os
import re
import shutil
import socket
import subprocess
import threading
import time
from ctypes import wintypes
from datetime import datetime
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, Form, Request
from fastapi.responses import HTMLResponse, RedirectResponse

app = FastAPI(docs_url=None, redoc_url=None)

CONFIG_FILE = Path(os.environ.get("SM_CONFIG", r"C:\ServiceMonitor\monitor-config.json"))
CHANGES_LOG = CONFIG_FILE.parent / "changes.log"

# Optional shared access token. Empty string => gate disabled (RDP login only).
ACCESS_TOKEN = os.environ.get("SM_TOKEN", "").strip()

_lock = threading.Lock()

DEFAULT_CONFIG = {
    "version": "1.2",
    "groups": {
        "dev":  {"label": "Dev Team",      "recipients": []},
        "support": {"label": "Support Team", "recipients": []},
    },
    "services": [],
}

# Standard Windows services excluded from the "add" panel
SYSTEM_SERVICES = frozenset({
    "AeLookupSvc","ALG","AppIDSvc","Appinfo","AppMgmt","AppReadiness",
    "AudioEndpointBuilder","AudioSrv","AxInstSV","BDESVC","BFE","BITS",
    "BrokerInfrastructure","CertPropSvc","ClipSVC","COMSysApp","CryptSvc",
    "DcomLaunch","defragsvc","DeviceAssociationService","DeviceInstall",
    "Dhcp","DiagTrack","DispBrokerDesktopSvc","Dnscache","DoSvc","DPS",
    "DsmSvc","Eaphost","EFS","EventLog","EventSystem","fdPHost","FDResPub",
    "FontCache","gpsvc","hidserv","IKEEXT","iphlpsvc","KeyIso","KtmRm",
    "LanmanServer","LanmanWorkstation","lfsvc","LicenseManager","lltdsvc",
    "lmhosts","LSM","MapsBroker","MMCSS","MpsSvc","MSDTC","NcaSvc",
    "NcbService","Netlogon","netprofm","NetSetupSvc","NlaSvc","nsi","NTDS",
    "PcaSvc","PlugPlay","PolicyAgent","Power","ProfSvc","RasAuto","RasMan",
    "RemoteAccess","RemoteRegistry","RpcEptMapper","RpcLocator","RpcSs",
    "SamSs","Schedule","seclogon","SENS","SessionEnv","SharedAccess",
    "ShellHWDetection","Spooler","SysMain","SystemEventsBroker","TapiSrv",
    "TermService","Themes","TimeBrokerSvc","TrkWks","TrustedInstaller",
    "UI0Detect","UmRdpService","upnphost","UserManager","UsoSvc","VaultSvc",
    "vds","W32Time","WbioSrvc","Wcmsvc","WdiServiceHost","WdiSystemHost",
    "WdNisSvc","WebClient","Wecsvc","WerSvc","WinDefend",
    "WinHttpAutoProxySvc","Winmgmt","WinRM","wlidsvc","WPDBusEnum",
    "wuauserv","wudfsvc","XblAuthManager","XblGameSave","XboxGipSvc",
})

# ---------------------------------------------------------------------------
# Config helpers
# ---------------------------------------------------------------------------

def read_config() -> dict:
    if not CONFIG_FILE.exists():
        write_config(DEFAULT_CONFIG)
        return dict(DEFAULT_CONFIG)
    # utf-8-sig tolerates a UTF-8 BOM if some other tool wrote one (PS 5.1's
    # Set-Content -Encoding UTF8 does); plain utf-8 would raise on the BOM.
    with open(CONFIG_FILE, encoding="utf-8-sig") as f:
        cfg = json.load(f)
    if migrate_config(cfg):
        write_config(cfg)
    return cfg


def migrate_config(cfg: dict) -> bool:
    """Rename the pre-1.4.0 mail group 'iver' to 'support'. Returns True if changed.

    ServiceMonitor.ps1 reads either key, so a config migrated here keeps working
    with an older monitor script, and an unmigrated one works with a newer script.
    """
    changed = False
    groups = cfg.setdefault("groups", {})
    if "iver" in groups and "support" not in groups:
        grp = groups.pop("iver")
        if grp.get("label") in (None, "", "Iver Support"):
            grp["label"] = "Support Team"
        groups["support"] = grp
        changed = True
    for svc in cfg.get("services", []):
        if svc.get("alerts") == "iver":
            svc["alerts"] = "support"
            changed = True
    return changed


def write_config(cfg: dict) -> None:
    tmp = CONFIG_FILE.with_suffix(".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2)
    shutil.move(str(tmp), str(CONFIG_FILE))


def log_change(user: str, action: str) -> None:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M")
    with open(CHANGES_LOG, "a", encoding="utf-8") as f:
        f.write(f"{ts}  {user:<20}  {action}\n")


def recent_changes(n: int = 20) -> list[str]:
    if not CHANGES_LOG.exists():
        return []
    with open(CHANGES_LOG, encoding="utf-8") as f:
        lines = [l.rstrip() for l in f if l.strip()]
    return lines[-n:][::-1]


# ---------------------------------------------------------------------------
# SMTP relay settings
#
# smtp.json lives next to monitor-config.json. ServiceMonitor.ps1 auto-loads it
# from the folder of -SmtpCredFile, else from beside the script - on a normal
# install both are C:\ServiceMonitor. Schema, as written by New-CredStore.ps1:
#     {"server": "host-only-no-port", "port": 587, "from": "...",
#      "user": "...", "useSsl": true}
# 'server' must be the host ONLY: ServiceMonitor.ps1 does
# New-Object SmtpClient($SmtpServer,$SmtpPort), so a "host:port" string is taken
# as a hostname and DNS resolution fails.
# 'preset' is an extra key this UI writes so it can re-select the same preset
# later. ServiceMonitor.ps1 reads named keys only and ignores it.
# ---------------------------------------------------------------------------

SMTP_FILE     = CONFIG_FILE.parent / "smtp.json"
CRED_FILE     = CONFIG_FILE.parent / "smtp.cred"
KEY_FILE      = CONFIG_FILE.parent / "smtp.key"
USER_FILE     = CONFIG_FILE.parent / "smtp.user"
CREDSTORE_PS1 = CONFIG_FILE.parent / "New-CredStore.ps1"
MONITOR_PS1   = CONFIG_FILE.parent / "ServiceMonitor.ps1"

_LABEL_RE = r"[A-Za-z0-9]([A-Za-z0-9-]{0,61}[A-Za-z0-9])?"
HOST_RE   = re.compile(rf"^(?=.{{1,253}}$){_LABEL_RE}(\.{_LABEL_RE})*$")
MAIL_RE   = re.compile(rf"^[^@\s,;:<>\"'\\]{{1,64}}@{_LABEL_RE}(\.{_LABEL_RE})+$")


def _is_ipv4(s: str) -> bool:
    parts = s.split(".")
    return (len(parts) == 4 and all(p.isdigit() and len(p) <= 3
                                    and 0 <= int(p) <= 255 for p in parts))


# Relay presets. 'auth' is what the relay expects, not what we enforce:
#   required - a username and a stored password are needed
#   optional - anonymous or authenticated, operator's choice
#   none     - the relay identifies this host some other way (IP / certificate)
# 'derive' means the server field is computed from the From domain by JS and is
# only a guess at the tenant MX host - the UI says so.
PRESETS = [
    {
        "id": "custom", "label": "Custom / LAN relay",
        "server": "", "port": 25, "ssl": False, "auth": "optional",
        "derive": False, "internal_only": False,
        "summary": "Any relay you run yourself - a Postfix or Exchange smart host, "
                   "an appliance, or an ISP relay. Nothing is assumed; the values "
                   "below are written to smtp.json verbatim.",
        "facts": [
            "Anonymous relay on port 25 with TLS off is fully supported, and is "
            "what this deployment uses today.",
            "Leave the username empty for an anonymous relay. Fill it in and store "
            "a password to use SMTP AUTH.",
        ],
        "warns": [],
    },
    {
        "id": "m365-auth", "label": "Microsoft 365 - SMTP AUTH (client submission)",
        "server": "smtp.office365.com", "port": 587, "ssl": True, "auth": "required",
        "derive": False, "internal_only": False,
        "summary": "Authenticates as a real, licensed Microsoft 365 mailbox and "
                   "sends on its behalf. The closest match to a classic "
                   "authenticated relay.",
        "facts": [
            "Endpoint smtp.office365.com, port 587, STARTTLS required (TLS 1.2 or "
            "1.3). Port 25 also works but is often blocked upstream.",
            "Delivers to internal and external recipients, and the message lands in "
            "the mailbox's Sent Items.",
            "Throttled at 10,000 recipients per day and 30 messages per minute.",
            "The From address must be the authenticating mailbox, or that mailbox "
            "needs Send As permission on it - otherwise mail bounces with "
            "5.7.60 Client does not have permissions to send as this sender.",
        ],
        "warns": [
            "ServiceMonitor can only use Basic authentication. It sends through "
            ".NET System.Net.Mail.SmtpClient, whose documented API has no way to "
            "supply an OAuth 2.0 token, so Modern auth is not available here.",
            "Microsoft is retiring Basic auth for SMTP AUTH. Per Message Center "
            "MC786329 as revised 2026-01-27: behaviour unchanged until December "
            "2026; disabled by default for existing tenants at the end of December "
            "2026 (admins can still re-enable it); unavailable by default for "
            "tenants created after that; a final removal date to be announced in "
            "the second half of 2027.",
            "SMTP AUTH is disabled by default for tenants created after January "
            "2020, and Entra ID security defaults disable it outright. Both have to "
            "be dealt with in the tenant before this preset can work.",
        ],
    },
    {
        "id": "m365-hve", "label": "Microsoft 365 - High Volume Email (HVE)",
        "server": "smtp.hve.mx.microsoft", "port": 587, "ssl": True, "auth": "required",
        "derive": False, "internal_only": True,
        "summary": "Microsoft's designated path for exactly this kind of traffic: "
                   "automated alerts from an application to people inside the "
                   "tenant. Uses a dedicated HVE account rather than a mailbox.",
        "facts": [
            "Endpoint smtp.hve.mx.microsoft, port 587, TLS required. The older name "
            "smtp-hve.office365.com still resolves but Microsoft says it will be "
            "deprecated.",
            "No recipient or message rate limits. Up to 50 recipients per message, "
            "10 MB maximum message size.",
            "HVE accounts support Basic authentication, so this preset works with "
            "ServiceMonitor as it is built today.",
            "Keeps working even when SmtpClientAuthenticationDisabled is True on the "
            "tenant, because HVE uses its own endpoint.",
        ],
        "warns": [
            "Internal recipients only. Mail to any address outside the tenant is "
            "rejected. Check the Mail groups view before choosing this.",
            "Requires Microsoft 365 pay-as-you-go billing with a billing policy "
            "assigned to the HVE account. Without one the account cannot send at "
            "all. Billing is per delivered recipient.",
            "Basic auth here still depends on Entra ID security defaults being off. "
            "Microsoft recommends OAuth, which ServiceMonitor cannot use.",
        ],
    },
    {
        "id": "m365-relay", "label": "Microsoft 365 - SMTP relay (inbound connector)",
        "server": "", "port": 25, "ssl": True, "auth": "none",
        "derive": True, "internal_only": False,
        "summary": "Relays through the tenant MX host with no SMTP credentials at "
                   "all. An inbound connector in Exchange recognises this server by "
                   "its public IP address or TLS certificate.",
        "facts": [
            "Endpoint is the tenant MX host, for example "
            "contoso-com.mail.protection.outlook.com. Port 25, STARTTLS.",
            "No username or password. Authentication is the connector matching this "
            "server's identity.",
            "Relays to external recipients, and the From address only needs to be in "
            "an accepted domain - it does not need a mailbox behind it.",
        ],
        "warns": [
            "Requires tenant-side setup this UI cannot do: an inbound connector in "
            "the Exchange admin center under Mail flow > Connectors. Without it the "
            "relay rejects the session.",
            "Needs a static, unshared public IP address - Microsoft does not support "
            "dynamic addresses. The certificate-based variant is not usable here: "
            "ServiceMonitor.ps1 never sets SmtpClient.ClientCertificates, so it "
            "presents no client certificate.",
            "The server value below is derived from the From domain and is a GUESS. "
            "Confirm it against that domain's real MX record in the Microsoft 365 "
            "admin center before relying on it.",
        ],
    },
    {
        "id": "m365-direct", "label": "Microsoft 365 - Direct Send",
        "server": "", "port": 25, "ssl": True, "auth": "none",
        "derive": True, "internal_only": True,
        "summary": "Delivers straight to the tenant MX host as if this server were "
                   "an external mail server. No credentials, no connector, and no "
                   "tenant configuration at all.",
        "facts": [
            "Endpoint is the tenant MX host, for example "
            "contoso-com.mail.protection.outlook.com. Port 25. Microsoft lists TLS "
            "as optional here; this preset leaves STARTTLS on.",
            "The From address must be in an accepted domain in the tenant. It does "
            "not need a mailbox.",
        ],
        "warns": [
            "Internal recipients only. Mail to Gmail, Outlook.com or any address "
            "outside the tenant is rejected.",
            "Microsoft's own guidance: \"Most customers don't need to use Direct "
            "Send. We're working on an option to disable Direct Send by default to "
            "protect customers.\" Treat it as a last resort.",
            "Mail arrives as anonymous internet mail and is fully spam-scanned. SPF "
            "must list this server's public IP, and DKIM and DMARC must be right, "
            "or alerts land in Junk.",
            "The server value below is derived from the From domain and is a GUESS. "
            "Confirm it against that domain's real MX record.",
        ],
    },
]

PRESET_BY_ID = {p["id"]: p for p in PRESETS}

# Paths deliberately NOT offered here, and why. Shown in the UI so the omission
# is visible rather than silent.
NOT_IMPLEMENTED = [
    "OAuth 2.0 / Modern authentication for any Microsoft 365 path - "
    "System.Net.Mail.SmtpClient exposes no way to present a bearer token, and "
    "Microsoft themselves say the class \"doesn't support many modern protocols\". "
    "Supporting it means replacing the send path with MailKit or moving to Graph.",
    "Azure Communication Services Email - a separate Azure product with its own "
    "endpoint, resource and connection string, not a Microsoft 365 relay setting.",
    "Certificate-based Microsoft 365 connectors - ServiceMonitor.ps1 never sets "
    "SmtpClient.ClientCertificates, so no client certificate is presented.",
]


def read_smtp() -> dict:
    """Current smtp.json, or {} when the file is absent or unreadable."""
    if not SMTP_FILE.exists():
        return {}
    try:
        # utf-8-sig: New-CredStore.ps1 writes without a BOM, but a hand-edit with
        # PS 5.1's Set-Content -Encoding UTF8 would add one and break plain utf-8.
        with open(SMTP_FILE, encoding="utf-8-sig") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def write_smtp(server: str, port: int, frm: str, user: str,
               use_ssl: bool, preset: str) -> None:
    """Write smtp.json in New-CredStore.ps1's exact key order. No secrets here."""
    doc = {"server": server, "port": port, "from": frm,
           "user": user, "useSsl": use_ssl}
    if preset:
        doc["preset"] = preset
    tmp = SMTP_FILE.with_suffix(".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(doc, f, indent=2)
    shutil.move(str(tmp), str(SMTP_FILE))


def creds_stored() -> bool:
    """True when a usable credential store exists. File existence ONLY - the
    password is never decrypted, here or anywhere else in this UI."""
    return CRED_FILE.exists() and KEY_FILE.exists()


def stored_cred_user() -> str:
    """The username from smtp.user (plaintext by design; usernames aren't secret)."""
    try:
        return USER_FILE.read_text(encoding="utf-8-sig").strip()
    except Exception:
        return ""


def infer_preset(smtp: dict) -> str:
    """Best guess at the preset for an smtp.json written before this UI existed."""
    explicit = str(smtp.get("preset", "") or "")
    if explicit in PRESET_BY_ID:
        return explicit
    server = str(smtp.get("server", "") or "").lower()
    if server == "smtp.office365.com":
        return "m365-auth"
    if server in ("smtp.hve.mx.microsoft", "smtp-hve.office365.com"):
        return "m365-hve"
    if server.endswith(".mail.protection.outlook.com"):
        # Direct Send and connector relay share an endpoint and are
        # indistinguishable from the file alone. Reported as unknown rather than
        # guessed, so the UI does not assert something it cannot know.
        return ""
    return "custom" if server else ""


def validate_smtp(server: str, port_raw: str, frm: str, user: str,
                  use_ssl: bool, preset: str, cfg: dict) -> tuple:
    """(errors, warnings, port). Errors block the write; warnings do not."""
    errors, warns = [], []

    server = server.strip()
    if not server:
        errors.append("Server is required.")
    elif "://" in server:
        errors.append("Server must be a bare hostname or IP - drop the "
                      f"scheme from {server!r}.")
    elif ":" in server:
        errors.append(f"Server must not contain a port ({server!r}). "
                      "ServiceMonitor passes it to SmtpClient as a hostname, so "
                      "'host:port' fails DNS resolution. Use the Port field.")
    elif "/" in server or any(c.isspace() for c in server):
        errors.append(f"Server contains an illegal character: {server!r}.")
    elif not (HOST_RE.match(server) or _is_ipv4(server)):
        errors.append(f"{server!r} is not a valid hostname or IPv4 address. "
                      "IPv6 literals are not supported by this form.")

    port = 0
    port_raw = port_raw.strip()
    if not port_raw.isdigit():
        errors.append(f"Port must be a number, got {port_raw!r}.")
    else:
        port = int(port_raw)
        if not 1 <= port <= 65535:
            errors.append(f"Port {port} is out of range - must be 1-65535.")

    frm = frm.strip()
    if not frm:
        errors.append("From address is required - relays silently drop mail from "
                      "unauthorized senders.")
    elif not MAIL_RE.match(frm):
        errors.append(f"{frm!r} does not look like an email address.")

    user = user.strip()
    if user and (any(c.isspace() for c in user) or len(user) > 320):
        errors.append("Username must be a single token of at most 320 characters.")

    if errors:
        return errors, warns, port

    # --- warnings: everything below is written anyway ----------------------
    if port == 587 and not use_ssl:
        warns.append("Port 587 is the submission port and effectively always "
                     "requires STARTTLS. With TLS off the relay will almost "
                     "certainly refuse the session.")
    if port == 465:
        warns.append("Port 465 is implicit TLS. ServiceMonitor sends through "
                     "System.Net.Mail.SmtpClient, whose EnableSsl negotiates "
                     "STARTTLS after connecting in the clear - it does not do "
                     "implicit TLS, so port 465 is unlikely to work.")

    p = PRESET_BY_ID.get(preset)
    if p:
        if p["auth"] == "required" and not user:
            warns.append(f"{p['label']} authenticates. With no username, "
                         "ServiceMonitor will connect anonymously and be rejected.")
        if p["auth"] == "none" and user:
            warns.append(f"{p['label']} does not authenticate. A username is set, "
                         "so if a credential store also exists ServiceMonitor will "
                         "attempt AUTH and the relay may reject the session. Clear "
                         "the username unless you know you need it.")
        if p["internal_only"]:
            domain = frm.split("@")[-1].lower()
            outside = sorted({
                r for g in cfg.get("groups", {}).values()
                for r in g.get("recipients", [])
                if "@" in r and r.rsplit("@", 1)[-1].lower() != domain
            })
            if outside:
                warns.append(
                    f"{p['label']} delivers to tenant recipients only, but these "
                    "addresses are not on " + domain + ": " + ", ".join(outside)
                    + ". Alerts to them will be rejected.")
    if user and not creds_stored():
        warns.append("A username is set but no password is stored. Save the "
                     "credentials below, or ServiceMonitor will connect "
                     "anonymously.")
    return errors, warns, port


# ---------------------------------------------------------------------------
# One-shot notice shown on the next /settings render.
#
# The POST-redirect-GET pattern here has nowhere to carry a result, and these
# messages can be long (a full SMTP error). Kept in memory rather than a cookie.
# Safe because this app is documented to run single-worker on loopback; with
# more than one worker a notice could surface in the wrong worker's next render.
# Never holds a password: nothing that touches one ever calls set_notice.
# ---------------------------------------------------------------------------

_notice = None


def set_notice(kind: str, text: str, detail: str = "") -> None:
    global _notice
    with _lock:
        _notice = {"kind": kind, "text": text, "detail": detail}


def take_notice() -> Optional[dict]:
    global _notice
    with _lock:
        n, _notice = _notice, None
    return n


def _ps(*args: str) -> list:
    return ["powershell", "-NoProfile", "-NonInteractive",
            "-ExecutionPolicy", "Bypass", *args]


def _tail(text: str, n: int = 14) -> str:
    lines = [l.rstrip() for l in (text or "").splitlines() if l.strip()]
    return "\n".join(lines[-n:])


def run_credstore(server: str, port: int, frm: str, user: str,
                  use_ssl: bool, preset: str, password: str):
    """Shell out to New-CredStore.ps1 with the password on stdin.

    The password is written to the child's stdin pipe and nowhere else: not the
    command line (which any user could read via Get-CimInstance Win32_Process),
    not a temp file, not a log, not the redirect URL. The existing script is
    reused rather than reimplementing the AES/ACL/BSTR handling in Python.

    New-CredStore.ps1 rewrites smtp.json, so the current connection settings are
    passed straight back in and the file round-trips unchanged.
    """
    args = _ps("-File", str(CREDSTORE_PS1),
               "-PasswordFromStdin", "-Force",
               "-SmtpUser", user, "-SmtpServer", server,
               "-SmtpPort", str(port), "-FromAddress", frm)
    if use_ssl:
        args.append("-SmtpUseSsl")
    if preset:
        args += ["-Preset", preset]
    return subprocess.run(args, input=password + "\n",
                          capture_output=True, text=True, timeout=90)


def run_test_email():
    """Run ServiceMonitor.ps1 -TestEmail with no SMTP overrides.

    Deliberately passes no -Smtp* arguments, so this exercises the exact
    resolution path the scheduled task uses: smtp.json beside the script, plus
    auto-discovery of smtp.cred / smtp.key. Testing a different path than the one
    that runs in production would prove nothing.
    """
    return subprocess.run(_ps("-File", str(MONITOR_PS1), "-TestEmail"),
                          capture_output=True, text=True, timeout=180)


# ---------------------------------------------------------------------------
# Resolve the interactive Windows user behind a localhost web request.
#
# The web service runs as LocalSystem, so it cannot read the RDP user's name
# from the environment (that yields the service account). Instead we map the
# request's client TCP port back to the owning process (the browser), then that
# process's session -> username:
#     GetExtendedTcpTable  (client port  -> owning PID)
#     ProcessIdToSessionId (PID          -> session id)
#     WTSQuerySessionInformationW (session id -> DOMAIN\user)
# This is correct even on a multi-user RDP host where every session hits the
# same 127.0.0.1. Best-effort only: any failure returns "" and the UI falls
# back to the manual "Your name" field.
# ---------------------------------------------------------------------------

_AF_INET                 = 2
_TCP_TABLE_OWNER_PID_ALL = 5
_WTS_CURRENT_SERVER      = 0
_WTS_USER_NAME           = 5
_WTS_DOMAIN_NAME         = 7


class _MIB_TCPROW_OWNER_PID(ctypes.Structure):
    _fields_ = [
        ("dwState",      wintypes.DWORD),
        ("dwLocalAddr",  wintypes.DWORD),
        ("dwLocalPort",  wintypes.DWORD),
        ("dwRemoteAddr", wintypes.DWORD),
        ("dwRemotePort", wintypes.DWORD),
        ("dwOwningPid",  wintypes.DWORD),
    ]


def _pid_for_loopback_conn(local_port: int, remote_port: int) -> Optional[int]:
    """PID owning the IPv4 loopback TCP connection with these local+remote ports."""
    iphlpapi = ctypes.windll.iphlpapi
    fn = iphlpapi.GetExtendedTcpTable
    fn.restype = wintypes.DWORD
    fn.argtypes = [ctypes.c_void_p, ctypes.POINTER(wintypes.DWORD), wintypes.BOOL,
                   wintypes.ULONG, ctypes.c_int, wintypes.ULONG]

    size = wintypes.DWORD(0)
    # First call (NULL buffer) reports the required size.
    fn(None, ctypes.byref(size), False, _AF_INET, _TCP_TABLE_OWNER_PID_ALL, 0)
    buf = ctypes.create_string_buffer(size.value)
    if fn(ctypes.cast(buf, ctypes.c_void_p), ctypes.byref(size), False,
          _AF_INET, _TCP_TABLE_OWNER_PID_ALL, 0) != 0:
        return None

    num = ctypes.cast(buf, ctypes.POINTER(wintypes.DWORD))[0]
    # The row array follows the leading DWORD count (4-byte aligned; all-DWORD rows).
    rows = ctypes.cast(ctypes.addressof(buf) + 4,
                       ctypes.POINTER(_MIB_TCPROW_OWNER_PID * num))[0]
    for i in range(num):
        row = rows[i]
        lp = socket.ntohs(row.dwLocalPort & 0xFFFF)
        rp = socket.ntohs(row.dwRemotePort & 0xFFFF)
        if lp == local_port and rp == remote_port:
            return int(row.dwOwningPid)
    return None


def _wts_session_string(session_id: int, info_class: int) -> str:
    wtsapi = ctypes.windll.wtsapi32
    query = wtsapi.WTSQuerySessionInformationW
    query.restype = wintypes.BOOL
    query.argtypes = [wintypes.HANDLE, wintypes.DWORD, ctypes.c_int,
                      ctypes.POINTER(wintypes.LPWSTR), ctypes.POINTER(wintypes.DWORD)]
    ptr = wintypes.LPWSTR()
    nbytes = wintypes.DWORD(0)
    if not query(_WTS_CURRENT_SERVER, session_id, info_class,
                 ctypes.byref(ptr), ctypes.byref(nbytes)):
        return ""
    try:
        return ptr.value or ""
    finally:
        wtsapi.WTSFreeMemory(ptr)


def resolve_request_user(request: Request) -> str:
    """Best-effort DOMAIN\\user of the interactive session behind this request."""
    if os.name != "nt":
        return ""
    try:
        client = request.client
        server = request.scope.get("server")
        if not client or not server:
            return ""
        # uvicorn binds IPv4 127.0.0.1; only that maps via the IPv4 TCP table.
        if client.host != "127.0.0.1":
            return ""
        pid = _pid_for_loopback_conn(int(client.port), int(server[1]))
        if not pid:
            return ""
        sid = wintypes.DWORD()
        pid2sid = ctypes.windll.kernel32.ProcessIdToSessionId
        pid2sid.restype = wintypes.BOOL
        pid2sid.argtypes = [wintypes.DWORD, ctypes.POINTER(wintypes.DWORD)]
        if not pid2sid(pid, ctypes.byref(sid)):
            return ""
        user = _wts_session_string(sid.value, _WTS_USER_NAME)
        if not user:
            return ""
        domain = _wts_session_string(sid.value, _WTS_DOMAIN_NAME)
        return f"{domain}\\{user}" if domain else user
    except Exception:
        return ""


def get_user(request: Request) -> str:
    # An explicit name (typed into the "Your name" field) always wins; otherwise
    # default to the resolved interactive RDP user so the audit log is attributed
    # automatically instead of falling back to "unknown".
    return request.cookies.get("sm_user", "") or resolve_request_user(request)


# ---------------------------------------------------------------------------
# Windows service helpers  (one PowerShell call per page load)
# ---------------------------------------------------------------------------

def get_all_services_with_status() -> dict[str, str]:
    """Returns {service_name: status_string}. Empty dict on failure."""
    try:
        cmd = (
            "Get-Service | "
            "Select-Object Name, @{N='S';E={$_.Status.ToString()}} | "
            "ConvertTo-Json -Compress"
        )
        r = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", cmd],
            capture_output=True, text=True, timeout=20,
        )
        data = json.loads(r.stdout.strip())
        if isinstance(data, dict):
            data = [data]
        return {item["Name"]: item["S"] for item in data}
    except Exception:
        return {}


# ---------------------------------------------------------------------------
# HTML builder
# ---------------------------------------------------------------------------

CSS = """
:root{
  color-scheme:dark;
  --bg:#242424; --chrome:#303030; --card:rgba(255,255,255,.08);
  --card-brd:rgba(255,255,255,.06); --sep:rgba(255,255,255,.08);
  --fg:#ffffff; --dim:rgba(255,255,255,.55); --dim2:rgba(255,255,255,.40);
  --btn:rgba(255,255,255,.10); --btn-h:rgba(255,255,255,.16);
  --shade:rgba(0,0,0,.40); --entry:rgba(0,0,0,.28);
  --accent:#3584e4; --accent-t:#78aeed;
  --ok-d:#33d17a; --ok-t:#78e9ab;
  --err-d:#f66151; --err-t:#ff7b63; --err-bg:#c01c28;
  --wrn-d:#f6d32d; --wrn-t:#ffbe6f;
  --sans:Cantarell,"Segoe UI",system-ui,-apple-system,sans-serif;
  --mono:"Source Code Pro","Cascadia Mono",Consolas,ui-monospace,monospace;
}
*{box-sizing:border-box;margin:0;padding:0}
html,body{height:100%}
body{background:var(--bg);color:var(--fg);font-family:var(--sans);font-size:15px}
a{color:var(--accent-t);text-decoration:none}
a:hover{color:#a5c8f2}

/* ---- shell ---------------------------------------------------------- */
.shell{display:flex;min-height:100vh}
.sidebar{width:260px;flex-shrink:0;background:var(--chrome);
  border-right:1px solid var(--shade);display:flex;flex-direction:column;
  position:sticky;top:0;height:100vh}
.sb-head{height:47px;flex-shrink:0;display:flex;align-items:center;gap:10px;
  padding:0 12px 0 14px;border-bottom:1px solid var(--shade)}
.sb-icon{width:22px;height:22px;border-radius:6px;background:var(--accent);
  color:#fff;display:flex;align-items:center;justify-content:center;flex-shrink:0}
.sb-name{font-size:15px;font-weight:700}
.sb-ver{margin-left:auto;font-size:12px;color:var(--dim2)}
.nav{display:flex;flex-direction:column;gap:2px;padding:10px 6px}
.nav-row{display:flex;align-items:center;gap:12px;height:38px;padding:0 10px;
  border-radius:6px;color:rgba(255,255,255,.82);font-size:14px}
.nav-row:hover{background:rgba(255,255,255,.06);color:var(--fg)}
.nav-row.active{background:rgba(255,255,255,.10);color:var(--fg);font-weight:700}
.nav-count{margin-left:auto;font-size:13px;color:var(--dim2);
  font-variant-numeric:tabular-nums}
.sb-foot{margin-top:auto;padding:14px;border-top:1px solid var(--sep);
  display:flex;flex-direction:column;gap:6px}
.sb-foot .lbl{display:flex;align-items:center;gap:8px;font-size:13px;font-weight:700}
.sb-foot .sub{font-size:12px;color:var(--dim);font-family:var(--mono)}

.content{flex:1;display:flex;flex-direction:column;min-width:0}
.topbar{height:47px;flex-shrink:0;display:flex;align-items:center;gap:16px;
  padding:0 16px;background:var(--chrome);border-bottom:1px solid var(--shade);
  position:sticky;top:0;z-index:5}
.topbar h1{font-size:15px;font-weight:700}
.user-form{margin-left:auto;display:flex;align-items:center;gap:8px;font-size:13px}
.user-form label{color:var(--dim)}
.view{flex:1;padding:24px}
.cols{display:grid;grid-template-columns:minmax(0,680px) minmax(0,1fr);
  gap:24px;align-items:start}
@media(max-width:1180px){.cols{grid-template-columns:minmax(0,1fr)}}

/* ---- preference groups ---------------------------------------------- */
.group{margin:0 0 12px}
.group h2{font-size:16px;font-weight:700;margin-bottom:2px}
.group p{font-size:13px;color:var(--dim)}
.card{background:var(--card);border:1px solid var(--card-brd);
  border-radius:12px;overflow:hidden}
.card.scroll{max-height:520px;overflow-y:auto}
.row{display:flex;align-items:center;gap:14px;min-height:54px;
  padding:0 8px 0 16px;border-bottom:1px solid var(--sep)}
.row:last-child{border-bottom:none}
.row:hover{background:rgba(255,255,255,.03)}
.row.compact{min-height:46px;gap:12px}
.svc-name{flex:1;min-width:0;font-family:var(--mono);font-size:14px;
  overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.state{width:78px;flex-shrink:0;font-size:13px;display:flex;align-items:center;gap:9px}
.dot{width:9px;height:9px;border-radius:50%;flex-shrink:0}
.st-running .dot{background:var(--ok-d)} .st-running{color:var(--ok-t)}
.st-stopped .dot{background:var(--err-d)} .st-stopped{color:var(--err-t)}
.st-paused  .dot{background:rgba(255,255,255,.32)} .st-paused{color:var(--dim)}
.st-unknown .dot{background:var(--wrn-d)} .st-unknown{color:var(--wrn-t)}
.actions{display:flex;gap:2px;margin-left:8px;flex-shrink:0}

/* ---- controls -------------------------------------------------------- */
.btn{display:inline-flex;align-items:center;justify-content:center;height:34px;
  padding:0 16px;border:none;border-radius:6px;background:var(--btn);
  color:var(--fg);font-family:inherit;font-size:14px;cursor:pointer}
.btn:hover{background:var(--btn-h)}
.btn-suggested{background:var(--accent);color:#fff;font-weight:700}
.btn-suggested:hover{background:#3c8ff0}
.btn-danger{background:var(--err-bg);color:#fff;font-weight:700}
.btn-danger:hover{background:#d02231}
.icon-btn{width:34px;height:34px;padding:0;border:none;border-radius:6px;
  background:transparent;color:rgba(255,255,255,.82);cursor:pointer;
  display:inline-flex;align-items:center;justify-content:center}
.icon-btn:hover{background:var(--btn)}
.icon-btn.danger{color:var(--err-t)}
.icon-btn.danger:hover{background:rgba(224,27,36,.20)}
.icon-btn.plus{width:30px;height:30px;background:var(--accent);color:#fff}
.icon-btn.plus:hover{background:#3c8ff0}
.btn-sm{height:30px;padding:0 12px;font-size:13px}
select,input[type=text],input[type=email],input[type=password]{
  height:34px;padding:0 10px;border-radius:6px;font-family:inherit;font-size:14px;
  color:var(--fg);border:1px solid var(--sep);background:var(--entry)}
select{background:var(--btn);border:none;padding-right:6px;cursor:pointer}
select:hover{background:var(--btn-h)}
input::placeholder{color:var(--dim2)}
input:focus,select:focus{outline:none;border-color:var(--accent);
  box-shadow:0 0 0 2px rgba(53,132,228,.28)}
.entry-lg{height:38px;width:100%}
form.inline{display:inline-flex}

/* ---- misc ------------------------------------------------------------ */
.filter-bar{display:flex;align-items:center;gap:10px;margin:0 0 12px}
.filter-bar .hint{flex:1;font-size:13px;color:var(--dim)}
.foot-note{margin-top:10px;padding:0 2px;font-size:12px;color:var(--dim2)}
.empty{padding:16px;font-size:13px;color:var(--dim)}
.banner{display:flex;align-items:center;gap:10px;background:rgba(245,194,17,.14);
  border:1px solid rgba(245,194,17,.35);color:var(--wrn-t);padding:10px 14px;
  font-size:13px;border-radius:12px;margin-bottom:20px}
.log-row{display:flex;align-items:center;gap:20px;min-height:40px;padding:0 16px;
  font-family:var(--mono);font-size:13px;border-bottom:1px solid var(--sep)}
.log-row:last-child{border-bottom:none}
.log-ts{width:150px;flex-shrink:0;color:var(--dim2)}
.log-user{width:96px;flex-shrink:0;color:var(--accent-t)}
.log-act{flex:1;min-width:0;color:rgba(255,255,255,.82);
  overflow:hidden;text-overflow:ellipsis;white-space:nowrap}

/* ---- settings -------------------------------------------------------- */
.row.tall{min-height:60px;padding:11px 16px;align-items:flex-start}
.set-lbl{flex:1;min-width:0;font-size:14px;padding-top:7px}
.set-lbl small{display:block;margin-top:3px;font-size:12px;color:var(--dim);
  line-height:1.5}
.set-ctl{flex-shrink:0;width:300px;display:flex;align-items:center;gap:9px}
.set-ctl>input[type=text],.set-ctl>input[type=email],
.set-ctl>input[type=password],.set-ctl>select{width:100%}
.set-ctl input[type=checkbox]{width:18px;height:18px;flex-shrink:0;
  accent-color:var(--accent);cursor:pointer}
.set-ctl .cbx{display:flex;align-items:center;gap:9px;font-size:14px;
  color:rgba(255,255,255,.82);cursor:pointer}
.card-act{display:flex;align-items:center;gap:12px;padding:12px 16px;
  border-top:1px solid var(--sep)}
.card-act .hint{flex:1;min-width:0;font-size:12px;color:var(--dim2);
  line-height:1.5}
.pre{margin:0;padding:14px 16px;font-family:var(--mono);font-size:12.5px;
  line-height:1.65;color:rgba(255,255,255,.82);white-space:pre;overflow-x:auto}
.banner.msg{display:block;line-height:1.6}
.banner.ok{background:rgba(51,209,122,.12);
  border-color:rgba(51,209,122,.35);color:var(--ok-t)}
.banner.err{background:rgba(192,28,40,.16);
  border-color:rgba(192,28,40,.45);color:var(--err-t)}
.banner.msg .pre{padding:10px 0 0;color:inherit;white-space:pre-wrap;
  font-size:12px;opacity:.92}
.note-body{padding:14px 16px;font-size:13px;line-height:1.6;
  color:rgba(255,255,255,.82)}
.note-body p{margin-bottom:10px}
.note-body ul{margin:0;padding-left:17px;list-style:disc}
.note-body li{margin-bottom:7px}
.note-body li:last-child{margin-bottom:0}
.note-body .cap{display:block;margin:12px 0 6px;font-size:12px;font-weight:700;
  letter-spacing:.04em;text-transform:uppercase;color:var(--dim2)}
.note-body .cap.wrn{color:var(--wrn-t)}
.note-body code{font-family:var(--mono);font-size:12px;
  background:rgba(0,0,0,.28);padding:1px 5px;border-radius:6px}
"""

# --- 16px stroke icons (no emoji: they scale and recolor) -------------------
_SVG = ('<svg width="16" height="16" viewBox="0 0 16 16" fill="none" '
        'stroke="currentColor" stroke-width="1.5" stroke-linecap="round" '
        'stroke-linejoin="round">{}</svg>')
I_RACK  = _SVG.format('<rect x="2.25" y="2.75" width="11.5" height="4.5" rx="1.25"/>'
                      '<rect x="2.25" y="8.75" width="11.5" height="4.5" rx="1.25"/>'
                      '<path d="M4.75 5h.01M4.75 11h.01"/>')
I_MAIL  = _SVG.format('<rect x="2" y="3.75" width="12" height="8.5" rx="1.25"/>'
                      '<path d="m2.75 4.75 5.25 3.9 5.25-3.9"/>')
I_CLOCK = _SVG.format('<circle cx="8" cy="8" r="5.75"/><path d="M8 4.6V8l2.4 1.5"/>')
I_PAUSE = _SVG.format('<path d="M6 3.75v8.5M10 3.75v8.5"/>')
I_PLAY  = _SVG.format('<path d="M5.5 3.6 12.2 8l-6.7 4.4z"/>')
I_X     = _SVG.format('<path d="m4.25 4.25 7.5 7.5M11.75 4.25l-7.5 7.5"/>')
I_PLUS  = _SVG.format('<path d="M8 3.5v9M3.5 8h9"/>')
I_LOCK  = _SVG.format('<rect x="3.25" y="7" width="9.5" height="6.25" rx="1.5"/>'
                      '<path d="M5.75 7V5.25a2.25 2.25 0 0 1 4.5 0V7"/>')
# Envelope with an arrow leaving it: mail on its way out through a relay.
I_RELAY = _SVG.format('<path d="M13.75 7.4V4.5a1.25 1.25 0 0 0-1.25-1.25h-9'
                      'A1.25 1.25 0 0 0 2.25 4.5v6a1.25 1.25 0 0 0 1.25 1.25h4.4"/>'
                      '<path d="m2.75 4.75 5.25 3.9 5.25-3.9"/>'
                      '<path d="M10.75 12.25h3.5m-1.55-1.55 1.55 1.55-1.55 1.55"/>')

VIEWS = {"/": "Services", "/groups": "Mail groups", "/log": "Recent changes",
         "/settings": "Relay settings"}


def safe_back(back: str) -> str:
    """Only ever redirect to one of our own views."""
    return back if back in VIEWS else "/"


def status_badge(status: str) -> str:
    s = (status or "").lower()
    if s == "running":
        cls, label = "st-running", "Running"
    elif s == "stopped":
        cls, label = "st-stopped", "Stopped"
    elif s in ("paused", "pause_pending"):
        cls, label = "st-paused", "Paused"
    else:
        cls, label = "st-unknown", (status or "Unknown")
    return f'<span class="state {cls}"><span class="dot"></span>{label}</span>'


def alerts_select(svc_name: str, current: str, back: str) -> str:
    opts = [("both", "Both"), ("dev", "Dev only"), ("support", "Support only"), ("none", "None")]
    inner = "".join(
        f'<option value="{v}"{" selected" if v == current else ""}>{label}</option>'
        for v, label in opts
    )
    return (
        '<form method="post" action="/services/set-alerts" class="inline">'
        f'<input type="hidden" name="name" value="{svc_name}">'
        f'<input type="hidden" name="back" value="{back}">'
        f'<select name="alerts" onchange="this.form.submit()" title="Alert routing">{inner}</select>'
        '</form>'
    )


def build_monitored_section(services: list, statuses: dict) -> str:
    if not services:
        return ('<div class="card"><p class="empty">No services monitored yet. '
                'Add some from the panel on the right.</p></div>')

    rows = ""
    for svc in services:
        name   = svc.get("name", "")
        paused = svc.get("paused", False)
        alerts = svc.get("alerts", "both")
        if paused:
            st_html   = status_badge("paused")
            icon, tip = I_PLAY, "Resume monitoring"
        else:
            st_html   = status_badge(statuses.get(name, ""))
            icon, tip = I_PAUSE, "Pause monitoring"
        pause_btn = (
            '<form method="post" action="/services/toggle-pause" class="inline">'
            f'<input type="hidden" name="name" value="{name}">'
            '<input type="hidden" name="back" value="/">'
            f'<button class="icon-btn" title="{tip}">{icon}</button></form>'
        )
        remove_btn = (
            '<form method="post" action="/services/remove" class="inline">'
            f'<input type="hidden" name="name" value="{name}">'
            '<input type="hidden" name="back" value="/">'
            f'<button class="icon-btn danger" title="Remove from monitoring" '
            f'onclick="return confirm(\'Remove {name} from monitoring?\')">{I_X}</button></form>'
        )
        rows += (
            '<div class="row">'
            f'<span class="svc-name">{name}</span>'
            f'{st_html}'
            f'{alerts_select(name, alerts, "/")}'
            f'<div class="actions">{pause_btn}{remove_btn}</div>'
            '</div>'
        )
    return f'<div class="card">{rows}</div>'


def build_available_section(monitored_names: set, statuses: dict) -> str:
    available = sorted(
        name for name in statuses
        if name not in monitored_names and name not in SYSTEM_SERVICES
    )
    rows = ""
    for name in available:
        rows += (
            '<form method="post" action="/services/add" class="svc-item">'
            f'<input type="hidden" name="name" value="{name}">'
            f'<input type="hidden" name="alerts" id="default-alerts-{name}">'
            '<input type="hidden" name="back" value="/">'
            '<div class="row compact">'
            f'<span class="svc-name">{name}</span>'
            f'<button type="submit" class="icon-btn plus" title="Add {name} to monitoring" '
            'onclick="'
            f"document.getElementById('default-alerts-{name}').value="
            "document.getElementById('default-alerts').value"
            f'">{I_PLUS}</button>'
            '</div></form>'
        )
    if not rows:
        rows  = ('<p class="empty">No additional services found '
                 '(or PowerShell unavailable).</p>')
        count = ""
    else:
        count = (f'<p class="foot-note">{len(available)} non-system '
                 'services detected on this host.</p>')

    return (
        '<input type="text" id="svc-filter" class="entry-lg" '
        'placeholder="Filter services..." oninput="filterSvcs(this.value)" '
        'style="margin-bottom:12px">'
        '<div class="filter-bar">'
        '<span class="hint">Alerts for newly added</span>'
        '<select id="default-alerts" title="Alert group for newly added services">'
        '<option value="both">Both</option>'
        '<option value="dev">Dev only</option>'
        '<option value="support">Support only</option>'
        '<option value="none">None</option>'
        '</select></div>'
        f'<div class="card scroll" id="svc-list">{rows}</div>{count}'
    )


def build_groups_section(groups: dict) -> str:
    cols = ""
    for gid, gdata in groups.items():
        label      = gdata.get("label", gid)
        recipients = gdata.get("recipients", [])
        rows = ""
        for email in recipients:
            rows += (
                '<div class="row compact">'
                '<span style="flex:1;min-width:0;font-size:14px;overflow:hidden;'
                f'text-overflow:ellipsis;white-space:nowrap">{email}</span>'
                '<form method="post" action="/recipients/remove" class="inline">'
                f'<input type="hidden" name="group" value="{gid}">'
                f'<input type="hidden" name="email" value="{email}">'
                '<input type="hidden" name="back" value="/groups">'
                f'<button class="icon-btn danger" title="Remove recipient" '
                f'onclick="return confirm(\'Remove {email}?\')">{I_X}</button>'
                '</form></div>'
            )
        if not rows:
            rows = '<p class="empty">No recipients in this group.</p>'
        add_form = (
            '<div class="row compact" style="gap:8px;padding:10px 10px 10px 16px">'
            '<form method="post" action="/recipients/add" '
            'style="display:flex;gap:8px;flex:1;min-width:0">'
            f'<input type="hidden" name="group" value="{gid}">'
            '<input type="hidden" name="back" value="/groups">'
            '<input type="email" name="email" placeholder="email@example.com" '
            'required style="flex:1;min-width:0">'
            '<button class="btn btn-suggested">Add</button>'
            '</form></div>'
        )
        cols += (
            '<div>'
            f'<div class="group"><h2>{label}</h2>'
            f'<p>{len(recipients)} recipient(s) &middot; group id '
            f'<span style="font-family:var(--mono)">{gid}</span></p></div>'
            f'<div class="card">{rows}{add_form}</div>'
            '</div>'
        )
    return ('<div style="display:grid;grid-template-columns:repeat(2,minmax(0,1fr));'
            f'gap:24px;align-items:start;max-width:1000px">{cols}</div>')


def build_changes_section(lines: list) -> str:
    if not lines:
        return '<div class="card"><p class="empty">No changes recorded yet.</p></div>'
    rows = ""
    for line in lines:
        # Log format: "2026-06-26 10:30  username              action"
        # Timestamp has a space -> need 3 splits to get [date, time, user, action]
        parts = line.split(None, 3)
        if len(parts) == 4:
            ts = f"{parts[0]} {parts[1]}"
            user, action = parts[2], parts[3]
            rows += (f'<div class="log-row"><span class="log-ts">{ts}</span>'
                     f'<span class="log-user">{user}</span>'
                     f'<span class="log-act">{action}</span></div>')
        else:
            rows += f'<div class="log-row"><span class="log-act">{line}</span></div>'
    return f'<div class="card">{rows}</div>'


VERSION = "1.4.0"


# The task state is read live rather than assumed from a default, because the
# polling interval is chosen at install time by -IntervalMinutes. Cached for
# TASK_CACHE_TTL seconds so paging around the UI does not spawn a PowerShell
# process per request.
TASK_CACHE_TTL = 30.0
_task_cache = {"at": 0.0, "val": None}


def scheduled_task_state() -> tuple:
    """(state, detail) for the ServiceMonitor scheduled task."""
    now = time.monotonic()
    cached = _task_cache["val"]
    if cached is not None and now - _task_cache["at"] < TASK_CACHE_TTL:
        return cached
    try:
        cmd = (
            "$t = Get-ScheduledTask -TaskName ServiceMonitor -ErrorAction Stop; "
            "$m = ($t.Triggers | ForEach-Object { $_.Repetition.Interval } | "
            "Select-Object -First 1); "
            "[pscustomobject]@{S=$t.State.ToString();I=$m;U=$t.Principal.UserId} | "
            "ConvertTo-Json -Compress"
        )
        r = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", cmd],
            capture_output=True, text=True, timeout=20,
        )
        data = json.loads(r.stdout.strip())
        state = data.get("S") or "Unknown"
        # ISO-8601 duration, e.g. PT5M -> "every 5 min"
        iso = (data.get("I") or "").upper()
        detail = ""
        if iso.startswith("PT"):
            body = iso[2:]
            if body.endswith("M") and body[:-1].isdigit():
                detail = "every " + body[:-1] + " min"
            elif body.endswith("H") and body[:-1].isdigit():
                detail = "every " + body[:-1] + " h"
        who = data.get("U") or ""
        parts = ["ServiceMonitor"] + [p for p in (detail, who) if p]
        result = (state, " &middot; ".join(parts))
    except Exception:
        result = ("", "Scheduled task not found")
    _task_cache["at"], _task_cache["val"] = now, result
    return result


def _task_footer() -> str:
    state, detail = scheduled_task_state()
    if state in ("Ready", "Running"):
        dot, label = "var(--ok-d)", "Scheduled task active"
    elif state == "Disabled":
        dot, label = "var(--wrn-d)", "Scheduled task disabled"
    elif state:
        dot, label = "var(--wrn-d)", "Scheduled task " + state
    else:
        dot, label = "rgba(255,255,255,.32)", "Scheduled task missing"
    return (
        '<div class="sb-foot">'
        f'<span class="lbl"><span class="dot" style="background:{dot}"></span>{label}</span>'
        f'<span class="sub">{detail}</span></div>'
    )


def _sidebar(active: str, n_services: int, n_groups: int) -> str:
    def row(path, icon, label, count):
        cls = "nav-row active" if path == active else "nav-row"
        cnt = f'<span class="nav-count">{count}</span>' if count != "" else ""
        return f'<a class="{cls}" href="{path}">{icon}<span>{label}</span>{cnt}</a>'

    return (
        '<aside class="sidebar">'
        '<div class="sb-head">'
        f'<span class="sb-icon">{I_RACK}</span>'
        '<span class="sb-name">ServiceMonitor</span>'
        f'<span class="sb-ver">v{VERSION}</span>'
        '</div><nav class="nav">'
        + row("/", I_RACK, "Services", n_services)
        + row("/groups", I_MAIL, "Mail groups", n_groups)
        + row("/log", I_CLOCK, "Recent changes", "")
        + row("/settings", I_RELAY, "Relay settings", "")
        + '</nav>' + _task_footer() + '</aside>'
    )


JS = r"""
<script>
function filterSvcs(q){
  q=q.toLowerCase();
  document.querySelectorAll('#svc-list .svc-item').forEach(function(el){
    el.style.display=el.querySelector('.svc-name').textContent.toLowerCase().includes(q)?'':'none';
  });
}
/* --- settings view. All of these no-op unless #preset is on the page. --- */
function smtpNotes(){
  var v=document.getElementById('preset').value;
  document.querySelectorAll('.preset-note').forEach(function(el){
    el.style.display=(el.getAttribute('data-preset')===v)?'':'none';
  });
}
function smtpDerive(){
  var d=SM_PRESETS[document.getElementById('preset').value];
  if(!d||!d.derive){return;}
  var f=document.getElementById('sender').value.trim(),at=f.lastIndexOf('@');
  if(at<0){return;}
  var dom=f.slice(at+1).toLowerCase();
  if(!dom){return;}
  document.getElementById('server').value=
    dom.replace(/\./g,'-')+'.mail.protection.outlook.com';
}
function smtpPreset(){
  var d=SM_PRESETS[document.getElementById('preset').value];
  smtpNotes();
  if(d){
    document.getElementById('port').value=d.port;
    document.getElementById('ssl').checked=d.ssl;
    if(d.derive){smtpDerive();}
    else if(d.server){document.getElementById('server').value=d.server;}
  }
  smtpPreview();
}
function smtpPreview(){
  var o={server:document.getElementById('server').value.trim(),
         port:parseInt(document.getElementById('port').value,10),
         from:document.getElementById('sender').value.trim(),
         user:SM_USER,
         useSsl:document.getElementById('ssl').checked};
  if(isNaN(o.port)){o.port=0;}
  var p=document.getElementById('preset').value;
  if(p){o.preset=p;}
  document.getElementById('preview').textContent=JSON.stringify(o,null,2);
}
document.addEventListener('DOMContentLoaded',function(){
  /* Note toggle + preview only. Never smtpPreset() on load: that would
     overwrite the saved values with the preset's defaults. */
  if(document.getElementById('preset')){smtpNotes();smtpPreview();}
});
</script>
"""


def build_shell(title: str, active: str, body: str, cfg: dict, user: str) -> str:
    n_services = len(cfg.get("services", []))
    n_groups   = len(cfg.get("groups", {}))

    banner = ""
    if not user:
        banner = ('<div class="banner">Set your name in the top bar so changes are '
                  'attributed in the audit log.</div>')

    user_form = (
        '<form class="user-form" method="post" action="/set-user">'
        '<label for="sm-user">You are</label>'
        f'<input id="sm-user" name="username" value="{user}" placeholder="Your name" '
        'maxlength="40" style="height:30px;width:150px">'
        f'<input type="hidden" name="back" value="{active}">'
        '<button class="btn btn-sm">Set</button>'
        '</form>'
    )

    return (
        "<!DOCTYPE html>"
        '<html lang="en"><head><meta charset="UTF-8">'
        '<meta name="viewport" content="width=device-width,initial-scale=1">'
        f'<title>{title} - ServiceMonitor</title>'
        f"<style>{CSS}</style></head><body>"
        '<div class="shell">'
        + _sidebar(active, n_services, n_groups) +
        '<div class="content">'
        f'<header class="topbar"><h1>{title}</h1>{user_form}</header>'
        f'<main class="view">{banner}{body}</main>'
        '</div></div>' + JS + '</body></html>'
    )


def build_page(cfg: dict, statuses: dict, user: str) -> str:
    """Services view - the landing page."""
    services      = cfg.get("services", [])
    monitored_set = {s["name"] for s in services}
    n_paused      = sum(1 for s in services if s.get("paused"))
    n_active      = len(services) - n_paused

    left = (
        '<div><div class="group"><h2>Monitored services</h2>'
        f'<p>{n_active} checked on every run &middot; {n_paused} paused</p></div>'
        f'{build_monitored_section(services, statuses)}</div>'
    )
    right = (
        '<div><div class="group"><h2>Add a service</h2>'
        '<p>Standard Windows services are hidden from this list.</p></div>'
        f'{build_available_section(monitored_set, statuses)}</div>'
    )
    body = f'<div class="cols">{left}{right}</div>'
    return build_shell("Services", "/", body, cfg, user)


def build_groups_page(cfg: dict, user: str) -> str:
    body = build_groups_section(cfg.get("groups", {}))
    return build_shell("Mail groups", "/groups", body, cfg, user)


def build_log_page(cfg: dict, user: str) -> str:
    body = (
        '<div style="max-width:900px">'
        '<div class="group"><h2>Recent changes</h2>'
        '<p>Appended to changes.log next to monitor-config.json. Last 20 entries.</p></div>'
        f'{build_changes_section(recent_changes())}</div>'
    )
    return build_shell("Recent changes", "/log", body, cfg, user)


# ---------------------------------------------------------------------------
# Relay settings view
# ---------------------------------------------------------------------------

def _esc(v) -> str:
    return html.escape(str(v), quote=True)


def _set_row(label: str, sub: str, control: str) -> str:
    """One preference row: label (+ explanation) on the left, control right."""
    sub_html = f"<small>{sub}</small>" if sub else ""
    return (f'<div class="row tall"><span class="set-lbl">{label}{sub_html}</span>'
            f'<span class="set-ctl">{control}</span></div>')


def _notice_html(n) -> str:
    if not n:
        return ""
    cls = {"ok": "banner msg ok", "err": "banner msg err"}.get(n.get("kind", ""),
                                                               "banner msg")
    out = f'<div class="{cls}"><strong>{_esc(n.get("text", ""))}</strong>'
    if n.get("detail"):
        out += f'<pre class="pre">{_esc(n["detail"])}</pre>'
    return out + "</div>"


def build_relay_card(smtp: dict, cur_preset: str) -> str:
    opts = ""
    if cur_preset not in PRESET_BY_ID:
        opts += ('<option value="" selected>Not recognised - choose one</option>')
    for pr in PRESETS:
        sel = " selected" if pr["id"] == cur_preset else ""
        opts += f'<option value="{pr["id"]}"{sel}>{_esc(pr["label"])}</option>'

    server = _esc(smtp.get("server", ""))
    port   = _esc(smtp.get("port", 25))
    sender = _esc(smtp.get("from", ""))
    ssl_on = " checked" if smtp.get("useSsl") else ""

    rows = (
        _set_row("Preset",
                 "Prefills the fields below. Nothing is hidden - every value is "
                 "shown and editable before you save.",
                 f'<select id="preset" name="preset" onchange="smtpPreset()">'
                 f'{opts}</select>')
        + _set_row("Server",
                   "Hostname or IPv4 address, <b>without</b> a port. "
                   "ServiceMonitor hands this straight to "
                   "<code>SmtpClient(host, port)</code>.",
                   f'<input type="text" id="server" name="server" value="{server}" '
                   'oninput="smtpPreview()" placeholder="relay.example.com" '
                   'spellcheck="false" autocapitalize="off">')
        + _set_row("Port", "25 for anonymous relay, 587 for STARTTLS submission.",
                   f'<input type="text" id="port" name="port" value="{port}" '
                   'inputmode="numeric" oninput="smtpPreview()">')
        + _set_row("STARTTLS",
                   "Sets <code>SmtpClient.EnableSsl</code>. This is STARTTLS on the "
                   "port above, not implicit TLS on 465.",
                   '<label class="cbx"><input type="checkbox" id="ssl" name="usessl" '
                   f'value="on"{ssl_on} onchange="smtpPreview()">Enabled</label>')
        + _set_row("From address",
                   "Must be on a domain the relay accepts mail for, or it will take "
                   "the message and silently drop it.",
                   f'<input type="email" id="sender" name="sender" value="{sender}" '
                   'oninput="smtpDerive();smtpPreview()" '
                   'placeholder="servicemonitor@example.com" spellcheck="false">')
    )
    return (
        '<form method="post" action="/settings/smtp">'
        '<input type="hidden" name="back" value="/settings">'
        f'<div class="card">{rows}'
        '<div class="card-act">'
        '<span class="hint">Written to smtp.json. The stored username and password '
        'are managed separately, below.</span>'
        '<button class="btn btn-suggested">Save relay settings</button>'
        '</div></div></form>'
    )


def build_credentials_card(smtp: dict) -> str:
    have = creds_stored()
    if have:
        who   = stored_cred_user() or str(smtp.get("user", "") or "") or "(unknown)"
        state = ('<span class="state st-running" style="width:auto">'
                 f'<span class="dot"></span>Stored</span>')
        detail = (f"Encrypted password on disk for <code>{_esc(who)}</code>. "
                  "The password is never decrypted or displayed by this UI - "
                  "only the presence of smtp.cred and smtp.key is checked.")
    else:
        state = ('<span class="state st-paused" style="width:auto">'
                 '<span class="dot"></span>Not stored</span>')
        detail = ("No smtp.cred / smtp.key pair on disk. ServiceMonitor will "
                  "connect to the relay anonymously.")

    user_val = _esc(stored_cred_user() or smtp.get("user", ""))
    rows = (
        _set_row("Credential store", detail, state)
        + _set_row("Username",
                   "For Microsoft 365 SMTP AUTH this is the mailbox address. "
                   "Saved in clear text to smtp.user - usernames are not secrets.",
                   f'<input type="text" name="username" value="{user_val}" '
                   'placeholder="alerts@example.com" spellcheck="false" '
                   'autocapitalize="off">')
        + _set_row("Password",
                   "Handed to New-CredStore.ps1 on stdin, encrypted with a random "
                   "AES-256 key, and never written to a log, a temp file or a "
                   "command line.",
                   '<input type="password" name="password" '
                   'placeholder="Not shown once stored" autocomplete="new-password">')
    )
    clear_btn = ""
    if have:
        clear_btn = (
            '<form method="post" action="/settings/credentials/clear" class="inline">'
            '<input type="hidden" name="back" value="/settings">'
            '<button class="btn btn-danger" onclick="return confirm('
            "'Delete smtp.key, smtp.cred and smtp.user? ServiceMonitor will fall "
            "back to anonymous relay.')\">Clear</button></form>"
        )
    return (
        '<form method="post" action="/settings/credentials">'
        '<input type="hidden" name="back" value="/settings">'
        f'<div class="card">{rows}'
        '<div class="card-act">'
        '<span class="hint">Saving runs New-CredStore.ps1, which also rewrites '
        'smtp.json from the values above - save the relay settings first.</span>'
        f'{clear_btn}'
        '<button class="btn btn-suggested">Store credentials</button>'
        '</div></div></form>'
    )


def build_test_card() -> str:
    return (
        '<form method="post" action="/settings/test-email">'
        '<input type="hidden" name="back" value="/settings">'
        '<div class="card">'
        + _set_row("Send test email",
                   "Runs <code>ServiceMonitor.ps1 -TestEmail</code> with no SMTP "
                   "overrides, so it resolves the relay exactly the way the "
                   "scheduled task does. Mails everyone in both groups. The result "
                   "appears at the top of this page.",
                   '<button class="btn btn-suggested">Send test</button>')
        + '</div></form>'
        '<p class="foot-note">A success means the relay <b>accepted</b> the message. '
        'It does not prove the message reached an inbox - check the mailbox too.</p>'
    )


def build_preset_notes(cur_preset: str) -> str:
    blocks = ""
    for pr in PRESETS:
        shown = "" if pr["id"] == cur_preset else ' style="display:none"'
        body = f'<p>{_esc(pr["summary"])}</p>'
        body += '<span class="cap">How it behaves</span><ul>'
        body += "".join(f"<li>{_esc(f)}</li>" for f in pr["facts"])
        body += "</ul>"
        if pr["warns"]:
            body += '<span class="cap wrn">Before you rely on it</span><ul>'
            body += "".join(f'<li>{_esc(w)}</li>' for w in pr["warns"])
            body += "</ul>"
        blocks += (f'<div class="preset-note" data-preset="{pr["id"]}"{shown}>'
                   f'<div class="note-body">{body}</div></div>')
    return f'<div class="card">{blocks}</div>'


def build_ondisk_card(smtp: dict) -> str:
    if not SMTP_FILE.exists():
        return ('<div class="card"><p class="empty">'
                f'{_esc(SMTP_FILE)} does not exist. ServiceMonitor.ps1 is running on '
                'its own built-in parameter defaults - send a test email and read '
                'the log to see which relay that actually is.</p></div>')
    try:
        raw = SMTP_FILE.read_text(encoding="utf-8-sig")
    except Exception as exc:
        return f'<div class="card"><p class="empty">Unreadable: {_esc(exc)}</p></div>'
    return f'<div class="card"><pre class="pre">{_esc(raw.strip())}</pre></div>'


def build_settings_page(cfg: dict, user: str) -> str:
    smtp       = read_smtp()
    cur_preset = infer_preset(smtp)
    js_presets = json.dumps({pr["id"]: {"server": pr["server"], "port": pr["port"],
                                        "ssl": pr["ssl"], "derive": pr["derive"]}
                             for pr in PRESETS})
    js_user = json.dumps(str(smtp.get("user", "") or ""))

    left = (
        '<div>'
        '<div class="group"><h2>Relay server</h2>'
        '<p>Connection details for the SMTP relay that carries the alerts. '
        f'Stored in <span style="font-family:var(--mono)">{_esc(SMTP_FILE)}</span>.</p>'
        '</div>'
        f'{build_relay_card(smtp, cur_preset)}'
        '<div class="group" style="margin-top:24px"><h2>Authentication</h2>'
        '<p>Optional. An anonymous relay needs none of this.</p></div>'
        f'{build_credentials_card(smtp)}'
        '<div class="group" style="margin-top:24px"><h2>Verify</h2>'
        '<p>The only honest proof that a relay configuration works.</p></div>'
        f'{build_test_card()}'
        '</div>'
    )

    not_impl = "".join(f"<li>{_esc(n)}</li>" for n in NOT_IMPLEMENTED)
    right = (
        '<div>'
        '<div class="group"><h2>About this preset</h2>'
        '<p>Sourced from Microsoft Learn. Read the second list before a customer '
        'tenant.</p></div>'
        f'{build_preset_notes(cur_preset)}'
        '<div class="group" style="margin-top:24px"><h2>What will be written</h2>'
        '<p>Live preview of the JSON the Save button produces.</p></div>'
        '<div class="card"><pre class="pre" id="preview"></pre></div>'
        '<div class="group" style="margin-top:24px"><h2>On disk now</h2>'
        '<p>The current file, verbatim.</p></div>'
        f'{build_ondisk_card(smtp)}'
        '<div class="group" style="margin-top:24px"><h2>Not offered here</h2>'
        '<p>Deliberate omissions, so they are visible rather than silent.</p></div>'
        f'<div class="card"><div class="note-body"><ul>{not_impl}</ul></div></div>'
        '</div>'
    )

    body = (
        f'<script>var SM_PRESETS={js_presets};var SM_USER={js_user};</script>'
        f'{_notice_html(take_notice())}'
        f'<div class="cols">{left}{right}</div>'
    )
    return build_shell("Relay settings", "/settings", body, cfg, user)


# ---------------------------------------------------------------------------
# Optional access-token gate (defense-in-depth on shared RDP hosts)
# ---------------------------------------------------------------------------

def _login_page(message: str = "") -> str:
    return (
        '<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">'
        '<meta name="viewport" content="width=device-width,initial-scale=1">'
        '<title>Sign in - ServiceMonitor</title>'
        f"<style>{CSS}</style></head><body>"
        '<div style="min-height:100vh;display:flex;align-items:center;'
        'justify-content:center;padding:24px">'
        '<div style="width:380px;background:#383838;'
        'border:1px solid rgba(255,255,255,.06);border-radius:12px;'
        'padding:28px 24px 24px;box-shadow:0 12px 32px rgba(0,0,0,.45)">'
        '<div style="width:44px;height:44px;border-radius:12px;background:var(--accent);'
        'color:#fff;display:flex;align-items:center;justify-content:center;'
        f'margin:0 auto 16px">{I_LOCK}</div>'
        '<h1 style="font-size:19px;font-weight:700;text-align:center;'
        'margin-bottom:4px">ServiceMonitor</h1>'
        '<p style="font-size:14px;color:var(--dim);text-align:center;'
        'margin-bottom:20px">Enter the access token to continue.</p>'
        f'{message}'
        '<form method="post" action="/auth">'
        '<input type="password" name="token" placeholder="Access token" autofocus '
        'class="entry-lg" style="margin-bottom:14px">'
        '<button class="btn btn-suggested" style="width:100%;height:38px">Sign in</button>'
        '</form>'
        '<p style="margin-top:20px;font-size:12px;color:var(--dim2);text-align:center;'
        'line-height:1.5">Bound to 127.0.0.1. This gate only appears when '
        '<span style="font-family:var(--mono)">SM_TOKEN</span> is set; otherwise the '
        'Windows login is the only layer.</p>'
        '</div></div></body></html>'
    )


def _token_matches(supplied: str) -> bool:
    # Compare as bytes: hmac.compare_digest raises TypeError on non-ASCII str,
    # which would 500 the whole UI if an admin chose a token with accented chars.
    return hmac.compare_digest(supplied.encode("utf-8"), ACCESS_TOKEN.encode("utf-8"))


def _is_authed(request: Request) -> bool:
    if not ACCESS_TOKEN:
        return True
    return _token_matches(request.cookies.get("sm_token", ""))


@app.middleware("http")
async def require_token(request: Request, call_next):
    # No token configured, or request is the auth POST itself => let it through.
    if not ACCESS_TOKEN or request.url.path == "/auth":
        return await call_next(request)
    if _is_authed(request):
        return await call_next(request)
    return HTMLResponse(_login_page(), status_code=401)


@app.post("/auth")
async def auth(token: str = Form(...)):
    if ACCESS_TOKEN and _token_matches(token.strip()):
        resp = RedirectResponse("/", status_code=302)
        resp.set_cookie(
            "sm_token", token.strip(),
            max_age=2_592_000, httponly=True, samesite="strict",
        )
        return resp
    msg = ('<p style="color:#dc2626;font-size:13px;margin-bottom:12px">'
           "Invalid token.</p>")
    return HTMLResponse(_login_page(msg), status_code=401)


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    cfg      = read_config()
    statuses = get_all_services_with_status()
    user     = get_user(request)
    return build_page(cfg, statuses, user)


@app.get("/groups", response_class=HTMLResponse)
async def groups_view(request: Request):
    return build_groups_page(read_config(), get_user(request))


@app.get("/log", response_class=HTMLResponse)
async def log_view(request: Request):
    return build_log_page(read_config(), get_user(request))


@app.get("/settings", response_class=HTMLResponse)
async def settings_view(request: Request):
    return build_settings_page(read_config(), get_user(request))


@app.post("/settings/smtp")
async def settings_smtp(request: Request, preset: str = Form(""),
                        server: str = Form(""), port: str = Form(""),
                        sender: str = Form(""), usessl: str = Form(""),
                        back: str = Form("/settings")):
    # The form field is 'sender', not 'from': 'from' is a Python keyword and
    # cannot be a parameter name. It is written to smtp.json as "from".
    user = get_user(request) or "unknown"
    if preset not in PRESET_BY_ID:
        preset = ""
    ssl_on = usessl == "on"
    cfg    = read_config()
    smtp   = read_smtp()
    # The username belongs to the credential store, which has its own form.
    # Carry it through untouched so saving relay settings never clears it.
    keep_user = str(smtp.get("user", "") or "")

    errors, warns, port_i = validate_smtp(server, port, sender, keep_user,
                                          ssl_on, preset, cfg)
    if errors:
        set_notice("err", "Nothing was saved - fix these first:",
                   "\n".join("- " + e for e in errors))
        return RedirectResponse(safe_back(back), status_code=302)

    server, sender = server.strip(), sender.strip()
    with _lock:
        write_smtp(server, port_i, sender, keep_user, ssl_on, preset)
        # Connection details only. There is no password anywhere in this line.
        log_change(user, f"SMTP relay set to {server}:{port_i} "
                         f"(tls {'on' if ssl_on else 'off'}, from {sender}, "
                         f"preset {preset or 'none'})")
    if warns:
        set_notice("warn", f"Saved to {SMTP_FILE}, but read these:",
                   "\n".join("- " + w for w in warns))
    else:
        set_notice("ok", f"Saved to {SMTP_FILE}.")
    return RedirectResponse(safe_back(back), status_code=302)


@app.post("/settings/credentials")
async def settings_credentials(request: Request, username: str = Form(""),
                               password: str = Form(""),
                               back: str = Form("/settings")):
    user     = get_user(request) or "unknown"
    username = username.strip()
    smtp     = read_smtp()

    if not CREDSTORE_PS1.exists():
        set_notice("err", f"{CREDSTORE_PS1} not found.",
                   "The credential store is created by that script; this UI only "
                   "drives it.")
    elif not smtp.get("server") or not smtp.get("from"):
        set_notice("err", "Save the relay server settings first.",
                   "New-CredStore.ps1 rewrites smtp.json, so it needs a server, "
                   "port and From address to write back.")
    elif not username or not password:
        set_notice("err", "A username and a password are both required.")
    elif any(c.isspace() for c in username) or len(username) > 320:
        set_notice("err", "Username must be one token of at most 320 characters.")
    elif "\n" in password or "\r" in password:
        set_notice("err", "The password cannot contain a line break.",
                   "It is passed to New-CredStore.ps1 as exactly one line on stdin.")
    else:
        try:
            r = run_credstore(str(smtp["server"]), int(smtp.get("port", 25) or 25),
                              str(smtp["from"]), username,
                              bool(smtp.get("useSsl")),
                              str(smtp.get("preset", "") or ""), password)
        except subprocess.TimeoutExpired:
            set_notice("err", "New-CredStore.ps1 timed out after 90 s.")
            return RedirectResponse(safe_back(back), status_code=302)
        except Exception as exc:
            set_notice("err", "Could not run New-CredStore.ps1.", str(exc))
            return RedirectResponse(safe_back(back), status_code=302)
        finally:
            # Best effort only: CPython strings are immutable, so this drops the
            # reference but cannot scrub the buffer. Stated plainly rather than
            # implied to be a wipe.
            password = ""

        if r.returncode == 0 and creds_stored():
            # Username only. The password never reaches the log.
            log_change(user, f"Stored SMTP credentials for user {username!r}")
            set_notice("ok", f"Credentials stored for {username}.", _tail(r.stdout))
        else:
            set_notice("err", f"New-CredStore.ps1 failed (exit {r.returncode}).",
                       _tail(r.stderr) or _tail(r.stdout) or "No output.")
    return RedirectResponse(safe_back(back), status_code=302)


@app.post("/settings/credentials/clear")
async def settings_credentials_clear(request: Request,
                                     back: str = Form("/settings")):
    user, removed = get_user(request) or "unknown", []
    with _lock:
        for path in (CRED_FILE, KEY_FILE, USER_FILE):
            try:
                if path.exists():
                    path.unlink()
                    removed.append(path.name)
            except Exception as exc:
                set_notice("err", f"Could not delete {path.name}.", str(exc))
                return RedirectResponse(safe_back(back), status_code=302)
        smtp = read_smtp()
        if smtp.get("user"):
            write_smtp(str(smtp.get("server", "")), int(smtp.get("port", 25) or 25),
                       str(smtp.get("from", "")), "", bool(smtp.get("useSsl")),
                       str(smtp.get("preset", "") or ""))
            removed.append("user in smtp.json")
    if removed:
        log_change(user, "Cleared stored SMTP credentials (" + ", ".join(removed) + ")")
        set_notice("ok", "Stored credentials cleared.",
                   "Removed: " + ", ".join(removed) +
                   "\nServiceMonitor now connects to the relay anonymously.")
    else:
        set_notice("warn", "Nothing to clear - no credential files were present.")
    return RedirectResponse(safe_back(back), status_code=302)


@app.post("/settings/test-email")
async def settings_test_email(request: Request, back: str = Form("/settings")):
    user = get_user(request) or "unknown"
    if not MONITOR_PS1.exists():
        set_notice("err", f"{MONITOR_PS1} not found.")
        return RedirectResponse(safe_back(back), status_code=302)
    try:
        r = run_test_email()
    except subprocess.TimeoutExpired:
        log_change(user, "Ran test email - timed out")
        set_notice("err", "ServiceMonitor.ps1 -TestEmail timed out after 180 s.")
        return RedirectResponse(safe_back(back), status_code=302)
    except Exception as exc:
        set_notice("err", "Could not run ServiceMonitor.ps1.", str(exc))
        return RedirectResponse(safe_back(back), status_code=302)

    out = _tail((r.stdout or "") + "\n" + (r.stderr or ""), 16)
    if r.returncode == 0:
        log_change(user, "Ran test email - relay accepted it")
        set_notice("ok", "The relay accepted the test message.", out)
    else:
        log_change(user, f"Ran test email - FAILED (exit {r.returncode})")
        set_notice("err", f"Test email FAILED (exit {r.returncode}).",
                   out or "No output.")
    return RedirectResponse(safe_back(back), status_code=302)


@app.post("/set-user")
async def set_user(username: str = Form(...), back: str = Form("/")):
    resp = RedirectResponse(safe_back(back), status_code=302)
    resp.set_cookie("sm_user", username.strip()[:40], max_age=2_592_000)
    return resp


@app.post("/services/add")
async def service_add(request: Request, name: str = Form(...), alerts: str = Form("both"), back: str = Form("/")):
    user = get_user(request) or "unknown"
    alerts = alerts if alerts in ("both", "dev", "support", "none") else "both"
    with _lock:
        cfg = read_config()
        names = {s["name"] for s in cfg["services"]}
        if name not in names:
            cfg["services"].append({"name": name, "paused": False, "alerts": alerts})
            write_config(cfg)
            log_change(user, f"Added {name!r}  alerts={alerts}")
    return RedirectResponse(safe_back(back), status_code=302)


@app.post("/services/remove")
async def service_remove(request: Request, name: str = Form(...), back: str = Form("/")):
    user = get_user(request) or "unknown"
    with _lock:
        cfg = read_config()
        before = len(cfg["services"])
        cfg["services"] = [s for s in cfg["services"] if s["name"] != name]
        if len(cfg["services"]) < before:
            write_config(cfg)
            log_change(user, f"Removed {name!r}")
    return RedirectResponse(safe_back(back), status_code=302)


@app.post("/services/toggle-pause")
async def service_toggle_pause(request: Request, name: str = Form(...), back: str = Form("/")):
    user = get_user(request) or "unknown"
    with _lock:
        cfg = read_config()
        for svc in cfg["services"]:
            if svc["name"] == name:
                svc["paused"] = not svc.get("paused", False)
                action = "paused" if svc["paused"] else "resumed"
                write_config(cfg)
                log_change(user, f"{action.capitalize()} {name!r}")
                break
    return RedirectResponse(safe_back(back), status_code=302)


@app.post("/services/set-alerts")
async def service_set_alerts(request: Request, name: str = Form(...), alerts: str = Form(...), back: str = Form("/")):
    user = get_user(request) or "unknown"
    if alerts not in ("both", "dev", "support", "none"):
        return RedirectResponse(safe_back(back), status_code=302)
    with _lock:
        cfg = read_config()
        for svc in cfg["services"]:
            if svc["name"] == name:
                old = svc.get("alerts", "both")
                if old != alerts:
                    svc["alerts"] = alerts
                    write_config(cfg)
                    log_change(user, f"Changed {name!r} alerts: {old} → {alerts}")
                break
    return RedirectResponse(safe_back(back), status_code=302)


@app.post("/recipients/add")
async def recipient_add(request: Request, group: str = Form(...), email: str = Form(...), back: str = Form("/")):
    user = get_user(request) or "unknown"
    email = email.strip().lower()
    if "@" not in email or group not in ("dev", "support"):
        return RedirectResponse(safe_back(back), status_code=302)
    with _lock:
        cfg = read_config()
        rcpts = cfg["groups"][group]["recipients"]
        if email not in rcpts:
            rcpts.append(email)
            write_config(cfg)
            log_change(user, f"Added recipient {email!r} → {group}")
    return RedirectResponse(safe_back(back), status_code=302)


@app.post("/recipients/remove")
async def recipient_remove(request: Request, group: str = Form(...), email: str = Form(...), back: str = Form("/")):
    user = get_user(request) or "unknown"
    if group not in ("dev", "support"):
        return RedirectResponse(safe_back(back), status_code=302)
    with _lock:
        cfg = read_config()
        before = cfg["groups"][group]["recipients"][:]
        cfg["groups"][group]["recipients"] = [r for r in before if r != email]
        if cfg["groups"][group]["recipients"] != before:
            write_config(cfg)
            log_change(user, f"Removed recipient {email!r} from {group}")
    return RedirectResponse(safe_back(back), status_code=302)
