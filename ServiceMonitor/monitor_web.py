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
import json
import os
import shutil
import socket
import subprocess
import threading
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
        "iver": {"label": "Iver Support",  "recipients": []},
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
        return json.load(f)


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
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:system-ui,-apple-system,sans-serif;background:#f1f5f9;color:#1e293b;font-size:14px}
a{color:#2563eb}
header{background:#1e3a5f;color:#fff;padding:12px 24px;display:flex;align-items:center;gap:16px}
header h1{font-size:18px;font-weight:700;flex:1}
.user-form{display:flex;align-items:center;gap:6px;font-size:13px}
.user-form input{padding:4px 8px;border-radius:4px;border:none;font-size:13px;width:140px}
.user-form button{padding:4px 10px;background:#2563eb;color:#fff;border:none;border-radius:4px;cursor:pointer}
main{display:grid;grid-template-columns:1fr 1fr;grid-template-rows:auto auto auto;gap:16px;padding:16px;max-width:1400px;margin:0 auto}
.card{background:#fff;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden}
.card-title{background:#f8fafc;border-bottom:1px solid #e2e8f0;padding:10px 14px;font-weight:600;font-size:13px;color:#475569;text-transform:uppercase;letter-spacing:.04em}
.card-body{padding:14px}
/* spans: monitored (left), add-panel (right), groups and log full width */
.span-full{grid-column:1/-1}
table{width:100%;border-collapse:collapse}
th{text-align:left;font-size:12px;color:#64748b;font-weight:600;padding:6px 8px;border-bottom:1px solid #e2e8f0;text-transform:uppercase}
td{padding:7px 8px;border-bottom:1px solid #f1f5f9;vertical-align:middle}
tr:last-child td{border-bottom:none}
tr:hover td{background:#f8fafc}
.badge{display:inline-block;padding:2px 8px;border-radius:999px;font-size:12px;font-weight:600}
.badge-running{background:#dcfce7;color:#16a34a}
.badge-stopped{background:#fee2e2;color:#dc2626}
.badge-paused{background:#f1f5f9;color:#64748b}
.badge-unknown{background:#fef3c7;color:#d97706}
select{padding:3px 6px;border:1px solid #cbd5e1;border-radius:4px;font-size:13px;cursor:pointer;background:#fff}
select:hover{border-color:#2563eb}
.btn{display:inline-block;padding:4px 10px;border:1px solid #cbd5e1;border-radius:4px;background:#fff;color:#475569;font-size:12px;cursor:pointer;text-decoration:none;white-space:nowrap}
.btn:hover{background:#f1f5f9;border-color:#94a3b8}
.btn-danger{border-color:#fca5a5;color:#dc2626}
.btn-danger:hover{background:#fee2e2}
.btn-primary{background:#2563eb;border-color:#2563eb;color:#fff}
.btn-primary:hover{background:#1d4ed8}
.btn-small{padding:2px 8px;font-size:11px}
.svc-name{font-weight:500;font-family:monospace;font-size:13px}
.actions{display:flex;gap:4px;align-items:center}
/* available services grid */
.svc-grid{display:flex;flex-wrap:wrap;gap:6px;margin-top:10px}
.svc-pill{display:flex;align-items:center;background:#f1f5f9;border:1px solid #e2e8f0;border-radius:6px;overflow:hidden;font-size:12px}
.svc-pill span{padding:4px 8px;font-family:monospace}
.svc-pill button{padding:4px 8px;background:#2563eb;color:#fff;border:none;cursor:pointer;font-size:11px}
.svc-pill button:hover{background:#1d4ed8}
.filter-bar{display:flex;align-items:center;gap:10px}
.filter-bar input{flex:1;padding:6px 10px;border:1px solid #cbd5e1;border-radius:4px;font-size:13px}
.filter-bar select{padding:5px 8px}
/* mail groups */
.groups-grid{display:grid;grid-template-columns:1fr 1fr;gap:0}
.group-col{padding:14px}
.group-col:first-child{border-right:1px solid #e2e8f0}
.group-label{font-weight:600;font-size:13px;margin-bottom:10px;color:#1e3a5f}
.recipient-row{display:flex;align-items:center;justify-content:space-between;padding:5px 0;border-bottom:1px solid #f1f5f9}
.recipient-row:last-of-type{border-bottom:none}
.recipient-email{font-size:13px;color:#334155}
.add-recipient{display:flex;gap:6px;margin-top:10px}
.add-recipient input{flex:1;padding:5px 8px;border:1px solid #cbd5e1;border-radius:4px;font-size:13px}
/* changes log */
.changes-table{width:100%;font-size:12px;font-family:monospace}
.changes-table td{padding:3px 8px;border-bottom:1px solid #f1f5f9;color:#475569}
.changes-table td:first-child{color:#94a3b8;white-space:nowrap}
.changes-table td:nth-child(2){color:#2563eb;white-space:nowrap}
.empty{color:#94a3b8;font-style:italic;font-size:13px;padding:8px 0}
.no-user-notice{background:#fef9c3;border:1px solid #fde68a;color:#92400e;padding:8px 14px;font-size:13px;border-radius:6px;margin-bottom:16px}
"""


def status_badge(status: str) -> str:
    s = (status or "").lower()
    if s == "running":
        return '<span class="badge badge-running">● Running</span>'
    if s == "stopped":
        return '<span class="badge badge-stopped">✗ Stopped</span>'
    if s in ("paused", "pause_pending"):
        return '<span class="badge badge-paused">⏸ Paused</span>'
    return f'<span class="badge badge-unknown">? {status or "Unknown"}</span>'


def alerts_select(svc_name: str, current: str) -> str:
    opts = [
        ("both", "Both groups"),
        ("dev",  "Dev only"),
        ("iver", "Iver only"),
        ("none", "None"),
    ]
    inner = "".join(
        f'<option value="{v}"{" selected" if v == current else ""}>{label}</option>'
        for v, label in opts
    )
    return (
        f'<form method="post" action="/services/set-alerts" style="display:inline">'
        f'<input type="hidden" name="name" value="{svc_name}">'
        f'<select name="alerts" onchange="this.form.submit()" title="Alert routing">{inner}</select>'
        f'</form>'
    )


def build_monitored_section(services: list, statuses: dict) -> str:
    if not services:
        return '<p class="empty">No services monitored yet. Add some from the panel on the right.</p>'

    rows = ""
    for svc in services:
        name    = svc.get("name", "")
        paused  = svc.get("paused", False)
        alerts  = svc.get("alerts", "both")
        if paused:
            st_html = '<span class="badge badge-paused">⏸ Paused</span>'
            pause_btn = (
                f'<form method="post" action="/services/toggle-pause" style="display:inline">'
                f'<input type="hidden" name="name" value="{name}">'
                f'<button class="btn btn-small" title="Resume">▶</button></form>'
            )
        else:
            raw_status = statuses.get(name, "")
            st_html    = status_badge(raw_status)
            pause_btn  = (
                f'<form method="post" action="/services/toggle-pause" style="display:inline">'
                f'<input type="hidden" name="name" value="{name}">'
                f'<button class="btn btn-small" title="Pause monitoring">⏸</button></form>'
            )
        remove_btn = (
            f'<form method="post" action="/services/remove" style="display:inline">'
            f'<input type="hidden" name="name" value="{name}">'
            f'<button class="btn btn-small btn-danger" title="Remove from monitoring" '
            f'onclick="return confirm(\'Remove {name} from monitoring?\')">✕</button></form>'
        )
        rows += (
            f"<tr><td class='svc-name'>{name}</td>"
            f"<td>{st_html}</td>"
            f"<td>{alerts_select(name, alerts)}</td>"
            f"<td><div class='actions'>{pause_btn}{remove_btn}</div></td></tr>"
        )

    return (
        "<table>"
        "<thead><tr>"
        "<th>Service</th><th>Status</th><th>Alert routing</th><th>Actions</th>"
        "</tr></thead>"
        f"<tbody>{rows}</tbody>"
        "</table>"
    )


def build_available_section(monitored_names: set, statuses: dict) -> str:
    available = sorted(
        name for name in statuses
        if name not in monitored_names and name not in SYSTEM_SERVICES
    )
    pills = ""
    for name in available:
        pills += (
            f'<form method="post" action="/services/add" style="display:inline">'
            f'<input type="hidden" name="name" value="{name}">'
            f'<input type="hidden" name="alerts" id="default-alerts-{name}">'
            f'<div class="svc-pill"><span>{name}</span>'
            f'<button type="submit" onclick="'
            f"document.getElementById('default-alerts-{name}').value="
            f"document.getElementById('default-alerts').value"
            f'" title="Add {name} to monitoring">+</button>'
            f'</div></form>'
        )
    if not pills:
        pills = '<p class="empty">No additional services found (or PowerShell unavailable).</p>'

    return (
        '<div class="filter-bar">'
        '<input type="text" id="svc-filter" placeholder="Filter services…" '
        'oninput="filterSvcs(this.value)">'
        '<select id="default-alerts" title="Alert group for newly added services">'
        '<option value="both">Alerts: Both</option>'
        '<option value="dev">Alerts: Dev only</option>'
        '<option value="iver">Alerts: Iver only</option>'
        '<option value="none">Alerts: None</option>'
        '</select>'
        '</div>'
        f'<div class="svc-grid" id="svc-grid">{pills}</div>'
    )


def build_groups_section(groups: dict) -> str:
    cols = ""
    for gid, gdata in groups.items():
        label      = gdata.get("label", gid)
        recipients = gdata.get("recipients", [])
        rows = ""
        for email in recipients:
            rows += (
                f'<div class="recipient-row">'
                f'<span class="recipient-email">{email}</span>'
                f'<form method="post" action="/recipients/remove" style="display:inline">'
                f'<input type="hidden" name="group" value="{gid}">'
                f'<input type="hidden" name="email" value="{email}">'
                f'<button class="btn btn-small btn-danger" '
                f'onclick="return confirm(\'Remove {email}?\')">✕</button>'
                f'</form></div>'
            )
        if not rows:
            rows = '<p class="empty">No recipients in this group.</p>'
        add_form = (
            f'<div class="add-recipient">'
            f'<form method="post" action="/recipients/add" style="display:flex;gap:6px;flex:1">'
            f'<input type="hidden" name="group" value="{gid}">'
            f'<input type="email" name="email" placeholder="email@example.com" required>'
            f'<button class="btn btn-primary btn-small">Add</button>'
            f'</form></div>'
        )
        cols += (
            f'<div class="group-col">'
            f'<div class="group-label">{label}</div>'
            f'{rows}{add_form}'
            f'</div>'
        )
    return f'<div class="groups-grid">{cols}</div>'


def build_changes_section(lines: list[str]) -> str:
    if not lines:
        return '<p class="empty">No changes recorded yet.</p>'
    rows = ""
    for line in lines:
        # Log format: "2026-06-26 10:30  username              action"
        # Timestamp has a space → need 3 splits to get [date, time, user, action]
        parts = line.split(None, 3)
        if len(parts) == 4:
            ts = f"{parts[0]} {parts[1]}"
            user, action = parts[2], parts[3]
            rows += f"<tr><td>{ts}</td><td>{user}</td><td>{action}</td></tr>"
        else:
            rows += f"<tr><td colspan='3'>{line}</td></tr>"
    return f'<table class="changes-table"><tbody>{rows}</tbody></table>'


def build_page(cfg: dict, statuses: dict, user: str) -> str:
    services      = cfg.get("services", [])
    groups        = cfg.get("groups", {})
    monitored_set = {s["name"] for s in services}
    changes       = recent_changes()

    user_notice = ""
    if not user:
        user_notice = (
            '<div class="no-user-notice">'
            'Set your name above so changes are attributed in the audit log.'
            '</div>'
        )

    user_input = (
        '<form class="user-form" method="post" action="/set-user">'
        f'<label>You are:</label>'
        f'<input name="username" value="{user}" placeholder="Your name" maxlength="40">'
        '<button>Set</button>'
        '</form>'
    )

    js = """
<script>
function filterSvcs(q){
  q=q.toLowerCase();
  document.querySelectorAll('#svc-grid .svc-pill').forEach(function(el){
    el.parentElement.style.display=el.querySelector('span').textContent.toLowerCase().includes(q)?'':'none';
  });
}
</script>
"""
    return (
        "<!DOCTYPE html>"
        '<html lang="en">'
        '<head><meta charset="UTF-8">'
        '<meta name="viewport" content="width=device-width,initial-scale=1">'
        '<title>ServiceMonitor Admin</title>'
        f"<style>{CSS}</style>"
        "</head><body>"
        f"<header><h1>ServiceMonitor Admin</h1>{user_input}</header>"
        f"<main>{user_notice}"
        # Row 1: monitored (left) + available (right)
        '<div class="card"><div class="card-title">Monitored Services</div>'
        f'<div class="card-body">{build_monitored_section(services, statuses)}</div></div>'
        '<div class="card"><div class="card-title">Add Service</div>'
        f'<div class="card-body">{build_available_section(monitored_set, statuses)}</div></div>'
        # Row 2: mail groups (full width)
        '<div class="card span-full"><div class="card-title">Mail Groups</div>'
        f'{build_groups_section(groups)}</div>'
        # Row 3: recent changes (full width)
        '<div class="card span-full"><div class="card-title">Recent Changes</div>'
        f'<div class="card-body">{build_changes_section(changes)}</div></div>'
        f"</main>{js}</body></html>"
    )


# ---------------------------------------------------------------------------
# Optional access-token gate (defense-in-depth on shared RDP hosts)
# ---------------------------------------------------------------------------

def _login_page(message: str = "") -> str:
    return (
        "<!DOCTYPE html><html lang=\"en\"><head><meta charset=\"UTF-8\">"
        '<meta name="viewport" content="width=device-width,initial-scale=1">'
        "<title>ServiceMonitor - Sign in</title>"
        f"<style>{CSS}</style></head><body>"
        '<div style="max-width:360px;margin:80px auto;background:#fff;'
        'border:1px solid #e2e8f0;border-radius:8px;padding:24px">'
        '<h1 style="font-size:18px;margin-bottom:16px;color:#1e3a5f">ServiceMonitor Admin</h1>'
        f"{message}"
        '<form method="post" action="/auth">'
        '<input type="password" name="token" placeholder="Access token" autofocus '
        'style="width:100%;padding:8px;border:1px solid #cbd5e1;border-radius:4px;margin-bottom:12px">'
        '<button class="btn btn-primary" style="width:100%">Sign in</button>'
        "</form></div></body></html>"
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


@app.post("/set-user")
async def set_user(username: str = Form(...)):
    resp = RedirectResponse("/", status_code=302)
    resp.set_cookie("sm_user", username.strip()[:40], max_age=2_592_000)
    return resp


@app.post("/services/add")
async def service_add(request: Request, name: str = Form(...), alerts: str = Form("both")):
    user = get_user(request) or "unknown"
    alerts = alerts if alerts in ("both", "dev", "iver", "none") else "both"
    with _lock:
        cfg = read_config()
        names = {s["name"] for s in cfg["services"]}
        if name not in names:
            cfg["services"].append({"name": name, "paused": False, "alerts": alerts})
            write_config(cfg)
            log_change(user, f"Added {name!r}  alerts={alerts}")
    return RedirectResponse("/", status_code=302)


@app.post("/services/remove")
async def service_remove(request: Request, name: str = Form(...)):
    user = get_user(request) or "unknown"
    with _lock:
        cfg = read_config()
        before = len(cfg["services"])
        cfg["services"] = [s for s in cfg["services"] if s["name"] != name]
        if len(cfg["services"]) < before:
            write_config(cfg)
            log_change(user, f"Removed {name!r}")
    return RedirectResponse("/", status_code=302)


@app.post("/services/toggle-pause")
async def service_toggle_pause(request: Request, name: str = Form(...)):
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
    return RedirectResponse("/", status_code=302)


@app.post("/services/set-alerts")
async def service_set_alerts(request: Request, name: str = Form(...), alerts: str = Form(...)):
    user = get_user(request) or "unknown"
    if alerts not in ("both", "dev", "iver", "none"):
        return RedirectResponse("/", status_code=302)
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
    return RedirectResponse("/", status_code=302)


@app.post("/recipients/add")
async def recipient_add(request: Request, group: str = Form(...), email: str = Form(...)):
    user = get_user(request) or "unknown"
    email = email.strip().lower()
    if "@" not in email or group not in ("dev", "iver"):
        return RedirectResponse("/", status_code=302)
    with _lock:
        cfg = read_config()
        rcpts = cfg["groups"][group]["recipients"]
        if email not in rcpts:
            rcpts.append(email)
            write_config(cfg)
            log_change(user, f"Added recipient {email!r} → {group}")
    return RedirectResponse("/", status_code=302)


@app.post("/recipients/remove")
async def recipient_remove(request: Request, group: str = Form(...), email: str = Form(...)):
    user = get_user(request) or "unknown"
    if group not in ("dev", "iver"):
        return RedirectResponse("/", status_code=302)
    with _lock:
        cfg = read_config()
        before = cfg["groups"][group]["recipients"][:]
        cfg["groups"][group]["recipients"] = [r for r in before if r != email]
        if cfg["groups"][group]["recipients"] != before:
            write_config(cfg)
            log_change(user, f"Removed recipient {email!r} from {group}")
    return RedirectResponse("/", status_code=302)
