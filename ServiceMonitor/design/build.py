#!/usr/bin/env python3
"""Emit the ServiceMonitor Adwaita-dark artboards as .dc.html files.

Palette + metrics are libadwaita dark named colors, not invented values:
  window_bg #242424 | headerbar/sidebar #303030 | card rgba(255,255,255,.08)
  accent_bg #3584e4 | success #33d17a/#78e9ab | error #f66151/#ff7b63
  warning #f6d32d/#ffbe6f | dim-label = 55% white
  card radius 12px | button/entry radius 6px | header bar 47px
"""
from pathlib import Path

OUT = Path(__file__).parent

BG      = "#242424"
CHROME  = "#303030"
CARD    = "rgba(255,255,255,0.08)"
CARDBRD = "rgba(255,255,255,0.06)"
SEP     = "rgba(255,255,255,0.08)"
FG      = "#ffffff"
DIM     = "rgba(255,255,255,0.55)"
DIM2    = "rgba(255,255,255,0.40)"
BTN     = "rgba(255,255,255,0.10)"
SHADE   = "rgba(0,0,0,0.40)"
OK_D, OK_T   = "#33d17a", "#78e9ab"
ERR_D, ERR_T = "#f66151", "#ff7b63"
WRN_D, WRN_T = "#f6d32d", "#ffbe6f"
ACCENT_T = "#78aeed"

SANS = "Cantarell, 'Segoe UI', system-ui, -apple-system, sans-serif"
MONO = "'Source Code Pro', 'Cascadia Mono', Consolas, ui-monospace, monospace"

FONTS = ('<link rel="stylesheet" href="https://fonts.googleapis.com/css2?'
         'family=Cantarell:wght@400;700&family=Source+Code+Pro:wght@400;500&display=swap">')


def svg(paths, extra=""):
    return ('<svg width="16" height="16" viewBox="0 0 16 16" fill="none" '
            'stroke="currentColor" stroke-width="1.5" stroke-linecap="round" '
            'stroke-linejoin="round" style="flex-shrink:0' + extra + '">' + paths + '</svg>')


I_RACK  = svg('<rect x="2.25" y="2.75" width="11.5" height="4.5" rx="1.25"></rect>'
              '<rect x="2.25" y="8.75" width="11.5" height="4.5" rx="1.25"></rect>'
              '<path d="M4.75 5h.01M4.75 11h.01"></path>')
I_MAIL  = svg('<rect x="2" y="3.75" width="12" height="8.5" rx="1.25"></rect>'
              '<path d="m2.75 4.75 5.25 3.9 5.25-3.9"></path>')
I_CLOCK = svg('<circle cx="8" cy="8" r="5.75"></circle><path d="M8 4.6V8l2.4 1.5"></path>')
I_FIND  = svg('<circle cx="7" cy="7" r="4.5"></circle><path d="m10.4 10.4 3.1 3.1"></path>')
I_PAUSE = svg('<path d="M6 3.75v8.5M10 3.75v8.5"></path>')
I_PLAY  = svg('<path d="M5.5 3.6 12.2 8l-6.7 4.4z"></path>')
I_X     = svg('<path d="m4.25 4.25 7.5 7.5M11.75 4.25l-7.5 7.5"></path>')
I_PLUS  = svg('<path d="M8 3.5v9M3.5 8h9"></path>')
I_CHEV  = svg('<path d="m4.75 6.5 3.25 3.25L11.25 6.5"></path>')
I_LOCK  = svg('<rect x="3.25" y="7" width="9.5" height="6.25" rx="1.5"></rect>'
              '<path d="M5.75 7V5.25a2.25 2.25 0 0 1 4.5 0V7"></path>')


def head(preview_w, preview_h, props_extra="", logic="class Component extends DCLogic {}"):
    props = '{"$preview":{"width":%d,"height":%d}%s}' % (preview_w, preview_h, props_extra)
    return props, logic


def page(body, preview_w, preview_h, props='', logic='class Component extends DCLogic {}'):
    if not props:
        props = '{"$preview":{"width":%d,"height":%d}}' % (preview_w, preview_h)
    return (
        '<!doctype html>\n<html>\n<head>\n  <meta charset="utf-8">\n'
        '  <script src="./support.js"></script>\n</head>\n<body>\n'
        '<x-dc>\n<helmet>\n  ' + FONTS + '\n  <style>\n'
        '    body { margin: 0; font-family: ' + SANS + '; }\n'
        '    a { color: ' + ACCENT_T + '; text-decoration: none; }\n'
        '    a:hover { color: #a5c8f2; }\n'
        '    .nav:hover { background: rgba(255,255,255,0.06); }\n'
        '    .row:hover { background: rgba(255,255,255,0.03); }\n'
        '    .btn:hover { background: rgba(255,255,255,0.16); }\n'
        '    .btn-flat:hover { background: rgba(255,255,255,0.10); }\n'
        '    .btn-danger:hover { background: rgba(224,27,36,0.20); }\n'
        '  </style>\n</helmet>\n' + body + '\n</x-dc>\n'
        "<script data-dc-script data-props='" + props + "'>\n" + logic + '\n</script>\n'
        '</body>\n</html>\n')


# --------------------------------------------------------------------------
# shared chrome
# --------------------------------------------------------------------------

def nav_row(icon, label, count, active=False):
    bg = "background:rgba(255,255,255,0.10);" if active else ""
    weight = "700" if active else "400"
    color = FG if active else "rgba(255,255,255,0.82)"
    cnt = ''
    if count:
        cnt = ('<span style="margin-left:auto;font-size:13px;color:' + DIM2 + ';'
               'font-variant-numeric:tabular-nums">' + count + '</span>')
    return ('<div class="nav" style="display:flex;align-items:center;gap:12px;height:38px;'
            'padding:0 10px;border-radius:6px;cursor:default;' + bg + 'color:' + color + '">'
            + icon + '<span style="font-size:14px;font-weight:' + weight + '">' + label
            + '</span>' + cnt + '</div>')


def sidebar(active, height):
    rows = (
        nav_row(I_RACK, "Services", "8", active == "services")
        + nav_row(I_MAIL, "Mail groups", "2", active == "groups")
        + nav_row(I_CLOCK, "Recent changes", "", active == "log")
    )
    return (
        '<aside style="width:260px;flex-shrink:0;background:' + CHROME + ';display:flex;'
        'flex-direction:column;border-right:1px solid ' + SHADE + '">'
        # sidebar header bar
        '<div style="height:47px;flex-shrink:0;display:flex;align-items:center;gap:10px;'
        'padding:0 12px 0 14px;border-bottom:1px solid ' + SHADE + '">'
        '<div style="width:22px;height:22px;border-radius:6px;background:{{accent}};color:#ffffff;'
        'display:flex;align-items:center;justify-content:center;flex-shrink:0">' + I_RACK + '</div>'
        '<span style="font-size:15px;font-weight:700;color:' + FG + '">ServiceMonitor</span>'
        '<span style="margin-left:auto;font-size:12px;color:' + DIM2 + '">v1.3.0</span>'
        '</div>'
        # nav
        '<nav style="display:flex;flex-direction:column;gap:2px;padding:10px 6px">' + rows + '</nav>'
        # footer: scheduled task state
        '<div style="margin-top:auto;padding:14px;border-top:1px solid ' + SEP + ';'
        'display:flex;flex-direction:column;gap:6px">'
        '<div style="display:flex;align-items:center;gap:8px">'
        '<span style="width:8px;height:8px;border-radius:50%;background:' + OK_D + '"></span>'
        '<span style="font-size:13px;color:' + FG + ';font-weight:700">Scheduled task active</span>'
        '</div>'
        '<span style="font-size:12px;color:' + DIM + ';font-family:' + MONO + '">'
        'ServiceMonitor &middot; every 5 min &middot; SYSTEM</span>'
        '</div>'
        '</aside>')


def user_control():
    return (
        '<div style="display:flex;align-items:center;gap:8px">'
        '<span style="font-size:13px;color:' + DIM + '">You are</span>'
        '<div style="display:flex;align-items:center;height:30px;padding:0 10px;border-radius:6px;'
        'background:rgba(0,0,0,0.28);border:1px solid ' + SEP + '">'
        '<span style="font-size:13px;color:' + FG + '">admin</span></div>'
        '<div class="btn" style="display:flex;align-items:center;height:30px;padding:0 12px;'
        'border-radius:6px;background:' + BTN + ';font-size:13px;color:' + FG + '">Set</div>'
        '</div>')


def shell(active, title, body, height=880):
    return (
        '<div style="display:flex;width:100%;height:' + str(height) + 'px;background:' + BG + ';'
        'color:' + FG + ';font-family:' + SANS + ';font-size:15px;overflow:hidden">'
        + sidebar(active, height) +
        '<div style="flex:1;display:flex;flex-direction:column;min-width:0">'
        '<header style="height:47px;flex-shrink:0;display:flex;align-items:center;'
        'padding:0 16px;gap:16px;background:' + CHROME + ';border-bottom:1px solid ' + SHADE + '">'
        '<span style="font-size:15px;font-weight:700;color:' + FG + '">' + title + '</span>'
        '<div style="margin-left:auto">' + user_control() + '</div>'
        '</header>'
        '<main style="flex:1;padding:24px;overflow:hidden">' + body + '</main>'
        '</div></div>')


def group_title(title, desc):
    return ('<div style="margin:0 0 12px">'
            '<h2 style="margin:0 0 2px;font-size:16px;font-weight:700;color:' + FG + '">'
            + title + '</h2>'
            '<p style="margin:0;font-size:13px;color:' + DIM + '">' + desc + '</p></div>')


def card(inner, extra=""):
    return ('<div style="background:' + CARD + ';border:1px solid ' + CARDBRD + ';'
            'border-radius:12px;overflow:hidden;' + extra + '">' + inner + '</div>')


def combo(label, width=112):
    return ('<div class="btn" style="display:flex;align-items:center;gap:6px;height:34px;'
            'width:' + str(width) + 'px;padding:0 8px 0 12px;border-radius:6px;background:' + BTN + ';'
            'color:' + FG + ';font-size:14px;flex-shrink:0">'
            '<span style="flex:1">' + label + '</span>'
            '<span style="color:' + DIM + '">' + I_CHEV + '</span></div>')


def icon_btn(icon, danger=False):
    color = ERR_T if danger else "rgba(255,255,255,0.82)"
    cls = "btn-danger" if danger else "btn-flat"
    return ('<div class="' + cls + '" style="width:34px;height:34px;border-radius:6px;'
            'display:flex;align-items:center;justify-content:center;flex-shrink:0;color:'
            + color + '">' + icon + '</div>')


ACCENT_OPTS = '"accent":{"editor":"color","default":"#3584e4","options":["#3584e4","#2190a4","#3a944a","#ed5b00"],"section":"Theme"}'
LOGIC = ("class Component extends DCLogic {\n"
         "  renderVals() {\n"
         "    return { accent: this.props.accent ?? '#3584e4' };\n"
         "  }\n"
         "}")


def props_for(w, h):
    return '{"$preview":{"width":%d,"height":%d},%s}' % (w, h, ACCENT_OPTS)


STATE = {
    "running": (OK_D, OK_T, "Running"),
    "stopped": (ERR_D, ERR_T, "Stopped"),
    "paused":  ("rgba(255,255,255,0.32)", DIM, "Paused"),
    "unknown": (WRN_D, WRN_T, "Unknown"),
}


def svc_row(name, state, alerts, last=False):
    dot, text, label = STATE[state]
    border = "" if last else "border-bottom:1px solid " + SEP + ";"
    resume = state == "paused"
    return ('<div class="row" style="display:flex;align-items:center;gap:14px;height:54px;'
            'padding:0 8px 0 16px;' + border + '">'
            '<span style="width:9px;height:9px;border-radius:50%;background:' + dot + ';'
            'flex-shrink:0"></span>'
            '<span style="flex:1;min-width:0;font-family:' + MONO + ';font-size:14px;'
            'color:' + FG + ';overflow:hidden;text-overflow:ellipsis;white-space:nowrap">'
            + name + '</span>'
            '<span style="width:76px;flex-shrink:0;font-size:13px;color:' + text + '">'
            + label + '</span>'
            + combo(alerts) +
            '<div style="display:flex;gap:2px;margin-left:8px">'
            + icon_btn(I_PLAY if resume else I_PAUSE)
            + icon_btn(I_X, danger=True) +
            '</div></div>')


def avail_row(name, last=False):
    border = "" if last else "border-bottom:1px solid " + SEP + ";"
    return ('<div class="row" style="display:flex;align-items:center;gap:12px;height:46px;'
            'padding:0 8px 0 16px;' + border + '">'
            '<span style="flex:1;min-width:0;font-family:' + MONO + ';font-size:14px;color:'
            + FG + ';overflow:hidden;text-overflow:ellipsis;white-space:nowrap">' + name + '</span>'
            '<div style="display:flex;align-items:center;justify-content:center;width:30px;'
            'height:30px;border-radius:6px;background:{{accent}};color:#ffffff;flex-shrink:0">'
            + I_PLUS + '</div></div>')


def entry(placeholder, icon="", width="100%", height=38):
    ico = ('<span style="color:' + DIM + ';display:flex">' + icon + '</span>') if icon else ""
    return ('<div style="display:flex;align-items:center;gap:10px;width:' + width + ';'
            'height:' + str(height) + 'px;padding:0 12px;border-radius:6px;'
            'background:rgba(0,0,0,0.28);border:1px solid ' + SEP + '">'
            + ico + '<span style="font-size:14px;color:' + DIM2 + '">' + placeholder + '</span></div>')


# --------------------------------------------------------------------------
# 1. Services (Main)
# --------------------------------------------------------------------------

MONITORED = [
    ("Spooler", "running", "Both"),
    ("W32Time", "running", "Both"),
    ("wuauserv", "stopped", "Dev only"),
    ("LanmanServer", "running", "Both"),
    ("LanmanWorkstation", "running", "Iver only"),
    ("Dnscache", "running", "Both"),
    ("BITS", "paused", "None"),
    ("WSearch", "paused", "None"),
]

AVAILABLE = ["MSSQLSERVER", "SQLSERVERAGENT", "VeeamBackupSvc", "nginx",
             "Redis", "PDQDeployService", "TeamViewer", "ZabbixAgent"]


def build_main():
    rows = "".join(svc_row(n, s, a, i == len(MONITORED) - 1)
                   for i, (n, s, a) in enumerate(MONITORED))
    left = (group_title("Monitored services",
                        "6 checked every 5 minutes &middot; 2 paused &middot; alerts route to Dev Team and Iver Support")
            + card(rows))

    arows = "".join(avail_row(n, i == len(AVAILABLE) - 1) for i, n in enumerate(AVAILABLE))
    right = (group_title("Add a service",
                         "Standard Windows services are hidden from this list.")
             + entry("Filter services…", I_FIND)
             + '<div style="display:flex;align-items:center;gap:10px;margin:12px 0">'
               '<span style="font-size:13px;color:' + DIM + ';flex:1">Alerts for newly added</span>'
             + combo("Both", 128) + '</div>'
             + card(arows)
             + '<p style="margin:10px 2px 0;font-size:12px;color:' + DIM2 + '">'
               '128 non-system services detected on this host.</p>')

    body = ('<div style="display:grid;grid-template-columns:680px 1fr;gap:24px;'
            'align-items:start">'
            '<div>' + left + '</div><div>' + right + '</div></div>')
    return page(shell("services", "Services", body), 1440, 880,
                props_for(1440, 880), LOGIC)


# --------------------------------------------------------------------------
# 2. Mail groups
# --------------------------------------------------------------------------

GROUPS = [
    ("Dev Team", "dev", ["dev.lead@example.com", "oncall@example.com", "dev@example.com"]),
    ("Iver Support", "iver", ["servicedesk@example.com", "support.lead@example.com"]),
]


def recipient_row(email):
    return ('<div class="row" style="display:flex;align-items:center;gap:12px;height:48px;'
            'padding:0 8px 0 16px;border-bottom:1px solid ' + SEP + '">'
            '<span style="flex:1;min-width:0;font-size:14px;color:' + FG + ';overflow:hidden;'
            'text-overflow:ellipsis;white-space:nowrap">' + email + '</span>'
            + icon_btn(I_X, danger=True) + '</div>')


def build_groups():
    cols = ""
    for label, gid, emails in GROUPS:
        rows = "".join(recipient_row(e) for e in emails)
        add = ('<div style="display:flex;align-items:center;gap:8px;padding:10px 10px 10px 16px">'
               + entry("email@example.com", "", "100%", 34) +
               '<div style="display:flex;align-items:center;height:34px;padding:0 14px;'
               'border-radius:6px;background:{{accent}};color:#ffffff;font-size:14px;'
               'font-weight:700;flex-shrink:0">Add</div></div>')
        cols += ('<div>' + group_title(label, str(len(emails)) + ' recipients &middot; group id <span style="font-family:'
                 + MONO + '">' + gid + '</span>')
                 + card(rows + add) + '</div>')
    body = ('<div style="display:grid;grid-template-columns:1fr 1fr;gap:24px;'
            'align-items:start;max-width:1000px">' + cols + '</div>')
    return page(shell("groups", "Mail groups", body), 1440, 880,
                props_for(1440, 880), LOGIC)


# --------------------------------------------------------------------------
# 3. Recent changes
# --------------------------------------------------------------------------

CHANGES = [
    ("2026-08-27 09:41", "admin", "Added 'VeeamBackupSvc'  alerts=both"),
    ("2026-08-27 09:38", "admin", "Changed 'wuauserv' alerts: both → dev"),
    ("2026-08-26 16:02", "linda", "Paused 'WSearch'"),
    ("2026-08-26 15:57", "linda", "Added recipient 'oncall@example.com' → dev"),
    ("2026-08-26 11:20", "admin", "Removed 'PrintNotify'"),
    ("2026-08-25 08:14", "admin", "Resumed 'BITS'"),
    ("2026-08-22 13:45", "admin", "Removed recipient 'dev@example.com' from iver"),
    ("2026-08-22 13:44", "admin", "Added recipient 'servicedesk@example.com' → iver"),
    ("2026-08-21 10:03", "admin", "Added 'Dnscache'  alerts=both"),
    ("2026-08-20 09:12", "admin", "Paused 'BITS'"),
]


def build_log():
    rows = ""
    for i, (ts, user, action) in enumerate(CHANGES):
        border = "" if i == len(CHANGES) - 1 else "border-bottom:1px solid " + SEP + ";"
        rows += ('<div class="row" style="display:flex;align-items:center;gap:20px;height:40px;'
                 'padding:0 16px;font-family:' + MONO + ';font-size:13px;' + border + '">'
                 '<span style="width:150px;flex-shrink:0;color:' + DIM2 + '">' + ts + '</span>'
                 '<span style="width:96px;flex-shrink:0;color:' + ACCENT_T + '">' + user + '</span>'
                 '<span style="flex:1;min-width:0;color:rgba(255,255,255,0.82);overflow:hidden;'
                 'text-overflow:ellipsis;white-space:nowrap">' + action + '</span></div>')
    body = ('<div style="max-width:900px">'
            + group_title("Recent changes",
                          "Appended to changes.log next to monitor-config.json. Last 20 entries.")
            + card(rows) + '</div>')
    return page(shell("log", "Recent changes", body), 1440, 880,
                props_for(1440, 880), LOGIC)


# --------------------------------------------------------------------------
# 4. Sign in
# --------------------------------------------------------------------------

def build_signin():
    body = ('<div style="width:100%;height:600px;background:' + BG + ';display:flex;'
            'align-items:center;justify-content:center;font-family:' + SANS + '">'
            '<div style="width:380px;background:#383838;border:1px solid rgba(255,255,255,0.06);'
            'border-radius:12px;padding:28px 24px 24px;box-shadow:0 12px 32px rgba(0,0,0,0.45);'
            'display:flex;flex-direction:column;align-items:stretch;gap:0">'
            '<div style="width:44px;height:44px;border-radius:12px;background:{{accent}};'
            'color:#ffffff;display:flex;align-items:center;justify-content:center;'
            'align-self:center;margin-bottom:16px">'
            '<svg width="22" height="22" viewBox="0 0 16 16" fill="none" stroke="currentColor" '
            'stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round">'
            '<rect x="3.25" y="7" width="9.5" height="6.25" rx="1.5"></rect>'
            '<path d="M5.75 7V5.25a2.25 2.25 0 0 1 4.5 0V7"></path></svg></div>'
            '<h1 style="margin:0 0 4px;font-size:19px;font-weight:700;color:' + FG + ';'
            'text-align:center">ServiceMonitor</h1>'
            '<p style="margin:0 0 20px;font-size:14px;color:' + DIM + ';text-align:center">'
            'Enter the access token to continue.</p>'
            '<div style="display:flex;align-items:center;height:38px;padding:0 12px;'
            'border-radius:6px;background:rgba(0,0,0,0.32);border:1px solid {{accent}};'
            'box-shadow:0 0 0 2px rgba(53,132,228,0.28);margin-bottom:14px">'
            '<span style="font-size:15px;color:' + FG + ';letter-spacing:3px">'
            '••••••••••</span>'
            '<span style="width:1px;height:17px;background:' + FG + ';margin-left:2px"></span></div>'
            '<div style="display:flex;align-items:center;justify-content:center;height:38px;'
            'border-radius:6px;background:{{accent}};color:#ffffff;font-size:15px;'
            'font-weight:700">Sign in</div>'
            '<p style="margin:20px 0 0;font-size:12px;color:' + DIM2 + ';text-align:center;'
            'line-height:1.5">Bound to 127.0.0.1:8080. This gate only appears when '
            '<span style="font-family:' + MONO + '">SM_TOKEN</span> is set; otherwise the '
            'RDP login is the only layer.</p>'
            '</div></div>')
    return page(body, 900, 600, props_for(900, 600), LOGIC)


# --------------------------------------------------------------------------
# 5. Token sheet
# --------------------------------------------------------------------------

def swatch(name, value, note=""):
    sub = note or value
    return ('<div style="display:flex;align-items:center;gap:12px;min-width:0">'
            '<div style="width:44px;height:44px;border-radius:8px;background:' + value + ';'
            'border:1px solid rgba(255,255,255,0.12);flex-shrink:0"></div>'
            '<div style="display:flex;flex-direction:column;gap:3px;min-width:0">'
            '<span style="font-size:13px;color:' + FG + '">' + name + '</span>'
            '<span style="font-size:12px;color:' + DIM + ';font-family:' + MONO + ';'
            'overflow:hidden;text-overflow:ellipsis;white-space:nowrap">' + sub + '</span>'
            '</div></div>')


def sheet_section(title, inner, mt=32):
    return ('<div style="margin-top:' + str(mt) + 'px">'
            '<h2 style="margin:0 0 14px;font-size:13px;font-weight:700;color:' + DIM + ';'
            'text-transform:uppercase;letter-spacing:0.08em">' + title + '</h2>'
            + inner + '</div>')


def metric(label, value):
    return ('<div style="display:flex;align-items:baseline;gap:12px;height:30px;'
            'border-bottom:1px solid ' + SEP + '">'
            '<span style="flex:1;font-size:13px;color:rgba(255,255,255,0.82)">' + label + '</span>'
            '<span style="font-family:' + MONO + ';font-size:13px;color:' + ACCENT_T + '">'
            + value + '</span></div>')


def build_components():
    grid4 = 'display:grid;grid-template-columns:repeat(4, minmax(0, 1fr));gap:18px 20px'
    grid2 = 'display:grid;grid-template-columns:repeat(2, minmax(0, 1fr));gap:0 40px'

    surfaces = ('<div style="' + grid4 + '">'
                + swatch("Window", "#242424")
                + swatch("Header bar / sidebar", "#303030")
                + swatch("Card / button fill", "#353535", "rgba(255,255,255,.08)")
                + swatch("Dialog / popover", "#383838")
                + swatch("Accent", "{{accent}}", "#3584e4 &middot; accent_bg_color")
                + swatch("Accent on dark", ACCENT_T, ACCENT_T + " &middot; links, log user")
                + swatch("Separator", "#3a3a3a", "rgba(255,255,255,.08)")
                + swatch("Sidebar edge", "#161616", "rgba(0,0,0,.40)")
                + '</div>')

    states = ('<div style="' + grid4 + '">'
              + swatch("Running &mdash; dot", OK_D, OK_D)
              + swatch("Running &mdash; label", OK_T, OK_T)
              + swatch("Stopped &mdash; dot", ERR_D, ERR_D)
              + swatch("Stopped &mdash; label", ERR_T, ERR_T)
              + swatch("Unknown &mdash; dot", WRN_D, WRN_D)
              + swatch("Unknown &mdash; label", WRN_T, WRN_T)
              + swatch("Paused &mdash; dot", "#575757", "rgba(255,255,255,.32)")
              + swatch("Dim label", "#8a8a8a", "rgba(255,255,255,.55)")
              + '</div>')

    type_rows = ('<div style="display:flex;flex-direction:column;gap:14px">'
                 '<div style="display:flex;align-items:baseline;gap:20px">'
                 '<span style="width:210px;flex-shrink:0;font-family:' + MONO + ';font-size:12px;'
                 'color:' + DIM + '">Cantarell 15 / 700</span>'
                 '<span style="font-size:15px;font-weight:700;color:' + FG + '">'
                 'Window &amp; sidebar titles</span></div>'
                 '<div style="display:flex;align-items:baseline;gap:20px">'
                 '<span style="width:210px;flex-shrink:0;font-family:' + MONO + ';font-size:12px;'
                 'color:' + DIM + '">Cantarell 16 / 700</span>'
                 '<span style="font-size:16px;font-weight:700;color:' + FG + '">'
                 'Group heading</span></div>'
                 '<div style="display:flex;align-items:baseline;gap:20px">'
                 '<span style="width:210px;flex-shrink:0;font-family:' + MONO + ';font-size:12px;'
                 'color:' + DIM + '">Cantarell 14 / 400</span>'
                 '<span style="font-size:14px;color:' + FG + '">Row labels, buttons, entries</span>'
                 '</div>'
                 '<div style="display:flex;align-items:baseline;gap:20px">'
                 '<span style="width:210px;flex-shrink:0;font-family:' + MONO + ';font-size:12px;'
                 'color:' + DIM + '">Cantarell 13 / 400</span>'
                 '<span style="font-size:13px;color:' + DIM + '">Descriptions and dim labels</span>'
                 '</div>'
                 '<div style="display:flex;align-items:baseline;gap:20px">'
                 '<span style="width:210px;flex-shrink:0;font-family:' + MONO + ';font-size:12px;'
                 'color:' + DIM + '">Source Code Pro 14</span>'
                 '<span style="font-family:' + MONO + ';font-size:14px;color:' + FG + '">'
                 'LanmanWorkstation</span></div>'
                 '<div style="display:flex;align-items:baseline;gap:20px">'
                 '<span style="width:210px;flex-shrink:0;font-family:' + MONO + ';font-size:12px;'
                 'color:' + DIM + '">Source Code Pro 13</span>'
                 '<span style="font-family:' + MONO + ';font-size:13px;color:' + DIM2 + '">'
                 '2026-08-27 09:41</span></div>'
                 '</div>')

    controls = ('<div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap">'
                '<div class="btn" style="display:flex;align-items:center;height:34px;padding:0 16px;'
                'border-radius:6px;background:' + BTN + ';color:' + FG + ';font-size:14px">Flat</div>'
                '<div style="display:flex;align-items:center;height:34px;padding:0 16px;'
                'border-radius:6px;background:{{accent}};color:#ffffff;font-size:14px;'
                'font-weight:700">Suggested</div>'
                '<div style="display:flex;align-items:center;height:34px;padding:0 16px;'
                'border-radius:6px;background:#c01c28;color:#ffffff;font-size:14px;'
                'font-weight:700">Destructive</div>'
                + combo("Both groups", 140)
                + icon_btn(I_PAUSE) + icon_btn(I_X, danger=True)
                + entry("Filter services…", I_FIND, "240px")
                + '</div>'
                '<div style="display:flex;align-items:center;gap:26px;margin-top:20px;'
                'flex-wrap:wrap">'
                + "".join(
                    '<div style="display:flex;align-items:center;gap:9px">'
                    '<span style="width:9px;height:9px;border-radius:50%;background:' + d + '"></span>'
                    '<span style="font-size:13px;color:' + t + '">' + l + '</span></div>'
                    for d, t, l in [STATE["running"], STATE["stopped"],
                                    STATE["paused"], STATE["unknown"]])
                + '</div>')

    metrics = ('<div style="' + grid2 + '">'
               '<div>'
               + metric("Header bar height", "47px")
               + metric("Sidebar width", "260px")
               + metric("Nav row height", "38px")
               + metric("Service row height", "54px")
               + metric("Recipient row height", "48px")
               + metric("Log row height", "40px")
               + '</div><div>'
               + metric("Card radius", "12px")
               + metric("Button / entry radius", "6px")
               + metric("Button height", "34px")
               + metric("Entry height", "38px")
               + metric("Content padding", "24px")
               + metric("Column gap", "24px")
               + '</div></div>')

    body = ('<div style="width:100%;min-height:1320px;background:' + BG + ';color:' + FG + ';'
            'font-family:' + SANS + ';padding:32px 36px 40px;box-sizing:border-box">'
            '<h1 style="margin:0 0 4px;font-size:22px;font-weight:700">Design tokens</h1>'
            '<p style="margin:0;font-size:14px;color:' + DIM + ';max-width:640px">'
            'libadwaita dark, as shipped by GNOME. Drop these straight into the '
            '<span style="font-family:' + MONO + '">CSS</span> string in '
            '<span style="font-family:' + MONO + '">monitor_web.py</span>.</p>'
            + sheet_section("Surfaces", surfaces, 30)
            + sheet_section("Service state", states)
            + sheet_section("Type", type_rows)
            + sheet_section("Controls", controls)
            + sheet_section("Metrics", metrics)
            + '</div>')
    return page(body, 1120, 1320, props_for(1120, 1320), LOGIC)


# --------------------------------------------------------------------------

CANVAS = """{
  "artboards": [
    { "file": "Main.dc.html",       "x": 0,    "y": 0,    "w": 1440, "h": 880 },
    { "file": "Groups.dc.html",     "x": 1560, "y": 0,    "w": 1440, "h": 880 },
    { "file": "Log.dc.html",        "x": 3120, "y": 0,    "w": 1440, "h": 880 },
    { "file": "SignIn.dc.html",     "x": 0,    "y": 1060, "w": 900,  "h": 600 },
    { "file": "Components.dc.html", "x": 1080, "y": 1060, "w": 1120, "h": 1320 }
  ],
  "annotations": [
    { "id": "note-source", "x": 0, "y": -190, "w": 420,
      "text": "Palette and metrics are real libadwaita dark values, not invented ones.\\nWindow #242424 / chrome #303030 / cards 8% white / accent #3584e4.\\nCards 12px, buttons and entries 6px, header bar 47px." },
    { "id": "note-sample", "x": 1560, "y": -190, "w": 420,
      "text": "Service names come from your services.txt. Statuses, recipients\\nand log lines are sample data for the mockup - the log lines do\\nuse the real format that log_change() writes." },
    { "id": "note-nochrome", "x": 3120, "y": -190, "w": 420,
      "text": "No fake window buttons: this runs in a browser at 127.0.0.1:8080,\\nso a painted minimise/maximise/close strip would sit doubled up\\nunder the real browser chrome." },
    { "id": "note-port", "x": 2300, "y": 1060, "w": 300,
      "text": "This sheet is the porting target - every value here maps to a rule\\nin the CSS string in monitor_web.py." }
  ],
  "launch": { "view": "canvas" }
}
"""


def main():
    files = {
        "Main.dc.html": build_main(),
        "Groups.dc.html": build_groups(),
        "Log.dc.html": build_log(),
        "SignIn.dc.html": build_signin(),
        "Components.dc.html": build_components(),
        "canvas.json": CANVAS,
    }
    for name, text in files.items():
        (OUT / name).write_text(text, encoding="utf-8")
        print("wrote %-22s %6d bytes" % (name, len(text.encode("utf-8"))))


if __name__ == "__main__":
    main()
