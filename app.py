"""Sunrise Live Operation Tracker - Flask backend.

One workbook feeds the whole dashboard. Sheets it reads:
  Tasks, IPRTracker (the "IPR Tracker" sheet), Alerts, Complaints, Events, FollowUps, SafetyTips, Escalation

Procurement rule: an IPR whose "PO Number" is blank (and that is not cancelled) is shown as "PO awaited".

Where the workbook comes from (first one that is configured wins):
  1. Microsoft Graph   : TENANT_ID, CLIENT_ID, CLIENT_SECRET + GRAPH_DRIVE_ID/GRAPH_ITEM_ID (or GRAPH_USER/GRAPH_FILE_PATH)
  2. Local/synced file : EXCEL_PATH   (a * is allowed, newest match wins; optional EXCEL_IPR_PATH for a separate IPR file)
  3. Share link        : EXCEL_LINK   (OneDrive/SharePoint "Anyone with the link")
  4. Google Sheet      : GSHEET_ID    (shared "Anyone with the link: Viewer")
  5. Demo data         : nothing configured

The browser asks /api/data every 60 s. The server answers from a cache and refreshes in a background
thread, so a slow or locked file can never freeze the screen.
"""
import csv
import glob
import hashlib
import io
import json
import os
import re
import shutil
import smtplib
import subprocess
import tempfile
import threading
import time
from datetime import date, datetime, timedelta, timezone
from datetime import time as dtime
from email.message import EmailMessage

import requests
from flask import Flask, jsonify, render_template

from demo_data import get_demo_sheets

app = Flask(__name__)  # templates/ is resolved next to this file

REFRESH_SECONDS = 60
SHEETS = ["Tasks", "IPRTracker", "Alerts", "Complaints", "Events", "FollowUps", "SafetyTips", "Escalation"]
SHEET_ALIASES = {"IPRTracker": ["IPR Tracker", "IPR Tracker ", "IPR"]}   # real sheet name has a trailing space
HEADER_HINT = {"IPRTracker": "IPR Reference"}                              # header row is not always row 1
PRIORITY_ORDER = {"High": 0, "Medium": 1, "Low": 2}
FOLLOWUP_COLS = ["Meeting", "Date", "Point", "Owner", "Due Date", "Status"]


# =============================================================================
#  Reading the workbook
# =============================================================================
def now_local():
    """Current time in the dashboard's time zone (default IST = UTC+5:30; override with TZ_OFFSET_MINUTES)."""
    off = int(os.environ.get("TZ_OFFSET_MINUTES", 330))
    return datetime.now(timezone(timedelta(minutes=off))).replace(tzinfo=None)


def today_local():
    return now_local().date()


def sheet_names(name):
    return [name] + SHEET_ALIASES.get(name, [])


def rows_to_records(values, hint=None):
    """List of rows -> list of dicts. Finds the header row, cleans names, de-duplicates repeated headers
    (a second 'Remarks' becomes 'Remarks#2')."""
    if not values:
        return []
    hi = 0
    if hint:
        for i, row in enumerate(values[:8]):
            if any(str(c).strip() == hint for c in row if c is not None):
                hi = i
                break
    if len(values) < hi + 2:
        return []
    header, seen = [], {}
    for h in values[hi]:
        h = " ".join(str(h if h is not None else "").split())
        seen[h] = seen.get(h, 0) + 1
        header.append(h if seen[h] == 1 else f"{h}#{seen[h]}")
    recs = []
    for row in values[hi + 1:]:
        if all(c in ("", None) for c in row):
            continue
        recs.append({header[i]: row[i] if i < len(row) else "" for i in range(len(header)) if header[i]})
    return recs


def sheet_rows(ws, max_blank=120):
    """Rows of an openpyxl sheet; stops after a long run of empty rows (a sheet can claim 1,000,000+ rows)."""
    rows, blank = [], 0
    for row in ws.iter_rows(values_only=True):
        if all(c in (None, "") for c in row):
            blank += 1
            if blank >= max_blank:
                break
            continue
        blank = 0
        rows.append(list(row))
    return rows


def parse_workbook(raw_bytes):
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True, read_only=True)
    by_norm = {t.strip().lower(): t for t in wb.sheetnames}
    out = {}
    for name in SHEETS:
        real = next((by_norm[c.strip().lower()] for c in sheet_names(name) if c.strip().lower() in by_norm), None)
        out[name] = rows_to_records(sheet_rows(wb[real]), HEADER_HINT.get(name)) if real else []
    return out


# ---- 1. Microsoft Graph -------------------------------------------------------
_token = {"value": None, "exp": 0}


def graph_configured():
    if not all(os.environ.get(k) for k in ("TENANT_ID", "CLIENT_ID", "CLIENT_SECRET")):
        return False
    return bool((os.environ.get("GRAPH_DRIVE_ID") and os.environ.get("GRAPH_ITEM_ID"))
                or (os.environ.get("GRAPH_USER") and os.environ.get("GRAPH_FILE_PATH")))


def graph_token():
    if _token["value"] and time.time() < _token["exp"] - 120:
        return _token["value"]
    r = requests.post(f"https://login.microsoftonline.com/{os.environ['TENANT_ID']}/oauth2/v2.0/token", data={
        "client_id": os.environ["CLIENT_ID"], "client_secret": os.environ["CLIENT_SECRET"],
        "scope": "https://graph.microsoft.com/.default", "grant_type": "client_credentials"}, timeout=20)
    r.raise_for_status()
    j = r.json()
    _token["value"], _token["exp"] = j["access_token"], time.time() + int(j.get("expires_in", 3600))
    return _token["value"]


def read_graph():
    drive, item = os.environ.get("GRAPH_DRIVE_ID"), os.environ.get("GRAPH_ITEM_ID")
    if drive and item:
        base = f"https://graph.microsoft.com/v1.0/drives/{drive}/items/{item}/workbook"
    else:
        user, path = os.environ["GRAPH_USER"], os.environ["GRAPH_FILE_PATH"].lstrip("/")
        base = f"https://graph.microsoft.com/v1.0/users/{user}/drive/root:/{path}:/workbook"
    headers = {"Authorization": f"Bearer {graph_token()}"}
    out = {}
    for name in SHEETS:
        out[name] = []
        for cand in sheet_names(name):
            r = requests.get(f"{base}/worksheets('{cand}')/usedRange(valuesOnly=true)?$select=values",
                             headers=headers, timeout=25)
            if r.status_code == 404:
                continue
            r.raise_for_status()
            out[name] = rows_to_records(r.json().get("values", []), HEADER_HINT.get(name))
            break
    return out


# ---- 2. Local / synced file ---------------------------------------------------
def resolve_file(pattern):
    """Path, or a pattern with * (newest match wins). Returns None when nothing matches."""
    pattern = (pattern or "").strip().strip('"')
    if not pattern:
        return None
    if "*" not in pattern:
        return pattern if os.path.exists(pattern) else None
    hits = [f for f in glob.glob(pattern) if not os.path.basename(f).startswith("~$")]
    return max(hits, key=os.path.getmtime) if hits else None


def read_file_bytes(path):
    """Excel/OneDrive can briefly lock a file: retry, then fall back to a copy."""
    last = None
    for _ in range(3):
        try:
            with open(path, "rb") as f:
                return f.read()
        except PermissionError as exc:
            last = exc
            time.sleep(1.5)
    tmp = os.path.join(tempfile.gettempdir(), "sunrise_snapshot_%d.xlsx" % (abs(hash(path)) % 9999))
    try:
        shutil.copyfile(path, tmp)
    except OSError:
        if os.name == "nt":
            subprocess.run(["cmd", "/c", "copy", "/y", path, tmp], capture_output=True, timeout=30)
    if os.path.exists(tmp):
        with open(tmp, "rb") as f:
            return f.read()
    raise PermissionError(f"{last} - close the file in Excel, wait for OneDrive to finish syncing, "
                          "and set it to 'Always keep on this device'")


def read_local():
    main = resolve_file(os.environ["EXCEL_PATH"])
    if not main:
        raise FileNotFoundError(f"Excel file not found: {os.environ['EXCEL_PATH']}")
    out = parse_workbook(read_file_bytes(main))
    ipr_file = resolve_file(os.environ.get("EXCEL_IPR_PATH"))     # optional separate IPR tracker file
    if ipr_file:
        out["IPRTracker"] = parse_workbook(read_file_bytes(ipr_file)).get("IPRTracker", [])
    return out


# ---- 3. Share link --------------------------------------------------------------
def read_link():
    link = os.environ["EXCEL_LINK"].strip()
    if "download=1" not in link:
        link += ("&" if "?" in link else "?") + "download=1"
    r = requests.get(link, timeout=40, allow_redirects=True, headers={"User-Agent": "Mozilla/5.0 (SunriseDashboard)"})
    r.raise_for_status()
    if not r.content.startswith(b"PK"):   # an .xlsx is a zip; anything else is a login/HTML page
        raise RuntimeError("Link did not return an Excel file - set sharing to 'Anyone with the link'")
    return parse_workbook(r.content)


# ---- 4. Google Sheet ------------------------------------------------------------
def read_gsheet():
    sid = os.environ["GSHEET_ID"].strip()
    out = {}
    for name in SHEETS:
        out[name] = []
        for cand in sheet_names(name):
            r = requests.get(f"https://docs.google.com/spreadsheets/d/{sid}/gviz/tq",
                             params={"tqx": "out:csv", "sheet": cand}, timeout=25)
            if r.status_code in (400, 404):
                continue
            r.raise_for_status()
            r.encoding = "utf-8"
            if r.text.lstrip().lower().startswith(("<!doctype", "<html")):
                raise RuntimeError("Google Sheet is not shared: set Share > Anyone with the link > Viewer")
            out[name] = rows_to_records(list(csv.reader(io.StringIO(r.text))), HEADER_HINT.get(name))
            break
    return out


def pick_source():
    """-> (name, reader function)"""
    if graph_configured():
        return "onedrive", read_graph
    if os.environ.get("EXCEL_PATH"):
        return "local-excel", read_local
    if os.environ.get("EXCEL_LINK"):
        return "excel-link", read_link
    if os.environ.get("GSHEET_ID"):
        return "googlesheet", read_gsheet
    return "demo", get_demo_sheets


# =============================================================================
#  Small helpers
# =============================================================================
DATE_FORMATS = ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%d-%b-%Y", "%d %b %Y", "%Y-%m-%d %H:%M:%S",
                "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d-%m-%Y %H:%M:%S", "%d-%b-%y", "%d/%m/%y",
                "%d-%b-%Y %H:%M:%S", "%Y-%m-%dT%H:%M:%S")


def pdate(v):
    if v in (None, ""):
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, (int, float)):  # Excel serial date
        return date(1899, 12, 30) + timedelta(days=int(v))
    s = str(v).strip()
    for f in DATE_FORMATS:
        try:
            return datetime.strptime(s, f).date()
        except ValueError:
            pass
    return None


def fmt(d):
    return d.strftime("%d %b") if d else ""


def ptime(v):
    """Time of day from an Excel time cell, a fraction of a day, or text like 09:30 / 9:30 AM."""
    if v in (None, ""):
        return None
    if isinstance(v, datetime):
        return v.time().replace(second=0, microsecond=0)
    if isinstance(v, dtime):
        return v.replace(second=0, microsecond=0)
    if isinstance(v, (int, float)) and 0 <= v < 1:
        mins = round(v * 1440)
        return dtime(mins // 60 % 24, mins % 60)
    s = str(v).strip()
    for f in ("%H:%M", "%H:%M:%S", "%I:%M %p", "%I:%M%p", "%I %p"):
        try:
            return datetime.strptime(s, f).time().replace(second=0)
        except ValueError:
            pass
    return None


def mask_email(addr):
    name, _, dom = addr.partition("@")
    return (name[:1] + "***@" + dom) if dom else addr


def norm_status(s):
    s = str(s or "").strip().lower()
    if any(k in s for k in ("complete", "done", "closed", "resolved")):
        return "Completed"
    if any(k in s for k in ("ongoing", "progress", "partial", "wip", "working", "started")):
        return "Ongoing"
    return "Open"


def norm_pri(p):
    p = str(p or "").strip().capitalize()
    return p if p in PRIORITY_ORDER else "Medium"


def num(v, default=0):
    try:
        return float(str(v).replace(",", "").replace("%", "").replace("₹", "").strip())
    except (TypeError, ValueError):
        return default


def is_blank(v):
    return v is None or str(v).strip() in ("", "0", "0.0", "-", "NA", "N/A")


def strip_private(rows):
    return [{k: v for k, v in r.items() if not k.startswith("_")} for r in rows]


# =============================================================================
#  Building the dashboard data - one function per module
# =============================================================================
def build_tasks(rows, today):
    tasks = []
    for r in rows:
        due, done = pdate(r.get("Due Date")), pdate(r.get("Completed Date"))
        st = norm_status(r.get("Status"))
        if done and st != "Completed":      # a filled Completed Date means the task is finished
            st = "Completed"
        overdue = st != "Completed" and due is not None and due < today
        tasks.append({"id": str(r.get("ID", "")), "task": str(r.get("Task", "")),
                      "owner": str(r.get("Owner", "") or "Unassigned"), "priority": norm_pri(r.get("Priority")),
                      "status": st, "due": fmt(due), "overdue": overdue,
                      "days_over": (today - due).days if overdue else 0, "_done": done, "_due": due})
    total = len(tasks)
    completed = sum(t["status"] == "Completed" for t in tasks)

    # cumulative completion % for the last 14 days
    trend = []
    for i in range(13, -1, -1):
        d = today - timedelta(days=i)
        c = sum(1 for t in tasks if t["status"] == "Completed" and t["_done"] and t["_done"] <= d)
        trend.append({"label": d.strftime("%d %b"), "value": round(100 * c / total, 1) if total else 0})

    # owner-wise: the 8 busiest owners, best completion first
    owners = {}
    for t in tasks:
        o = owners.setdefault(t["owner"], {"owner": t["owner"], "total": 0, "done": 0, "overdue": 0})
        o["total"] += 1
        o["done"] += t["status"] == "Completed"
        o["overdue"] += t["overdue"]
    owner_list = sorted(({**o, "pct": round(100 * o["done"] / o["total"])} for o in owners.values()),
                        key=lambda x: (-x["total"], -x["pct"]))[:8]
    owner_list.sort(key=lambda x: -x["pct"])

    open_tasks = [t for t in tasks if t["status"] != "Completed"]
    open_tasks.sort(key=lambda t: (not t["overdue"], PRIORITY_ORDER[t["priority"]], t["_due"] or date.max))
    return {
        "open_tasks": open_tasks, "trend": trend, "owners": owner_list,
        "total": total, "completed": completed, "pct": round(100 * completed / total) if total else 0,
        "ongoing": sum(t["status"] == "Ongoing" for t in tasks),
        "open": sum(t["status"] == "Open" for t in tasks),
        "overdue": sum(t["overdue"] for t in tasks),
    }


def build_procurement(rows, today):
    """IPR -> PO. Blank PO number (and not cancelled) = PO awaited."""
    def short_ref(v):
        parts = str(v or "").split("|")
        return ("IPR-" + parts[-1].strip()) if len(parts) > 1 else str(v or "")

    items = []
    for r in rows:
        ref = r.get("IPR Reference")
        if is_blank(ref):
            continue
        d = pdate(r.get("IPR Date"))
        status = str(r.get("Status") or "").strip()
        ia = str(r.get("IA Team - Accept / Reject") or "").strip()
        combined = str(r.get("PO & IPR Status") or "")
        note = str(r.get("Remarks#2") or r.get("PO & IPR Status") or r.get("Remarks") or "").strip()
        po_raw = r.get("PO Number")
        has_po = not is_blank(po_raw)
        cancelled = any("cancel" in x.lower() for x in (status, ia, combined, note))
        if cancelled:
            stage = "Cancelled"
        elif has_po:
            stage = "PO received"
        elif status.lower() == "hold":
            stage = "On hold"
        elif ia.lower().startswith("accept"):
            stage = "Approved - no PO"
        elif status.lower() == "approved":
            stage = "With IA team"
        else:
            stage = "Approval pending"
        po_text = ""
        if has_po:
            po_text = str(int(num(po_raw))) if num(po_raw) else str(po_raw).strip()
        items.append({
            "ref": short_ref(ref), "date": fmt(d), "age": (today - d).days if d else 0,
            "dept": str(r.get("Department") or "").replace("&", " & ").replace("  ", " ").strip(),
            "item": str(r.get("Item Description") or "").strip(), "vendor": str(r.get("Vendor Name") or "").strip(),
            "amount": num(r.get("IPR Amount (Excluding GST)")), "po": po_text,
            "po_awaited": (not has_po) and not cancelled, "stage": stage,
            "note": "" if (note in ("PO Received", "None") or stage == "PO received") else note})

    awaited = sorted((i for i in items if i["po_awaited"]), key=lambda i: -i["age"])
    stages, depts = {}, {}
    for i in awaited:
        stages[i["stage"]] = stages.get(i["stage"], 0) + 1
        depts[i["dept"] or "Other"] = depts.get(i["dept"] or "Other", 0) + 1
    return awaited, {
        "awaited": len(awaited), "awaited_value": sum(i["amount"] for i in awaited),
        "over30": sum(1 for i in awaited if i["age"] > 30),
        "stages": sorted(({"stage": k, "n": v} for k, v in stages.items()), key=lambda x: -x["n"]),
        "depts": sorted(({"dept": k, "n": v} for k, v in depts.items()), key=lambda x: -x["n"])[:8],
        "rows": awaited[:12],
    }


def build_complaints(rows, today):
    out = []
    for r in rows:
        d = pdate(r.get("Date"))
        st = norm_status(r.get("Status"))
        out.append({"ticket": str(r.get("Ticket", "")), "date": fmt(d), "dept": str(r.get("Department", "")),
                    "text": str(r.get("Complaint", "")), "assigned": str(r.get("Assigned To", "") or "Unassigned"),
                    "priority": norm_pri(r.get("Priority")), "status": st, "raw_status": str(r.get("Status", "")),
                    "age_hours": (today - d).days * 24 if d else 0})
    open_c = [c for c in out if c["status"] != "Completed"]
    open_c.sort(key=lambda c: (PRIORITY_ORDER[c["priority"]], -c["age_hours"]))
    return open_c


def build_events(rows, today):
    out = []
    for r in rows:
        d = pdate(r.get("Date"))
        if d is None or d < today:
            continue
        out.append({"event": str(r.get("Event", "")), "date": fmt(d), "days": (d - today).days,
                    "venue": str(r.get("Venue", "")), "owner": str(r.get("Owner", "")),
                    "prep": int(num(r.get("Preparation %"))), "_d": d})
    out.sort(key=lambda e: e["_d"])
    return out


def build_followups(rows, today):
    out = []
    for r in rows:
        for i, nm in enumerate(FOLLOWUP_COLS, start=1):      # accept 'Column 1..6' headings too
            if nm not in r and f"Column {i}" in r:
                r[nm] = r[f"Column {i}"]
        if norm_status(r.get("Status")) == "Completed":
            continue
        due = pdate(r.get("Due Date"))
        out.append({"meeting": str(r.get("Meeting", "")), "point": str(r.get("Point", "")),
                    "owner": str(r.get("Owner", "")), "due": fmt(due),
                    "overdue": bool(due and due < today), "_d": due or date.max})
    out.sort(key=lambda f: f["_d"])
    return out


def safety_tip(rows, today):
    for r in rows:
        if pdate(r.get("Date")) == today and r.get("Tip"):
            return str(r["Tip"])
    pool = [str(r["Tip"]) for r in rows if r.get("Tip") and not pdate(r.get("Date"))]
    return pool[today.toordinal() % len(pool)] if pool else ""


def build_escalations(rows, open_tasks, open_complaints):
    matrix = [{"priority": norm_pri(r.get("Priority")), "level": int(num(r.get("Level"), 1)),
               "to": str(r.get("Escalate To", "")), "hours": num(r.get("Deadline Hours"))} for r in rows]
    matrix.sort(key=lambda m: (PRIORITY_ORDER[m["priority"]], m["level"]))

    def level_for(priority, age_hours):
        reached = None
        for m in matrix:
            if m["priority"] == priority and age_hours >= m["hours"]:
                reached = m
        return reached

    active = []
    for t in open_tasks:
        if t["overdue"]:
            m = level_for(t["priority"], t["days_over"] * 24)
            if m:
                active.append({"kind": "Task", "ref": t["id"], "text": t["task"], "owner": t["owner"],
                               "priority": t["priority"], "level": m["level"], "to": m["to"],
                               "age": f"{t['days_over']}d overdue"})
    for c in open_complaints:
        m = level_for(c["priority"], c["age_hours"])
        if m:
            active.append({"kind": "Complaint", "ref": c["ticket"], "text": c["text"], "owner": c["assigned"],
                           "priority": c["priority"], "level": m["level"], "to": m["to"],
                           "age": f"{c['age_hours'] // 24}d open"})
    active.sort(key=lambda a: (-a["level"], PRIORITY_ORDER[a["priority"]]))
    return matrix, active


def build_ticker(open_tasks, awaited, open_complaints, events, followups, alerts=()):
    t = []
    for a in [x for x in alerts if x["due"]][:4]:
        t.append({"tag": "ALERT", "cls": "red", "text": a["text"]})
    for x in [x for x in open_tasks if x["overdue"]][:6]:
        t.append({"tag": "OVERDUE", "cls": "red", "text": f"{x['task']} ({x['owner']}) +{x['days_over']}d"})
    for i in awaited[:4]:
        t.append({"tag": "PO AWAITED", "cls": "amber", "text": f"{i['ref']} {i['item'][:50]} ({i['age']}d)"})
    for c in [x for x in open_complaints if x["priority"] == "High"][:3]:
        t.append({"tag": "HELPDESK", "cls": "red", "text": f"{c['ticket']} {c['text']}"})
    for e in events[:2]:
        t.append({"tag": "EVENT", "cls": "blue", "text": f"{e['event']} in {e['days']}d"})
    for f in followups[:3]:
        t.append({"tag": "FOLLOW-UP", "cls": "amber" if f["overdue"] else "green",
                  "text": f"{f['point']} ({f['owner']})"})
    return t


# ---------------------------------------------------------------- alerts + e-mail
_mail = {"sent": None, "fail": {}, "busy": False, "error": ""}
_mail_lock = threading.Lock()
MAIL_MAX_AGE_DAYS = 2       # only alerts due within the last 2 days are e-mailed (no flood of old rows)
MAIL_RETRY_SECONDS = 600


def smtp_ready():
    return bool(os.environ.get("SMTP_HOST") and (os.environ.get("SMTP_USER") or os.environ.get("SMTP_FROM")))


def _log_path():
    return os.environ.get("ALERT_LOG_PATH") or os.path.join(os.path.dirname(os.path.abspath(__file__)), "alerts_sent.json")


def _sent_log():
    if _mail["sent"] is None:
        try:
            with open(_log_path(), encoding="utf-8") as f:
                _mail["sent"] = json.load(f)
        except (OSError, ValueError):
            _mail["sent"] = {}
    return _mail["sent"]


def _save_sent_log():
    try:
        with open(_log_path(), "w", encoding="utf-8") as f:
            json.dump(_mail["sent"], f)
    except OSError as exc:
        app.logger.error("Could not save alert log: %s", exc)


def recipients(raw):
    raw = raw or os.environ.get("ALERT_EMAIL_TO", "")
    return [x for x in re.split(r"[,;\s]+", str(raw)) if "@" in x]


def send_mail(to_list, subject, body):
    host, port = os.environ["SMTP_HOST"], int(os.environ.get("SMTP_PORT", 587))
    user, pwd = os.environ.get("SMTP_USER"), os.environ.get("SMTP_PASSWORD")
    msg = EmailMessage()
    msg["From"] = os.environ.get("SMTP_FROM") or user
    msg["To"] = ", ".join(to_list)
    msg["Subject"] = subject
    msg.set_content(body)
    if port == 465:
        server = smtplib.SMTP_SSL(host, port, timeout=30)
    else:
        server = smtplib.SMTP(host, port, timeout=30)
        if os.environ.get("SMTP_TLS", "1") != "0":
            server.starttls()
    try:
        if user and pwd:
            server.login(user, pwd)
        server.send_message(msg)
    finally:
        try:
            server.quit()
        except Exception:
            pass


def _mail_worker(jobs):
    try:
        for it in jobs:
            try:
                dash = os.environ.get("DASHBOARD_URL", "")
                body = (f"ALERT: {it['text']}\n\nDate: {it['date']} {it['time']}".rstrip() +
                        f"\nPriority: {it['priority']}\n" + (f"\nDashboard: {dash}\n" if dash else "") +
                        "\nSent automatically by the Live Operation Tracker.")
                send_mail(it["_to"], f"[ALERT] {it['text']} - {it['date']}", body)
                _sent_log()[it["key"]] = now_local().isoformat(timespec="seconds")
                _save_sent_log()
                _mail["fail"].pop(it["key"], None)
                _mail["error"] = ""
            except Exception as exc:     # remember the failure and retry later
                _mail["fail"][it["key"]] = time.time()
                _mail["error"] = f"{type(exc).__name__}: {exc}"[:120]
                app.logger.error("Alert e-mail failed: %s", exc)
    finally:
        _mail["busy"] = False


def dispatch_alert_mail(items, today):
    """E-mail each due alert once (in a background thread so a slow mail server never delays the screen)."""
    if not smtp_ready() or _mail["busy"]:
        return
    sent = _sent_log()
    jobs = [i for i in items if i["due"] and i["key"] not in sent and i["_to"] and i["_age"] <= MAIL_MAX_AGE_DAYS
            and time.time() - _mail["fail"].get(i["key"], 0) > MAIL_RETRY_SECONDS]
    if jobs:
        _mail["busy"] = True
        threading.Thread(target=_mail_worker, args=(jobs,), daemon=True).start()


def build_alerts(rows, now):
    today = now.date()
    sent = _sent_log()
    items = []
    for r in rows:
        text = str(r.get("Alert / Task") or r.get("Alert") or r.get("Task") or "").strip()
        d = pdate(r.get("Alert Date"))
        if not text or not d:
            continue
        if any(k in str(r.get("Status") or "").lower() for k in ("done", "closed", "complete", "cancel", "stop")):
            continue
        t = ptime(r.get("Alert Time"))
        due = datetime.combine(d, t or dtime(0, 0)) <= now
        to_list = recipients(r.get("Email To"))
        key = hashlib.sha1(f"{d}|{t}|{text}|{','.join(to_list)}".encode()).hexdigest()[:16]
        age = (today - d).days
        if key in sent:
            mail = "sent"
        elif not smtp_ready():
            mail = "off"
        elif not to_list:
            mail = "norecipient"
        elif not due:
            mail = "scheduled"
        elif age > MAIL_MAX_AGE_DAYS:
            mail = "skipped"
        elif key in _mail["fail"]:
            mail = "failed"
        else:
            mail = "pending"
        items.append({
            "key": key, "date": fmt(d), "time": t.strftime("%H:%M") if t else "", "text": text,
            "priority": norm_pri(r.get("Priority")), "due": due, "days": (d - today).days,
            "state": ("OVERDUE" if d < today else "DUE NOW") if due else "UPCOMING",
            "to": ", ".join(mask_email(x) for x in to_list), "mail": mail,
            "_to": to_list, "_age": age, "_d": datetime.combine(d, t or dtime(0, 0))})
    items.sort(key=lambda i: (not i["due"], i["_d"]))
    return items


def build_payload(sheets, source):
    now = now_local()
    today = now.date()
    tk = build_tasks(sheets.get("Tasks", []), today)
    awaited, ipr = build_procurement(sheets.get("IPRTracker", []), today)
    alert_items = build_alerts(sheets.get("Alerts", []), now)
    dispatch_alert_mail(alert_items, today)
    complaints = build_complaints(sheets.get("Complaints", []), today)
    events = build_events(sheets.get("Events", []), today)
    followups = build_followups(sheets.get("FollowUps", []), today)
    matrix, escalations = build_escalations(sheets.get("Escalation", []), tk["open_tasks"], complaints)

    return {
        "source": source, "updated": now.strftime("%H:%M:%S"), "refresh": REFRESH_SECONDS,
        "kpis": {"total": tk["total"], "open": tk["open"], "ongoing": tk["ongoing"], "completed": tk["completed"],
                 "overdue": tk["overdue"], "pct": tk["pct"], "target": int(os.environ.get("TARGET_PCT", 85)),
                 "po_pending": ipr["awaited"], "po_value": ipr["awaited_value"],
                 "complaints_open": len(complaints),
                 "complaints_high": sum(c["priority"] == "High" for c in complaints),
                 "events_upcoming": len(events), "followups_pending": len(followups),
                 "escalations": len(escalations), "alerts_due": sum(1 for a in alert_items if a["due"])},
        "tasks": strip_private(tk["open_tasks"][:60]),
        "ipr": ipr,
        "alerts": {"items": strip_private(alert_items[:14]), "due": strip_private([a for a in alert_items if a["due"]][:6]),
                   "due_count": sum(1 for a in alert_items if a["due"]),
                   "upcoming_count": sum(1 for a in alert_items if not a["due"] and a["days"] <= 7),
                   "mail_ready": smtp_ready(), "mail_error": _mail["error"]},
        "complaints": complaints[:14], "events": strip_private(events[:8]),
        "followups": strip_private(followups[:10]), "tip": safety_tip(sheets.get("SafetyTips", []), today),
        "matrix": matrix, "escalations": escalations[:10],
        "trend": tk["trend"], "owners": tk["owners"],
        "ticker": build_ticker(tk["open_tasks"], awaited, complaints, events, followups, alert_items),
    }


# =============================================================================
#  Cache: answer instantly, refresh in the background
# =============================================================================
_cache = {"ts": 0, "payload": None}
_lock = threading.Lock()
_state = {"refreshing": False}


def compute_payload():
    name, reader = pick_source()
    try:
        payload = build_payload(reader(), name)
        payload["error"] = None
    except Exception as exc:   # keep the last good data on screen and show what went wrong
        app.logger.error("Data refresh failed: %s", exc)
        last = _cache["payload"]
        payload = dict(last) if last else build_payload(get_demo_sheets(), "demo")
        payload["error"] = f"{type(exc).__name__}: {exc}"[:200]
    return payload


def _refresh():
    try:
        payload = compute_payload()
        with _lock:
            _cache.update(ts=time.time(), payload=payload)
    finally:
        _state["refreshing"] = False


def get_payload():
    if _cache["payload"] is None:          # very first request has to wait once
        with _lock:
            if _cache["payload"] is None:
                _cache.update(ts=time.time(), payload=compute_payload())
        return _cache["payload"]
    if time.time() - _cache["ts"] >= REFRESH_SECONDS - 2 and not _state["refreshing"]:
        _state["refreshing"] = True
        threading.Thread(target=_refresh, daemon=True).start()
    return _cache["payload"]


# =============================================================================
#  Routes
# =============================================================================
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/data")
def api_data():
    resp = jsonify(get_payload())
    resp.headers["Cache-Control"] = "no-store"
    return resp


@app.route("/alerts/test-email")
def alerts_test_email():
    """Sends one test e-mail to ALERT_EMAIL_TO. Disabled unless ALERT_ADMIN_KEY is set; call ?key=<that key>."""
    from flask import request
    key = os.environ.get("ALERT_ADMIN_KEY", "")
    if not key or request.args.get("key") != key:
        return {"ok": False, "error": "disabled or wrong key"}, 403
    if not smtp_ready():
        return {"ok": False, "error": "SMTP_HOST and SMTP_USER/SMTP_FROM are not set"}
    to = recipients("")
    if not to:
        return {"ok": False, "error": "ALERT_EMAIL_TO is not set"}
    try:
        send_mail(to, "[TEST] Live Operation Tracker alert e-mail", "If you can read this, alert e-mails are working.")
        return {"ok": True, "sent_to": [mask_email(x) for x in to]}
    except Exception as exc:
        return {"ok": False, "error": f"{type(exc).__name__}: {exc}"[:200]}


@app.route("/health")
def health():
    return {"status": "ok", "source": pick_source()[0]}


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=False, threaded=True)
