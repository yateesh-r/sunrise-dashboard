# Sunrise Live Operation Tracker

A Flask web app that turns one Excel / Google Sheet workbook into a live, auto-refreshing TV dashboard
(16:9, rotating pages, refreshes every 60 seconds).

## Pages
1. Overview - completion %, overdue, PO awaited, complaints, events, escalations, trend, owner performance
2. Task Tracker - open and overdue tasks (flips through them 15 at a time)
3. Procurement (IPR -> PO) - IPRs whose **PO Number is blank** are shown as **PO awaited**
4. Helpdesk Complaints
5. Events & Follow-ups
6. Escalation Matrix
7. Alerts - reminders you create by date; shown as a red pop-up + banner and e-mailed once

Pages with nothing to show (no complaints, no events/follow-ups, no alerts) are skipped automatically.
Keys / TV remote: left and right arrows change page, OK or Space pauses, 1-7 jump to a page.
URL options: `?page=3` start page, `?rotate=0` no auto-rotation, `?lite=1` no animations (slow TVs).

## The workbook (one file, these sheets)
Tasks, **IPR Tracker**, Alerts, Complaints, Events, FollowUps, SafetyTips, Escalation.
Do not rename the sheets or the header names. The old *PendingPO* sheet is no longer used - delete it.

IPR Tracker columns the dashboard reads: IPR Reference, IPR Date, Department, Item Description, Vendor Name,
IPR Amount (Excluding GST), Status, IA Team - Accept / Reject, PO Number, PO & IPR Status, Remarks.
All other columns are ignored. Cancelled IPRs (status contains "cancel") are not shown.

`python make_template.py` creates a fresh `MasterWorkbook.xlsx` (with sample rows; `--blank` for headers only).

## Where the data comes from (first one configured wins)
| Setting | Meaning |
|---|---|
| `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET` + `GRAPH_DRIVE_ID`, `GRAPH_ITEM_ID` | Microsoft Graph (needs IT approval) |
| `EXCEL_PATH` | Excel file on the computer running the app. `*` allowed, newest match wins |
| `EXCEL_IPR_PATH` | Optional: IPR tracker kept in a separate file |
| `EXCEL_LINK` | OneDrive/SharePoint "Anyone with the link" share link |
| `GSHEET_ID` | Google Sheet shared "Anyone with the link: Viewer" (set the sheet locale to India) |
| none | Demo data |

## Run on a laptop / office PC (Windows)
Put `app.py`, `demo_data.py`, `requirements.txt`, `run_local.bat` and the `templates` folder in one folder,
edit the path inside `run_local.bat`, and double-click it. Open http://localhost:5000
(for the TV use http://YOUR-PC-IP:5000).

## Run on Render
GitHub repo layout (`index.html` must be inside a folder named `templates`):
```
app.py  demo_data.py  make_template.py  requirements.txt  render.yaml  Procfile
templates/index.html
```
Build: `pip install -r requirements.txt`
Start: `gunicorn app:app --bind 0.0.0.0:$PORT --workers 1 --threads 4 --timeout 60`
Render cannot see a laptop or OneDrive folder, so set `GSHEET_ID` (or `EXCEL_LINK`) there.
`/health` shows which source is in use.

## Alerts and e-mail
Alerts sheet columns: Alert Date, Alert Time (optional), Alert / Task, Priority, Email To (optional), Status (Active / Done).
- From the alert date/time the TV shows a red banner, pops up once, and lists it on the Alerts page until Status = Done.
- If SMTP is configured the alert is e-mailed **once**, to "Email To" (several ids separated by , or ;) or to `ALERT_EMAIL_TO`.
- Only alerts due in the last 2 days are e-mailed (old rows never cause a flood). Failed e-mails retry every 10 minutes.
- Sent alerts are remembered in `alerts_sent.json` next to `app.py`. On Render's free plan that file is wiped on each
  redeploy/restart, so an alert may be e-mailed again once after a restart.
- E-mail addresses are masked on the public screen (a***@gmail.com).
- Settings: SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASSWORD, ALERT_EMAIL_TO (see `.env.example`).
- Times use India time (UTC+5:30); change with `TZ_OFFSET_MINUTES`.
