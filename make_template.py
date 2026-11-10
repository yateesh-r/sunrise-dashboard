"""Generates MasterWorkbook.xlsx - the Excel file you keep in OneDrive.
Sheet names and column headers MUST match what app.py reads. Run:  python make_template.py
Pass --blank to create an empty template (headers only)."""
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter
from datetime import date, datetime

from demo_data import get_demo_sheets

LAYOUT = {
    "Tasks": ["ID", "Task", "Owner", "Priority", "Status", "Due Date", "Completed Date"],
    "Complaints": ["Ticket", "Date", "Department", "Complaint", "Assigned To", "Priority", "Status"],
    "Events": ["Event", "Date", "Venue", "Owner", "Preparation %", "Status"],
    "FollowUps": ["Meeting", "Date", "Point", "Owner", "Due Date", "Status"],
    "SafetyTips": ["Date", "Tip"],
    "Escalation": ["Priority", "Level", "Escalate To", "Deadline Hours"],
    "Alerts": ["Alert Date", "Alert Time", "Alert / Task", "Priority", "Email To", "Status"],
    "IPRTracker": ["IPR Reference", "IPR Date", "Department", "Requester", "Item Description", "Vendor Name",
                   "IPR Amount (Excluding GST)", "Status", "IA Team - Accept / Reject", "PO Number", "PO Date",
                   "PO & IPR Status", "Remarks"],
}
DATE_COLS = {"Due Date", "Completed Date", "Date", "IPR Date", "Alert Date"}
DROPDOWNS = {
    "Tasks": {"Priority": "High,Medium,Low", "Status": "Open,Ongoing,Completed"},
    "IPRTracker": {"Status": "Approved,Pending,Hold,Cancelled"},
    "Alerts": {"Priority": "High,Medium,Low", "Status": "Active,Done"},
    "Complaints": {"Priority": "High,Medium,Low", "Status": "Open,In Progress,Resolved"},
    "Events": {"Status": "Planned,Completed,Cancelled"},
    "FollowUps": {"Status": "Pending,In Progress,Done"},
    "Escalation": {"Priority": "High,Medium,Low"},
}


def to_date(v):
    if isinstance(v, str) and v:
        try:
            return datetime.strptime(v, "%Y-%m-%d").date()
        except ValueError:
            return v
    return v or None


def build(path="MasterWorkbook.xlsx", blank=False):
    demo = {} if blank else get_demo_sheets()
    wb = Workbook()
    wb.remove(wb.active)
    red = PatternFill("solid", bgColor="FFC7CE", fgColor="FFC7CE")

    for name, cols in LAYOUT.items():
        ws = wb.create_sheet(name)
        ws.append(cols)
        rows = demo.get(name, [])
        for r in rows:
            ws.append([to_date(r.get(c, "")) if c in DATE_COLS else r.get(c, "") for c in cols])
        n = max(len(rows), 1) + 1
        ref = f"A1:{get_column_letter(len(cols))}{n}"
        t = Table(displayName=f"t{name}", ref=ref)
        t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(t)
        for i, c in enumerate(cols, start=1):
            letter = get_column_letter(i)
            ws.column_dimensions[letter].width = 44 if c in ("Task", "Complaint", "Point", "Tip", "Description") else 18
            if c in DATE_COLS:
                for row in range(2, 502):
                    ws[f"{letter}{row}"].number_format = "dd-mmm-yyyy"
        for col, items in DROPDOWNS.get(name, {}).items():
            idx = cols.index(col) + 1
            L = get_column_letter(idx)
            dv = DataValidation(type="list", formula1=f'"{items}"', allow_blank=True)
            ws.add_data_validation(dv)
            dv.add(f"{L}2:{L}500")
        ws.freeze_panes = "A2"

    # highlight overdue tasks in the Tasks sheet (due < today and not completed)
    ws = wb["Tasks"]
    ws.conditional_formatting.add("A2:G500", FormulaRule(
        formula=['AND($F2<>"",$F2<TODAY(),$E2<>"Completed")'], fill=red))

    # instructions sheet
    ins = wb.create_sheet("READ ME", 0)
    for line in [
        "SUNRISE OPERATIONS DASHBOARD - master workbook",
        "",
        "1. Keep this file in company OneDrive. Do NOT rename sheets or column headers.",
        "2. Add one row per item. Dates: dd-mmm-yyyy. Pick Priority/Status from the dropdowns.",
        "3. Tasks: set Status = Completed and fill Completed Date to feed the trend graph.",
        "4. SafetyTips: leave Date blank for a rotating pool, or enter a date to pin a tip to that day.",
        "5. IPRTracker: leave PO Number BLANK until the PO is raised - blank PO shows on the TV as PO AWAITED.",
        "6. Alerts: Alert Date (+ optional Alert Time) and the reminder text. On that date the TV pops it up and e-mails it. Status = Done stops it.",
        "5b. Escalation: Deadline Hours = hours (or overdue hours) after which that level is triggered.",
        "6. The TV refreshes every 60 seconds. Save the file (autosave in OneDrive) and wait a minute.",
    ]:
        ins.append([line])
    ins["A1"].font = Font(bold=True, size=14)
    ins.column_dimensions["A"].width = 100
    wb.save(path)
    print("saved", path)


if __name__ == "__main__":
    build(blank="--blank" in sys.argv)
