"""Demo data used until the OneDrive workbook is connected.
Dates are relative to today so the demo always looks current."""
import random
from datetime import date, timedelta

OWNERS = ["Rajesh", "Anita", "Suresh", "Priya", "Mohan", "Deepa"]


def _d(offset, today):
    return (today + timedelta(days=offset)).isoformat()


def get_demo_sheets(today=None):
    today = today or date.today()
    rnd = random.Random(today.toordinal())

    tasks_src = [
        ("Prepare monthly MIS report", "High"),
        ("Fire extinguisher inspection", "High"),
        ("Update vendor contact list", "Low"),
        ("Conduct toolbox talk - Shift A", "Medium"),
        ("Close audit observations (Ph-1)", "High"),
        ("Renew AMC for DG set", "Medium"),
        ("Stores stock reconciliation", "Medium"),
        ("Submit statutory compliance returns", "High"),
        ("Housekeeping audit - Block C", "Low"),
        ("Training calendar for Q4", "Medium"),
        ("Review PPE inventory", "Medium"),
        ("Update emergency contact board", "Low"),
        ("Pump house preventive maintenance", "High"),
        ("Visitor pass process revision", "Low"),
        ("Quarterly safety committee minutes", "Medium"),
        ("Vendor payment follow-up", "High"),
        ("Fleet insurance renewal", "Medium"),
        ("Canteen hygiene inspection", "Medium"),
        ("CCTV coverage review", "Low"),
        ("First-aid kit refill", "Medium"),
    ]
    tasks = []
    for i, (name, pr) in enumerate(tasks_src, start=1):
        state = rnd.choices(["Completed", "Ongoing", "Open"], [45, 25, 30])[0]
        due = rnd.randint(-6, 10)
        comp = ""
        if state == "Completed":
            comp = _d(-rnd.randint(0, 13), today)
            due = rnd.randint(-10, 5)
        tasks.append({
            "ID": f"T-{100+i}", "Task": name, "Owner": rnd.choice(OWNERS),
            "Priority": pr, "Status": state, "Due Date": _d(due, today),
            "Completed Date": comp,
        })

    complaints_src = [
        ("AC not working - Admin office", "Facilities"),
        ("Printer jam - Accounts", "IT"),
        ("Water leakage near canteen", "Maintenance"),
        ("Biometric reader fault - Gate 2", "Security"),
        ("Street light out - Parking", "Electrical"),
        ("Wi-Fi drop in Conference Room", "IT"),
        ("Washroom cleaning delay - Block B", "Housekeeping"),
    ]
    complaints = []
    for i, (txt, dept) in enumerate(complaints_src, start=1):
        complaints.append({
            "Ticket": f"HD-{2200+i}", "Date": _d(-rnd.randint(0, 6), today),
            "Department": dept, "Complaint": txt, "Assigned To": rnd.choice(OWNERS),
            "Priority": rnd.choice(["High", "Medium", "Low", "Medium"]),
            "Status": rnd.choice(["Open", "In Progress", "Resolved", "Open"]),
        })

    events = [
        ("Safety Week Kick-off", 3, "Main Hall", 70),
        ("Customer Visit - Audit Team", 6, "Conference Room", 45),
        ("Annual Day Rehearsal", 12, "Auditorium", 30),
        ("Fire Mock Drill", 8, "Plant Area", 85),
        ("Blood Donation Camp", 18, "Canteen Block", 20),
    ]
    ev = [{"Event": n, "Date": _d(d, today), "Venue": v, "Owner": rnd.choice(OWNERS),
           "Preparation %": p, "Status": "Planned"} for n, d, v, p in events]

    fu_src = ["Share revised production schedule", "Confirm vendor delivery dates",
              "Circulate safety audit action list", "Finalise budget for Q4 training",
              "Update escalation contacts", "Review pending PO approvals"]
    followups = [{"Meeting": rnd.choice(["Weekly Ops Review", "Safety Committee", "Management Review"]),
                  "Date": _d(-rnd.randint(1, 8), today), "Point": p, "Owner": rnd.choice(OWNERS),
                  "Due Date": _d(rnd.randint(-3, 7), today),
                  "Status": rnd.choice(["Pending", "Pending", "In Progress", "Done"])}
                 for p in fu_src]

    tips = [
        "Stop. Think. Act. Take 10 seconds to check your surroundings before you start.",
        "Wear your PPE every time - it only protects you when it is on.",
        "Keep walkways clear. A tidy workplace is a safe workplace.",
        "Report every near-miss. It prevents tomorrow's accident.",
        "Never bypass a safety guard or interlock - not even for a quick job.",
        "Lift with your legs, not your back. Ask for help with heavy loads.",
        "Know your nearest exit and assembly point.",
        "Switch off and lock out before maintenance. Always.",
        "Hydrate regularly - fatigue and heat reduce alertness.",
        "Phones down while walking in plant areas.",
    ]
    safety = [{"Date": "", "Tip": t} for t in tips]

    escalation = [
        {"Priority": "High", "Level": 1, "Escalate To": "Section Head", "Deadline Hours": 24},
        {"Priority": "High", "Level": 2, "Escalate To": "Department Manager", "Deadline Hours": 48},
        {"Priority": "High", "Level": 3, "Escalate To": "Plant Head", "Deadline Hours": 72},
        {"Priority": "Medium", "Level": 1, "Escalate To": "Section Head", "Deadline Hours": 48},
        {"Priority": "Medium", "Level": 2, "Escalate To": "Department Manager", "Deadline Hours": 96},
        {"Priority": "Medium", "Level": 3, "Escalate To": "Plant Head", "Deadline Hours": 168},
        {"Priority": "Low", "Level": 1, "Escalate To": "Section Head", "Deadline Hours": 96},
        {"Priority": "Low", "Level": 2, "Escalate To": "Department Manager", "Deadline Hours": 192},
        {"Priority": "Low", "Level": 3, "Escalate To": "Plant Head", "Deadline Hours": 336},
    ]

    ipr_src = [
        ("Wall crack filling and painting proposal", "M&E", "RJ Facility Management", 198000, -150, "Approved", "", "Awaiting confirmation"),
        ("Busduct rectification work - Tower 30", "M&E", "Medium Voltage Technologies", 338000, -148, "Approved", "", "Submitted For PO Creation"),
        ("Supply of pantry consumables", "SS", "Bookmark Office Automation", 683000, -63, "Approved", "", "Budget issue to share the PO"),
        ("Grease trap sludge removal", "SS", "Asianet Groups", 33600, -72, "Approved", "", "Submitted For PO Creation"),
        ("Fire extinguisher refilling", "EHS", "Mahalaxmi Safety Solutions", 125000, -20, "Approved", "", "Submitted For PO Creation"),
        ("Housekeeping machine spares", "HK", "Clean Tech Services", 54000, -12, "", "", "Sent to approver"),
        ("CCTV camera replacement", "IT", "Secure Vision", 242000, -9, "Approved", "", ""),
        ("Diwali gift hampers for clients", "Admin", "SKRP Enterprise", 14573, -3, "Approved", "", ""),
        ("Panel space heaters for transformers", "M&E", "SKRP Enterprise", 27625, -2, "", "", "Sent to approver"),
        ("Hazardous waste authorization renewal", "EHS", "RR Enterprises", 310000, -190, "Approved", "4700037265", "PO Received"),
        ("Coffee machine consumables", "SS", "SAR Coffee India", 99620, -185, "Approved", "4700037295", "PO Received"),
        ("Emergency window signage", "EHS", "Mahalaxmi Safety Solutions", 23500, -180, "Approved", "4700037349", "PO Received"),
        ("Member hub consumables", "SS", "Bookmark Office Automation", 15000, -175, "canceled", "", "Shared IA Team For Cancelling The IPR"),
    ]
    iprs = []
    for i, (item, dept, vendor, amt, off, st, po, note) in enumerate(ipr_src, start=1):
        iprs.append({"IPR Reference": f"ECOWORLD|BLR|2026-27|IPR|{i:03d}", "IPR Date": _d(off, today),
                     "Department": dept, "Requester": rnd.choice(OWNERS), "Item Description": item,
                     "Vendor Name": vendor, "IPR Amount (Excluding GST)": amt, "Status": st,
                     "IA Team - Accept / Reject": "Accepted" if st == "Approved" else "",
                     "PO Number": po, "PO Date": "", "PO & IPR Status": note, "Remarks": note})

    alerts = [
        {"Alert Date": _d(0, today), "Alert Time": "09:00", "Alert / Task": "Submit statutory compliance returns", "Priority": "High", "Email To": "", "Status": "Active"},
        {"Alert Date": _d(-1, today), "Alert Time": "", "Alert / Task": "Renew AMC for DG set", "Priority": "Medium", "Email To": "", "Status": "Active"},
        {"Alert Date": _d(2, today), "Alert Time": "10:30", "Alert / Task": "Safety committee meeting - share agenda", "Priority": "Medium", "Email To": "", "Status": "Active"},
        {"Alert Date": _d(5, today), "Alert Time": "", "Alert / Task": "Fire mock drill preparation check", "Priority": "High", "Email To": "", "Status": "Active"},
    ]

    return {"Alerts": alerts, "IPRTracker": iprs, "Tasks": tasks, "Complaints": complaints, "Events": ev,
            "FollowUps": followups, "SafetyTips": safety, "Escalation": escalation}
