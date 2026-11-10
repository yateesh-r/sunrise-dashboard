@echo off
title Sunrise Dashboard (local)
cd /d "%~dp0"
rem The * means "any characters", so OPERATIONS Tracker.xlsx / OPERATIONS_Tracker.xlsx both work. Change the folder if needed.
set "EXCEL_PATH=D:\OneDrive - Smart Services\OPERATIONS*Tracker*.xlsx"
rem OPTIONAL: only if the IPR tracker is kept in a separate file (remove "rem" to turn on):
rem set "EXCEL_IPR_PATH=D:\OneDrive - Smart Services\Updated IPR Tracker*.xlsx"
rem ---------------- E-MAIL ALERTS (remove the word "rem" at the start of these 5 lines and fill them in) ----------------
rem Gmail: Google Account > Security > 2-Step Verification > App passwords > create one, paste the 16 letters below.
rem set "SMTP_HOST=smtp.gmail.com"
rem set "SMTP_PORT=587"
rem set "SMTP_USER=yourname@gmail.com"
rem set "SMTP_PASSWORD=your-16-letter-app-password"
rem set "ALERT_EMAIL_TO=person1@company.com,person2@company.com"
rem set "ALERT_ADMIN_KEY=choose-any-secret-word"
echo Installing requirements (first time only)...
python -m pip install -q flask requests openpyxl
echo.
echo Dashboard running. Open http://localhost:5000 on this laptop.
echo For the TV, open http://YOUR-LAPTOP-IP:5000  (find the IP with: ipconfig)
echo Keep this window open. Press Ctrl+C to stop.
echo.
python app.py
pause
