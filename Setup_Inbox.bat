@echo off
cd /d C:\Users\admin\Documents\Claude\Projects\work-inbox
echo Downloading latest inbox setup script from GitHub...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$t=[DateTimeOffset]::UtcNow.ToUnixTimeSeconds(); Invoke-WebRequest -UseBasicParsing ('https://raw.githubusercontent.com/begb0037admin/work-inbox/main/setup_inbox.py?t='+$t) -OutFile 'setup_inbox.py'"
echo Running Outlook inbox setup (folders + rules)...
python setup_inbox.py
echo.
echo Done. Check Outlook folder list and Rules and Alerts to verify.
pause
