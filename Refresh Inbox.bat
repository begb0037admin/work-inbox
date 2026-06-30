@echo off
cd /d C:\Users\admin\Documents\Claude\Projects\work-inbox
echo Downloading latest script from GitHub...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$t=[DateTimeOffset]::UtcNow.ToUnixTimeSeconds(); Invoke-WebRequest -UseBasicParsing ('https://raw.githubusercontent.com/begb0037admin/work-inbox/main/fetch_inbox.py?t='+$t) -OutFile 'fetch_inbox.py'"
echo Running inbox briefing...
python fetch_inbox.py
echo.
echo Done. Refresh your dashboard.
pause
