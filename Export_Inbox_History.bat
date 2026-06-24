@echo off
cd /d C:\Users\admin\Documents\Claude\Projects\work-inbox
echo Downloading latest export script from GitHub...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$t=[DateTimeOffset]::UtcNow.ToUnixTimeSeconds(); Invoke-WebRequest -UseBasicParsing ('https://raw.githubusercontent.com/begb0037admin/work-inbox/main/export_inbox_history.py?t='+$t) -OutFile 'export_inbox_history.py'"
echo Running inbox history export (6 months)...
python export_inbox_history.py
echo.
echo Done. Check data/inbox_export.json in GitHub.
pause
