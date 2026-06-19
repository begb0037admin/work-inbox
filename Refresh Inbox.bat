@echo off
cd /d C:\Users\admin\Documents\Claude\Projects\work-inbox
powershell -NoProfile -ExecutionPolicy Bypass -Command "$t=[DateTimeOffset]::UtcNow.ToUnixTimeSeconds(); Invoke-WebRequest -UseBasicParsing ('https://raw.githubusercontent.com/begb0037admin/work-inbox/main/fetch_inbox.py?t='+$t) -OutFile 'fetch_inbox.py'"
python fetch_inbox.py
echo.
pause
