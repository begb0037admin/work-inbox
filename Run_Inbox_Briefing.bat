@echo off
title Inbox Briefing - Manual Refresh
cd /d C:\Users\admin\Documents\Claude\Projects\work-inbox
echo Downloading latest script from GitHub...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$t=[DateTimeOffset]::UtcNow.ToUnixTimeSeconds(); Invoke-WebRequest -UseBasicParsing ('https://raw.githubusercontent.com/begb0037admin/work-inbox/main/fetch_inbox.py?t='+$t) -OutFile 'fetch_inbox.py'"
if errorlevel 1 (
    echo WARNING - could not download latest script. Running existing local copy.
)
echo.
echo Running inbox briefing...
echo.
python fetch_inbox.py
if %errorlevel% equ 0 (
    echo.
    echo Done. Refresh your dashboard.
    pause
) else (
    echo.
    echo ERROR - check output above.
    pause
)
