@echo off
title Inbox Briefing - Manual Refresh
cd /d C:\Users\admin\Documents\Claude\Projects\work-inbox
echo Downloading latest script from GitHub...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$t=[DateTimeOffset]::UtcNow.ToUnixTimeSeconds(); Invoke-WebRequest -UseBasicParsing ('https://raw.githubusercontent.com/begb0037admin/work-inbox/main/fetch_inbox.py?t='+$t) -OutFile 'fetch_inbox.py'"
if errorlevel 1 (
    echo WARNING - could not download latest script. Running existing local copy.
)
rem Connector-only, standing rule (Kevin, 15 Sep 2026): COM and IMAP are
rem retired as live options on every machine. Never leave these unset --
rem fetch_inbox.py defaults to MAIL_BACKEND=com / CAL_BACKEND=com /
rem TEAMS_BACKEND=off when unset, which is exactly the fallback Kevin wants
rem gone for good.
set MAIL_BACKEND=connector
set CAL_BACKEND=connector
set TEAMS_BACKEND=connector
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
