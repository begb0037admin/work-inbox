@echo off
title Inbox Briefing - Manual Refresh
cd /d C:\Users\admin\Documents\Claude\Projects\work-inbox
git config gc.auto 0
echo Pulling latest script from GitHub...
git fetch origin
git checkout origin/main -- fetch_inbox.py
echo.
echo Updating inbox briefing...
echo.
python fetch_inbox.py
if %errorlevel% equ 0 (
    echo.
    echo Done. Dashboard will show updated data on next open.
    timeout /t 4 /nobreak >nul
) else (
    echo.
    echo ERROR - check output above.
    pause
)
