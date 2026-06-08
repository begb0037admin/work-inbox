@echo off
schtasks /run /tn "WorkInbox-Briefing"
echo Briefing run triggered. Wait 90 seconds then refresh the dashboard.
pause
