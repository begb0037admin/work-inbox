@echo off
echo Creating Work Inbox Task Scheduler tasks...
echo.

schtasks /create /tn "WorkInbox-0900" /tr "cmd.exe /c \"C:\Users\admin\Desktop\Run Inbox Briefing.bat\"" /sc weekly /d MON,TUE,WED,THU,FRI /st 09:00 /rl highest /f
schtasks /create /tn "WorkInbox-1200" /tr "cmd.exe /c \"C:\Users\admin\Desktop\Run Inbox Briefing.bat\"" /sc weekly /d MON,TUE,WED,THU,FRI /st 12:00 /rl highest /f
schtasks /create /tn "WorkInbox-1500" /tr "cmd.exe /c \"C:\Users\admin\Desktop\Run Inbox Briefing.bat\"" /sc weekly /d MON,TUE,WED,THU,FRI /st 15:00 /rl highest /f

echo.
echo Done. Check above for any errors.
pause
