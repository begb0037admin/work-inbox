' Run Classic Outlook Keepalive Hidden.vbs
' Launches Ensure-ClassicOutlook.ps1 with a fully hidden window.
' Added 2026-08-28 (Drew). Same hidden-launch pattern as
' "Run Inbox Briefing Hidden.vbs" -- WScript.Shell.Run windowStyle=0 is the
' only reliable way to run the keepalive every 10 minutes with no window
' flash. Fire-and-forget (third arg False) is fine here: the keepalive is
' best-effort and Task Scheduler's ExecutionTimeLimit still bounds it.
Set objShell = CreateObject("WScript.Shell")
objShell.CurrentDirectory = "D:\OneDrive - lelitte.com\Desktop"
objShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -File ""D:\OneDrive - lelitte.com\Desktop\Ensure-ClassicOutlook.ps1""", 0, False
