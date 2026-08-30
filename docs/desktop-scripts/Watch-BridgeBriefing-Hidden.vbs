' Watch-BridgeBriefing-Hidden.vbs
' ===============================
' Hidden launcher for the "Work Inbox Briefing Toast Watcher" scheduled task
' (DESKTOP-MJDJM64 / admin). Task Scheduler's own "Hidden" property AND
' powershell.exe -WindowStyle Hidden do NOT suppress the initial console-window
' flash for a direct powershell.exe task action -- and this task fires every
' 5 minutes, so the flash is constantly on screen. WScript.Shell.Run with
' windowStyle = 0 launches the process (and its children) with no window at all.
' This matches the "Run X Hidden.vbs" pattern every other work-inbox desktop
' task already uses (Run Inbox Briefing Hidden.vbs, Run Draft Diff Capture
' Hidden.vbs, etc.).
'
' Synchronous (Run's 3rd arg = True) + WScript.Quit(rc): keeps the task
' instance alive until the poll finishes (Task Scheduler tears down the job
' object the moment the action process exits, which would kill a fire-and-
' forget child mid-run) and propagates the real exit code to LastTaskResult.
'
' Register-BridgeBriefingToastWatcher.ps1 (re)generates this file next to the
' live .ps1 on every run -- this repo copy is the reference / for inspection.

Set sh = CreateObject("WScript.Shell")
sh.CurrentDirectory = "D:\OneDrive - lelitte.com\Desktop"
rc = sh.Run("powershell.exe -NoProfile -ExecutionPolicy Bypass -File ""D:\OneDrive - lelitte.com\Desktop\Watch-BridgeBriefing.ps1""", 0, True)
WScript.Quit(rc)
