<#
Register-BridgeBriefingToastWatcher.ps1
======================================
Registers the TEMPORARY desktop task "Work Inbox Briefing Toast Watcher" on
DESKTOP-MJDJM64 (user admin). Every 5 minutes it runs Watch-BridgeBriefing.ps1,
which toasts when GitHub gets a new "chore: update briefing" commit for
data/briefing.json (produced by the laptop bridge). Independent of the desktop's
own "Work Inbox Briefing" task (which stays Disabled during the bridge).

Run ONCE in a normal (non-elevated) PowerShell as the desktop 'admin' user.

Expects the live script at:
  D:\OneDrive - lelitte.com\Desktop\Watch-BridgeBriefing.ps1

UNREGISTER (end of bridge):
  Unregister-ScheduledTask -TaskName 'Work Inbox Briefing Toast Watcher' -Confirm:$false
#>
$ErrorActionPreference = 'Stop'
$taskName = 'Work Inbox Briefing Toast Watcher'
$live     = 'D:\OneDrive - lelitte.com\Desktop\Watch-BridgeBriefing.ps1'

if (-not (Test-Path $live)) {
  throw "live script not found: $live  --  download it there first (see Drew's handover)."
}

$action = New-ScheduledTaskAction -Execute 'powershell.exe' `
  -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$live`""

# every 5 minutes, indefinitely
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date) `
  -RepetitionInterval (New-TimeSpan -Minutes 5) -RepetitionDuration (New-TimeSpan -Days 3650)

$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" `
  -LogonType Interactive -RunLevel Limited

$settings = New-ScheduledTaskSettingsSet `
  -MultipleInstances IgnoreNew `
  -ExecutionTimeLimit (New-TimeSpan -Minutes 3) `
  -StartWhenAvailable `
  -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
  -DontStopOnIdleEnd

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger `
  -Principal $principal -Settings $settings -Force `
  -Description 'TEMPORARY bridge-period watcher: toasts on a new GitHub "chore: update briefing" commit while the work-inbox pipeline runs on the laptop. Read-only GitHub poll, no Outlook/M365. Remove when the pipeline has a permanent home.'

Write-Host "Registered '$taskName' (every 5 min, as $env:USERDOMAIN\$env:USERNAME, run-only-when-logged-on)."
Write-Host "Smoke test now:  Start-ScheduledTask -TaskName '$taskName'  ;  Get-Content `"$env:LOCALAPPDATA\WorkInboxAI\bridge_toast_watcher.log`" -Tail 5"
Write-Host "(first run seeds the last-seen SHA silently; the next NEW briefing commit produces the first toast.)"
