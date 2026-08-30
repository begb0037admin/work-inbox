<#
Register-LaptopBridgeBriefing.ps1
================================
Registers the TEMPORARY scheduled task "Work Inbox Bridge Briefing" on Kevin's
Oxford laptop (101L-DE013193 / begb0037.AD-OAK) for the duration of the desktop
M365 outage. It runs "Run Laptop Bridge Briefing.ps1" (full pipeline: IMAP mail
pull -> claude -p triage -> Phase 4 GitHub push -> Phase 5 command-centre sync
-> best-effort publishers). No calendar.

RUN THIS ONCE, in a normal (NON-elevated) PowerShell, signed in as ad-oak\begb0037
(the PRT-holding standard user -- NOT begb0037-a).

  -Cadence Bridge   (default)  09:00 / 12:00 / 15:00 Mon-Fri   (3x/day -- safer for a single Pro account)
  -Cadence Full                06:00 / 09:00 / 12:00 / 15:00 / 18:00 Mon-Fri   (matches the old desktop task)

LogonType = Interactive: the task only runs while ad-oak\begb0037 is logged on.
That is required -- the MSAL broker silent-token path and the periodic one-click
browser re-auth both need the interactive session. Keep the laptop docked + logged in.

UNREGISTER (end of bridge):
  Unregister-ScheduledTask -TaskName 'Work Inbox Bridge Briefing' -Confirm:$false
Then on the admin desktop:
  Enable-ScheduledTask -TaskName 'Work Inbox Briefing'
#>
param(
  [ValidateSet('Bridge','Full')] [string]$Cadence = 'Bridge'
)

$ErrorActionPreference = 'Stop'
$taskName = 'Work Inbox Bridge Briefing'
$root     = Join-Path $env:USERPROFILE 'work-inbox'
$wrapper  = Join-Path $root 'Run Laptop Bridge Briefing.ps1'

if (-not (Test-Path $wrapper)) {
  throw "wrapper not found: $wrapper  --  copy 'Run Laptop Bridge Briefing.ps1' into $root first (see the download line in Drew's handover)."
}
if (-not (Test-Path 'C:\WorkInboxAI\kevin\.credentials.json')) {
  throw "C:\WorkInboxAI\kevin\.credentials.json missing -- kevin@ isolated Claude config not logged in. Fix that before registering (the task would fail every run)."
}

$times = if ($Cadence -eq 'Full') { '06:00','09:00','12:00','15:00','18:00' } else { '09:00','12:00','15:00' }

$action = New-ScheduledTaskAction -Execute 'powershell.exe' `
  -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$wrapper`""

$triggers = foreach ($tm in $times) {
  New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At $tm
}

$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" `
  -LogonType Interactive -RunLevel Limited

$settings = New-ScheduledTaskSettingsSet `
  -MultipleInstances IgnoreNew `
  -ExecutionTimeLimit (New-TimeSpan -Minutes 20) `
  -StartWhenAvailable `
  -DontStopOnIdleEnd `
  -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $triggers `
  -Principal $principal -Settings $settings -Force `
  -Description "TEMPORARY work-inbox bridge during the desktop M365 outage. Real mail-only briefing: IMAP pull -> claude -p triage -> Phase 4 push -> Phase 5 command-centre sync. No calendar. Unregister + re-enable the desktop 'Work Inbox Briefing' task when the desktop is fixed."

Write-Host "Registered '$taskName'  (cadence: $Cadence -> $($times -join ', ') Mon-Fri, as $env:USERDOMAIN\$env:USERNAME, run-only-when-logged-on)."
(Get-ScheduledTask -TaskName $taskName).Triggers | Format-Table -AutoSize
Write-Host ""
Write-Host "Smoke test now:  Start-ScheduledTask -TaskName '$taskName'"
Write-Host "Then watch:      Get-Content `"$root\logs\bridge_briefing_last_run.log`" -Wait"

# never-sleep so the scheduled slots always fire while docked
powercfg /change standby-timeout-ac 0
powercfg /change hibernate-timeout-ac 0
Write-Host "powercfg: AC standby + hibernate timeouts set to 0 (never)."
