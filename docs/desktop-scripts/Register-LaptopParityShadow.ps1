# Register-LaptopParityShadow.ps1
# work-inbox laptop migration, Phase 4. Registers the scheduled task
# "Work Inbox Laptop Parity Shadow" on Kevin's Oxford laptop.
#
# Runs "Run Laptop Parity Shadow.ps1" at 07:00 / 09:00 / 11:00 / 13:00 / 15:00 /
# 17:00 Mon-Fri -- matching the live desktop "Work Inbox Briefing" cadence -- as
# ad-oak\begb0037 (the PRT-holding standard user; NOT begb0037-a).
#
# PARALLEL ONLY: the wrapper writes data\parallel\* + logs\parity_shadow.log,
# never pushes, never touches data\briefing.json, never opens classic Outlook.
#
# Run this ONCE, in a normal (non-elevated) PowerShell as ad-oak\begb0037.
# Unregister:  Unregister-ScheduledTask -TaskName 'Work Inbox Laptop Parity Shadow' -Confirm:$false

$ErrorActionPreference = 'Stop'
$taskName = 'Work Inbox Laptop Parity Shadow'
$root     = Join-Path $env:USERPROFILE 'work-inbox'
$wrapper  = Join-Path $root 'Run Laptop Parity Shadow.ps1'

if (-not (Test-Path $wrapper)) {
  throw "wrapper not found: $wrapper  -- copy 'Run Laptop Parity Shadow.ps1' into $root first"
}

$action = New-ScheduledTaskAction -Execute 'powershell.exe' `
  -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$wrapper`""

$times = '07:00','09:00','11:00','13:00','15:00','17:00'
$triggers = foreach ($tm in $times) {
  New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At $tm
}

# Interactive = runs only when ad-oak\begb0037 is logged on. The laptop stays
# docked + logged in (Kevin's commitment), and this keeps the MSAL broker/WAM
# silent-token path (proven in Phase 2(i)) available. If you ever need it to run
# with no session, change -LogonType to S4U and re-test the silent token.
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" `
  -LogonType Interactive -RunLevel Limited

$settings = New-ScheduledTaskSettingsSet `
  -MultipleInstances IgnoreNew `
  -ExecutionTimeLimit (New-TimeSpan -Minutes 15) `
  -StartWhenAvailable `
  -DontStopOnIdleEnd `
  -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $triggers `
  -Principal $principal -Settings $settings -Force `
  -Description 'work-inbox laptop migration Phase 4 shadow: IMAP mail capture + parity vs live briefing.json. Parallel only - writes data\parallel\*, never pushes.'

Write-Host "Registered '$taskName'. Triggers:"
(Get-ScheduledTask -TaskName $taskName).Triggers | Format-Table -AutoSize
Write-Host "`nTest now:  Start-ScheduledTask -TaskName '$taskName'  ; then check %USERPROFILE%\work-inbox\logs\parity_shadow.log"

# never-sleep so the scheduled slots always fire
powercfg /change standby-timeout-ac 0
powercfg /change hibernate-timeout-ac 0
Write-Host "powercfg: AC standby + hibernate timeouts set to 0 (never)."
