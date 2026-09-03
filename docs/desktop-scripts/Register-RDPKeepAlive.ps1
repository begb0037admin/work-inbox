<#
Register-RDPKeepAlive.ps1
==========================
Registers "RDP Keep-Alive (idle-lock bypass)" on 101L-DE013193 -- fires
RDP-Keepalive-Nudge.ps1 every 4 minutes from logon, for as long as
AD-OAK\begb0037 stays logged on. See that script's own header for the full
why / scoping / kill-switch / rollback -- this file only registers it.

Run ONCE. Can be run as begb0037-a (local admin) targeting the
AD-OAK\begb0037 principal, same pattern already used for this laptop's other
Interactive scheduled tasks (Bridge Briefing, Parity Shadow, etc.) -- admin
can register a task FOR another local/domain user without being that user.

Unregister:  Unregister-ScheduledTask -TaskName 'RDP Keep-Alive (idle-lock bypass)' -Confirm:$false
Pause without unregistering:  Disable-ScheduledTask -TaskName 'RDP Keep-Alive (idle-lock bypass)'
#>
$ErrorActionPreference = 'Stop'
$taskName = 'RDP Keep-Alive (idle-lock bypass)'
$root     = 'C:\Users\begb0037.AD-OAK\work-inbox'
$script   = Join-Path $root 'RDP-Keepalive-Nudge.ps1'

if (-not (Test-Path $script)) {
  throw "script not found: $script -- copy RDP-Keepalive-Nudge.ps1 into $root first"
}

$action = New-ScheduledTaskAction -Execute 'powershell.exe' `
  -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$script`""

$trigger = New-ScheduledTaskTrigger -AtLogOn -User 'AD-OAK\begb0037'
# New-ScheduledTaskTrigger's returned object doesn't have a settable
# Repetition CIM instance by default -- build one explicitly and assign it
# (the naive "$trigger.Repetition.Interval = ..." fails with
# PropertyNotFound because $trigger.Repetition starts out $null).
$repClass = Get-CimClass -ClassName MSFT_TaskRepetitionPattern -Namespace Root/Microsoft/Windows/TaskScheduler
$repetition = New-CimInstance -CimClass $repClass -ClientOnly
$repetition.Interval = 'PT4M'
$repetition.Duration = ''   # indefinite -- repeats for the life of the logon
                             # session (this trigger re-fires fresh at every new logon anyway)
$repetition.StopAtDurationEnd = $false
$trigger.Repetition = $repetition

$principal = New-ScheduledTaskPrincipal -UserId 'AD-OAK\begb0037' -LogonType Interactive -RunLevel Limited

$settings = New-ScheduledTaskSettingsSet `
  -MultipleInstances IgnoreNew `
  -ExecutionTimeLimit (New-TimeSpan -Seconds 30) `
  -DontStopOnIdleEnd `
  -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger `
  -Principal $principal -Settings $settings -Force `
  -Description ('Simulates a harmless 1px mouse nudge every ~4 min so the domain-GPO-enforced ' +
                 '10-min secure screensaver lock never fires. Kevin explicit go-ahead 3 Sept 2026 ' +
                 '(deliberate circumvention of a domain policy, not a settings change -- see ' +
                 'HANDOVER.md). Only nudges while `query session` shows begb0037 Active (connected); ' +
                 'skips if Disc. Kill switch: create KEEPALIVE_DISABLE next to the script, or ' +
                 'Disable-ScheduledTask.') `
  | Out-Null

Write-Host "Registered '$taskName'."
(Get-ScheduledTask -TaskName $taskName).Triggers | Format-List
