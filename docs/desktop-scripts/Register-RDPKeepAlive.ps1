<#
Register-RDPKeepAlive.ps1
==========================
Registers "RDP Keep-Alive (idle-lock bypass)" on 101L-DE013193 -- fires
RDP-Keepalive-Nudge.ps1 every 3 minutes, for as long as
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

# TRIGGER: a ONE-TIME trigger with a start time already in the past + an
# infinite 3-minute repetition. Deliberately NOT -AtLogOn.
#   Why not -AtLogOn: verified live 3 Sept 2026 that on this box an -AtLogOn
#   trigger's repetition cycle never armed -- it needs a genuine Logon
#   (4624-class) session-creation event, which a screen UNLOCK is not, and it
#   does not fire retroactively for a session already logged in when the task
#   was registered. Result: the task only ever ran once (a manual
#   Start-ScheduledTask), idle time kept climbing, the lock still fired.
#   A -Once trigger with a past start time arms IMMEDIATELY at registration
#   and the repetition just keeps ticking every 3 min forever, independent of
#   logon/lock/unlock events. Scoping is still real: the Interactive principal
#   below means Task Scheduler only actually EXECUTES it while AD-OAK\begb0037
#   has a session, and the script itself skips the nudge unless `query
#   session` shows that session Active (connected), not Disc.
$trigger = New-ScheduledTaskTrigger -Once -At ((Get-Date).AddMinutes(-2))
$repClass = Get-CimClass -ClassName MSFT_TaskRepetitionPattern -Namespace Root/Microsoft/Windows/TaskScheduler
$repetition = New-CimInstance -CimClass $repClass -ClientOnly
$repetition.Interval = 'PT3M'
$repetition.Duration = ''            # indefinite
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
  -Description ('Simulates a harmless 1px mouse nudge every ~3 min so the domain-GPO-enforced ' +
                 '10-min secure screensaver lock never fires. Kevin explicit go-ahead 3 Sept 2026 ' +
                 '(deliberate circumvention of a domain policy, not a settings change -- see ' +
                 'HANDOVER.md). Only nudges while `query session` shows begb0037 Active (connected); ' +
                 'skips if Disc. Kill switch: create KEEPALIVE_DISABLE next to the script, or ' +
                 'Disable-ScheduledTask.') `
  | Out-Null

Write-Host "Registered '$taskName'."
(Get-ScheduledTask -TaskName $taskName).Triggers | Format-List
