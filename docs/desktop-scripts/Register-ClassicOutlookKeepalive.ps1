# Register-ClassicOutlookKeepalive.ps1
# Creates the "Classic Outlook Keepalive" scheduled task:
#   - runs at logon, then repeats every 10 minutes, indefinitely
#   - runs as the current interactive user (Outlook needs a desktop session)
#   - action = Run Classic Outlook Keepalive Hidden.vbs -> Ensure-ClassicOutlook.ps1
# Idempotent: unregisters any existing task of the same name first.
# PowerShell 5.1 compatible. Does NOT require elevation for a per-user task.
#
# Restore / removal:
#   Unregister-ScheduledTask -TaskName 'Classic Outlook Keepalive' -Confirm:$false
#
# Added 2026-08-28 (Drew).

$ErrorActionPreference = 'Stop'
$taskName = 'Classic Outlook Keepalive'
$vbs      = 'D:\OneDrive - lelitte.com\Desktop\Run Classic Outlook Keepalive Hidden.vbs'

if (-not (Test-Path $vbs)) { throw "wrapper not found: $vbs" }

if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    Write-Output "removed existing task '$taskName'"
}

$action = New-ScheduledTaskAction -Execute 'wscript.exe' -Argument ('"{0}"' -f $vbs)

$tLogon = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
$tRepeat = New-ScheduledTaskTrigger -Once -At (Get-Date).Date.AddMinutes(1) `
            -RepetitionInterval (New-TimeSpan -Minutes 10) `
            -RepetitionDuration (New-TimeSpan -Days 3650)

$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 5) `
    -RestartCount 0

$principal = New-ScheduledTaskPrincipal -UserId ("{0}\{1}" -f $env:USERDOMAIN, $env:USERNAME) `
    -LogonType Interactive -RunLevel Limited

Register-ScheduledTask -TaskName $taskName -Action $action `
    -Trigger @($tLogon, $tRepeat) -Settings $settings -Principal $principal `
    -Description 'Keeps classic Outlook (OUTLOOK.EXE) running + MAPI-ready for the work-inbox pipelines. Relaunches it if closed / only New Outlook is up; raises a desktop toast (1/hour) if it is stuck on an interactive Oxford sign-in. Added 2026-08-28 (Drew).' | Out-Null

Write-Output "registered '$taskName'"
Get-ScheduledTask -TaskName $taskName | Select-Object TaskName, State
(Get-ScheduledTask -TaskName $taskName).Triggers | Format-List
