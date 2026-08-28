# Unregister-ClassicOutlookKeepalive.ps1
# Removes the "Classic Outlook Keepalive" scheduled task. The Desktop
# scripts (Ensure-ClassicOutlook.ps1, the VBS wrapper) are left in place;
# they are harmless without the task and are still used as the .bat
# preflight. Added 2026-08-28 (Drew).
$ErrorActionPreference = 'Stop'
$taskName = 'Classic Outlook Keepalive'
if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    Write-Output "removed scheduled task '$taskName'"
} else {
    Write-Output "no scheduled task '$taskName' found - nothing to do"
}
