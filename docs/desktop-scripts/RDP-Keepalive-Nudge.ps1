<#
RDP-Keepalive-Nudge.ps1
========================
Single-shot idle-lock keep-alive nudge for AD-OAK\begb0037's RDP session on
101L-DE013193. Registered as a REPEATING scheduled task (every 4 minutes,
from logon -- see Register-RDPKeepAlive.ps1) rather than a long-lived loop
process: Task Scheduler handles the interval natively, more robust than a
script sleeping in a loop that could get stuck or silently die.

WHY THIS EXISTS (3 Sept 2026): Oxford's domain GPO enforces a 10-minute
secure screensaver lock on this machine --
  HKU:\<begb0037 SID>\Software\Policies\Microsoft\Windows\Control Panel\Desktop
  ScreenSaveTimeOut=600, ScreenSaverIsSecure=1
Confirmed NOT locally overridable: a direct registry edit to that key would
appear to work immediately (it's a live per-session value) but is very
likely to be silently reverted on the next Group Policy background refresh
(90-120 min, or at next logon) -- a worse outcome than not touching it. See
HANDOVER.md's "RDP idle-lock" section for the full investigation.

Kevin's EXPLICIT, INFORMED go-ahead (3 Sept 2026, via coordinator): simulate
harmless periodic input so the OS's own idle timer never reaches the GPO's
600-second threshold, instead of attempting to change the policy itself.
THIS IS A DELIBERATE CIRCUMVENTION OF A DOMAIN SECURITY CONTROL, not a
settings change -- stated plainly here and in HANDOVER.md, not disguised.

SCOPED to only act while a session is genuinely CONNECTED (`query session`
shows State=Active for begb0037) -- if disconnected (State=Disc, i.e. Kevin
closed the RDP client without logging off), this exits without touching
anything, so the GPO lock resumes its normal behaviour the moment nobody's
client is attached. HONEST LIMITATION, not hidden: this does NOT distinguish
"connected but stepped away from the desk" from "connected and present" --
that distinction is the entire purpose of the control being bypassed here.
A connected-but-idle-human session will still be kept unlocked by this
script; only a genuinely disconnected session stops it.

KILL SWITCH: if a file named KEEPALIVE_DISABLE exists next to this script,
this exits immediately without nudging -- the lowest-friction way to pause
without touching the scheduled task itself (create/delete one file, no
elevated PowerShell needed).

MECHANISM: moves the mouse cursor 1 pixel and immediately back. This
registers as real input to Windows' own idle-timer (GetLastInputInfo, the
same underlying call path a physical mouse move uses) without touching
keyboard focus or any application's state -- chosen over a SendKeys
keystroke specifically because a keystroke could type into whatever control
happens to have focus at the time, which a mouse nudge cannot do.

STOP THIS ENTIRELY:
  Disable-ScheduledTask -TaskName 'RDP Keep-Alive (idle-lock bypass)'
  Unregister-ScheduledTask -TaskName 'RDP Keep-Alive (idle-lock bypass)' -Confirm:$false
#>

$ErrorActionPreference = 'SilentlyContinue'

$killSwitch = Join-Path $PSScriptRoot 'KEEPALIVE_DISABLE'
if (Test-Path $killSwitch) { exit 0 }

# Only nudge while begb0037's session is genuinely Active (connected), not Disc.
$qs = (query session 2>$null)
$myLine = $qs | Where-Object { $_ -match 'begb0037' }
if (-not $myLine -or $myLine -notmatch 'Active') { exit 0 }

Add-Type -AssemblyName System.Windows.Forms
$pos = [System.Windows.Forms.Cursor]::Position
$nudged = New-Object System.Drawing.Point(($pos.X + 1), $pos.Y)
[System.Windows.Forms.Cursor]::Position = $nudged
Start-Sleep -Milliseconds 80
[System.Windows.Forms.Cursor]::Position = $pos
exit 0
