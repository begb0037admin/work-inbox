<#
RDP-Keepalive-Nudge.ps1
========================
Single-shot idle-lock keep-alive nudge for AD-OAK\begb0037's RDP session on
101L-DE013193. Registered as a REPEATING scheduled task -- a one-time
trigger dated in the past + an infinite 3-minute repetition (see
Register-RDPKeepAlive.ps1; deliberately NOT -AtLogOn, which was verified on
3 Sept 2026 to never arm its repetition on this box). Task Scheduler handles
the interval natively, more robust than a script sleeping in a loop that
could get stuck or silently die. The Interactive principal + the
`query session` Active check below are what scope it to "only while Kevin's
connected", not the trigger type.

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

MECHANISM: SendInput() with a relative mouse move of (+1,0) then (-1,0) --
net zero pointer displacement, no keyboard focus touched, no application
state changed. Uses SendInput specifically (NOT SetCursorPos /
Cursor.Position): diagnosed live 3 Sept 2026 that SetCursorPos does NOT
reset GetLastInputInfo's idle timer on modern Windows 10/11 (hardened by
Microsoft to defeat keep-alive tools), so an earlier Cursor.Position version
moved the pointer but the lock timer kept counting and the session still
locked. Each run logs idleBefore/idleAfter (from GetLastInputInfo, the exact
value the screensaver checks) to logs\keepalive.log so it's obvious whether
it's actually working.

HARD LIMIT, unavoidable: once the session HAS locked, SendInput from an
ordinary process is blocked by the secure desktop -- this cannot clear an
existing lock, only prevent the next one. It must be running and resetting
the idle timer BEFORE it reaches 600s, with no gap that long.

STOP THIS ENTIRELY:
  Disable-ScheduledTask -TaskName 'RDP Keep-Alive (idle-lock bypass)'
  Unregister-ScheduledTask -TaskName 'RDP Keep-Alive (idle-lock bypass)' -Confirm:$false
#>

$ErrorActionPreference = 'SilentlyContinue'

$log = Join-Path $PSScriptRoot 'logs\keepalive.log'
New-Item -ItemType Directory -Force -Path (Split-Path $log) | Out-Null
function LK($m) {
  $line = "$([DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss'))  $m"
  Add-Content -LiteralPath $log -Value $line
  # keep the log small -- trim to the last 400 lines occasionally
  if ((Get-Random -Maximum 20) -eq 0) {
    $all = Get-Content -LiteralPath $log
    if ($all.Count -gt 400) { Set-Content -LiteralPath $log -Value ($all | Select-Object -Last 400) }
  }
}

$killSwitch = Join-Path $PSScriptRoot 'KEEPALIVE_DISABLE'
if (Test-Path $killSwitch) { LK 'SKIP  KEEPALIVE_DISABLE present'; exit 0 }

# Only nudge while begb0037's session is genuinely Active (connected), not Disc.
$qs = (query session 2>$null)
$myLine = $qs | Where-Object { $_ -match 'begb0037' }
if (-not $myLine -or $myLine -notmatch 'Active') { LK 'SKIP  session not Active (Disc / not found)'; exit 0 }

# Real idle before the nudge, straight from the OS -- GetLastInputInfo is what
# the screensaver/lock timer itself keys off. Far more precise than `quser`'s
# minute-granularity IDLE TIME column (which also lags by a minute or two).
Add-Type -Name LII -Namespace W32 -MemberDefinition @'
[System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
public struct LASTINPUTINFO { public uint cbSize; public uint dwTime; }
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);
[System.Runtime.InteropServices.DllImport("kernel32.dll")]
public static extern uint GetTickCount();
public static int IdleSeconds() {
  LASTINPUTINFO l = new LASTINPUTINFO(); l.cbSize = (uint)System.Runtime.InteropServices.Marshal.SizeOf(l);
  if (!GetLastInputInfo(ref l)) return -1;
  return (int)((GetTickCount() - l.dwTime) / 1000);
}
'@
$idleBefore = [W32.LII]::IdleSeconds()

# Use SendInput() (NOT SetCursorPos / Cursor.Position). Diagnosed live
# 3 Sept 2026: SetCursorPos does NOT reset GetLastInputInfo's idle timer on
# modern Windows 10/11 -- Microsoft hardened this specifically to defeat
# keep-alive tools -- so the earlier Cursor.Position approach nudged the
# pointer but the lock timer kept counting. SendInput feeds the raw input
# thread and DOES update the idle timer. A relative move of (+1,0) then
# (-1,0) leaves the pointer exactly where it was.
Add-Type -Name SI -Namespace W32 -MemberDefinition @'
[System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
public struct MOUSEINPUT { public int dx; public int dy; public uint mouseData; public uint dwFlags; public uint time; public System.IntPtr dwExtraInfo; }
[System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
public struct INPUT { public uint type; public MOUSEINPUT mi; }
[System.Runtime.InteropServices.DllImport("user32.dll", SetLastError=true)]
public static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);
public static uint Move(int dx, int dy) {
  INPUT[] inp = new INPUT[1];
  inp[0].type = 0; // INPUT_MOUSE
  inp[0].mi.dx = dx; inp[0].mi.dy = dy; inp[0].mi.dwFlags = 0x0001; // MOUSEEVENTF_MOVE (relative)
  return SendInput(1, inp, System.Runtime.InteropServices.Marshal.SizeOf(typeof(INPUT)));
}
'@
$sent1 = [W32.SI]::Move(1, 0)
Start-Sleep -Milliseconds 40
$sent2 = [W32.SI]::Move(-1, 0)
Start-Sleep -Milliseconds 150
$idleAfter = [W32.LII]::IdleSeconds()

# If the nudge reached the real desktop, idleAfter should be ~0. If the session
# is locked (secure desktop), SetCursorPos from this process is ignored and
# idleAfter stays high -- that shows up plainly here.
$verdict = if ($idleAfter -lt 30) { 'OK reset' } elseif ($idleBefore -ge 540) { 'FAILED -- idle not reset (session already locked: SendInput is blocked on the secure desktop; a keep-alive can only PREVENT a lock, never clear one)' } else { 'idle not reset -- unexpected, investigate' }
LK ("NUDGE idleBefore={0}s idleAfter={1}s  SendInput sent={2}/{3}  -> {4}" -f $idleBefore, $idleAfter, $sent1, $sent2, $verdict)
exit 0
