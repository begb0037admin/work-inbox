# Ensure-ClassicOutlook.ps1
# ---------------------------------------------------------------------------
# Keeps CLASSIC Outlook (OUTLOOK.EXE) running and MAPI-usable so the
# work-inbox pipelines (fetch_inbox.py Phase 1, draft_final_diff_capture.py)
# can automate it over COM. New Outlook (olk.exe) has NO COM interface and
# cannot stand in.
#
# Used in two places, same behaviour, always exits 0:
#   1. Preflight, called by "Run Inbox Briefing.bat" / "Run Draft Diff
#      Capture.bat" just before their Python step.
#   2. Keepalive, run by the "Classic Outlook Keepalive" scheduled task
#      (at logon, then every 10 min) -- see Register-ClassicOutlookKeepalive.ps1.
#
# Health model:
#   - classic Outlook running + a quick MAPI probe succeeds  -> healthy, exit.
#   - classic Outlook not running (or only olk.exe up)        -> launch it.
#   - launched / still starting                               -> poll MAPI up
#     to 120s.
#   - still not MAPI-ready after that                         -> it is almost
#     certainly sitting on an interactive Windows Security / Oxford sign-in
#     prompt that only Kevin can complete -> raise ONE desktop toast (rate-
#     limited to 1/hour) and exit 0.
#
# Background (28 Aug 2026, Drew): a 13:30 reboot left classic Outlook closed;
# every scheduled run failed at its first Outlook COM call. Root cause of the
# *interactive prompt* is that this device is Azure-AD-registered (Workplace
# Joined) with NO Primary Refresh Token, so Office cannot silently renew an
# expired Oxford token -- it must prompt. That recurs periodically (not every
# reboot -- the pipeline survived two reboots on 27 Aug). This script cannot
# fix that prompt; it makes the common cases self-heal and makes the prompt
# case loud instead of silent.
# ---------------------------------------------------------------------------

$ErrorActionPreference = 'SilentlyContinue'

$NotifyScript = 'D:\OneDrive - lelitte.com\Desktop\Show-TaskNotification.ps1'
$StampDir     = Join-Path $env:LOCALAPPDATA 'WorkInboxAI'
$ToastStamp   = Join-Path $StampDir 'classic_outlook_signin_toast.stamp'
$ToastEveryMin = 60
$PollSeconds  = 120

function plog($m) { Write-Output ("[{0}] ClassicOutlook {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $m) }

function Test-MapiReady {
    try {
        $o = New-Object -ComObject Outlook.Application
        $null = $o.GetNamespace('MAPI').GetDefaultFolder(6).Items.Count
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch { }
        return $true
    } catch { return $false }
}

function Get-ClassicExe {
    @(
        'C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE',
        'C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE',
        'C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE',
        'C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE'
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1
}

function Send-SignInToast {
    if (-not (Test-Path $NotifyScript)) { plog "toast skipped - notify script not found."; return }
    try {
        if (Test-Path $ToastStamp) {
            $age = (New-TimeSpan -Start (Get-Item $ToastStamp).LastWriteTime -End (Get-Date)).TotalMinutes
            if ($age -lt $ToastEveryMin) { plog ("toast suppressed - last one {0:N0} min ago." -f $age); return }
        }
        if (-not (Test-Path $StampDir)) { New-Item -ItemType Directory -Path $StampDir -Force | Out-Null }
        $detail = Join-Path $StampDir 'classic_outlook_signin.log'
        "[{0}] Classic Outlook is open but not connected to Exchange - it is waiting on an interactive Windows Security / Oxford sign-in prompt. Click it, sign in, approve MFA, confirm 'Connected to: Microsoft Exchange', and leave Outlook running. The work-inbox briefing + Draft Diff Capture cannot run until this is done." -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | Set-Content -Path $detail -Encoding UTF8
        powershell.exe -NoProfile -WindowStyle Hidden -File $NotifyScript -Status Failure -TaskName 'Classic Outlook needs sign-in' -LogPath $detail | Out-Null
        (Get-Date).ToString('o') | Set-Content -Path $ToastStamp
        plog "sign-in toast raised."
    } catch { plog ("toast error: " + $_.Exception.Message) }
}

# --- 1. healthy fast path -------------------------------------------------
$classic = Get-Process OUTLOOK -ErrorAction SilentlyContinue
if ($classic -and (Test-MapiReady)) {
    plog ("healthy - classic Outlook running (pid {0}) and MAPI ready." -f $classic.Id)
    if (Test-Path $ToastStamp) { Remove-Item $ToastStamp -ErrorAction SilentlyContinue }
    exit 0
}

# --- 2. launch if needed ------------------------------------------------
if (-not $classic) {
    if (Get-Process olk -ErrorAction SilentlyContinue) {
        plog "WARNING: New Outlook (olk.exe) is running but CLASSIC Outlook is not. The pipeline needs classic Outlook - New Outlook has no COM interface."
    }
    $exe = Get-ClassicExe
    if (-not $exe) { plog "WARNING: could not locate OUTLOOK.EXE to start."; exit 0 }
    plog ("classic Outlook not running - starting " + $exe)
    # Launch via explorer so Outlook starts under the shell, NOT inside a
    # Task Scheduler job object (torn down, killing children, at task end).
    Start-Process explorer.exe -ArgumentList "`"$exe`""
} else {
    plog ("classic Outlook running (pid {0}) but MAPI not ready yet - waiting." -f $classic.Id)
}

# --- 3. poll MAPI readiness -------------------------------------------
$deadline = (Get-Date).AddSeconds($PollSeconds)
$ready = $false
while ((Get-Date) -lt $deadline) {
    Start-Sleep -Seconds 5
    if (Test-MapiReady) { $ready = $true; break }
}

if ($ready) {
    plog "classic Outlook MAPI is now ready."
    if (Test-Path $ToastStamp) { Remove-Item $ToastStamp -ErrorAction SilentlyContinue }
} else {
    plog ("WARNING: classic Outlook did not become MAPI-ready within {0}s - almost certainly waiting on an interactive Windows Security / Oxford sign-in prompt that only Kevin can complete." -f $PollSeconds)
    Send-SignInToast
}
exit 0
