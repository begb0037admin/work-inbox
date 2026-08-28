# Ensure-ClassicOutlook.ps1 -- preflight for "Run Inbox Briefing.bat"
# Added 2026-08-28 (Drew).
#
# fetch_inbox.py's Phase 1 pull needs CLASSIC Outlook (OUTLOOK.EXE) running
# and connected to Exchange -- it automates it over COM. New Outlook
# (olk.exe) has NO COM interface and cannot stand in for it.
#
# A 13:30 reboot on 28 Aug 2026 left classic Outlook closed and every
# scheduled briefing run failed at Phase 1 with pywintypes.com_error
# -2147352567 / inner -2147221231 ("The file <profile>.ost cannot be
# accessed. You must connect to Microsoft Exchange at least once...").
#
# This starts classic Outlook if it is not running and waits for MAPI to
# become usable before fetch_inbox.py runs. Best-effort ONLY: it always
# exits 0 so it can never fail the briefing run -- fetch_inbox.py has its
# own matching detection, auto-launch, and a specific desktop toast.

$ErrorActionPreference = 'SilentlyContinue'
function plog($m) { Write-Output ("[{0}] PREFLIGHT {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $m) }

$classic = Get-Process OUTLOOK -ErrorAction SilentlyContinue
if ($classic) {
    plog "classic Outlook already running (pid $($classic.Id))."
} else {
    if (Get-Process olk -ErrorAction SilentlyContinue) {
        plog "WARNING: New Outlook (olk.exe) is running but CLASSIC Outlook is not. The pipeline needs classic Outlook -- New Outlook has no COM interface."
    }
    $exe = @(
        'C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE',
        'C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE',
        'C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE',
        'C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE'
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $exe) {
        plog "WARNING: could not locate OUTLOOK.EXE to start."
        exit 0
    }
    plog "starting classic Outlook -- $exe"
    # Launch via explorer so Outlook starts under the shell, NOT inside this
    # task's job object (Task Scheduler tears that down, killing child
    # processes, the moment the task's action process exits).
    Start-Process explorer.exe -ArgumentList "`"$exe`""
}

# Wait for MAPI to be genuinely usable (Exchange store mounted), up to ~120s.
$deadline = (Get-Date).AddSeconds(120)
$ready = $false
while ((Get-Date) -lt $deadline) {
    Start-Sleep -Seconds 5
    try {
        $o = New-Object -ComObject Outlook.Application
        $null = $o.GetNamespace('MAPI').GetDefaultFolder(6).Items.Count
        $ready = $true
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch { }
        break
    } catch { }
}
if ($ready) {
    plog "classic Outlook MAPI is ready."
} else {
    plog "WARNING: classic Outlook did not become MAPI-ready within 120s. It may be waiting on an interactive Windows Security / Oxford sign-in prompt that only Kevin can complete. fetch_inbox.py will retry and, if still stuck, raise a specific toast."
}
exit 0
