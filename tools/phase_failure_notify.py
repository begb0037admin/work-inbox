"""
phase_failure_notify.py
------------------------
Shared desktop-toast helper for the standalone `tools/` pipeline scripts
(publish_needs_reply.py, publish_drafted_replies.py). Both already catch
their own top-level exception in `__main__` and print "FATAL: ..." plus
exit 1 -- but that's log-only. The scheduled chain (`Run Inbox Briefing
Hidden.vbs` -> fetch_inbox.py -> publish_needs_reply.py ->
publish_drafted_replies.py) treats a non-zero exit from either downstream
step as non-fatal to the overall briefing run, so today a real failure here
produces nothing Kevin will actually see -- the exact same silent-failure
gap fetch_inbox.py's own Phase 3.6 had before 20 Aug 2026 (commit
ab1f6bb4). See work-inbox HANDOVER.md, Phase 2 (20-21 Aug 2026), for the
full stability-plan context.

This reuses that exact same notification mechanism (Show-TaskNotification.ps1
/ BurntToast) instead of inventing a second one. Deliberately NOT imported
by fetch_inbox.py itself -- that script's own `_notify_phase_failure` stays
inline and untouched, to avoid any risk to the already-shipped, already
live-verified fix. This module exists purely so the two NEWER standalone
tools/ scripts don't each hand-roll a third near-duplicate copy of the same
~20 lines.

Writes a small dedicated one-line detail file per caller (via
`log_filename`) rather than pointing Show-TaskNotification.ps1 at either
script's own last-run log -- same reasoning as fetch_inbox.py's version:
a shared/growing log's content isn't guaranteed to match
Get-LogTailDetail's regex, and a dedicated file makes the detail text
deterministic regardless of what runs afterward or in parallel.

Best-effort only, by design: a failure to raise the toast must never mask
or replace the original exception already being handled by the caller's
own except block, and must never itself crash the run.
"""

import os
import subprocess
from datetime import datetime

NOTIFY_SCRIPT_PATH = r"D:\OneDrive - lelitte.com\Desktop\Show-TaskNotification.ps1"


def notify_phase_failure(task_name, detail, log_filename="phase_failure_last.log"):
    """Write a one-line detail file next to this module and fire a real
    desktop toast via Show-TaskNotification.ps1 (BurntToast). Mirrors
    fetch_inbox.py's own _notify_phase_failure exactly -- see the module
    docstring above for why this is a separate shared copy rather than a
    direct import from fetch_inbox.py."""
    try:
        detail_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), log_filename)
        with open(detail_path, "w", encoding="utf-8") as f:
            f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {task_name} failed: {detail}\n")
        if os.path.exists(NOTIFY_SCRIPT_PATH):
            subprocess.run(
                ["powershell.exe", "-NoProfile", "-WindowStyle", "Hidden", "-File", NOTIFY_SCRIPT_PATH,
                 "-Status", "Failure", "-TaskName", task_name, "-LogPath", detail_path],
                timeout=20, capture_output=True
            )
        else:
            print(f"WARNING: failure toast skipped for '{task_name}' - notification script not found at {NOTIFY_SCRIPT_PATH}")
    except Exception as notify_err:
        print(f"WARNING: failure toast for '{task_name}' could not be sent - {notify_err}")
