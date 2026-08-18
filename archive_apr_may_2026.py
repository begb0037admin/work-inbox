"""
archive_apr_may_2026.py -- one-off Outlook COM archiving tool.

Built 18 Aug 2026 (Drew), per Kevin's explicit request: move every email
currently sitting in the live Inbox with a ReceivedTime in April 2026 or
May 2026 into the classic-Outlook Archive folder. June, July and August
2026 must be left untouched. No other folder is read or written. Nothing
is ever deleted, only moved.

SCOPE EXPANDED same day, still before any execution: Kevin also wants
everything in the Inbox older than 1 April 2026 (no lower bound -- all
older mail) archived in the same operation. The move window is therefore
"everything with ReceivedTime before 1 June 2026" -- June 2026 onward stays
in the Inbox. The report below still breaks the count down into
pre-April-2026 / April-2026 / May-2026 subtotals plus a combined total, so
each piece stays individually auditable even though they move together.

DRY-RUN BY DEFAULT. Pass --execute to actually move anything, and only
after Kevin has reviewed the dry-run report and given explicit go-ahead.

Reuses connect_to_outlook() byte-for-byte in spirit from fetch_inbox.py
(begb0037admin/work-inbox) -- late-bound Dispatch + GetNamespace + first
GetDefaultFolder(6) call wrapped in a 3-attempt/45s retry, because Outlook's
COM layer has twice (11 Aug 2026) rejected the very first call of a run
with a transient "Call was rejected by callee" com_error. See
begb0037admin/drew memory/outlook-com-connection-retry.md.

Deliberately does NOT use Items.Restrict() for the date filter. fetch_inbox.py's
restrict_date() (see its docstring, ~line 229) proved live on 12 Aug 2026 that
Outlook COM's Restrict() parses a filter string's embedded date using this
machine's UK locale (dd/mm) regardless of what order the string itself
writes the numbers in -- a mm/dd/yyyy-formatted cutoff was silently misread
as dd/mm and shifted the effective date bound by months, with Restrict()
still "succeeding" (no exception, a plausible Count). For a real,
hard-to-reverse mailbox move, that whole class of bug is not worth
re-risking just to use Restrict()'s speed. This script instead does a full
manual iteration of the Inbox and compares plain Python datetimes (via the
same dt() COM-time-to-datetime helper fetch_inbox.py already uses), which
is immune to the locale issue entirely.
"""

import sys
import time
import json
import os
from datetime import datetime

import win32com.client
import pywintypes

# Windows console default codepage (cp1252) chokes on real Outlook subject-
# line characters (non-breaking hyphen U+2011, em-dash, etc.) -- confirmed
# live 18 Aug 2026, crashed mid-report on a genuine subject line. Force
# UTF-8 stdout with a safe fallback instead of silently mangling or losing
# report output.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def log(msg):
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}")


def connect_to_outlook(max_attempts=3, retry_wait_seconds=45):
    """Same pattern as fetch_inbox.py's connect_to_outlook() -- late-bound
    Dispatch to avoid a corrupt win32com.gen_py cache, then GetNamespace,
    then touch GetDefaultFolder(6) (Inbox) since that's the exact call site
    that failed transiently on 11 Aug 2026, not Dispatch/GetNamespace
    themselves."""
    last_error = None
    for attempt in range(1, max_attempts + 1):
        try:
            outlook_app = win32com.client.dynamic.Dispatch("Outlook.Application")
            mapi_ns = outlook_app.GetNamespace("MAPI")
            inbox_folder = mapi_ns.GetDefaultFolder(6)
            if attempt > 1:
                log(f"Outlook COM connection succeeded on attempt {attempt}/{max_attempts}.")
            return outlook_app, mapi_ns, inbox_folder
        except pywintypes.com_error as e:
            last_error = e
            log(f"Outlook COM connection attempt {attempt}/{max_attempts} failed: {e}")
            if attempt < max_attempts:
                log(f"Outlook automation layer appears busy (transient). Waiting {retry_wait_seconds}s before retrying...")
                time.sleep(retry_wait_seconds)
    log(f"Outlook COM connection failed after {max_attempts} attempts. Giving up.")
    raise last_error


def dt(com_time):
    try:
        return datetime(com_time.year, com_time.month, com_time.day,
                         com_time.hour, com_time.minute, com_time.second)
    except Exception:
        return None


def get_archive_folder(mapi_ns, inbox_folder):
    """Read-only resolution of the destination Archive folder.

    GetDefaultFolder(23) (olFolderArchive) was tried first and CONFIRMED
    WRONG live on this machine/mailbox on 18 Aug 2026 -- it resolved to
    the Junk Email folder, not Archive. Dropped entirely; do not
    reintroduce without re-verifying live first.

    This Outlook session has FIVE separate stores/mailboxes attached
    (confirmed live 18 Aug 2026: 'HR Functional Analysis Team', 'People
    Department - HR Systems', 'Begbroke IT Support', Kevin's own primary
    'kevin.lelitte@admin.ox.ac.uk', and 'University of Oxford Recruitment
    Support'), and EVERY one of them has its own folder literally named
    'Archive'. A naive search across every store for the first folder
    named 'Archive' would silently pick a different mailbox's Archive
    folder (enumeration order puts 'HR Functional Analysis Team' before
    Kevin's own primary store) -- caught here before any move was
    attempted. The only correct destination is the Archive folder that is
    a direct sibling of the live Inbox folder we are actually reading
    from, i.e. scoped to inbox_folder.Parent (the same store Inbox itself
    lives in), never a mailbox-wide search."""
    mailbox_root = inbox_folder.Parent
    log(f"Inbox's own mailbox root: '{mailbox_root.Name}' -- scoping Archive search to this store only.")
    for f in mailbox_root.Folders:
        if f.Name.strip().lower() == "archive":
            log(f"Archive folder resolved (scoped to Inbox's own store): '{f.FolderPath}' (current item count: {f.Items.Count})")
            return f
    raise RuntimeError(
        f"No folder named 'Archive' found as a direct sibling of Inbox in store "
        f"'{mailbox_root.Name}'. Refusing to guess a move destination -- stopping."
    )


# APRIL_START is a classification boundary only (for the pre-April / April /
# May subtotal breakdown in the report) -- it is NOT a lower bound on what
# gets matched/moved. ARCHIVE_CUTOFF is the real filter: everything with
# ReceivedTime before 1 June 2026 is in scope, with no lower bound at all,
# per Kevin's same-day scope expansion (archive Apr/May 2026 AND everything
# older, leaving June 2026 onward untouched).
APRIL_START = datetime(2026, 4, 1, 0, 0, 0)
ARCHIVE_CUTOFF = datetime(2026, 6, 1, 0, 0, 0)  # exclusive -- everything from June onward is excluded


def find_items_to_archive(inbox_folder, cutoff=ARCHIVE_CUTOFF):
    """Full manual scan of the live Inbox (see module docstring for why
    Restrict() is deliberately not used). Returns a plain Python list of
    dicts holding the live COM item reference plus the fields needed for
    reporting/verification. Collecting into a separate list before any
    Move() call is deliberate -- moving items while still traversing the
    same Items collection with GetNext() corrupts the traversal.

    No lower bound is applied here -- any item with ReceivedTime < cutoff
    matches, regardless of how old it is. This deliberately mirrors the
    scope: "archive Apr/May 2026 AND everything older than that too.\""""
    items = inbox_folder.Items
    items.Sort("[ReceivedTime]", False)  # ascending, cheap ordering for the report
    matches = []
    total_scanned = 0
    unreadable = 0
    item = items.GetFirst()
    while item is not None:
        total_scanned += 1
        try:
            received = dt(item.ReceivedTime)
            subject = item.Subject
            entry_id = item.EntryID
            if received is not None and received < cutoff:
                matches.append({
                    "com_item": item,
                    "subject": subject,
                    "received": received,
                    "entry_id": entry_id,
                })
        except Exception as e:
            unreadable += 1
            log(f"WARNING: skipped one Inbox item during scan (unreadable ReceivedTime/Subject/EntryID): {e}")
        item = items.GetNext()
    log(f"Scanned {total_scanned} Inbox items total ({unreadable} unreadable/skipped).")
    return matches, unreadable


def count_in_range(inbox_folder, start, end):
    items = inbox_folder.Items
    item = items.GetFirst()
    n = 0
    while item is not None:
        try:
            received = dt(item.ReceivedTime)
            if received is not None and start <= received < end:
                n += 1
        except Exception:
            pass
        item = items.GetNext()
    return n


def report(matches, unreadable_count=0):
    if not matches:
        log("No matching items found in Inbox. Nothing to archive.")
        return
    pre_april = [m for m in matches if m["received"] < APRIL_START]
    by_month = {}
    for m in matches:
        key = m["received"].strftime("%Y-%m")
        by_month.setdefault(key, []).append(m)
    log(f"{len(matches)} Inbox items match the archive window (everything before 1 June 2026):")
    log(f"  pre-April 2026 (no lower bound): {len(pre_april)} items")
    for key in sorted(by_month):
        if key < "2026-04":
            continue  # already folded into the pre-April 2026 subtotal above
        log(f"  {key}: {len(by_month[key])} items")
    log(f"  COMBINED TOTAL: {len(matches)} items "
        f"(pre-April {len(pre_april)} + April {len(by_month.get('2026-04', []))} + May {len(by_month.get('2026-05', []))})")
    if unreadable_count:
        log(f"  NOTE: {unreadable_count} item(s) were unreadable during this scan (skipped, excluded from the move) "
            f"-- see the WARNING lines above for detail on each.")
    log("Full list (oldest to newest):")
    for m in sorted(matches, key=lambda x: x["received"]):
        log(f"  [{m['received'].strftime('%Y-%m-%d %H:%M')}] {m['subject']}")
    oldest = min(m["received"] for m in matches)
    newest = max(m["received"] for m in matches)
    log(f"Date range of matched items: {oldest.strftime('%Y-%m-%d')} (oldest item in Inbox) to {newest.strftime('%Y-%m-%d')}")
    if pre_april:
        oldest_pre_april = min(m["received"] for m in pre_april)
        log(f"Oldest pre-April-2026 item specifically: {oldest_pre_april.strftime('%Y-%m-%d %H:%M')}")


def write_report_json(matches, unreadable_count, path):
    pre_april = [m for m in matches if m["received"] < APRIL_START]
    april = [m for m in matches if m["received"].strftime("%Y-%m") == "2026-04"]
    may = [m for m in matches if m["received"].strftime("%Y-%m") == "2026-05"]
    payload = {
        "generated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "lower_bound": None,
        "cutoff_exclusive": ARCHIVE_CUTOFF.isoformat(),
        "count_total": len(matches),
        "count_pre_april_2026": len(pre_april),
        "count_april_2026": len(april),
        "count_may_2026": len(may),
        "unreadable_skipped": unreadable_count,
        "items": [
            {"subject": m["subject"], "received": m["received"].isoformat(), "entry_id": m["entry_id"]}
            for m in sorted(matches, key=lambda x: x["received"])
        ]
    }
    with open(path, "wb") as f:
        f.write(json.dumps(payload, indent=2, ensure_ascii=False).encode("utf-8"))
    log(f"Report written to {path}")


def execute_move(matches, archive_folder):
    moved = 0
    failed = []
    for m in matches:
        try:
            m["com_item"].Move(archive_folder)
            moved += 1
        except Exception as e:
            failed.append({"subject": m["subject"], "entry_id": m["entry_id"], "error": str(e)})
            log(f"FAILED to move '{m['subject']}' ({m['entry_id']}): {e}")
    log(f"Move complete: {moved}/{len(matches)} succeeded, {len(failed)} failed.")
    return moved, failed


def main():
    execute = "--execute" in sys.argv
    log(f"archive_apr_may_2026.py starting -- mode: {'EXECUTE' if execute else 'DRY RUN'}")

    outlook, mapi_ns, inbox_folder = connect_to_outlook()
    log(f"Connected to Outlook. Inbox total item count right now: {inbox_folder.Items.Count}")

    archive_folder = get_archive_folder(mapi_ns, inbox_folder)

    matches, unreadable_count = find_items_to_archive(inbox_folder)
    report(matches, unreadable_count)

    out_name = "archive_apr_may_2026_executed.json" if execute else "archive_apr_may_2026_dryrun.json"
    write_report_json(matches, unreadable_count, os.path.join(os.path.dirname(os.path.abspath(__file__)), out_name))

    if not execute:
        log("DRY RUN complete. Nothing was moved. Re-run with --execute only after Kevin's explicit go-ahead.")
        return

    if not matches:
        log("Nothing to move. Exiting.")
        return

    log(f"EXECUTING move of {len(matches)} items to '{archive_folder.FolderPath}'...")
    moved, failed = execute_move(matches, archive_folder)

    log("Post-run verification...")
    remaining, _ = find_items_to_archive(inbox_folder)
    if remaining:
        log(f"WARNING: {len(remaining)} pre-June-2026 items (pre-April/April/May) still present in Inbox after the move.")
    else:
        log("Verified: 0 pre-June-2026 items (pre-April/April/May) remain in Inbox.")

    june_aug_count = count_in_range(inbox_folder, datetime(2026, 6, 1), datetime(2026, 9, 1))
    log(f"Verified: {june_aug_count} June/July/August 2026 items remain in Inbox (expected -- untouched).")
    log(f"Inbox item count after move: {inbox_folder.Items.Count}")
    log(f"Archive folder ('{archive_folder.FolderPath}') item count after move: {archive_folder.Items.Count}")

    if failed:
        log(f"{len(failed)} item(s) FAILED to move -- see the JSON report and re-run investigation before assuming this is done.")


if __name__ == "__main__":
    main()
