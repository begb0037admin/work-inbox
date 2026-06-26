"""
scan_inbox_by_year.py  v5
READ-ONLY diagnostic scan. Does NOT move or modify anything.
Scans Inbox and all subfolders recursively (skips _Archive).
Reports email counts by year with sample subjects.

Uses Restrict() + dt() patterns proven in fetch_inbox.py.
Run via PowerShell — always downloads latest from GitHub before running.
"""

import win32com.client
from collections import defaultdict
from datetime import datetime

print("Script version: v5")

SAMPLE_SIZE = 5
CUTOFF = "01/01/2026 12:00 AM"   # match fetch_inbox.py date format


def dt(com_time):
    """Safely extract a Python datetime from a COM datetime object."""
    try:
        return datetime(com_time.year, com_time.month, com_time.day,
                        com_time.hour, com_time.minute, com_time.second)
    except Exception:
        return None


def scan_folder(folder, archive_entry_id, results, counts):
    """Recursively scan a folder for pre-2026 mail items. Read-only."""
    if folder.EntryID == archive_entry_id:
        return

    try:
        # Server-side restriction — Exchange evaluates ReceivedTime,
        # returns only matching items with properties already resolved.
        # Pattern taken directly from fetch_inbox.py restrict_date().
        restricted = folder.Items.Restrict(
            "[ReceivedTime] < '" + CUTOFF + "'"
        )
        counts["found"] += restricted.Count

        for item in restricted:
            try:
                received = dt(item.ReceivedTime)
                if received:
                    results[received.year].append(item.Subject)
                else:
                    counts["no_date"] += 1
            except Exception:
                counts["item_error"] += 1

    except Exception as e:
        counts["folder_error"] += 1

    for subfolder in folder.Folders:
        scan_folder(subfolder, archive_entry_id, results, counts)


def main():
    print("Connecting to Outlook...")
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    inbox = ns.GetDefaultFolder(6)  # olFolderInbox

    archive_entry_id = None
    for folder in inbox.Folders:
        if folder.Name == "_Archive":
            archive_entry_id = folder.EntryID
            break

    if archive_entry_id is None:
        print("NOTE: No _Archive folder found — scanning everything under Inbox.")

    print(f"Scanning for emails before {CUTOFF} (read-only)...\n")
    results = defaultdict(list)
    counts = {"found": 0, "no_date": 0, "item_error": 0, "folder_error": 0}

    scan_folder(inbox, archive_entry_id, results, counts)

    total_read = sum(len(v) for v in results.values())
    print(f"Items matched by Restrict(): {counts['found']}")
    print(f"Items successfully read:     {total_read}")
    if counts["no_date"]:
        print(f"Items with unreadable date:  {counts['no_date']}")
    if counts["item_error"]:
        print(f"Item errors:                 {counts['item_error']}")
    if counts["folder_error"]:
        print(f"Folder errors:               {counts['folder_error']}")

    if not results:
        print("\nNo pre-2026 emails found in inbox.")
    else:
        print("\nBreakdown by year:")
        print("-" * 50)
        for year in sorted(results.keys(), reverse=True):
            subjects = results[year]
            print(f"\n  {year}: {len(subjects)} emails")
            for s in subjects[:SAMPLE_SIZE]:
                print(f"    - {s}")
            if len(subjects) > SAMPLE_SIZE:
                print(f"    ... and {len(subjects) - SAMPLE_SIZE} more")

    print()
    input("Press Enter to close...")


if __name__ == "__main__":
    main()
