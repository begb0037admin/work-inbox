"""
scan_inbox_by_year.py
READ-ONLY diagnostic scan. Does NOT move or modify anything.
Scans Inbox and all subfolders recursively.
Reports email counts by year with sample subjects.
Run via PowerShell — always downloads latest from GitHub before running.
"""

import win32com.client
from collections import defaultdict

SAMPLE_SIZE = 5  # number of example subjects to show per year


def collect_items(folder, archive_entry_id, results, scanned, depth=0):
    """Recursively collect (year, subject) tuples, skipping _Archive."""
    if folder.EntryID == archive_entry_id:
        return

    items = folder.Items
    count = items.Count
    scanned[0] += count

    for i in range(1, count + 1):
        try:
            item = items[i]
            year = item.ReceivedTime.year
            subject = item.Subject
            results[year].append(subject)
        except Exception:
            pass  # skip contacts, calendar items, corrupted entries

    for subfolder in folder.Folders:
        collect_items(subfolder, archive_entry_id, results, scanned, depth + 1)


def main():
    print("Connecting to Outlook...")
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    inbox = ns.GetDefaultFolder(6)  # olFolderInbox

    # Locate _Archive to skip it
    archive_entry_id = None
    for folder in inbox.Folders:
        if folder.Name == "_Archive":
            archive_entry_id = folder.EntryID
            break

    if archive_entry_id is None:
        print("NOTE: No _Archive folder found — will scan everything under Inbox.")

    print("Scanning Inbox and all subfolders (read-only)...\n")
    results = defaultdict(list)
    scanned = [0]
    collect_items(inbox, archive_entry_id, results, scanned)

    print(f"Total items scanned: {scanned[0]}")
    print(f"Total emails found with a ReceivedTime: {sum(len(v) for v in results.values())}")
    print()

    if not results:
        print("No emails found.")
    else:
        print("Breakdown by year:")
        print("-" * 50)
        for year in sorted(results.keys(), reverse=True):
            subjects = results[year]
            print(f"\n  {year}: {len(subjects)} emails")
            print(f"  Sample subjects:")
            for s in subjects[:SAMPLE_SIZE]:
                print(f"    - {s}")
            if len(subjects) > SAMPLE_SIZE:
                print(f"    ... and {len(subjects) - SAMPLE_SIZE} more")

    print()
    input("Press Enter to close...")


if __name__ == "__main__":
    main()
