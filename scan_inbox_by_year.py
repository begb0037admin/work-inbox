"""
scan_inbox_by_year.py
READ-ONLY diagnostic scan. Does NOT move or modify anything.
Scans Inbox and all subfolders recursively.
Reports email counts by year with sample subjects.
Run via PowerShell — always downloads latest from GitHub before running.
"""

import win32com.client
from collections import defaultdict

SAMPLE_SIZE = 5    # number of example subjects to show per year
OL_MAIL_ITEM = 43  # Outlook mail item class constant


def collect_items(folder, archive_entry_id, results, scanned, type_counts):
    """Recursively collect (year, subject) tuples, skipping _Archive."""
    if folder.EntryID == archive_entry_id:
        return

    items = folder.Items
    scanned[0] += items.Count

    # Direct iteration holds the COM reference correctly (index access releases it)
    for item in items:
        try:
            item_class = item.Class
            type_counts[item_class] = type_counts.get(item_class, 0) + 1

            if item_class == OL_MAIL_ITEM:
                year = item.ReceivedTime.year
                subject = item.Subject
                results[year].append(subject)

        except Exception as e:
            type_counts["error"] = type_counts.get("error", 0) + 1

    for subfolder in folder.Folders:
        collect_items(subfolder, archive_entry_id, results, scanned, type_counts)


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
    type_counts = {}
    collect_items(inbox, archive_entry_id, results, scanned, type_counts)

    print(f"Total items scanned: {scanned[0]}")
    print(f"Total mail items found: {sum(len(v) for v in results.values())}")

    print("\nItem type breakdown:")
    print("  43=Mail  45=Appointment  46=Contact  48=Task")
    for cls in sorted(type_counts.keys(), key=lambda x: str(x)):
        print(f"  Class {cls}: {type_counts[cls]}")

    if not results:
        print("\nNo mail items found.")
    else:
        print("\nMail breakdown by year:")
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
