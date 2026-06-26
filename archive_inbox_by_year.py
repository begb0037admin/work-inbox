"""
archive_inbox_by_year.py
Full recursive sweep of Inbox and all subfolders.
Moves any email received before 2026 to Inbox/_Archive/<year>/.
Skips _Archive itself. Creates year folders as needed.
Run via bat file — always downloads latest from GitHub before running.
"""

import win32com.client

CUTOFF_YEAR = 2026


def get_or_create_folder(parent, name):
    for folder in parent.Folders:
        if folder.Name == name:
            return folder
    return parent.Folders.Add(name)


def collect_items(folder, archive_entry_id, results, scanned):
    """Recursively collect (item, year) tuples, skipping _Archive."""
    if folder.EntryID == archive_entry_id:
        return

    items = folder.Items
    count = items.Count
    scanned[0] += count

    for i in range(1, count + 1):
        try:
            item = items[i]
            year = item.ReceivedTime.year
            if year < CUTOFF_YEAR:
                results.append((item, year))
        except Exception:
            pass  # skip contacts, calendar items, corrupted entries

    for subfolder in folder.Folders:
        collect_items(subfolder, archive_entry_id, results, scanned)


def main():
    print("Connecting to Outlook...")
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    inbox = ns.GetDefaultFolder(6)  # olFolderInbox

    # Find _Archive folder
    archive = None
    for folder in inbox.Folders:
        if folder.Name == "_Archive":
            archive = folder
            break

    if archive is None:
        print("ERROR: _Archive folder not found inside Inbox. Aborting.")
        input("Press Enter to close...")
        return

    print(f"Scanning all Inbox folders recursively for emails before {CUTOFF_YEAR}...")
    to_move = []
    scanned = [0]
    collect_items(inbox, archive.EntryID, to_move, scanned)

    print(f"Scanned {scanned[0]} items. Found {len(to_move)} emails to archive.")

    if not to_move:
        print("Nothing to do.")
        input("Press Enter to close...")
        return

    print("Moving emails...")
    year_counts = {}
    errors = 0

    for i, (item, year) in enumerate(to_move, 1):
        year_str = str(year)
        year_folder = get_or_create_folder(archive, year_str)
        try:
            item.Move(year_folder)
            year_counts[year_str] = year_counts.get(year_str, 0) + 1
        except Exception as e:
            errors += 1
            try:
                subject = item.Subject
            except Exception:
                subject = "(unknown)"
            print(f"  [!] Failed to move '{subject}': {e}")

        if i % 100 == 0:
            print(f"  ... {i}/{len(to_move)} processed")

    print("\nDone. Summary:")
    for year_str in sorted(year_counts.keys(), reverse=True):
        print(f"  _Archive/{year_str}: {year_counts[year_str]} emails moved")
    if errors:
        print(f"  Errors: {errors} (check output above)")

    input("\nPress Enter to close...")


if __name__ == "__main__":
    main()
