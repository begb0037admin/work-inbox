"""create_missing_folders.py
Creates missing Outlook mail folders needed by rules that errored.
Read-only check first — only creates folders that don't already exist.
"""
import win32com.client

FOLDERS_TO_CREATE = [
    "ICT Mailing Lists",
    "Bodleian & Sector",
]

print("Connecting to Outlook...")
outlook = win32com.client.Dispatch("Outlook.Application")
ns = outlook.GetNamespace("MAPI")
inbox = ns.GetDefaultFolder(6)  # olFolderInbox
parent = inbox.Parent  # top-level mailbox folder

existing = [f.Name for f in parent.Folders]
print(f"Existing top-level folders: {existing}\n")

for name in FOLDERS_TO_CREATE:
    if name in existing:
        print(f"  SKIP  '{name}' — already exists")
    else:
        try:
            parent.Folders.Add(name)
            print(f"  OK    '{name}' — created")
        except Exception as e:
            print(f"  ERROR '{name}' — {e}")

print("\nDone. Re-run your Outlook rules now (Home → Rules → Run Rules Now).")
input("\nPress Enter to close...")
