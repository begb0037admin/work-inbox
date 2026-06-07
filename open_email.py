import sys
import win32com.client

if len(sys.argv) < 2:
    print("Usage: open_email.py <EntryID>")
    sys.exit(1)

entry_id = sys.argv[1]

try:
    outlook = win32com.client.Dispatch("Outlook.Application")
    mapi = outlook.GetNamespace("MAPI")
    item = mapi.GetItemFromID(entry_id)
    item.Display()
except Exception as e:
    print(f"Error: {e}")
    sys.exit(1)
