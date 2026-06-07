import sys
import win32com.client
from datetime import datetime

LOG = r"C:\Users\admin\work-inbox\data\openmail.log"

def log(msg):
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(f"{datetime.now().isoformat()} {msg}\n")

if len(sys.argv) < 2:
    log("ERROR: no argument passed")
    sys.exit(1)

raw = sys.argv[1]
log(f"RAW ARG: {raw}")
entry_id = raw.replace("openmail://", "").rstrip("/")
log(f"ENTRY ID: {entry_id}")

try:
    outlook = win32com.client.Dispatch("Outlook.Application")
    mapi = outlook.GetNamespace("MAPI")
    item = mapi.GetItemFromID(entry_id)
    item.Display()
    log("SUCCESS")
except Exception as e:
    log(f"ERROR: {e}")
    sys.exit(1)
