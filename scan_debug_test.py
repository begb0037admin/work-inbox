"""
scan_debug_test.py
Minimal debug test — READ ONLY. Does NOT move or modify anything.
Tests Restrict() + dt() on the root inbox only (no recursion).
Checks the first 3 matched items and reports exactly what it reads.
"""

import win32com.client
from datetime import datetime

CUTOFF = "01/01/2026 12:00 AM"
TEST_LIMIT = 3

def dt(com_time):
    try:
        return datetime(com_time.year, com_time.month, com_time.day,
                        com_time.hour, com_time.minute, com_time.second)
    except Exception as e:
        return None

print("=== scan_debug_test.py ===")
print("Connecting to Outlook...")
outlook = win32com.client.Dispatch("Outlook.Application")
ns = outlook.GetNamespace("MAPI")
inbox = ns.GetDefaultFolder(6)

print(f"\nStep 1: Running Restrict() on root Inbox for emails before {CUTOFF}...")
try:
    restricted = inbox.Items.Restrict("[ReceivedTime] < '" + CUTOFF + "'")
    print(f"  OK — Restrict() returned {restricted.Count} items")
except Exception as e:
    print(f"  FAILED — {e}")
    input("\nPress Enter to close...")
    exit()

print(f"\nStep 2: Reading first {TEST_LIMIT} items using dt() helper...")
checked = 0
for item in restricted:
    if checked >= TEST_LIMIT:
        break
    try:
        received = dt(item.ReceivedTime)
        subject  = item.Subject
        if received:
            print(f"  [{checked+1}] OK — {received.year}/{received.month:02d}/{received.day:02d}  '{subject}'")
        else:
            print(f"  [{checked+1}] dt() returned None — ReceivedTime type: {type(item.ReceivedTime)}")
    except Exception as e:
        print(f"  [{checked+1}] ERROR — {e}")
    checked += 1

if checked == 0:
    print("  No items to check — Restrict() returned empty collection.")

print("\n=== Test complete ===")
input("Press Enter to close...")
