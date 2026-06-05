import json, os
from datetime import datetime, timedelta
import win32com.client

OUTPUT_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "inbox_raw.json")

outlook = win32com.client.Dispatch("Outlook.Application")
mapi = outlook.GetNamespace("MAPI")

cutoff = datetime.now() - timedelta(days=7)
today = datetime.now().date()
tomorrow = today + timedelta(days=1)

def dt(com_time):
    try:
        return datetime(com_time.year, com_time.month, com_time.day,
                       com_time.hour, com_time.minute, com_time.second)
    except: return None

# Inbox
inbox = []
for msg in mapi.GetDefaultFolder(6).Items:
    try:
        t = dt(msg.ReceivedTime)
        if t and t >= cutoff:
            inbox.append({
                "subject": msg.Subject,
                "from": msg.SenderName,
                "from_email": msg.SenderEmailAddress,
                "received": str(msg.ReceivedTime),
                "is_read": not msg.UnRead,
                "body_preview": (msg.Body or "")[:500],
                "has_attachments": msg.Attachments.Count > 0,
                "importance": msg.Importance
            })
    except: continue

# Sent items
sent = []
for msg in mapi.GetDefaultFolder(5).Items:
    try:
        t = dt(msg.SentOn)
        if t and t >= cutoff:
            sent.append({
                "subject": msg.Subject,
                "to": msg.To,
                "sent": str(msg.SentOn),
                "body_preview": (msg.Body or "")[:200]
            })
    except: continue

# Calendar
calendar = []
for item in mapi.GetDefaultFolder(9).Items:
    try:
        t = dt(item.Start)
        if t and today <= t.date() <= tomorrow:
            calendar.append({
                "subject": item.Subject,
                "start": str(item.Start),
                "end": str(item.End),
                "location": item.Location,
                "organizer": item.Organizer,
                "body_preview": (item.Body or "")[:200],
                "all_day": item.AllDayEvent
            })
    except: continue

output = {
    "pulled_at": datetime.now().isoformat(),
    "inbox": inbox,
    "sent": sent,
    "calendar": calendar
}

os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
    json.dump(output, f, indent=2, ensure_ascii=False)

print(f"\n✅ Done. Output: {OUTPUT_FILE}")
print(f"   Inbox: {len(inbox)} | Sent: {len(sent)} | Calendar: {len(calendar)}")