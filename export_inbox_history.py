import json, os, base64, urllib.request, urllib.error
from datetime import datetime, timedelta
import win32com.client

GITHUB_REPO  = "begb0037admin/work-inbox"
GITHUB_PATH  = "data/inbox_export.json"
GITHUB_PAT   = os.environ.get("GITHUB_PAT", "")

MONTHS_BACK  = 6
cutoff       = datetime.now() - timedelta(days=MONTHS_BACK * 30)

print(f"Exporting inbox + sent items from {cutoff.strftime('%d %b %Y')} to today...")

outlook = win32com.client.Dispatch("Outlook.Application")
mapi    = outlook.GetNamespace("MAPI")

def to_dt(com_time):
    try:
        return datetime(com_time.year, com_time.month, com_time.day,
                        com_time.hour, com_time.minute, com_time.second)
    except:
        return None

# Pull inbox -- sort newest first so we can stop as soon as we pass the cutoff
print("Phase 1 - pulling inbox...")
inbox_items = []
items = mapi.GetDefaultFolder(6).Items
items.Sort("[ReceivedTime]", True)  # descending
for msg in items:
    try:
        t = to_dt(msg.ReceivedTime)
        if not t:
            continue
        if t < cutoff:
            break  # sorted descending, so nothing older will appear
        inbox_items.append({
            "subject":    (msg.Subject or "").strip(),
            "from":       (msg.SenderName or "").strip(),
            "from_email": (msg.SenderEmailAddress or "").lower().strip(),
            "received":   t.strftime("%Y-%m-%d %H:%M"),
            "is_read":    not msg.UnRead
        })
        if len(inbox_items) % 200 == 0:
            print(f"  ...{len(inbox_items)} inbox emails so far")
    except:
        continue

print(f"Phase 1 done - {len(inbox_items)} inbox emails")

# Pull sent items -- sort newest first, stop at cutoff
print("Phase 2 - pulling sent items...")
sent_items = []
sitems = mapi.GetDefaultFolder(5).Items
sitems.Sort("[SentOn]", True)  # descending
for msg in sitems:
    try:
        t = to_dt(msg.SentOn)
        if not t:
            continue
        if t < cutoff:
            break
        sent_items.append({
            "subject": (msg.Subject or "").strip(),
            "to":      (msg.To or "").strip(),
            "sent":    t.strftime("%Y-%m-%d %H:%M")
        })
        if len(sent_items) % 200 == 0:
            print(f"  ...{len(sent_items)} sent emails so far")
    except:
        continue

print(f"Phase 2 done - {len(sent_items)} sent emails")

export = {
    "exported_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
    "period_from": cutoff.strftime("%Y-%m-%d"),
    "period_to":   datetime.now().strftime("%Y-%m-%d"),
    "inbox_count": len(inbox_items),
    "sent_count":  len(sent_items),
    "inbox":       inbox_items,
    "sent":        sent_items
}

if not GITHUB_PAT:
    print("ERROR: GITHUB_PAT env var not set - cannot push.")
else:
    print("Phase 3 - pushing to GitHub...")
    api_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_PATH}"
    headers = {
        "Authorization": f"token {GITHUB_PAT}",
        "Content-Type":  "application/json",
        "User-Agent":    "inbox-export-script"
    }
    sha = None
    try:
        req = urllib.request.Request(api_url, headers=headers)
        with urllib.request.urlopen(req) as r:
            sha = json.loads(r.read()).get("sha")
    except urllib.error.HTTPError as e:
        if e.code != 404:
            raise

    content_b64 = base64.b64encode(
        json.dumps(export, indent=2, ensure_ascii=False).encode("utf-8")
    ).decode("ascii")
    payload = {
        "message": f"chore: inbox export {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        "content": content_b64
    }
    if sha:
        payload["sha"] = sha

    req = urllib.request.Request(api_url, data=json.dumps(payload).encode("utf-8"),
                                 headers=headers, method="PUT")
    with urllib.request.urlopen(req) as r:
        result = json.loads(r.read())
        commit = result.get("commit", {}).get("sha", "?")[:7]
        print(f"Phase 3 done - pushed to GitHub (commit: {commit})")

    print(f"\nExport complete:")
    print(f"  Inbox:  {len(inbox_items)} emails")
    print(f"  Sent:   {len(sent_items)} emails")
    print(f"  Period: {cutoff.strftime('%d %b %Y')} to today")
    print(f"  File:   {GITHUB_REPO}/{GITHUB_PATH}")
