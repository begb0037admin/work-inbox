"""
search_michael_evidence.py
Searches Outlook inbox and sent items from 2026 for evidence of Michael O'Sullivan's
work beyond WFM Wave 3 and clinical pay. Pushes full email bodies to GitHub.

Run on admin Windows machine:
    git fetch origin && git checkout origin/main -- search_michael_evidence.py && python search_michael_evidence.py
"""

import json, os, base64, urllib.request, urllib.error
from datetime import datetime
import win32com.client

GITHUB_REPO = "begb0037admin/work-inbox"
GITHUB_PATH = "data/michael_evidence.json"
GITHUB_PAT  = os.environ.get("GITHUB_PAT", "")

# Keywords targeting Michael's OTHER 2026 work
# WFM wave 3 and clinical pay already covered - excluded here to keep output focused
SEARCH_TERMS = [
    "tupe", "liverpool",
    "ndm", "workgroup", "work group",
    "osps", "contribution rate",
    "sickness absence", "absence code", "absence request",
    "college staff", "dtp1092",
    "calendar view", "employee calendar",
    "ess", "employee self service", "self service",
    "ref2029", "ref 2029", "pay admin",
    "ultrix",
    "rostering", "roster",
    "volunteering", "volunteer hours",
    "hesa", "research data",
    "non sickness", "non-sickness",
    "rolled up", "rolled-up", "holiday pay",
    "variable hours", "unrostered",
    "michael o'sullivan", "michael o sullivan",
]

CUTOFF = datetime(2026, 1, 1)

def com_dt(com_time):
    try:
        return datetime(com_time.year, com_time.month, com_time.day,
                        com_time.hour, com_time.minute, com_time.second)
    except:
        return None

def matches(subj, body):
    text = ((subj or "") + " " + (body or "")[:3000]).lower()
    return any(term in text for term in SEARCH_TERMS)

print("Connecting to Outlook...")
outlook = win32com.client.Dispatch("Outlook.Application")
mapi    = outlook.GetNamespace("MAPI")

results = []

# --- Inbox ---
print("Searching inbox from Jan 2026...")
inbox_folder = mapi.GetDefaultFolder(6)
try:
    filter_str = "[ReceivedTime] >= '01/01/2026 12:00 AM'"
    items = inbox_folder.Items.Restrict(filter_str)
    print(f"  Restrict() returned {items.Count} items")
except Exception as e:
    print(f"  Restrict() failed ({e}), falling back to full iteration")
    items = inbox_folder.Items

for msg in items:
    try:
        t = com_dt(msg.ReceivedTime)
        if not t or t < CUTOFF:
            continue
        subj = msg.Subject or ""
        body = msg.Body or ""
        if matches(subj, body):
            results.append({
                "direction": "received",
                "subject":   subj,
                "from":      msg.SenderName,
                "from_email": msg.SenderEmailAddress,
                "to":        msg.To,
                "date":      str(msg.ReceivedTime)[:16],
                "body":      body[:4000],
                "entry_id":  msg.EntryID,
            })
    except:
        continue

inbox_count = len(results)
print(f"  Inbox matches: {inbox_count}")

# --- Sent Items ---
print("Searching sent items from Jan 2026...")
sent_folder = mapi.GetDefaultFolder(5)
try:
    filter_str = "[SentOn] >= '01/01/2026 12:00 AM'"
    sent_items = sent_folder.Items.Restrict(filter_str)
    print(f"  Restrict() returned {sent_items.Count} items")
except Exception as e:
    print(f"  Restrict() failed ({e}), falling back to full iteration")
    sent_items = sent_folder.Items

for msg in sent_items:
    try:
        t = com_dt(msg.SentOn)
        if not t or t < CUTOFF:
            continue
        subj = msg.Subject or ""
        body = msg.Body or ""
        if matches(subj, body):
            results.append({
                "direction": "sent",
                "subject":   subj,
                "from":      "Kevin (sent)",
                "to":        msg.To,
                "date":      str(msg.SentOn)[:16],
                "body":      body[:4000],
                "entry_id":  msg.EntryID,
            })
    except:
        continue

sent_count = len(results) - inbox_count
print(f"  Sent matches: {sent_count}")
print(f"  Total: {len(results)}")

results.sort(key=lambda x: x["date"], reverse=True)

output = {
    "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
    "search_terms":  SEARCH_TERMS,
    "total":         len(results),
    "emails":        results,
}

if not GITHUB_PAT:
    out_file = "michael_evidence.json"
    with open(out_file, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2, ensure_ascii=False)
    print(f"GITHUB_PAT not set - saved locally as {out_file}")
else:
    print(f"Pushing {len(results)} emails to GitHub ({GITHUB_REPO}/{GITHUB_PATH})...")
    api_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_PATH}"
    headers = {
        "Authorization": f"token {GITHUB_PAT}",
        "Content-Type":  "application/json",
        "User-Agent":    "search-michael-evidence",
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
        json.dumps(output, indent=2, ensure_ascii=False).encode("utf-8")
    ).decode("ascii")
    payload = {"message": f"chore: michael evidence {datetime.now().strftime('%Y-%m-%d %H:%M')}",
               "content": content_b64}
    if sha:
        payload["sha"] = sha

    req = urllib.request.Request(api_url,
                                  data=json.dumps(payload).encode("utf-8"),
                                  headers=headers, method="PUT")
    with urllib.request.urlopen(req) as r:
        result = json.loads(r.read())
        commit = result.get("commit", {}).get("sha", "?")[:7]
        print(f"Done - commit {commit}")
        print(f"Claude can now read: https://github.com/{GITHUB_REPO}/blob/main/{GITHUB_PATH}")
