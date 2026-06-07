import json, os, base64, urllib.request, urllib.error
from datetime import datetime, timedelta
import win32com.client
import anthropic

OUTPUT_RAW      = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "inbox_raw.json")
OUTPUT_BRIEFING = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "briefing.json")
GITHUB_REPO     = "begb0037admin/work-inbox"
GITHUB_PATH     = "data/briefing.json"
GITHUB_PAT      = os.environ.get("GITHUB_PAT", "")

outlook = win32com.client.Dispatch("Outlook.Application")
mapi    = outlook.GetNamespace("MAPI")
cutoff  = datetime.now() - timedelta(days=7)
today   = datetime.now().date()

def next_workday(d):
    d = d + timedelta(days=1)
    while d.weekday() >= 5:
        d = d + timedelta(days=1)
    return d

tomorrow = next_workday(today)

def dt(com_time):
    try:
        return datetime(com_time.year, com_time.month, com_time.day,
                        com_time.hour, com_time.minute, com_time.second)
    except:
        return None

def restrict_date(folder, cutoff_dt):
    filter_str = "[ReceivedTime] >= '" + cutoff_dt.strftime("%m/%d/%Y %H:%M %p") + "'"
    try:
        return folder.Items.Restrict(filter_str)
    except:
        return folder.Items

print("Phase 1 - pulling Outlook data...")
inbox = []
for msg in restrict_date(mapi.GetDefaultFolder(6), cutoff):
    try:
        is_read = not msg.UnRead
        entry = {
            "subject":         msg.Subject,
            "from":            msg.SenderName,
            "from_email":      msg.SenderEmailAddress,
            "received":        str(msg.ReceivedTime),
            "is_read":         is_read,
            "has_attachments": msg.Attachments.Count > 0,
            "importance":      msg.Importance,
            "entry_id":        msg.EntryID
        }
        if not is_read:
            entry["body_preview"] = (msg.Body or "")[:150]
        inbox.append(entry)
    except:
        continue

inbox.sort(key=lambda x: (x["is_read"], x["received"]))

sent = []
for msg in mapi.GetDefaultFolder(5).Items:
    try:
        t = dt(msg.SentOn)
        if t and t >= cutoff:
            sent.append({
                "subject":      msg.Subject,
                "to":           msg.To,
                "sent":         str(msg.SentOn),
                "body_preview": (msg.Body or "")[:100]
            })
    except:
        continue

calendar = []
for item in mapi.GetDefaultFolder(9).Items:
    try:
        t = dt(item.Start)
        if t and t.weekday() < 5 and today <= t.date() <= tomorrow:
            calendar.append({
                "subject":      item.Subject,
                "start":        str(item.Start),
                "end":          str(item.End),
                "location":     item.Location,
                "organizer":    item.Organizer,
                "body_preview": (item.Body or "")[:100],
                "all_day":      item.AllDayEvent
            })
    except:
        continue

raw = {
    "pulled_at": datetime.now().isoformat(),
    "inbox":     inbox,
    "sent":      sent,
    "calendar":  calendar
}

os.makedirs(os.path.dirname(OUTPUT_RAW), exist_ok=True)
with open(OUTPUT_RAW, "w", encoding="utf-8") as f:
    json.dump(raw, f, indent=2, ensure_ascii=False)

unread_count = sum(1 for m in inbox if not m["is_read"])
print(f"Phase 1 done - inbox:{len(inbox)} (unread:{unread_count}) sent:{len(sent)} calendar:{len(calendar)}")
print("Phase 2 - calling Anthropic API...")

now = datetime.now()
today_str    = now.strftime("%A") + " " + str(now.day) + " " + now.strftime("%B %Y")
tomorrow_str = tomorrow.strftime("%A") + " " + str(tomorrow.day) + " " + tomorrow.strftime("%B %Y")

cal_today    = [c for c in calendar if datetime.fromisoformat(c["start"]).date() == today]
cal_tomorrow = [c for c in calendar if datetime.fromisoformat(c["start"]).date() == tomorrow]

inbox_for_api = [{k:v for k,v in m.items() if k != "entry_id"} for m in inbox]

SYSTEM = """You are Kevin's morning inbox triage assistant at Oxford University Personnel Services.
Analyse the inbox, sent items, and calendar data provided and return ONLY a valid JSON object - no preamble, no markdown, no code fences.
Use only plain ASCII punctuation: use - instead of dashes, use ' instead of curly quotes, use ... instead of ellipsis.

The JSON must match this exact schema:
{
  "date": "<Day D Month YYYY>",
  "subtitle": "<one short phrase describing the day, optional>",
  "context": "<2-3 sentence summary of the current work situation>",
  "urgent": [{"title":"...","sub":"...","badge":"...","badgeType":"red|yellow|blue|gray","subject":"..."}],
  "needs":  [{"title":"...","sub":"...","badge":"...","badgeType":"red|yellow|blue|gray","subject":"..."}],
  "fyi":    [{"title":"...","sub":"...","badge":"...","badgeType":"red|yellow|blue|gray","subject":"..."}],
  "low":    [{"title":"...","sub":"...","badge":"...","badgeType":"red|yellow|blue|gray","subject":"..."}],
  "calToday":    [{"time":"...","title":"...","sub":"...","alert":"..."}],
  "calTomorrow": [{"time":"...","title":"...","sub":"...","alert":"..."}],
  "absences": ["Name - reason"],
  "priorities": [{"text":"...","date":"...","dateType":"red|yellow|blue|gray","subject":"..."}]
}

Rules:
- urgent = must act today; needs = act within 48hrs; fyi = no action; low = noise/admin
- badge is a short deadline or action label (e.g. "Deadline 11 June", "Reply today")
- priorities is an ordered list of the top actions across all categories
- sub fields may contain <strong> tags for emphasis
- omit alert from calToday/calTomorrow items unless there is a genuine conflict or required action
- absences: only include people confirmed absent, inferred from out-of-office replies or calendar blocks
- calendar shows working days only (Monday to Friday)
- subject field: copy the exact email subject this item relates to (used to match EntryID for opening in Outlook). If not from a specific email, omit the subject field.
"""

USER = f"""Today is {today_str}. Tomorrow (next working day) is {tomorrow_str}.

INBOX ({len(inbox_for_api)} emails, last 7 days - unread have body_preview, read do not):
{json.dumps(inbox_for_api, indent=2, ensure_ascii=True)}

SENT ({len(sent)} items, last 7 days):
{json.dumps(sent, indent=2, ensure_ascii=True)}

CALENDAR TODAY:
{json.dumps(cal_today, indent=2, ensure_ascii=True)}

CALENDAR TOMORROW:
{json.dumps(cal_tomorrow, indent=2, ensure_ascii=True)}
"""

client = anthropic.Anthropic(timeout=60.0)
response = client.messages.create(
    model      = "claude-haiku-4-5",
    max_tokens = 4096,
    system     = SYSTEM,
    messages   = [{"role": "user", "content": USER}]
)

raw_text = response.content[0].text.strip()
if raw_text.startswith("```"):
    raw_text = "\n".join(raw_text.split("\n")[1:])
if raw_text.endswith("```"):
    raw_text = "\n".join(raw_text.split("\n")[:-1])

briefing = json.loads(raw_text)

subject_to_entryid = {}
for m in inbox:
    if m.get("entry_id") and m.get("subject"):
        subject_to_entryid[m["subject"]] = m["entry_id"]

def inject_entry_ids(items):
    if not items:
        return items
    for item in items:
        subj = item.get("subject", "")
        if subj and subj in subject_to_entryid:
            item["entry_id"] = subject_to_entryid[subj]
    return items

for section in ["urgent", "needs", "fyi", "low", "priorities"]:
    inject_entry_ids(briefing.get(section, []))

with open(OUTPUT_BRIEFING, "w", encoding="utf-8") as f:
    json.dump(briefing, f, indent=2, ensure_ascii=False)

print(f"Phase 2 done - briefing written to {OUTPUT_BRIEFING}")
print("Phase 3 - pushing briefing to GitHub...")

if not GITHUB_PAT:
    print("WARNING: GITHUB_PAT not set - skipping GitHub push")
else:
    try:
        api_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_PATH}"
        headers = {
            "Authorization": f"token {GITHUB_PAT}",
            "Content-Type": "application/json",
            "User-Agent": "work-inbox-script"
        }
        sha = None
        try:
            req = urllib.request.Request(api_url, headers=headers)
            with urllib.request.urlopen(req) as r:
                existing = json.loads(r.read())
                sha = existing.get("sha")
        except urllib.error.HTTPError as e:
            if e.code != 404:
                raise
        content_b64 = base64.b64encode(
            json.dumps(briefing, indent=2, ensure_ascii=False).encode("utf-8")
        ).decode("ascii")
        payload = {
            "message": f"chore: update briefing {datetime.now().strftime('%Y-%m-%d %H:%M')}",
            "content": content_b64
        }
        if sha:
            payload["sha"] = sha
        data = json.dumps(payload).encode("utf-8")
        req = urllib.request.Request(api_url, data=data, headers=headers, method="PUT")
        with urllib.request.urlopen(req) as r:
            result = json.loads(r.read())
            print(f"Phase 3 done - briefing pushed to GitHub (commit: {result.get('commit',{}).get('sha','?')[:7]})")
    except Exception as e:
        print(f"Phase 3 WARNING - GitHub push failed: {e}")
        print("Briefing is still available locally.")
