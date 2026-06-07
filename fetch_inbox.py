import json, os, base64, webbrowser
from datetime import datetime, timedelta
import win32com.client
import anthropic

OUTPUT_RAW      = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "inbox_raw.json")
OUTPUT_BRIEFING = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "briefing.json")
DASHBOARD_URL   = "https://begb0037admin.github.io/work-inbox/"

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
            "importance":      msg.Importance
        }
        # Unread: include body preview. Read: subject + sender only.
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

SYSTEM = """You are Kevin's morning inbox triage assistant at Oxford University Personnel Services.
Analyse the inbox, sent items, and calendar data provided and return ONLY a valid JSON object - no preamble, no markdown, no code fences.
Use only plain ASCII punctuation: use - instead of dashes, use ' instead of curly quotes, use ... instead of ellipsis.

Note: inbox items with no body_preview field have already been read - triage by subject and sender only.
Items with body_preview are unread and should be given priority consideration.

The JSON must match this exact schema:
{
  "date": "<Day D Month YYYY>",
  "subtitle": "<one short phrase describing the day, optional>",
  "context": "<2-3 sentence summary of the current work situation>",
  "urgent": [{"title":"...","sub":"...","badge":"...","badgeType":"red|amber|blue|gray"}],
  "needs":  [{"title":"...","sub":"...","badge":"...","badgeType":"red|amber|blue|gray"}],
  "fyi":    [{"title":"...","sub":"...","badge":"...","badgeType":"red|amber|blue|gray"}],
  "low":    [{"title":"...","sub":"...","badge":"...","badgeType":"red|amber|blue|gray"}],
  "calToday":    [{"time":"...","title":"...","sub":"...","alert":"..."}],
  "calTomorrow": [{"time":"...","title":"...","sub":"...","alert":"..."}],
  "absences": ["Name - reason"],
  "priorities": [{"text":"...","date":"...","dateType":"red|amber|blue|gray"}]
}

Rules:
- urgent = must act today; needs = act within 48hrs; fyi = no action; low = noise/admin
- badge is a short deadline or action label (e.g. "Deadline 11 June", "Reply today")
- priorities is an ordered list of the top actions across all categories
- sub fields may contain <strong> tags for emphasis
- omit alert from calToday/calTomorrow items unless there is a genuine conflict or required action
- absences: only include people confirmed absent, inferred from out-of-office replies or calendar blocks
- calendar shows working days only (Monday to Friday)
"""

USER = f"""Today is {today_str}. Tomorrow (next working day) is {tomorrow_str}.

INBOX ({len(inbox)} emails, last 7 days - unread have body_preview, read do not):
{json.dumps(inbox, indent=2, ensure_ascii=True)}

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

with open(OUTPUT_BRIEFING, "w", encoding="utf-8") as f:
    json.dump(briefing, f, indent=2, ensure_ascii=False)

print(f"Phase 2 done - briefing written to {OUTPUT_BRIEFING}")

encoded = base64.urlsafe_b64encode(json.dumps(briefing, ensure_ascii=False).encode("utf-8")).decode("ascii")
url     = DASHBOARD_URL + "#load=" + encoded
webbrowser.open(url)
print(f"Dashboard opened.")