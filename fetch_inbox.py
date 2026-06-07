import json, os, base64, webbrowser
from datetime import datetime, timedelta
import win32com.client
import anthropic

OUTPUT_RAW      = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "inbox_raw.json")
OUTPUT_BRIEFING = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "briefing.json")
DASHBOARD_URL   = "https://begb0037admin.github.io/work-inbox/"

# ── Phase 1: pull from Outlook ───────────────────────────────────────────────
outlook = win32com.client.Dispatch("Outlook.Application")
mapi    = outlook.GetNamespace("MAPI")
cutoff  = datetime.now() - timedelta(days=7)
today   = datetime.now().date()
tomorrow = today + timedelta(days=1)

def dt(com_time):
    try:
        return datetime(com_time.year, com_time.month, com_time.day,
                        com_time.hour, com_time.minute, com_time.second)
    except:
        return None

inbox = []
for msg in mapi.GetDefaultFolder(6).Items:
    try:
        t = dt(msg.ReceivedTime)
        if t and t >= cutoff:
            inbox.append({
                "subject":       msg.Subject,
                "from":          msg.SenderName,
                "from_email":    msg.SenderEmailAddress,
                "received":      str(msg.ReceivedTime),
                "is_read":       not msg.UnRead,
                "body_preview":  (msg.Body or "")[:500],
                "has_attachments": msg.Attachments.Count > 0,
                "importance":    msg.Importance
            })
    except:
        continue

sent = []
for msg in mapi.GetDefaultFolder(5).Items:
    try:
        t = dt(msg.SentOn)
        if t and t >= cutoff:
            sent.append({
                "subject":      msg.Subject,
                "to":           msg.To,
                "sent":         str(msg.SentOn),
                "body_preview": (msg.Body or "")[:200]
            })
    except:
        continue

calendar = []
for item in mapi.GetDefaultFolder(9).Items:
    try:
        t = dt(item.Start)
        if t and today <= t.date() <= tomorrow:
            calendar.append({
                "subject":      item.Subject,
                "start":        str(item.Start),
                "end":          str(item.End),
                "location":     item.Location,
                "organizer":    item.Organizer,
                "body_preview": (item.Body or "")[:200],
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

print(f"Phase 1 done — inbox:{len(inbox)} sent:{len(sent)} calendar:{len(calendar)}")

# ── Phase 2: triage via Anthropic API ────────────────────────────────────────
today_str    = datetime.now().strftime("%A %-d %B %Y")
tomorrow_str = (datetime.now() + timedelta(days=1)).strftime("%A %-d %B %Y")

cal_today    = [c for c in calendar if datetime.fromisoformat(c["start"]).date() == today]
cal_tomorrow = [c for c in calendar if datetime.fromisoformat(c["start"]).date() == tomorrow]

SYSTEM = """You are Kevin's morning inbox triage assistant at Oxford University Personnel Services.
Analyse the inbox, sent items, and calendar data provided and return ONLY a valid JSON object — no preamble, no markdown, no code fences.

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
  "absences": ["Name — reason"],
  "priorities": [{"text":"...","date":"...","dateType":"red|amber|blue|gray"}]
}

Rules:
- urgent = must act today; needs = act within 48hrs; fyi = no action; low = noise/admin
- badge is a short deadline or action label (e.g. "Deadline 11 June", "Reply today")
- priorities is an ordered list of the top actions across all categories
- sub fields may contain <strong> tags for emphasis
- omit alert from calToday/calTomorrow items unless there is a genuine conflict or required action
- absences: only include people confirmed absent, inferred from out-of-office replies or calendar blocks
"""

USER = f"""Today is {today_str}. Tomorrow is {tomorrow_str}.

INBOX ({len(inbox)} emails, last 7 days):
{json.dumps(inbox, indent=2, ensure_ascii=False)}

SENT ({len(sent)} items, last 7 days):
{json.dumps(sent, indent=2, ensure_ascii=False)}

CALENDAR TODAY:
{json.dumps(cal_today, indent=2, ensure_ascii=False)}

CALENDAR TOMORROW:
{json.dumps(cal_tomorrow, indent=2, ensure_ascii=False)}
"""

client   = anthropic.Anthropic()
response = client.messages.create(
    model      = "claude-sonnet-4-20250514",
    max_tokens = 4096,
    system     = SYSTEM,
    messages   = [{"role": "user", "content": USER}]
)

raw_text = response.content[0].text.strip()
# Strip any accidental code fences
if raw_text.startswith("```"):
    raw_text = "\n".join(raw_text.split("\n")[1:])
if raw_text.endswith("```"):
    raw_text = "\n".join(raw_text.split("\n")[:-1])

briefing = json.loads(raw_text)

with open(OUTPUT_BRIEFING, "w", encoding="utf-8") as f:
    json.dump(briefing, f, indent=2, ensure_ascii=False)

print(f"Phase 2 done — briefing written to {OUTPUT_BRIEFING}")

# ── Phase 3: open dashboard with hash-loaded briefing ────────────────────────
encoded = base64.urlsafe_b64encode(json.dumps(briefing, ensure_ascii=False).encode()).decode()
url     = DASHBOARD_URL + "#load=" + encoded
webbrowser.open(url)
print(f"Dashboard opened.")