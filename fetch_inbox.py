import json, os, base64, html, urllib.request, urllib.error
from datetime import datetime, timedelta
import win32com.client
import anthropic

GITHUB_REPO = "begb0037admin/work-inbox"
GITHUB_PATH = "data/briefing.json"
GITHUB_PAT  = os.environ.get("GITHUB_PAT", "")

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
    filter_str = "[ReceivedTime] >= '" + cutoff_dt.strftime("%m/%d/%Y %I:%M %p") + "'"
    try:
        restricted = folder.Items.Restrict(filter_str)
        if restricted.Count > 200:
            raise Exception("Filter returned too many items - likely failed")
        return restricted
    except:
        items = folder.Items
        items.Sort("[ReceivedTime]", True)
        return items

# Ã¢ÂÂÃ¢ÂÂ Phase 1 Ã¢ÂÂ pull Outlook data Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂ
print("Phase 1 - pulling Outlook data...")
inbox = []
unread_count = 0
read_count   = 0
MAX_UNREAD   = 50
MAX_READ     = 30

for msg in restrict_date(mapi.GetDefaultFolder(6), cutoff):
    try:
        if unread_count >= MAX_UNREAD and read_count >= MAX_READ:
            break
        is_read = not msg.UnRead
        if is_read and read_count >= MAX_READ:
            continue
        if not is_read and unread_count >= MAX_UNREAD:
            continue
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
            unread_count += 1
        else:
            read_count += 1
        inbox.append(entry)
    except:
        continue

inbox.sort(key=lambda x: (not x["is_read"], x["received"]), reverse=True)

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

week_end = today + timedelta(days=6)
lookback  = today - timedelta(days=30)  # catch multi-day absences spanning today
calendar = []
for item in mapi.GetDefaultFolder(9).Items:
    try:
        t = dt(item.Start)
        if not t:
            continue
        if lookback <= t.date() <= week_end:
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

unread_total = sum(1 for m in inbox if not m["is_read"])
print(f"Phase 1 done - inbox:{len(inbox)} (unread:{unread_total}) sent:{len(sent)} calendar:{len(calendar)}")

# Ã¢ÂÂÃ¢ÂÂ Phase 2 Ã¢ÂÂ AI writes context paragraph only Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂ
print("Phase 2 - calling Anthropic API for context...")

now          = datetime.now()
today_str    = now.strftime("%A") + " " + str(now.day) + " " + now.strftime("%B %Y")
tomorrow_str = tomorrow.strftime("%A") + " " + str(tomorrow.day) + " " + tomorrow.strftime("%B %Y")

cal_today    = [c for c in calendar if datetime.fromisoformat(c["start"]).date() == today]
cal_tomorrow = [c for c in calendar if datetime.fromisoformat(c["start"]).date() == tomorrow]

inbox_for_api = [{k: v for k, v in m.items() if k != "entry_id"} for m in inbox]

SYSTEM = """You are Kevin's morning inbox briefing assistant at Oxford University Personnel Services.
Your ONLY job is to write the context paragraph. You do not categorise emails. You do not produce cards.
Return ONLY a valid JSON object with exactly two fields - no preamble, no markdown, no code fences.
Use only plain ASCII punctuation: use - instead of dashes, use ' instead of curly quotes.

Return exactly this:
{
  "context": "<A dense, specific 5-7 sentence morning briefing for Kevin. Must include: full names and exact return dates of every absent colleague; which specific projects, systems or cases are blocked because of those absences; any emails waiting more than 48 hours without a response; the most time-critical deadline this week with its exact date; the one thing Kevin should open first. Use real names, real dates, real case numbers and real project names from the data. Every sentence must contain at least one specific proper noun. Do not generalise. Do not mention GitHub, CI/CD, or workflow authentication issues.>",
  "subtitle": "<one short phrase describing the day>"
}"""

USER = f"""Today is {today_str}. Tomorrow (next working day) is {tomorrow_str}.

INBOX ({len(inbox_for_api)} emails, last 7 days):
{json.dumps(inbox_for_api, indent=2, ensure_ascii=True)}

SENT ({len(sent)} items, last 7 days):
{json.dumps(sent, indent=2, ensure_ascii=True)}

CALENDAR TODAY:
{json.dumps(cal_today, indent=2, ensure_ascii=True)}

CALENDAR TOMORROW:
{json.dumps(cal_tomorrow, indent=2, ensure_ascii=True)}
"""

client   = anthropic.Anthropic(timeout=60.0)
response = client.messages.create(
    model      = "claude-haiku-4-5",
    max_tokens = 1024,
    system     = SYSTEM,
    messages   = [{"role": "user", "content": USER}]
)

raw_text = response.content[0].text.strip()
if raw_text.startswith("```"):
    raw_text = "\n".join(raw_text.split("\n")[1:])
if raw_text.endswith("```"):
    raw_text = "\n".join(raw_text.split("\n")[:-1])

ai_output = json.loads(raw_text)
context  = ai_output.get("context", "")
subtitle = ai_output.get("subtitle", "")
print("Phase 2 done - context written")

# Ã¢ÂÂÃ¢ÂÂ Phase 3 Ã¢ÂÂ Python builds every card Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂ
print("Phase 3 - building cards from inbox...")

# Categorisation rules Ã¢ÂÂ applied in order, first match wins
# importance: 0=low, 1=normal, 2=high
URGENT_SENDERS   = []  # add sender email fragments here if needed
URGENT_SUBJECTS  = ["major incident", "priority 1", "p1", "urgent", "critical", "security vulnerab"]
NEEDS_SUBJECTS   = ["re:", "fw:", "fwd:", "action", "required", "please", "timeline", "update",
                    "chasing", "waiting", "overdue", "follow", "scoping", "handover", "error",
                    "import", "failed", "issue", "case ", "support"]
FYI_SUBJECTS     = ["fyi", "notification", "scheduled", "maintenance", "summary", "workshop",
                    "invitation", "invite", "digest", "recap", "newsletter", "annual leave",
                    "out of office", "automatic reply", "accepted:", "declined:", "cancelled:"]
LOW_SUBJECTS     = ["unsubscribe", "noreply", "no-reply", "do not reply", "automated",
                    "github", "pages", "build", "deploy", "run failed", "wisp"]

def categorise(msg):
    subj    = (msg.get("subject") or "").lower()
    sender  = (msg.get("from_email") or "").lower()
    is_read = msg.get("is_read", True)
    imp     = msg.get("importance", 1)

    # High importance flag always pushes to urgent
    if imp == 2:
        return "urgent"

    # Subject keyword matching
    for kw in LOW_SUBJECTS:
        if kw in subj or kw in sender:
            return "low"
    for kw in URGENT_SUBJECTS:
        if kw in subj:
            return "urgent"
    # Unread + needs keywords Ã¢ÂÂ needs response
    if not is_read:
        for kw in NEEDS_SUBJECTS:
            if kw in subj:
                return "needs"
    for kw in FYI_SUBJECTS:
        if kw in subj:
            return "fyi"
    # Unread with no other match Ã¢ÂÂ needs response
    if not is_read:
        return "needs"
    # Read with no match Ã¢ÂÂ fyi
    return "fyi"

def badge_for(msg, category):
    imp  = msg.get("importance", 1)
    received = msg.get("received", "")
    age_hrs = 0
    try:
        t = datetime.fromisoformat(received.split("+")[0].split(" (")[0].strip())
        age_hrs = (datetime.now() - t).total_seconds() / 3600
    except:
        pass

    if category == "urgent":
        return "Act today", "red"
    if category == "needs":
        if age_hrs > 48:
            return "Overdue", "red"
        return "Reply within 48hrs", "yellow"
    if category == "fyi":
        return "FYI", "gray"
    return "", "gray"

def make_card(msg, category):
    subj    = msg.get("subject") or "(no subject)"
    sender  = msg.get("from") or ""
    preview = (msg.get("body_preview") or "").strip()
    badge, badge_type = badge_for(msg, category)

    title = subj
    sub   = f"From <strong>{sender}</strong>."
    if preview:
        sub += f" {html.escape(preview[:120])}"

    received_str = ""
    try:
        rec = msg.get("received", "")
        rec_dt = datetime.fromisoformat(rec.split("+")[0].split(" (")[0].strip())
        received_str = str(rec_dt.day) + rec_dt.strftime(" %b")
    except:
        pass
    card = {
        "title":     title,
        "sub":       sub,
        "badge":     badge,
        "badgeType": badge_type,
        "subject":   subj,
        "entry_id":  msg.get("entry_id", ""),
        "received":  received_str
    }
    return card

urgent = []
needs  = []
fyi    = []
low    = []

for msg in inbox:
    cat  = categorise(msg)
    card = make_card(msg, cat)
    if cat == "urgent":
        urgent.append(card)
    elif cat == "needs":
        needs.append(card)
    elif cat == "fyi":
        fyi.append(card)
    else:
        low.append(card)

print(f"Phase 3 done - urgent:{len(urgent)} needs:{len(needs)} fyi:{len(fyi)} low:{len(low)}")

# Ã¢ÂÂÃ¢ÂÂ Calendar post-processing Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂ
KNOWN_ABSENCES = [
    {
        "triggers": ["marie", "cooksey"],
        "title": "Annual Leave - Marie Cooksey",
        "sub": "Marie is on leave 8-13 June. Any items requiring her approval or sign-off must wait until she returns. Kevin and Chris are covering H&S support queue and OSM escalations.",
        "alert": "Marie unavailable all week - action DTP1092 comments and volunteer reporting queries independently"
    },
    {
        "triggers": ["james", "salas", "guillen"],
        "title": "Annual Leave - James Salas Guillen",
        "sub": "James is on leave until 18 June. DSE/Cardinus archiving, SBS users in feed and applicant data work all on hold. Handover document received Fri 6 Jun.",
        "alert": "James away until 18 June - Kevin and Chris covering OSM tickets and H&S support queue"
    }
]

def build_cal_items(items):
    result = []
    for item in items:
        start = item.get("start", "")
        try:
            t = datetime.fromisoformat(start)
            time_str = "All day" if item.get("all_day") else t.strftime("%H:%M")
        except:
            time_str = "All day" if item.get("all_day") else ""

        title = item.get("subject", "")
        sub   = item.get("organizer", "") or ""
        alert = ""

        # Check known absences
        title_lower = title.lower()
        for absence in KNOWN_ABSENCES:
            if any(tr in title_lower for tr in absence["triggers"]):
                time_str = "All day"
                title    = absence["title"]
                sub      = absence["sub"]
                alert    = absence["alert"]
                break

        cal_item = {"time": time_str, "title": title, "sub": sub}
        if alert:
            cal_item["alert"] = alert
        result.append(cal_item)
    return result

# Detect absences from calendar Ã¢ÂÂ all-day leave events spanning today
ABSENCE_KEYWORDS = ["annual leave", "a/l", "on leave", "out of office", "holiday"]
absence_set = set()
for item in calendar:
    if not item.get("all_day"):
        continue
    subj_lower = (item.get("subject") or "").lower()
    if not any(kw in subj_lower for kw in ABSENCE_KEYWORDS):
        continue
    try:
        item_start = datetime.fromisoformat(item["start"]).date()
        item_end   = datetime.fromisoformat(item["end"]).date()
        # Outlook all-day end date is exclusive (midnight next day), so use < not <=
        if item_start <= today < item_end:
            # Extract name from subject Ã¢ÂÂ strip the keyword portion
            name = item.get("subject", "")
            for kw in ["- Annual Leave", "- A/L", "- On Leave", "- Out of Office", "- Holiday",
                       "Annual Leave -", "A/L -", "Annual Leave", "A/L"]:
                name = name.replace(kw, "").strip()
            if name:
                absence_set.add(name)
    except:
        continue

absences = sorted(list(absence_set))

# Always include hardcoded known absences if their date range covers today
KNOWN_ABSENCE_DATES = [
    {"name": "Marie Cooksey",       "from": "2026-06-08", "to": "2026-06-13"},
    {"name": "James Salas Guillen", "from": "2026-06-08", "to": "2026-06-18"},
]
for ka in KNOWN_ABSENCE_DATES:
    ka_start = datetime.strptime(ka["from"], "%Y-%m-%d").date()
    ka_end   = datetime.strptime(ka["to"],   "%Y-%m-%d").date()
    if ka_start <= today <= ka_end:
        # Remove partial matches so full name replaces partial calendar extract
        absences = [a for a in absences if not ka["name"].startswith(a) and not a.startswith(ka["name"])]
        absences.append(ka["name"])
absences = sorted(absences)

# Priority actions Ã¢ÂÂ pulled from Command Centre tasks.json
COMMAND_CENTRE_REPO = "begb0037admin/command-centre"
COMMAND_CENTRE_PATH = "data/tasks.json"
priorities_today = []
priorities_week  = []
try:
    cc_url     = f"https://api.github.com/repos/{COMMAND_CENTRE_REPO}/contents/{COMMAND_CENTRE_PATH}"
    cc_headers = {
        "Authorization": f"token {GITHUB_PAT}",
        "Content-Type":  "application/json",
        "User-Agent":    "work-inbox-script"
    }
    cc_req = urllib.request.Request(cc_url, headers=cc_headers)
    with urllib.request.urlopen(cc_req) as r:
        cc_data    = json.loads(r.read())
        cc_content = json.loads(base64.b64decode(cc_data["content"]).decode("utf-8"))
    for task in cc_content.get("tasks", []):
        tier = task.get("tier", "")
        entry = {
            "text":        task.get("title", ""),
            "description": task.get("description", ""),
            "actions":     task.get("actions", []),
            "source":      task.get("source", ""),
            "dateType":    "red" if tier == "today" else "orange"
        }
        if tier == "today":
            priorities_today.append(entry)
        elif tier == "week":
            priorities_week.append(entry)
    print(f"Command Centre loaded - today:{len(priorities_today)} week:{len(priorities_week)}")
except Exception as e:
    print(f"WARNING: Could not load Command Centre tasks - {e}")
    priorities_today = []
    priorities_week  = []

# Ã¢ÂÂÃ¢ÂÂ Assemble final briefing Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂ
briefing = {
    "date":         today_str,
    "subtitle":     subtitle,
    "context":      context,
    "urgent":       urgent,
    "needs":        needs,
    "fyi":          fyi,
    "low":          low,
    "calToday":     build_cal_items(cal_today),
    "calTomorrow":  build_cal_items(cal_tomorrow),
    "absences":     absences,
    "prioritiesToday": priorities_today,
    "prioritiesWeek":  priorities_week,
    "refreshed_at": datetime.now().strftime("%A %d %B · %H:%M")
}

# Ã¢ÂÂÃ¢ÂÂ Phase 4 Ã¢ÂÂ push to GitHub Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂ
print("Phase 4 - pushing briefing to GitHub...")

if not GITHUB_PAT:
    print("ERROR: GITHUB_PAT env var not set - cannot push.")
else:
    try:
        api_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_PATH}"
        headers = {
            "Authorization": f"token {GITHUB_PAT}",
            "Content-Type":  "application/json",
            "User-Agent":    "work-inbox-script"
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
            print(f"Phase 4 done - briefing pushed to GitHub (commit: {result.get('commit',{}).get('sha','?')[:7]})")
    except Exception as e:
        print(f"Phase 4 FAILED - {e}")
        raise
