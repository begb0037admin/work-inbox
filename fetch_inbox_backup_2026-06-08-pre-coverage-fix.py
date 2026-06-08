import json, os, base64, urllib.request, urllib.error
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

# ── Phase 1 — pull Outlook data ──────────────────────────────────────────────
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

unread_total = sum(1 for m in inbox if not m["is_read"])
print(f"Phase 1 done - inbox:{len(inbox)} (unread:{unread_total}) sent:{len(sent)} calendar:{len(calendar)}")

# ── Phase 2 — Anthropic triage ───────────────────────────────────────────────
print("Phase 2 - calling Anthropic API...")

now          = datetime.now()
today_str    = now.strftime("%A") + " " + str(now.day) + " " + now.strftime("%B %Y")
tomorrow_str = tomorrow.strftime("%A") + " " + str(tomorrow.day) + " " + tomorrow.strftime("%B %Y")

cal_today    = [c for c in calendar if datetime.fromisoformat(c["start"]).date() == today]
cal_tomorrow = [c for c in calendar if datetime.fromisoformat(c["start"]).date() == tomorrow]

inbox_for_api = [{k: v for k, v in m.items() if k != "entry_id"} for m in inbox]

SYSTEM = """You are Kevin's morning inbox triage assistant at Oxford University Personnel Services.
Analyse the inbox, sent items, and calendar data provided and return ONLY a valid JSON object - no preamble, no markdown, no code fences.
Use only plain ASCII punctuation: use - instead of dashes, use ' instead of curly quotes, use ... instead of ellipsis.

The JSON must match this exact schema:
{
  "date": "<Day D Month YYYY>",
  "subtitle": "<one short phrase describing the day, optional>",
  "context": "<A dense, specific 5-7 sentence morning briefing written for Kevin. Must include: full names and exact return dates of every absent colleague; which specific projects, systems or cases are blocked or at risk because of those absences; any emails waiting more than 48 hours without a response; the most time-critical deadline this week with its exact date; and the one thing Kevin should open first. Use real names, real dates, real case numbers and real project names from the inbox data. Every sentence must contain at least one specific proper noun from the data. Do not generalise. Do not use vague phrases like several items require attention. Do not mention GitHub PAT exposure, token revocation, workflow authentication failures, or CI/CD pipeline issues - these are handled separately.",
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
- calToday/calTomorrow time field: use "All day" for all-day events, never use date ranges like "All week 8-13 June"
- calToday/calTomorrow title: format as "Event Type - Full Name" e.g. "Annual Leave - Marie Cooksey"
- calToday/calTomorrow sub: be specific - state exact dates, what is blocked, who is covering. e.g. "Marie is on leave 8-13 June. Any items involving her approval or input must wait until she returns. Kevin and Chris are covering H&S support queue and OSM escalations." Never write vague text like "Check handover documents and note absence for escalations."
- calToday/calTomorrow alert: if the person is absent and has active dependencies, name the specific projects or actions affected. e.g. "Marie unavailable all week - action DTP1092 comments and volunteer reporting queries independently". Never write generic text like "Colleague absent" or "Team member absent".
- omit alert only if there are genuinely no active dependencies or actions affected

CALENDAR ITEM EXAMPLE - for an absent manager or colleague, output MUST look exactly like this:
CORRECT:
{"time":"All day","title":"Annual Leave - Marie Cooksey","sub":"Marie is on leave 8-13 June. Any items requiring her approval or sign-off must wait. Kevin and Chris are covering H&S support queue and OSM escalations in her absence.","alert":"Marie unavailable all week - action DTP1092 comments and volunteer reporting queries independently"}

NEVER produce output like this - it is wrong and useless:
{"time":"All week","title":"Marie Cooksey annual leave","sub":"Away 8-12 June. Handover documentation provided.","alert":"Team member absent"}

The difference: correct output names specific projects, specific coverage arrangements, specific actions. Wrong output is generic and tells Kevin nothing he does not already know. Every calendar item for an absent colleague must meet the CORRECT standard above.
- absences: only include people confirmed absent, inferred from out-of-office replies or calendar blocks
- calendar shows working days only (Monday to Friday)
- calToday/calTomorrow: if a colleague is known to be absent from OOO replies, handover emails or inbox context - even if no calendar block exists - include them as an All day entry. Use the full name from email context, not abbreviated calendar titles like "Annual Leave - Marie". Cross-reference all data sources to build the most complete and accurate picture.
- subject field: you MUST copy the exact email subject verbatim - character for character - from the inbox data provided. Do NOT paraphrase, summarise or invent a subject. The subject field is used to look up the Outlook EntryID and open the exact email. If you change even one character it will fail. If the item does not relate to a single specific email, omit the subject field entirely.
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

client   = anthropic.Anthropic(timeout=60.0)
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

# Inject EntryIDs for click-through
subject_to_entryid = {}
for m in inbox:
    if m.get("entry_id") and m.get("subject"):
        subject_to_entryid[m["subject"]] = m["entry_id"]

def inject_entry_ids(items):
    if not items:
        return items
    for item in items:
        subj = item.get("subject", "")
        if not subj:
            continue
        if subj in subject_to_entryid:
            item["entry_id"] = subject_to_entryid[subj]
        else:
            subj_lower = subj.lower()
            for inbox_subj, entry_id in subject_to_entryid.items():
                inbox_lower = inbox_subj.lower()
                if subj_lower in inbox_lower or inbox_lower in subj_lower:
                    item["entry_id"] = entry_id
                    break
    return items

for section in ["urgent", "needs", "fyi", "low", "priorities"]:
    inject_entry_ids(briefing.get(section, []))

# Post-process calendar — rewrite sub/alert for known absent colleagues
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

def fix_cal_items(items):
    if not items:
        return items
    for item in items:
        title = (item.get("title") or "").lower()
        for absence in KNOWN_ABSENCES:
            if any(trigger in title for trigger in absence["triggers"]):
                item["time"]  = "All day"
                item["title"] = absence["title"]
                item["sub"]   = absence["sub"]
                item["alert"] = absence["alert"]
                break
    return items

for cal_key in ["calToday", "calTomorrow"]:
    fix_cal_items(briefing.get(cal_key, []))

briefing["refreshed_at"] = datetime.now().strftime("%A %d %B · %H:%M")

print("Phase 2 done - briefing assembled in memory")

# ── Phase 3 — push briefing to GitHub ────────────────────────────────────────
print("Phase 3 - pushing briefing to GitHub...")

if not GITHUB_PAT:
    print("ERROR: GITHUB_PAT env var not set - cannot push. Set it and re-run.")
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
        print(f"Phase 3 FAILED - GitHub push error: {e}")
        raise
