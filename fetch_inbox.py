import json, os, base64, re, urllib.request, urllib.error
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
            raise Exception("Filter too large")
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
                "subject":   item.Subject,
                "start":     str(item.Start),
                "end":       str(item.End),
                "location":  item.Location,
                "organizer": item.Organizer,
                "all_day":   item.AllDayEvent
            })
    except:
        continue

unread_total = sum(1 for m in inbox if not m["is_read"])
print(f"Phase 1 done - inbox:{len(inbox)} (unread:{unread_total}) sent:{len(sent)} calendar:{len(calendar)}")

# ── Phase 2 — AI writes context paragraph only ───────────────────────────────
print("Phase 2 - calling Anthropic API for context...")

now          = datetime.now()
today_str    = now.strftime("%A") + " " + str(now.day) + " " + now.strftime("%B %Y")
tomorrow_str = tomorrow.strftime("%A") + " " + str(tomorrow.day) + " " + tomorrow.strftime("%B %Y")

cal_today    = [c for c in calendar if datetime.fromisoformat(c["start"]).date() == today]
cal_tomorrow = [c for c in calendar if datetime.fromisoformat(c["start"]).date() == tomorrow]

inbox_for_api = [{k: v for k, v in m.items() if k != "entry_id"} for m in inbox]

SYSTEM = """You are Kevin's morning inbox briefing assistant at Oxford University Personnel Services.
Your ONLY job is to write the context paragraph and subtitle. You do not categorise emails. You do not produce cards. You do not decide what Kevin sees.
Return ONLY a valid JSON object with exactly two fields. No preamble, no markdown, no code fences.
Use only plain ASCII punctuation: use - instead of dashes, use ' instead of curly quotes.

Return exactly this structure:
{
  "context": "<5-7 sentence morning briefing for Kevin. Must include: full names and exact return dates of every absent colleague; which specific projects, systems or cases are blocked; any emails waiting 48hrs+ without response; most time-critical deadline with exact date; one thing Kevin should open first. Use real names, real dates, real case numbers from the data. Every sentence must contain at least one specific proper noun. Do not generalise. Do not mention GitHub, CI/CD, or workflow issues.>",
  "subtitle": "<one short phrase describing the day>"
}"""

USER = f"""Today is {today_str}. Tomorrow is {tomorrow_str}.

INBOX ({len(inbox_for_api)} emails):
{json.dumps(inbox_for_api, indent=2, ensure_ascii=True)}

SENT ({len(sent)} items):
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

ai_out   = json.loads(raw_text)
context  = ai_out.get("context", "")
subtitle = ai_out.get("subtitle", "")
print("Phase 2 done - context written")

# ── Phase 3 — Python builds every card ───────────────────────────────────────
print("Phase 3 - grouping threads and building cards...")

def base_subject(subj):
    s = (subj or "").strip()
    s = re.sub(r'^(re|fw|fwd)\s*:\s*', '', s, flags=re.IGNORECASE)
    s = re.sub(r'^(re|fw|fwd)\s*:\s*', '', s, flags=re.IGNORECASE)
    return s.strip().lower()

# Group by thread
threads = {}
for msg in inbox:
    key = base_subject(msg.get("subject", ""))
    if key not in threads:
        threads[key] = []
    threads[key].append(msg)

# Sort each thread newest first
thread_list = []
for key, msgs in threads.items():
    msgs.sort(key=lambda x: x.get("received", ""), reverse=True)
    thread_list.append(msgs)

# Sort threads: unread first, then newest
thread_list.sort(
    key=lambda msgs: (msgs[0].get("is_read", True), msgs[0].get("received", "")),
    reverse=True
)

URGENT_KW = ["major incident", "priority 1", "p1 ", "urgent", "critical", "security vulnerab"]
NEEDS_KW  = ["action", "required", "please", "timeline", "chasing", "waiting", "overdue",
             "follow", "scoping", "handover", "error", "import", "failed", "issue",
             "case ", "support", "re:", "fw:", "fwd:"]
FYI_KW    = ["fyi", "notification", "scheduled", "maintenance", "summary", "workshop",
             "invitation", "invite", "digest", "recap", "newsletter", "annual leave",
             "out of office", "automatic reply", "accepted:", "declined:", "cancelled:"]
LOW_KW    = ["unsubscribe", "noreply", "no-reply", "do not reply", "automated",
             "github", "pages", "build", "deploy", "run failed", "wisp", "bulletin"]

def categorise(msg):
    subj   = (msg.get("subject") or "").lower()
    sender = (msg.get("from_email") or "").lower()
    is_read = msg.get("is_read", True)
    imp    = msg.get("importance", 1)
    if imp == 2:
        return "urgent"
    for kw in LOW_KW:
        if kw in subj or kw in sender:
            return "low"
    for kw in URGENT_KW:
        if kw in subj:
            return "urgent"
    if not is_read:
        for kw in NEEDS_KW:
            if kw in subj:
                return "needs"
    for kw in FYI_KW:
        if kw in subj:
            return "fyi"
    if not is_read:
        return "needs"
    return "fyi"

CAT_RANK = {"urgent": 0, "needs": 1, "fyi": 2, "low": 3}

def thread_category(msgs):
    return min([categorise(m) for m in msgs], key=lambda c: CAT_RANK[c])

def badge_for(category, msgs):
    if category == "urgent":
        return "Act today", "red"
    if category == "needs":
        try:
            oldest = min(msgs, key=lambda x: x.get("received", ""))
            t = datetime.fromisoformat(oldest["received"].split("+")[0].split(" (")[0].strip())
            if (datetime.now() - t).total_seconds() / 3600 > 48:
                return "Overdue", "red"
        except:
            pass
        return "Reply within 48hrs", "yellow"
    if category == "fyi":
        return "FYI", "gray"
    return "", "gray"

def make_card(msgs, category):
    latest  = msgs[0]
    subj    = latest.get("subject") or "(no subject)"
    sender  = latest.get("from") or ""
    preview = (latest.get("body_preview") or "").strip()
    count   = len(msgs)
    badge, badge_type = badge_for(category, msgs)

    title = base_subject(subj)
    title = title[0].upper() + title[1:] if title else subj

    sub = f"From <strong>{sender}</strong>"
    if count > 1:
        sub += f" &middot; <span style='color:#6b7280;font-size:12px'>{count} emails</span>"
    sub += "."
    if preview:
        sub += f" {preview[:120]}"

    return {
        "title":     title,
        "sub":       sub,
        "badge":     badge,
        "badgeType": badge_type,
        "subject":   latest.get("subject", ""),
        "entry_id":  latest.get("entry_id", "")
    }

urgent, needs, fyi, low = [], [], [], []

for msgs in thread_list:
    cat  = thread_category(msgs)
    card = make_card(msgs, cat)
    {"urgent": urgent, "needs": needs, "fyi": fyi, "low": low}[cat].append(card)

print(f"Phase 3 done - {len(thread_list)} threads: urgent:{len(urgent)} needs:{len(needs)} fyi:{len(fyi)} low:{len(low)}")

# ── Calendar processing ───────────────────────────────────────────────────────
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
        try:
            t = datetime.fromisoformat(item["start"])
            time_str = "All day" if item.get("all_day") else t.strftime("%H:%M")
        except:
            time_str = "All day" if item.get("all_day") else ""

        title = item.get("subject", "")
        sub   = item.get("organizer", "") or ""
        alert = ""

        for absence in KNOWN_ABSENCES:
            if any(tr in title.lower() for tr in absence["triggers"]):
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

# Detect absences from OOO emails
absence_set = set()
for msg in inbox:
    subj    = (msg.get("subject") or "").lower()
    preview = (msg.get("body_preview") or "").lower()
    for kw in ["out of office", "annual leave", "on leave", "away until", "a/l"]:
        if kw in subj or kw in preview:
            absence_set.add(msg.get("from", ""))
            break

# Top 5 priorities from urgent + needs
priorities = []
for card in (urgent + needs)[:5]:
    priorities.append({
        "text":     card["title"],
        "date":     card["badge"],
        "dateType": card["badgeType"],
        "subject":  card.get("subject", ""),
        "entry_id": card.get("entry_id", "")
    })

# ── Assemble briefing ─────────────────────────────────────────────────────────
briefing = {
    "date":        today_str,
    "subtitle":    subtitle,
    "context":     context,
    "urgent":      urgent,
    "needs":       needs,
    "fyi":         fyi,
    "low":         low,
    "calToday":    build_cal_items(cal_today),
    "calTomorrow": build_cal_items(cal_tomorrow),
    "absences":    sorted(list(absence_set)),
    "priorities":  priorities,
    "refreshed_at": datetime.now().strftime("%A %d %B · %H:%M")
}

# ── Phase 4 — push to GitHub ──────────────────────────────────────────────────
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
        req = urllib.request.Request(api_url, data=json.dumps(payload).encode("utf-8"),
                                     headers=headers, method="PUT")
        with urllib.request.urlopen(req) as r:
            result = json.loads(r.read())
            print(f"Phase 4 done - briefing pushed to GitHub (commit: {result.get('commit',{}).get('sha','?')[:7]})")
    except Exception as e:
        print(f"Phase 4 FAILED - {e}")
        raise
