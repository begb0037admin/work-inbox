import json, os, base64, html, re, urllib.request, urllib.error, subprocess
from datetime import datetime, timedelta
import win32com.client
import anthropic

# Suppress Windows git gc --auto interactive prompts
subprocess.run(["git", "config", "gc.auto", "0"], capture_output=True,
               cwd=os.path.dirname(os.path.abspath(__file__)))

GITHUB_REPO = "begb0037admin/work-inbox"
GITHUB_PATH = "data/briefing.json"
GITHUB_PAT  = os.environ.get("GITHUB_PAT", "")
GITHUB_TIMEOUT = 30

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

# -- Phase 1 -- pull Outlook data --
print("Phase 1 - pulling Outlook data...")
inbox = []
unread_count = 0
read_count   = 0
MAX_UNREAD   = 50
MAX_READ     = 30

# VIP senders -- always captured regardless of cap
VIP_NAMES = {
    'Athena Artuso','Marie Cooksey','Sarah Rowles','Simon Burford',
    'Asta Palmer','James Salas Guillen',"Michael O'Sullivan",
    'Anna Carter-Windle','Anthony Kong','Beth Gray','Christopher Sanders',
    'David Johnson','Emma Fitz-Gibbon','Henry Acheampong','Iyanuloluwa Akinsanya',
    'Julie Hickman','Marie King','Michelle Williams','Nathan Kirwan',
    'Susan Pratt','Anne Mortimer','Nicholas Chandler','Steve McBrearty',
}
VIP_EMAILS = {
    'tony.boydell@it.ox.ac.uk','erika.braverman@it.ox.ac.uk',
    'hr.systems@admin.ox.ac.uk','support.access@theaccessgroup.com',
    'edward.demetillo@cority.com','crispin.muncaster@it.ox.ac.uk',
    'christopher.sanders@admin.ox.ac.uk','henry.acheampong@admin.ox.ac.uk',
    'iyanuloluwa.akinsanya@tss.ox.ac.uk',
}

def is_vip(msg):
    try:
        return (msg.SenderName or '').strip() in VIP_NAMES or \
               (msg.SenderEmailAddress or '').lower().strip() in VIP_EMAILS
    except:
        return False

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

# VIP sweep -- pick up any VIP emails missed by the cap
captured_ids = {e["entry_id"] for e in inbox}
for msg in restrict_date(mapi.GetDefaultFolder(6), cutoff):
    try:
        if msg.EntryID in captured_ids:
            continue
        if not is_vip(msg):
            continue
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
        captured_ids.add(msg.EntryID)
    except:
        continue

inbox.sort(key=lambda x: (not x["is_read"], x["received"]), reverse=True)
print(f"Phase 1 VIP sweep done - total inbox now: {len(inbox)}")

sent = []
for msg in mapi.GetDefaultFolder(5).Items:
    try:
        t = dt(msg.SentOn)
        if t and t >= cutoff:
            sent.append({
                "subject":      msg.Subject,
                "to":           msg.To,
                "sent":         str(msg.SentOn),
                "body_preview": (msg.Body or "")[:100],
                "entry_id":     msg.EntryID
            })
    except:
        continue

week_end = today + timedelta(days=6)
lookback  = today - timedelta(days=30)  # catch multi-day absences spanning today
calendar = []
_cal_items = mapi.GetDefaultFolder(9).Items
_cal_items.IncludeRecurrences = True
_cal_items.Sort("[Start]")
for item in _cal_items:
    try:
        t = dt(item.Start)
        if not t:
            continue
        if t.date() > week_end:
            break
        if t.date() < lookback:
            continue
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

# -- Phase 2 -- AI writes context paragraph only --
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

# -- Phase 3 -- Python builds every card --
print("Phase 3 - building cards from inbox...")

# Categorisation rules -- applied in order, first match wins
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
    # Unread + needs keywords -- needs response
    if not is_read:
        for kw in NEEDS_SUBJECTS:
            if kw in subj:
                return "needs"
    for kw in FYI_SUBJECTS:
        if kw in subj:
            return "fyi"
    # Unread with no other match -- needs response
    if not is_read:
        return "needs"
    # Read with no match -- fyi
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
    preview = re.sub(r"<\?\s*https?://\S+>?", "[link]", preview)
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
        "from":      sender,
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

# -- Calendar post-processing --
KNOWN_ABSENCES = []

def build_cal_items(items):
    result = []
    items = sorted(items, key=lambda x: x.get("start", ""))
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

# Detect absences from calendar -- today's leave and next working day
ABSENCE_KEYWORDS = ["annual leave", "a/l", "on leave", "out of office", "holiday"]
absence_set = set()
is_weekend = today.weekday() >= 5
today_suffix = " (next week)" if is_weekend else ""
tomorrow_label = "tomorrow" if tomorrow == today + timedelta(days=1) else "next week"

def _extract_absence_name(subject):
    name = subject
    for kw in ["- Annual Leave", "- A/L", "- On Leave", "- Out of Office", "- Holiday",
               "Annual Leave -", "A/L -", "Annual Leave", "A/L"]:
        name = name.replace(kw, "").strip()
    return name

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
            name = _extract_absence_name(item.get("subject", ""))
            if name:
                absence_set.add(name + today_suffix)
        elif item_start == tomorrow:
            name = _extract_absence_name(item.get("subject", ""))
            if name:
                absence_set.add(f"{name} ({tomorrow_label})")
    except:
        continue

absences = sorted(list(absence_set))

# Always include hardcoded known absences if their date range covers today
KNOWN_ABSENCE_DATES = []
for ka in KNOWN_ABSENCE_DATES:
    ka_start = datetime.strptime(ka["from"], "%Y-%m-%d").date()
    ka_end   = datetime.strptime(ka["to"],   "%Y-%m-%d").date()
    if ka_start <= today <= ka_end:
        absences = [a for a in absences if not ka["name"].startswith(a) and not a.startswith(ka["name"])]
        absences.append(ka["name"])
absences = sorted(absences)

# Priority actions -- pulled from Command Centre tasks.json
COMMAND_CENTRE_REPO = "begb0037admin/command-centre"
COMMAND_CENTRE_PATH = "data/tasks.json"
priorities_today    = []
priorities_tomorrow = []
priorities_week     = []
cc_content = []
try:
    cc_url     = f"https://api.github.com/repos/{COMMAND_CENTRE_REPO}/contents/{COMMAND_CENTRE_PATH}"
    cc_headers = {
        "Authorization": f"token {GITHUB_PAT}",
        "Content-Type":  "application/json",
        "User-Agent":    "work-inbox-script"
    }
    cc_req = urllib.request.Request(cc_url, headers=cc_headers)
    with urllib.request.urlopen(cc_req, timeout=GITHUB_TIMEOUT) as r:
        cc_data    = json.loads(r.read())
        cc_content = json.loads(base64.b64decode(cc_data["content"]).decode("utf-8"))
    task_list = cc_content if isinstance(cc_content, list) else cc_content.get("tasks", [])
    for task in task_list:
        if task.get("done"):
            continue
        tier = task.get("tier", "")
        entry = {
            "id":          task.get("id", ""),
            "text":        task.get("title", ""),
            "description": task.get("description", ""),
            "actions":     task.get("actions", []),
            "source":      task.get("source", ""),
            "dateType":    "red" if tier == "today" else "orange"
        }
        if tier == "today":
            priorities_today.append(entry)
        elif tier == "tomorrow":
            priorities_tomorrow.append(entry)
        elif tier == "week":
            priorities_week.append(entry)
    print(f"Command Centre loaded - today:{len(priorities_today)} tomorrow:{len(priorities_tomorrow)} week:{len(priorities_week)}")
except Exception as e:
    print(f"WARNING: Could not load Command Centre tasks - {e}")
    priorities_today    = []
    priorities_tomorrow = []
    priorities_week     = []

# Phase 3.5 - AI triage: which emails should become Command Centre tasks
print("Phase 3.5 - triaging inbox for task suggestions...")
suggestions = {
    "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
    "new_tasks":    [],
    "task_updates": []
}
# Dedupe ledger - emails already applied to Command Centre tasks
ledger = {"applied": {}}
if GITHUB_PAT:
    try:
        _lurl = f"https://api.github.com/repos/{GITHUB_REPO}/contents/data/triage_ledger.json"
        _lreq = urllib.request.Request(_lurl, headers={"Authorization": f"token {GITHUB_PAT}", "User-Agent": "work-inbox-script"})
        with urllib.request.urlopen(_lreq) as r:
            ledger = json.loads(base64.b64decode(json.loads(r.read())["content"]).decode("utf-8"))
        if "applied" not in ledger:
            ledger["applied"] = {}
    except Exception:
        pass
try:
    task_summaries = []
    task_list = cc_content if isinstance(cc_content, list) else cc_content.get("tasks", [])
    for t in task_list:
        task_summaries.append({
            "id":          t.get("id", ""),
            "title":       t.get("title", ""),
            "description": (t.get("description") or "")[:300],
            "emailRef":    t.get("emailRef", "")
        })
    if not task_summaries:
        raise Exception("Command Centre tasks unavailable - skipping triage")

    email_candidates = []
    for m in inbox:
        if categorise(m) in ("urgent", "needs"):
            email_candidates.append({
                "subject":      m.get("subject", ""),
                "from":         m.get("from", ""),
                "received":     (m.get("received", "") or "")[:16],
                "body_preview": re.sub(r"<\?\s*https?://\S+>?", "[link]", (m.get("body_preview") or ""))[:150],
                "entry_id":     m.get("entry_id", "")
            })

    for s in sent[:30]:
        email_candidates.append({
            "subject":      s.get("subject", ""),
            "from":         "Kevin (sent to: " + (s.get("to") or "") + ")",
            "received":     (s.get("sent", "") or "")[:16],
            "body_preview": re.sub(r"<\?\s*https?://\S+>?", "[link]", (s.get("body_preview") or ""))[:150],
            "entry_id":     s.get("entry_id", ""),
            "direction":    "sent"
        })

    api_emails = [{"n": i, "direction": e.get("direction", "received"),
                   "subject": e["subject"], "from": e["from"],
                   "received": e["received"], "body_preview": e["body_preview"]}
                  for i, e in enumerate(email_candidates)]

    TRIAGE_SYSTEM = (
        "You are Kevin's task triage assistant at Oxford University Personnel Services.\n"
        "You receive his existing Command Centre task list, his recent action-required received emails, and emails Kevin himself sent (direction: sent).\n"
        "Identify:\n"
        "1. new_tasks - emails that represent real, actionable work for Kevin that is NOT covered by any existing task. Be selective. Max 5. "
        "If an email concerns work that any existing task already covers - even partially, even if you would mention that task in your description - it belongs in task_updates with that task's id, NEVER in new_tasks.\n"
        "2. task_updates - emails that are progress, replies or new information on an EXISTING task. Max 8. "
        "A task_update must clearly concern that specific task - same case number, same named project, or same people AND topic. "
        "If no existing task is a clear match, do NOT force one: either propose it under new_tasks or omit it entirely.\n"
        "Return ONLY a valid JSON object - no preamble, no markdown, no code fences. Plain ASCII punctuation only.\n"
        "{\n"
        '  "new_tasks": [{"email_n": <n>, "title": "<short imperative task title>", "tier": "today|tomorrow|week", "description": "<2-3 sentences: what the work is and why, drawn from the email>"}],\n'
        '  "task_updates": [{"email_n": <n>, "task_id": "<existing task id>", "note": "<one sentence: what this email adds to the task>"}]\n'
        "}\n"
        'Rules: tier "today" only if the deadline is today or overdue; "tomorrow" if it must happen the next working day; otherwise "week". '
        "Never invent case numbers or names. Automated notifications, newsletters, calendar "
        "accept/decline messages and out-of-office replies are never tasks. "
        "Use direction=sent emails to log Kevin's own actions on existing tasks as task_updates "
        "(e.g. 'Kevin replied to Reenu with the requested staff list') so the action log shows "
        "both sides of the conversation. Never propose a new task for work that a sent email "
        "shows Kevin has already handled."
    )

    triage_user = (
        f"Today is {today_str}. Tomorrow (next working day) is {tomorrow_str}.\n\n"
        f"EXISTING TASKS:\n{json.dumps(task_summaries, indent=1, ensure_ascii=True)}\n\n"
        f"EMAILS (received urgent/needs + sent by Kevin, last 7 days):\n{json.dumps(api_emails, indent=1, ensure_ascii=True)}"
    )

    t_resp = client.messages.create(
        model      = "claude-haiku-4-5",
        max_tokens = 1500,
        system     = TRIAGE_SYSTEM,
        messages   = [{"role": "user", "content": triage_user}]
    )
    t_raw = t_resp.content[0].text.strip()
    if t_raw.startswith("```"):
        t_raw = "\n".join(t_raw.split("\n")[1:])
    if t_raw.endswith("```"):
        t_raw = "\n".join(t_raw.split("\n")[:-1])
    t_out = json.loads(t_raw)

    task_by_id = {t["id"]: t for t in task_summaries}
    for nt in t_out.get("new_tasks", [])[:5]:
        i = nt.get("email_n")
        if not isinstance(i, int) or not (0 <= i < len(email_candidates)):
            continue
        src = email_candidates[i]
        suggestions["new_tasks"].append({
            "title":         nt.get("title", ""),
            "tier":          nt.get("tier") if nt.get("tier") in ("today", "tomorrow", "week") else "week",
            "description":   nt.get("description", ""),
            "email_subject": src["subject"],
            "email_from":    src["from"],
            "received":      src["received"],
            "entry_id":      src["entry_id"]
        })
    for tu in t_out.get("task_updates", [])[:8]:
        i   = tu.get("email_n")
        tid = tu.get("task_id", "")
        if not isinstance(i, int) or not (0 <= i < len(email_candidates)) or tid not in task_by_id:
            continue
        if email_candidates[i]["entry_id"] + "_" + tid in ledger.get("applied", {}):
            continue
        src = email_candidates[i]
        suggestions["task_updates"].append({
            "task_id":       tid,
            "task_title":    task_by_id[tid]["title"],
            "note":          tu.get("note", ""),
            "email_subject": src["subject"],
            "email_from":    src["from"],
            "received":      src["received"],
            "entry_id":      src["entry_id"]
        })
    print(f"Phase 3.5 done - new:{len(suggestions['new_tasks'])} updates:{len(suggestions['task_updates'])}")
except Exception as e:
    print(f"WARNING: Phase 3.5 triage failed - {e}")


# -- Assemble final briefing --

# Phase 3.6 - apply task updates directly to Command Centre tasks.json
def _gh_get(url, headers):
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=GITHUB_TIMEOUT) as r:
        return json.loads(r.read())

def _gh_put(url, headers, message, content_bytes, sha=None):
    payload = {"message": message,
               "content": base64.b64encode(content_bytes).decode("ascii")}
    if sha:
        payload["sha"] = sha
    req = urllib.request.Request(url, data=json.dumps(payload).encode("utf-8"),
                                 headers=headers, method="PUT")
    with urllib.request.urlopen(req, timeout=GITHUB_TIMEOUT) as r:
        return json.loads(r.read())

if GITHUB_PAT and suggestions["task_updates"]:
    try:
        gh_headers = {"Authorization": f"token {GITHUB_PAT}",
                      "Content-Type":  "application/json",
                      "User-Agent":    "work-inbox-script"}
        cc_tasks_url = f"https://api.github.com/repos/{COMMAND_CENTRE_REPO}/contents/{COMMAND_CENTRE_PATH}"
        cc_meta   = _gh_get(cc_tasks_url, gh_headers)
        tasks_doc = json.loads(base64.b64decode(cc_meta["content"]).decode("utf-8"))

        # Mandatory daily backup before any write to tasks.json
        backup_path = f"Archive/tasks_backup_{datetime.now().strftime('%Y%m%d')}.json"
        backup_url  = f"https://api.github.com/repos/{COMMAND_CENTRE_REPO}/contents/{backup_path}"
        try:
            _gh_get(backup_url, gh_headers)
        except urllib.error.HTTPError as e:
            if e.code != 404:
                raise
            _gh_put(backup_url, gh_headers,
                    f"backup: tasks.json {datetime.now().strftime('%Y-%m-%d')}",
                    base64.b64decode(cc_meta["content"]))
            print(f"Phase 3.6 - daily backup created: {backup_path}")

        stamp   = datetime.now().strftime("%d %b %Y")
        applied = 0
        task_list = tasks_doc if isinstance(tasks_doc, list) else tasks_doc.get("tasks", [])
        for task in task_list:
            for upd in suggestions["task_updates"]:
                if task.get("id") == upd["task_id"]:
                    action_text = f"[{stamp}] {upd['note']} (email: {upd['email_from']} - {upd['email_subject']})"
                    actions = task.setdefault("actions", [])
                    if action_text in actions:
                        break
                    actions.append(action_text)
                    task["entryId"] = upd["entry_id"]
                    applied += 1
                    break
        if applied:
            _gh_put(cc_tasks_url, gh_headers,
                    f"inbox: apply {applied} task update(s) {datetime.now().strftime('%Y-%m-%d %H:%M')}",
                    json.dumps(tasks_doc, indent=2, ensure_ascii=False).encode("utf-8"),
                    cc_meta["sha"])
            print(f"Phase 3.6 done - {applied} update(s) applied to Command Centre")
            for u in suggestions["task_updates"]:
                ledger["applied"][u["entry_id"] + "_" + u["task_id"]] = datetime.now().strftime("%Y-%m-%d")
            ledger_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/data/triage_ledger.json"
            l_sha = None
            try:
                l_sha = _gh_get(ledger_url, gh_headers).get("sha")
            except urllib.error.HTTPError as e:
                if e.code != 404:
                    raise
            _gh_put(ledger_url, gh_headers, "chore: update triage ledger",
                    json.dumps(ledger, indent=1).encode("utf-8"), l_sha)
        suggestions["applied_updates"] = suggestions.pop("task_updates")
    except Exception as e:
        print(f"WARNING: Phase 3.6 apply failed - {e}")


# Phase 3.7 - AI summaries for priority tasks
print("Phase 3.7 - generating AI task summaries...")
all_priorities = priorities_today + priorities_tomorrow + priorities_week
if all_priorities:
    try:
        tasks_for_summary = [
            {
                "id":          e["id"],
                "title":       e["text"],
                "description": (e.get("description") or "")[:300],
                "actions":     e.get("actions", [])[-5:]
            }
            for e in all_priorities if e.get("id")
        ]
        SUMMARY_SYSTEM = (
            "You are Kevin's task briefing assistant at Oxford University Personnel Services.\n"
            "For each task, write a 1-2 sentence status summary: current state, what needs to happen next, any blockers.\n"
            "Be specific - use names, dates and case numbers from the data. Plain ASCII punctuation only.\n"
            "Return ONLY a valid JSON object mapping task id to summary string - no preamble, no markdown.\n"
            "Example: {\"task-001\": \"Awaiting response from Jane Smith re HRIS migration. Next: chase by Friday 20 Jun.\"}"
        )
        summary_user = (
            f"Today is {today_str}.\n\n"
            f"TASKS:\n{json.dumps(tasks_for_summary, indent=1, ensure_ascii=True)}"
        )
        s_resp = client.messages.create(
            model      = "claude-haiku-4-5",
            max_tokens = 4096,
            system     = SUMMARY_SYSTEM,
            messages   = [{"role": "user", "content": summary_user}]
        )
        s_raw = s_resp.content[0].text.strip()
        if s_raw.startswith("```"):
            s_raw = "\n".join(s_raw.split("\n")[1:])
        if s_raw.endswith("```"):
            s_raw = "\n".join(s_raw.split("\n")[:-1])
        summaries = json.loads(s_raw)
        for entry in all_priorities:
            tid = entry.get("id", "")
            if tid in summaries:
                entry["ai_summary"] = summaries[tid]
        print(f"Phase 3.7 done - {len(summaries)} summaries generated")
    except Exception as e:
        print(f"WARNING: Phase 3.7 AI summaries failed - {e}")


# Pre-build cal items so Phase 3.8 can annotate them before briefing dict is assembled
cal_today_items    = build_cal_items(cal_today)
cal_tomorrow_items = build_cal_items(cal_tomorrow)

# -- Phase 3.7b -- Fetch recent Granola meeting notes for calendar context --
GRANOLA_API_KEY = os.environ.get("GRANOLA_API_KEY", "")
_granola_context = {}  # "day_idx" -> {"note_title": str, "summary": str}

def _granola_keywords(title):
    t = re.sub(r'\b\d{1,2}/\d{2}\b', '', title)   # remove DD/MM dates
    t = re.sub(r'\b\d{4}\b', '', t)                # remove years
    t = re.sub(r'[—\-&]', ' ', t)             # dashes and ampersands to spaces
    t = re.sub(r'[^\w\s]', '', t)                  # strip remaining punctuation
    return set(w.lower() for w in t.split() if len(w) >= 2)

def _granola_fetch(url):
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {GRANOLA_API_KEY}"})
    with urllib.request.urlopen(req, timeout=15) as r:
        return json.loads(r.read().decode())

if GRANOLA_API_KEY:
    try:
        _lookback = (datetime.now() - timedelta(days=10)).strftime("%Y-%m-%dT%H:%M:%SZ")
        _g_data   = _granola_fetch(f"https://public-api.granola.ai/v1/notes?created_after={_lookback}")
        _g_notes  = _g_data.get("notes", [])

        # Build the calendar items list for matching (today + tomorrow non-all-day)
        _cal_candidates = [
            {"idx": i, "day": "today",    "title": c["title"]}
            for i, c in enumerate(cal_today_items) if c.get("time", "").lower() != "all day"
        ] + [
            {"idx": i, "day": "tomorrow", "title": c["title"]}
            for i, c in enumerate(cal_tomorrow_items) if c.get("time", "").lower() != "all day"
        ]

        for cal_item in _cal_candidates:
            cal_kw = _granola_keywords(cal_item["title"])
            if not cal_kw:
                continue
            best_note, best_score = None, 0
            for note in _g_notes:
                score = len(cal_kw & _granola_keywords(note.get("title", "")))
                if score > best_score:
                    best_score, best_note = score, note
            if best_note and best_score >= 1:
                # ?include=transcript required for the API to return the summary field
                detail   = _granola_fetch(f"https://public-api.granola.ai/v1/notes/{best_note['id']}?include=transcript")
                _raw_sum = detail.get("summary") or ""
                if isinstance(_raw_sum, dict):
                    summary = (_raw_sum.get("text") or _raw_sum.get("content") or "").strip()
                else:
                    summary = str(_raw_sum).strip()
                if not summary:
                    summary = (detail.get("summary_text") or detail.get("summary_markdown") or "").strip()
                if summary:
                    key = f"{cal_item['day']}_{cal_item['idx']}"
                    _granola_context[key] = {"note_title": best_note.get("title", ""), "summary": summary[:1500]}

        print(f"Phase 3.7 done - Granola context for {len(_granola_context)} meetings")
    except Exception as e:
        print(f"WARNING: Phase 3.7 Granola fetch failed - {e}")
else:
    print("Phase 3.7 skipped - GRANOLA_API_KEY not set")

# -- Phase 3.8 -- AI prep summaries for today/tomorrow calendar items --
_cal_for_summary = [
    {
        "idx": i, "day": "today", "time": c["time"], "title": c["title"],
        "organizer": c.get("sub", ""),
        "prev_meeting_notes": _granola_context.get(f"today_{i}", {}).get("summary", "")
    }
    for i, c in enumerate(cal_today_items) if c.get("time", "").lower() != "all day"
] + [
    {
        "idx": i, "day": "tomorrow", "time": c["time"], "title": c["title"],
        "organizer": c.get("sub", ""),
        "prev_meeting_notes": _granola_context.get(f"tomorrow_{i}", {}).get("summary", "")
    }
    for i, c in enumerate(cal_tomorrow_items) if c.get("time", "").lower() != "all day"
]
if _cal_for_summary:
    try:
        CAL_SUM_SYSTEM = (
            "You are Kevin's briefing assistant at Oxford University HR Systems.\n"
            "For each meeting, write 2-3 concise sentences of prep context Kevin needs before walking in.\n"
            "Where 'prev_meeting_notes' is provided, use it as your primary source -- it is the AI summary from the last time this meeting ran.\n"
            "Prioritise: carry-forwards and open actions from last time, any live decision or blocker, who Kevin needs to speak to, and the most useful detail Kevin should remember.\n"
            "Plain ASCII punctuation only. No filler like 'This meeting is about...'. Be direct and specific.\n"
            "Return ONLY valid JSON: {\"day_idx\": \"2-3 concise sentences\"} where day_idx is 'today_0', 'today_1', 'tomorrow_0' etc.\n"
            "Example: {\"today_0\": \"Pick up the evaluation scoring from last week -- Helen still needs a decision on weightings. Confirm whether James has resolved the reporting extract and agree the next owner before Friday.\"}"
        )
        _cal_user = (
            f"Today is {today_str}.\n\n"
            f"MEETINGS:\n{json.dumps(_cal_for_summary, indent=1, ensure_ascii=True)}"
        )
        _cs_resp = client.messages.create(
            model      = "claude-haiku-4-5",
            max_tokens = 900,
            system     = CAL_SUM_SYSTEM,
            messages   = [{"role": "user", "content": _cal_user}]
        )
        _cs_raw = _cs_resp.content[0].text.strip()
        if _cs_raw.startswith("```"): _cs_raw = "\n".join(_cs_raw.split("\n")[1:])
        if _cs_raw.endswith("```"):   _cs_raw = "\n".join(_cs_raw.split("\n")[:-1])
        _cs_map = json.loads(_cs_raw)
        for item in _cal_for_summary:
            key = f"{item['day']}_{item['idx']}"
            if key in _cs_map:
                target = cal_today_items if item["day"] == "today" else cal_tomorrow_items
                target[item["idx"]]["summary"] = _cs_map[key]
        print(f"Phase 3.8 done - {len(_cs_map)} calendar summaries generated")
    except Exception as e:
        print(f"WARNING: Phase 3.8 calendar summaries failed - {e}")

# Build calFull -- Mon through Fri of the current working week
def _week_workdays(ref):
    mon = ref - timedelta(days=ref.weekday())
    return [mon + timedelta(days=i) for i in range(5)]

calFull = []
for _wd in _week_workdays(today):
    _day_items = [c for c in calendar if datetime.fromisoformat(c["start"]).date() == _wd]
    calFull.append({
        "date":    _wd.strftime("%Y-%m-%d"),
        "label":   _wd.strftime("%A") + " " + str(_wd.day) + " " + _wd.strftime("%b"),
        "items":   build_cal_items(_day_items),
        "isToday": _wd == today
    })

briefing = {
    "date":         today_str,
    "subtitle":     subtitle,
    "context":      context,
    "urgent":       urgent,
    "needs":        needs,
    "fyi":          fyi,
    "low":          low,
    "calToday":     cal_today_items,
    "calTomorrow":  cal_tomorrow_items,
    "calFull":      calFull,
    "absences":     absences,
    "prioritiesToday":    priorities_today,
    "prioritiesTomorrow": priorities_tomorrow,
    "prioritiesWeek":     priorities_week,
    "refreshed_at": datetime.now().strftime("%A %d %B · %H:%M")
}

# -- Phase 4 -- push to GitHub --
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
            with urllib.request.urlopen(req, timeout=GITHUB_TIMEOUT) as r:
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
        with urllib.request.urlopen(req, timeout=GITHUB_TIMEOUT) as r:
            result = json.loads(r.read())
            print(f"Phase 4 done - briefing pushed to GitHub (commit: {result.get('commit',{}).get('sha','?')[:7]})")
    except Exception as e:
        print(f"Phase 4 FAILED - {e}")
        raise


# Phase 5 - push task suggestions to GitHub (consumed by Command Centre dashboard)
if GITHUB_PAT:
    try:
        sug_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/data/inbox_suggestions.json"
        headers = {
            "Authorization": f"token {GITHUB_PAT}",
            "Content-Type":  "application/json",
            "User-Agent":    "work-inbox-script"
        }
        sha = None
        try:
            req = urllib.request.Request(sug_url, headers=headers)
            with urllib.request.urlopen(req, timeout=GITHUB_TIMEOUT) as r:
                sha = json.loads(r.read()).get("sha")
        except urllib.error.HTTPError as e:
            if e.code != 404:
                raise
        content_b64 = base64.b64encode(
            json.dumps(suggestions, indent=2, ensure_ascii=False).encode("utf-8")
        ).decode("ascii")
        payload = {
            "message": f"chore: update task suggestions {datetime.now().strftime('%Y-%m-%d %H:%M')}",
            "content": content_b64
        }
        if sha:
            payload["sha"] = sha
        data = json.dumps(payload).encode("utf-8")
        req = urllib.request.Request(sug_url, data=data, headers=headers, method="PUT")
        with urllib.request.urlopen(req, timeout=GITHUB_TIMEOUT) as r:
            result = json.loads(r.read())
            print(f"Phase 5 done - suggestions pushed (commit: {result.get('commit',{}).get('sha','?')[:7]})")
    except Exception as e:
        print(f"Phase 5 FAILED - {e}")
