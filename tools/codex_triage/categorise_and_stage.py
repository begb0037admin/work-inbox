"""
Ports fetch_inbox.py's deterministic categorise()/badge_for()/make_card()
logic VERBATIM (same keyword lists, same order-of-precedence, same rules)
onto Codex's own live connector pull, so the split into urgent/needs/fyi/low
is identical business logic to the real pipeline -- only the underlying
data source (Codex-Graph connector vs Outlook COM) and the downstream AI
judgement model differ. This is deterministic code, not a model call, so it
is reused rather than re-derived or left to Codex's own judgement.
"""
import json, re, html
from datetime import datetime, timedelta

URGENT_SUBJECTS = ["major incident", "priority 1", "p1", "urgent", "critical", "security vulnerab"]
NEEDS_SUBJECTS = ["re:", "fw:", "fwd:", "action", "required", "please", "timeline", "update",
                  "chasing", "waiting", "overdue", "follow", "scoping", "handover", "error",
                  "import", "failed", "issue", "case ", "support"]
FYI_SUBJECTS = ["fyi", "notification", "scheduled", "maintenance", "summary", "workshop",
                "invitation", "invite", "digest", "recap", "newsletter", "annual leave",
                "out of office", "automatic reply", "accepted:", "declined:", "cancelled:"]
LOW_SUBJECTS = ["unsubscribe", "noreply", "no-reply", "do not reply", "automated",
                "github", "pages", "build", "deploy", "run failed", "wisp"]

_IMPORTANCE_MAP = {"high": 2, "normal": 1, "low": 0, None: 1}

def categorise(msg):
    subj = (msg.get("subject") or "").lower()
    sender = (msg.get("from_email") or "").lower()
    is_read = msg.get("is_read", True)
    imp = _IMPORTANCE_MAP.get(msg.get("importance"), 1)

    if imp == 2:
        return "urgent"
    for kw in LOW_SUBJECTS:
        if kw in subj or kw in sender:
            return "low"
    for kw in URGENT_SUBJECTS:
        if kw in subj:
            return "urgent"
    if not is_read:
        for kw in NEEDS_SUBJECTS:
            if kw in subj:
                return "needs"
    for kw in FYI_SUBJECTS:
        if kw in subj:
            return "fyi"
    if not is_read:
        return "needs"
    return "fyi"

def badge_for(msg, category):
    received = msg.get("received_utc", "")
    age_hrs = 0
    try:
        t = datetime.fromisoformat((received or "").replace("Z", "+00:00"))
        now = datetime.now(t.tzinfo)
        age_hrs = (now - t).total_seconds() / 3600
    except Exception:
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

def make_card(msg, cat):
    subj = msg.get("subject") or "(no subject)"
    sender = msg.get("from_name") or msg.get("from_email") or ""
    preview = (msg.get("body_preview") or "").strip()
    preview = re.sub(r"<\?\s*https?://\S+>?", "[link]", preview)
    badge, badge_type = badge_for(msg, cat)
    card = {
        "title": subj,
        "sub": f"From <strong>{sender}</strong>. {html.escape(preview[:120])}" if preview else f"From <strong>{sender}</strong>.",
        "badge": badge, "badgeType": badge_type,
        "subject": subj, "from": sender,
        "graph_id": msg.get("id", ""),
        "received_raw": msg.get("received_utc", ""),
        "kevin_is_primary_recipient": msg.get("kevin_is_primary_recipient", True),
    }
    return card

inbox = json.load(open("call_inbox_result.json", encoding="utf-8")).get("inbox", [])

urgent, needs, fyi, low = [], [], [], []
for m in inbox:
    cat = categorise(m)
    card = make_card(m, cat)
    {"urgent": urgent, "needs": needs, "fyi": fyi, "low": low}[cat].append(card)

print(f"categorise() done - urgent:{len(urgent)} needs:{len(needs)} fyi:{len(fyi)} low:{len(low)}")

# Stage Phase-2-equivalent (email summaries) candidates -- urgent+needs, same
# fields as fetch_inbox.py's real summary_candidates builder, verbatim logic.
def _age_days(card):
    try:
        rec_dt = datetime.fromisoformat((card.get("received_raw") or "").replace("Z", "+00:00"))
        now = datetime.now(rec_dt.tzinfo)
        return (now - rec_dt).days
    except Exception:
        return None

summary_candidates = urgent + needs
emails_for_summary = [
    {
        "id": str(i),
        "subject": c["subject"],
        "from": c["from"],
        "preview": (c.get("sub") or "")[:250],
        "kevin_is_primary_recipient": c.get("kevin_is_primary_recipient", True),
        "age_days": _age_days(c),
    }
    for i, c in enumerate(summary_candidates)
]
json.dump(emails_for_summary, open("stage_email_summary_candidates.json", "w", encoding="utf-8"), indent=1, ensure_ascii=True)
print(f"staged {len(emails_for_summary)} email-summary candidates")

# Save the raw urgent/needs card lists (with graph_id) so the diff script can
# map Codex's verdicts (keyed 0..N matching emails_for_summary above) back
# onto real cards after Call 2 runs.
json.dump({"urgent": urgent, "needs": needs, "fyi": fyi, "low": low},
          open("stage_categorised_cards.json", "w", encoding="utf-8"), indent=1, ensure_ascii=True)

# Stage Phase-3.5-equivalent (task-suggestion triage) candidates -- urgent+
# needs received mail plus sent mail, same shape as fetch_inbox.py's
# email_candidates/api_emails builder.
sent = json.load(open("call_sent_result.json", encoding="utf-8")).get("sent", [])
email_candidates = []
for c in (urgent + needs):
    email_candidates.append({
        "subject": c["subject"], "from": c["from"],
        "received": (c.get("received_raw") or "")[:16],
        "body_preview": re.sub(r"<\?\s*https?://\S+>?", "[link]", (c.get("sub") or ""))[:150],
        "graph_id": c.get("graph_id", ""),
    })
for s in sent[:30]:
    email_candidates.append({
        "subject": s.get("subject", ""),
        "from": "Kevin (sent to: " + (s.get("to_name") or "") + ")",
        "received": (s.get("sent_utc") or "")[:16],
        "body_preview": re.sub(r"<\?\s*https?://\S+>?", "[link]", (s.get("body_preview") or ""))[:150],
        "graph_id": s.get("id", ""), "direction": "sent",
    })
api_emails = [
    {"n": i, "direction": e.get("direction", "received"), "subject": e["subject"],
     "from": e["from"], "received": e["received"], "body_preview": e["body_preview"]}
    for i, e in enumerate(email_candidates)
]
json.dump(api_emails, open("stage_triage_api_emails.json", "w", encoding="utf-8"), indent=1, ensure_ascii=True)
json.dump(email_candidates, open("stage_triage_email_candidates.json", "w", encoding="utf-8"), indent=1, ensure_ascii=True)
print(f"staged {len(api_emails)} triage candidates ({len(urgent+needs)} received + {len(sent[:30])} sent)")

# Stage Phase-3.8-equivalent calendar candidates from CODEX'S OWN calendar
# pull (not the real pipeline's), same idx/real_idx offset-bug-fixed shape,
# with Granola context reused from the earlier build_granola_context.py run
# where keyword overlap allows (recomputed against Codex's own titles).
calendar = json.load(open("call_calendar_result.json", encoding="utf-8")).get("calendar", [])

def _granola_keywords(title):
    t = re.sub(r'\b\d{1,2}/\d{2}\b', '', title or '')
    t = re.sub(r'\b\d{4}\b', '', t)
    t = re.sub(r'[—\-&]', ' ', t)
    t = re.sub(r'[^\w\s]', '', t)
    return set(w.lower() for w in t.split() if len(w) >= 2)

import os, urllib.request
GRANOLA_API_KEY = os.environ.get("GRANOLA_API_KEY", "")
_g_notes = []
if GRANOLA_API_KEY:
    try:
        req = urllib.request.Request(
            f"https://public-api.granola.ai/v1/notes?created_after={(datetime.now()-timedelta(days=10)).strftime('%Y-%m-%dT%H:%M:%SZ')}",
            headers={"Authorization": f"Bearer {GRANOLA_API_KEY}"})
        with urllib.request.urlopen(req, timeout=15) as r:
            _g_notes = json.loads(r.read().decode()).get("notes", [])
    except Exception as e:
        print(f"WARNING: granola list fetch failed - {e}")

def _granola_fetch_detail(note_id):
    req = urllib.request.Request(f"https://public-api.granola.ai/v1/notes/{note_id}?include=transcript",
                                  headers={"Authorization": f"Bearer {GRANOLA_API_KEY}"})
    with urllib.request.urlopen(req, timeout=15) as r:
        return json.loads(r.read().decode())

by_day = {"today": [], "tomorrow": []}
for c in calendar:
    if c.get("day") in by_day:
        by_day[c["day"]].append(c)

def _non_all_day(items):
    out = []
    idx = 0
    for real_idx, c in enumerate(items):
        if c.get("all_day"):
            continue
        out.append({"idx": idx, "real_idx": real_idx, "title": c.get("title", ""),
                    "organizer": c.get("organizer_name", ""),
                    "start_utc": c.get("start_utc", "")})
        idx += 1
    return out

cal_for_summary = []
granola_hits = 0
for day in ("today", "tomorrow"):
    for cand in _non_all_day(by_day[day]):
        cal_kw = _granola_keywords(cand["title"])
        prev_notes = ""
        if cal_kw and _g_notes:
            best_note, best_score = None, 0
            for note in _g_notes:
                score = len(cal_kw & _granola_keywords(note.get("title", "")))
                if score > best_score:
                    best_score, best_note = score, note
            if best_note and best_score >= 1:
                try:
                    detail = _granola_fetch_detail(best_note["id"])
                    raw_sum = detail.get("summary") or ""
                    summary = (raw_sum.get("text") or raw_sum.get("content") or "").strip() if isinstance(raw_sum, dict) else str(raw_sum).strip()
                    if not summary:
                        summary = (detail.get("summary_text") or detail.get("summary_markdown") or "").strip()
                    if summary:
                        prev_notes = summary[:1500]
                        granola_hits += 1
                except Exception as e:
                    print(f"WARNING: granola detail fetch failed for {cand['title']} - {e}")
        cal_for_summary.append(dict(cand, day=day, prev_meeting_notes=prev_notes))

json.dump(cal_for_summary, open("stage_cal_for_summary.json", "w", encoding="utf-8"), indent=1, ensure_ascii=True)
print(f"staged {len(cal_for_summary)} calendar candidates ({granola_hits} with Granola context)")
