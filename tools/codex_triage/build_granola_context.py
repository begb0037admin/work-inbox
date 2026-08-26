"""
Ports fetch_inbox.py's Phase 3.7b Granola-fetch logic verbatim (same lookback,
same keyword-match algorithm, same idx/real_idx offset-bug fix) so the Codex
dry run gets equivalent calendar-prep context to the real pipeline, without
needing Codex itself to hold GRANOLA_API_KEY or make the REST call.
Read-only. Writes granola_context.json locally only.
"""
import json, os, re, urllib.request
from datetime import datetime, timedelta

GRANOLA_API_KEY = os.environ.get("GRANOLA_API_KEY", "")

d = json.load(open("real_briefing.json", encoding="utf-8"))
cal_today_items = d.get("calToday", [])
cal_tomorrow_items = d.get("calTomorrow", [])

def _granola_keywords(title):
    t = re.sub(r'\b\d{1,2}/\d{2}\b', '', title)
    t = re.sub(r'\b\d{4}\b', '', t)
    t = re.sub(r'[—\-&]', ' ', t)
    t = re.sub(r'[^\w\s]', '', t)
    return set(w.lower() for w in t.split() if len(w) >= 2)

def _granola_fetch(url):
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {GRANOLA_API_KEY}"})
    with urllib.request.urlopen(req, timeout=15) as r:
        return json.loads(r.read().decode())

def _non_all_day_candidates(items, day_label):
    candidates = []
    model_idx = 0
    for real_idx, c in enumerate(items):
        if (c.get("time") or "").lower() == "all day":
            continue
        candidates.append({
            "idx": model_idx, "real_idx": real_idx, "day": day_label,
            "time": c["time"], "title": c["title"], "organizer": c.get("sub", "")
        })
        model_idx += 1
    return candidates

_all_day_candidates = (
    _non_all_day_candidates(cal_today_items, "today") +
    _non_all_day_candidates(cal_tomorrow_items, "tomorrow")
)

_granola_context = {}
if GRANOLA_API_KEY:
    try:
        _lookback = (datetime.now() - timedelta(days=10)).strftime("%Y-%m-%dT%H:%M:%SZ")
        _g_data = _granola_fetch(f"https://public-api.granola.ai/v1/notes?created_after={_lookback}")
        _g_notes = _g_data.get("notes", [])
        for cal_item in _all_day_candidates:
            cal_kw = _granola_keywords(cal_item["title"])
            if not cal_kw:
                continue
            best_note, best_score = None, 0
            for note in _g_notes:
                score = len(cal_kw & _granola_keywords(note.get("title", "")))
                if score > best_score:
                    best_score, best_note = score, note
            if best_note and best_score >= 1:
                detail = _granola_fetch(f"https://public-api.granola.ai/v1/notes/{best_note['id']}?include=transcript")
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
        print(f"Granola context built for {len(_granola_context)} meetings (of {len(_all_day_candidates)} non-all-day candidates)")
    except Exception as e:
        print(f"WARNING: Granola fetch failed - {e}")
else:
    print("Skipped - GRANOLA_API_KEY not set")

# Emit calendar candidates + granola context together, in the exact shape
# Phase 3.8's own AI call consumes (idx-keyed, prev_meeting_notes attached).
cal_for_summary = [
    dict(c, prev_meeting_notes=_granola_context.get(f"{c['day']}_{c['idx']}", {}).get("summary", ""))
    for c in _all_day_candidates
]
json.dump(cal_for_summary, open("cal_for_summary.json", "w", encoding="utf-8"), indent=1, ensure_ascii=True)
print(f"Staged {len(cal_for_summary)} calendar candidates to cal_for_summary.json")
