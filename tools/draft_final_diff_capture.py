"""
draft_final_diff_capture.py
----------------------------
Ongoing draft/final diff capture for the cross-agent style-learning pipeline
(begb0037admin/agent-commons issue #3, item 3 "corpus approach" -- the
forward-going half; sent_corpus_pull.py is the one-time/periodic historical
half).

Design proposed and confirmed with Kevin, 10 Aug 2026 (see agent-commons
issue #3 comments for the full feasibility investigation this is built on).

MECHANISM -- periodic snapshot + correlate, not an event-driven listener:
  Outlook has no draft/sent diffing primitive and EntryID is NOT a safe
  correlation key (sending a draft mints a new MAPI entry). ConversationID
  IS reliable -- confirmed live against real Drafts (103 items) and Sent
  (1585 items): present on 40/40 sampled items in both folders. Each run of
  this script:
    1. Snapshots the current Drafts folder (Class == olMail only -- Drafts
       also holds meeting-related non-mail items, same as Sent Items).
    2. Diffs against the previous run's local ledger to find drafts that
       have vanished (sent or discarded) since last seen.
    3. For each vanished draft, searches Sent Items for a mail item sharing
       the same ConversationID, sent within a bounded window after the
       draft's last-seen time (default 72h) -- if found, that's the
       draft/final pair; if not, the draft was discarded, not sent, and is
       dropped (not a diff-pair candidate).
    4. Applies the SAME redaction discipline as sent_corpus_pull.py to BOTH
       sides of the pair -- if EITHER draft or final text matches a
       sensitive category, the WHOLE PAIR is excluded, not just one side.
    5. Computes recipient_tier from the final Sent item's To field, reusing
       the exact live mapping (direct-report/senior-management/other).
    6. Classifies edit_type (tone/trim/factual-correction/add-context/
       register-shift) and generates a short note via claude-haiku-4-5,
       given the REDACTED (already passed the filter) draft/final pair --
       confirmed OK with Kevin 10 Aug 2026, consistent with fetch_inbox.py's
       existing use of the same model for triage.

KNOWN LIMITATION, stated plainly: the captured "draft" is the state as of
the last poll before it vanished, not necessarily the literal last edit
before hitting send -- an inherent ceiling of polling vs. an event-driven
Outlook ItemSend hook, which was considered and explicitly not chosen for
v1 (would require a persistently-running listener process, a different
architecture from every other script in this system).

FIRST RUN PRODUCES ZERO PAIRS, BY DESIGN: this is a forward-looking
mechanism -- it can only diff a draft against its final version once that
draft has actually vanished (sent or discarded) BETWEEN two runs. A first
run only establishes the baseline ledger. This is not a bug.

HARD RULE, same as sent_corpus_pull.py: writes ONLY to a local-only staging
directory that must never be inside a git working copy and must never be
committed. Nothing this script produces is pushed anywhere by the script
itself. Pushing to the durable home (agent-commons/corpus/draft-final-diffs/)
is a separate, explicit, reviewed step.
"""

import argparse
import json
import os
import sys
from datetime import datetime, timedelta

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from style_corpus_common import is_sensitive, recipient_tier, OL_MAIL_CLASS, REDACTION_PATTERNS

DEFAULT_LEDGER_PATH = r"C:\Users\admin\Documents\CorpusStaging\draft_watch\ledger.json"
DEFAULT_OUT_DIR = r"C:\Users\admin\Documents\CorpusStaging\draft_watch"


# ---------------------------------------------------------------------------
# Ledger persistence (local-only, unredacted -- never pushed anywhere)
# ---------------------------------------------------------------------------

def load_ledger(path):
    if not os.path.exists(path):
        return {}
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_ledger(path, ledger):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(ledger, f, ensure_ascii=True, indent=2)


# ---------------------------------------------------------------------------
# edit_type / note classification via claude-haiku-4-5 -- same model/pattern
# fetch_inbox.py already uses for triage. Called only on already-redacted-
# clean pairs (the sensitive-category check happens before this).
# ---------------------------------------------------------------------------

EDIT_TYPES = ["tone", "trim", "factual-correction", "add-context", "register-shift"]

CLASSIFY_SYSTEM = (
    "You compare a draft version of an email to the final sent version and "
    "classify the type of edit that was made, for a writing-style training "
    "corpus. Respond with ONLY a JSON object, no markdown fences, no other text: "
    '{"edit_type": "<one of: tone, trim, factual-correction, add-context, register-shift>", '
    '"note": "<one short sentence describing what changed, no names, no sensitive detail>"}. '
    "tone = wording softened/sharpened without changing length much. "
    "trim = final is meaningfully shorter, content removed. "
    "factual-correction = a fact, date, name, or figure was corrected. "
    "add-context = final adds explanatory detail not in the draft. "
    "register-shift = formality/warmth level changed (e.g. added a greeting, "
    "dropped/added pleasantries)."
)


def classify_edit(client, draft_text, final_text):
    user = (
        f"DRAFT:\n{draft_text}\n\n---\n\nFINAL (as sent):\n{final_text}\n\n"
        "Classify the edit per the system instructions."
    )
    try:
        response = client.messages.create(
            model="claude-haiku-4-5",
            max_tokens=200,
            system=CLASSIFY_SYSTEM,
            messages=[{"role": "user", "content": user}],
        )
        raw_text = response.content[0].text.strip()
        if raw_text.startswith("```"):
            raw_text = "\n".join(raw_text.split("\n")[1:])
        if raw_text.endswith("```"):
            raw_text = "\n".join(raw_text.split("\n")[:-1])
        parsed = json.loads(raw_text)
        edit_type = parsed.get("edit_type")
        note = parsed.get("note", "")
        if edit_type not in EDIT_TYPES:
            return None, None, f"invalid edit_type returned: {edit_type!r}"
        return edit_type, note, None
    except Exception as e:
        return None, None, f"{type(e).__name__}: {e}"


# ---------------------------------------------------------------------------
# Outlook COM snapshot + correlation
# ---------------------------------------------------------------------------

def snapshot_drafts(drafts_folder):
    """Returns {conversation_id: {conversation_topic, subject, to, body, last_seen}}
    for current Drafts items. Class-filtered to real mail only (Drafts also
    holds meeting-related non-mail items, same as Sent Items)."""
    snapshot = {}
    now_iso = datetime.now().isoformat()
    for msg in drafts_folder.Items:
        try:
            if msg.Class != OL_MAIL_CLASS:
                continue
            conv_id = msg.ConversationID
            if not conv_id:
                continue  # no reliable correlation key -- can't track this one
            snapshot[conv_id] = {
                "conversation_topic": msg.ConversationTopic or "",
                "subject": msg.Subject or "",
                "to": msg.To or "",
                "body": msg.Body or "",
                "last_seen": now_iso,
            }
        except Exception:
            continue
    return snapshot


def find_sent_match(sent_folder, conversation_id, after_dt, window_hours):
    """Search Sent Items for a mail item with the same ConversationID, sent
    within [after_dt, after_dt + window_hours]. Returns the closest match by
    time, or None. Does NOT fall back to any looser match -- ambiguous or
    absent correlation means the pair is dropped, not guessed."""
    window_end = after_dt + timedelta(hours=window_hours)
    candidates = []
    for msg in sent_folder.Items:
        try:
            if msg.Class != OL_MAIL_CLASS:
                continue
            if msg.ConversationID != conversation_id:
                continue
            sent_on = msg.SentOn
            # Normalize COM datetime -> python datetime for comparison
            sent_dt = datetime(sent_on.year, sent_on.month, sent_on.day,
                                sent_on.hour, sent_on.minute, sent_on.second)
            if after_dt <= sent_dt <= window_end:
                candidates.append((sent_dt, msg))
        except Exception:
            continue
    if not candidates:
        return None
    candidates.sort(key=lambda c: c[0])
    return candidates[0][1]  # earliest match after the draft was last seen


# ---------------------------------------------------------------------------
# Main run
# ---------------------------------------------------------------------------

def run(ledger_path, out_dir, window_hours=72, use_ai=True, stats_only=False):
    import win32com.client.dynamic

    outlook = win32com.client.dynamic.Dispatch("Outlook.Application")
    mapi = outlook.GetNamespace("MAPI")
    drafts_folder = mapi.GetDefaultFolder(16)  # olFolderDrafts
    sent_folder = mapi.GetDefaultFolder(5)     # olFolderSentMail

    previous_ledger = load_ledger(ledger_path)
    current_snapshot = snapshot_drafts(drafts_folder)

    vanished_keys = set(previous_ledger.keys()) - set(current_snapshot.keys())

    pairs_found = 0
    pairs_redacted = 0
    pairs_classification_failed = 0
    diffs = []
    redaction_log = []
    classification_failures = []
    abandoned_count = 0

    client = None
    if use_ai and not stats_only:
        import anthropic
        client = anthropic.Anthropic(timeout=60.0)

    for conv_id in vanished_keys:
        draft = previous_ledger[conv_id]
        try:
            last_seen_dt = datetime.fromisoformat(draft["last_seen"])
        except Exception:
            continue
        final_msg = find_sent_match(sent_folder, conv_id, last_seen_dt, window_hours)
        if final_msg is None:
            abandoned_count += 1
            continue

        try:
            final_subject = final_msg.Subject or ""
            final_body = final_msg.Body or ""
            final_to = final_msg.To or ""
            final_entry_id = final_msg.EntryID
            final_sent = str(final_msg.SentOn)
        except Exception:
            continue

        pairs_found += 1

        draft_cats = is_sensitive(draft["subject"], draft["body"])
        final_cats = is_sensitive(final_subject, final_body)
        combined_cats = sorted(set(draft_cats) | set(final_cats))
        if combined_cats:
            pairs_redacted += 1
            redaction_log.append({
                "conversation_id": conv_id,
                "final_entry_id": final_entry_id,
                "sent": final_sent,
                "categories": combined_cats,
            })
            continue

        tier = recipient_tier(final_to)

        if stats_only:
            diffs.append({"conversation_id": conv_id, "recipient_tier": tier, "stats_only": True})
            continue

        edit_type, note, err = (None, None, "AI classification disabled") if not use_ai \
            else classify_edit(client, draft["body"], final_body)

        if edit_type is None:
            pairs_classification_failed += 1
            classification_failures.append({
                "conversation_id": conv_id, "final_entry_id": final_entry_id,
                "sent": final_sent, "error": err,
            })
            continue

        diffs.append({
            "channel": "email",
            "draft": draft["body"],
            "final": final_body,
            "edit_type": edit_type,
            "note": note,
            "recipient_tier": tier,
            "confirmed_via": (
                f"draft_final_diff_capture.py, matched via ConversationID + sent within "
                f"{window_hours}h of last-seen draft snapshot (last_seen={draft['last_seen']}, "
                f"sent={final_sent}), final_entry_id={final_entry_id}"
            ),
        })

    # Update ledger: drop resolved (vanished) keys, refresh/add current drafts
    new_ledger = dict(current_snapshot)
    save_ledger(ledger_path, new_ledger)

    stats = {
        "run_time": datetime.now().isoformat(),
        "drafts_tracked_now": len(current_snapshot),
        "drafts_vanished_since_last_run": len(vanished_keys),
        "abandoned_or_discarded": abandoned_count,
        "draft_final_pairs_found": pairs_found,
        "pairs_excluded_by_redaction": pairs_redacted,
        "pairs_classification_failed": pairs_classification_failed,
        "pairs_published": len(diffs) if not stats_only else 0,
        "window_hours": window_hours,
    }

    if not stats_only:
        os.makedirs(out_dir, exist_ok=True)
        diffs_path = os.path.join(out_dir, "draft_final_diffs.json")
        log_path = os.path.join(out_dir, "draft_final_redaction_log.json")
        fail_path = os.path.join(out_dir, "draft_final_classification_failures.json")

        # Append-not-overwrite: multiple runs accumulate pairs over time.
        existing_diffs = []
        if os.path.exists(diffs_path):
            with open(diffs_path, "r", encoding="utf-8") as f:
                existing_diffs = json.load(f)
        existing_diffs.extend(diffs)
        with open(diffs_path, "w", encoding="utf-8") as f:
            json.dump(existing_diffs, f, ensure_ascii=True, indent=2)

        existing_log = []
        if os.path.exists(log_path):
            with open(log_path, "r", encoding="utf-8") as f:
                existing_log = json.load(f)
        existing_log.extend(redaction_log)
        with open(log_path, "w", encoding="utf-8") as f:
            json.dump(existing_log, f, ensure_ascii=True, indent=2)

        if classification_failures:
            existing_fail = []
            if os.path.exists(fail_path):
                with open(fail_path, "r", encoding="utf-8") as f:
                    existing_fail = json.load(f)
            existing_fail.extend(classification_failures)
            with open(fail_path, "w", encoding="utf-8") as f:
                json.dump(existing_fail, f, ensure_ascii=True, indent=2)

        stats["diffs_path"] = diffs_path
        stats["total_diffs_accumulated"] = len(existing_diffs)

    return stats


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--ledger-path", default=DEFAULT_LEDGER_PATH)
    parser.add_argument("--out-dir", default=DEFAULT_OUT_DIR)
    parser.add_argument("--window-hours", type=int, default=72)
    parser.add_argument("--no-ai", action="store_true", help="Skip edit_type/note classification (diagnostic only)")
    parser.add_argument("--stats-only", action="store_true", help="No writes, no AI calls, aggregate counts only")
    args = parser.parse_args()

    result = run(
        args.ledger_path, args.out_dir,
        window_hours=args.window_hours,
        use_ai=not args.no_ai,
        stats_only=args.stats_only,
    )
    print(json.dumps(result, indent=2))

# ---------------------------------------------------------------------------
# Hand-off to durable storage (NOT done by this script) -- same discipline as
# sent_corpus_pull.py:
# 1. Run repeatedly over time (Task Scheduler, proposing hourly during working
#    hours) -- real pairs only appear once a tracked draft has actually vanished
#    between two runs.
# 2. Review draft_final_redaction_log.json (safe -- category/date/entry_id
#    only) and draft_final_classification_failures.json.
# 3. Spot-check a sample of draft_final_diffs.json locally.
# 4. Only after review: push to agent-commons/corpus/draft-final-diffs/ via
#    the GitHub Contents/Data API, per explicit go-ahead.
# ---------------------------------------------------------------------------
