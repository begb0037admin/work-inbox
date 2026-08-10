"""
sent_corpus_pull.py
--------------------
Bulk Sent-Mail corpus pull for the cross-agent style-learning pipeline
(begb0037admin/agent-commons issue #3, item 3 "corpus approach" + item 4 "mailbox access").

Reuses the proven Outlook COM connection pattern from work-inbox/fetch_inbox.py
(win32com.client.dynamic.Dispatch("Outlook.Application") -> GetNamespace("MAPI") ->
GetDefaultFolder(5) for olFolderSentMail) but is intentionally a SEPARATE script,
not a modification of the live 6x/day briefing pipeline:
- fetch_inbox.py's existing Sent Mail read (line ~292) pulls only the last 7 days as
  100-char previews, purely as ephemeral AI-triage context. This script pulls full
  body text over an arbitrary historical window for a durable corpus -- a different
  job with different performance/retention/sensitivity characteristics that has no
  business being inside the live scheduled pipeline.

HARD RULE -- read before running:
  This script writes ONLY to a local-only staging directory that must never be
  inside a git working copy and must never be committed. Nothing this script
  produces (raw OR redacted) is pushed anywhere by the script itself. Pushing the
  redacted corpus to its durable home (proposed: begb0037admin/agent-commons,
  corpus/sent-items/) is a separate, explicit, reviewed step -- see the bottom of
  this file for that hand-off note. This mirrors work-inbox's own standing rule
  ("Never commit raw email data or API keys").

Redaction approach (automated, not manual review, per Kevin's decision 10 Aug 2026):
  Keyword/pattern-based, case-insensitive, on subject+body combined. ANY match in
  any category means the WHOLE message is excluded from the corpus (not partially
  redacted in-place) -- simplest, most conservative option: losing a few borderline
  Sent items from a large corpus costs little; leaving a sensitive fragment in a
  corpus that gets read by another agent costs a lot. Excluded items are logged to
  a separate redaction ledger with category + entry_id + date only -- never with
  the matched text or surrounding context, so the ledger itself carries no
  sensitive content.

  Categories: health, bereavement, hr_case, absence. See REDACTION_PATTERNS below
  for the exact term lists. This is a keyword/pattern floor, not NLP/NER -- it will
  have both false positives (excludes some genuinely clean mail, acceptable) and
  false negatives (a message that references a real sensitive situation without
  using any listed term will NOT be caught -- this is a real, inherent limit of a
  pattern-based pass and should be stated plainly, not implied to be complete).
"""

import argparse
import json
import os
import sys
from datetime import datetime, timedelta

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from style_corpus_common import (
    REDACTION_PATTERNS, is_sensitive, mentions_named_person, OL_MAIL_CLASS,
)

# Redaction patterns, entity list, and the Outlook item-Class constant now
# live in style_corpus_common.py -- shared with draft_final_diff_capture.py,
# which needs the identical classification logic. Was a reasonable duplication
# when this was the only script; a second consumer made a shared import the
# right call instead of a second "keep in sync" copy.

# ---------------------------------------------------------------------------
# Outlook COM pull (requires local Outlook desktop client + pywin32; same
# machine/account constraint as the rest of work-inbox)
# ---------------------------------------------------------------------------

def month_ranges(start_date, end_date):
    cur = start_date.replace(day=1)
    while cur <= end_date:
        if cur.month == 12:
            nxt = cur.replace(year=cur.year + 1, month=1)
        else:
            nxt = cur.replace(month=cur.month + 1)
        chunk_start = max(cur, start_date)
        chunk_end = min(nxt - timedelta(seconds=1), end_date)
        yield (chunk_start, chunk_end)
        cur = nxt


def restrict_sent(folder, start_dt, end_dt):
    filter_str = (
        "[SentOn] >= '" + start_dt.strftime("%m/%d/%Y %I:%M %p") + "' AND "
        "[SentOn] <= '" + end_dt.strftime("%m/%d/%Y %I:%M %p") + "'"
    )
    try:
        restricted = folder.Items.Restrict(filter_str)
        return list(restricted)
    except Exception:
        # Deliberately do NOT fall back to full-folder iteration inside a
        # per-chunk loop -- unlike fetch_inbox.py's restrict_date() (safe for
        # a single 7-day pull), doing that here once per month chunk over a
        # multi-year backfill risks repeatedly walking the entire Sent folder.
        # A failed chunk is skipped and reported, not silently masked.
        return None


def pull_sent_corpus(start_date, end_date, out_dir, dry_run_stats_only=False):
    import win32com.client.dynamic

    if not dry_run_stats_only:
        os.makedirs(out_dir, exist_ok=True)
        out_path = os.path.join(out_dir, "sent_corpus_clean.json")
        log_path = os.path.join(out_dir, "sent_corpus_redaction_log.json")

    outlook = win32com.client.dynamic.Dispatch("Outlook.Application")
    mapi = outlook.GetNamespace("MAPI")
    sent_folder = mapi.GetDefaultFolder(5)  # olFolderSentMail

    # OL_MAIL_CLASS (43, olMail) imported from style_corpus_common -- Sent
    # Items also holds meeting requests/responses/cancellations (Class
    # 53/54/55/56/57) Kevin has sent, which have neither a mail-style Body
    # nor a real "To" the same way and aren't email correspondence anyway --
    # filter them explicitly instead of letting them fall through as an
    # unclassified exception.

    clean_entries = []
    redaction_log = []
    failed_chunks = []
    total_seen = 0
    # Diagnostics only -- item Class (COM integer constant) and exception
    # type name are safe metadata, never message content.
    skipped_non_mail_by_class = {}
    skipped_by_exc = {}

    for chunk_start, chunk_end in month_ranges(start_date, end_date):
        items = restrict_sent(sent_folder, chunk_start, chunk_end)
        if items is None:
            failed_chunks.append(chunk_start.strftime("%Y-%m"))
            continue
        for msg in items:
            try:
                item_class = msg.Class
            except Exception:
                item_class = None
            if item_class != OL_MAIL_CLASS:
                skipped_non_mail_by_class[item_class] = skipped_non_mail_by_class.get(item_class, 0) + 1
                continue

            total_seen += 1
            try:
                subject = msg.Subject or ""
                body = msg.Body or ""
                sent_on = str(msg.SentOn)
                to = msg.To or ""
                entry_id = msg.EntryID

                cats = is_sensitive(subject, body)
                if cats:
                    redaction_log.append({
                        "entry_id": entry_id,
                        "sent": sent_on,
                        "categories": cats,
                        "named_person_mentioned": mentions_named_person(f"{subject}\n{body}"),
                    })
                    continue

                clean_entries.append({
                    "entry_id": entry_id,
                    "subject": subject,
                    "to": to,
                    "sent": sent_on,
                    "body": body,
                })
            except Exception as e:
                skipped_by_exc[type(e).__name__] = skipped_by_exc.get(type(e).__name__, 0) + 1
                continue

    stats = {
        "date_range": [start_date.isoformat(), end_date.isoformat()],
        "total_seen": total_seen,
        "clean_count": len(clean_entries),
        "redacted_count": len(redaction_log),
        "redacted_by_category": {
            cat: sum(1 for r in redaction_log if cat in r["categories"])
            for cat in REDACTION_PATTERNS
        },
        "failed_chunks": failed_chunks,
        "skipped_non_mail_count": sum(skipped_non_mail_by_class.values()),
        "skipped_non_mail_by_item_class": {str(k): v for k, v in skipped_non_mail_by_class.items()},
        "unexpected_errors_on_mail_items": skipped_by_exc,
    }

    if not dry_run_stats_only:
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(clean_entries, f, ensure_ascii=True, indent=2)
        with open(log_path, "w", encoding="utf-8") as f:
            json.dump(redaction_log, f, ensure_ascii=True, indent=2)
        stats["out_path"] = out_path
        stats["log_path"] = log_path
    else:
        # Aggregate-only mode: prove the pull + redaction pass work without
        # writing any real content to disk at all -- used for verification
        # runs against real mail before anything is trusted.
        stats["out_path"] = None
        stats["log_path"] = None

    return stats


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--start", required=True, help="YYYY-MM-DD")
    parser.add_argument("--end", required=True, help="YYYY-MM-DD")
    parser.add_argument(
        "--out-dir",
        default=r"C:\Users\admin\Documents\CorpusStaging\sent_items",
        help="Local-only staging directory. MUST NOT be inside any git working copy.",
    )
    parser.add_argument(
        "--stats-only", action="store_true",
        help="Dry run: pull and classify but write nothing to disk, print aggregate counts only.",
    )
    args = parser.parse_args()

    start_date = datetime.strptime(args.start, "%Y-%m-%d")
    end_date = datetime.strptime(args.end, "%Y-%m-%d")

    result = pull_sent_corpus(start_date, end_date, args.out_dir, dry_run_stats_only=args.stats_only)
    print(json.dumps(result, indent=2))

# ---------------------------------------------------------------------------
# Hand-off to durable storage (NOT done by this script):
#
# 1. Run this script locally (requires Outlook desktop client signed in --
#    same admin / begb0037.AD-OAK machine constraint as fetch_inbox.py).
# 2. Review sent_corpus_redaction_log.json's aggregate category counts --
#    it contains no sensitive text, only entry_id/date/category/flag, so it
#    is safe to review directly.
# 3. Spot-check a sample of sent_corpus_clean.json locally (never paste
#    real bodies into a chat transcript, ticket, or any durable memory file).
# 4. Only after that review: push sent_corpus_clean.json's contents to the
#    proposed durable home (begb0037admin/agent-commons, corpus/sent-items/)
#    via the GitHub Contents API, per Kevin's explicit go-ahead -- this is a
#    separate, reviewed step from running this script.
# ---------------------------------------------------------------------------
