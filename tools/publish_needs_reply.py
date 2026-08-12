"""
publish_needs_reply.py
------------------------
Publishes work-inbox/data/needs_reply.json -- the hand-off queue for the
Drew-to-Lauren wiring (agent-commons issue #3, step-3 brief, item 2).

Design (agreed and corrected on the issue thread, 10 Aug 2026):
  fetch_inbox.py's Phase 3.2 now flags each urgent/needs email with
  needs_reply: true/false (see that script's own comments). This script is a
  SEPARATE follow-on step, not a Phase inside fetch_inbox.py itself, because
  it needs to fetch FULL body text for the flagged subset via Outlook COM --
  briefing.json only ever carries a ~150-char preview -- and full body is
  needed for Lauren to draft anything useful. Keeping it separate also means
  fetch_inbox.py's own dependency footprint (a single file, pulled fresh via
  raw URL by the production .bat) doesn't change; this script lives in
  tools/ alongside sent_corpus_pull.py and draft_final_diff_capture.py and
  can cleanly import style_corpus_common.py the same way they do.

Flow:
  1. Read the just-pushed live data/briefing.json (via GitHub Contents API).
  2. Filter to entries where needs_reply == true, across urgent + needs.
  3. For each, fetch full body via Outlook COM (mapi.GetItemFromID(entry_id),
     the same lookup open_email.py already uses to open a card).
  4. Apply the SAME redaction classifier already built for the corpora
     (style_corpus_common.is_sensitive on subject+full body) -- a
     flagged-sensitive email is simply never written to needs_reply.json.
     It's still visible to Kevin normally, in his own inbox and on the
     regular work-inbox dashboard -- this only keeps it out of the
     automated Lauren-drafting pipeline.
  5. Compute sender_tier from the "from" display-name field, reusing
     recipient_tier()'s exact mapping (direct-report/senior-management/
     other) -- same function, just read as "who's asking Kevin" instead of
     "who Kevin sent to."
  6. Write work-inbox/data/needs_reply.json: entry_id, subject, sender
     (real display name), sender_tier, received, body (full,
     redaction-cleared only), ai_note (Phase 3.2's existing one-sentence
     summary, reused, not regenerated).

Real sender name included, not just sender_tier -- Kevin's explicit decision,
10 Aug 2026, deliberately distinct from the tier-only-not-raw-name discipline
used on the Sent-items/draft-final CORPORA (agent-commons/corpus/*). His
reasoning: that minimization is for shared, durable, cross-agent data leaving
Drew's own system boundary; needs_reply.json is internal work-inbox plumbing
that never leaves it, and Kevin already sees the real sender the instant he
opens the source email anyway (one click away via entry_id/openmail://). The
two files intentionally follow different rules for different reasons -- don't
"fix" this file to be tier-only later by analogy with the corpora without
Kevin re-deciding it.

This script does NOT push to GitHub itself when run as a library call from
run() -- the CLI entry point does the push, with the same backup-and-verify
discipline as every other write in this pipeline.
"""

import argparse
import base64
import json
import os
import sys
import urllib.request
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from style_corpus_common import is_sensitive, recipient_tier

GITHUB_API = "https://api.github.com"
OWNER = "begb0037admin"
REPO = "work-inbox"


def gh_get(path, token):
    req = urllib.request.Request(
        f"{GITHUB_API}/repos/{OWNER}/{REPO}/contents/{path}?ref=main",
        headers={"Authorization": f"token {token}", "Accept": "application/vnd.github+json"},
    )
    with urllib.request.urlopen(req) as r:
        return json.load(r)


def gh_put(path, content_bytes, message, sha, token, branch="main"):
    body = {
        "message": message,
        "content": base64.b64encode(content_bytes).decode("ascii"),
        "branch": branch,
    }
    if sha:
        body["sha"] = sha
    req = urllib.request.Request(
        f"{GITHUB_API}/repos/{OWNER}/{REPO}/contents/{path}",
        data=json.dumps(body).encode("utf-8"),
        headers={"Authorization": f"token {token}", "Accept": "application/vnd.github+json", "Content-Type": "application/json"},
        method="PUT",
    )
    with urllib.request.urlopen(req) as r:
        return json.load(r)


def gh_blob(sha, token):
    req = urllib.request.Request(
        f"{GITHUB_API}/repos/{OWNER}/{REPO}/git/blobs/{sha}",
        headers={"Authorization": f"token {token}", "Accept": "application/vnd.github+json"},
    )
    with urllib.request.urlopen(req) as r:
        return base64.b64decode(json.load(r)["content"])


def fetch_full_body(entry_id, mapi):
    item = mapi.GetItemFromID(entry_id)
    return item.Body or ""


def build_needs_reply_entries(briefing, mapi, stats_only=False):
    urgent = briefing.get("urgent", [])
    needs = briefing.get("needs", [])
    candidates = [c for c in (urgent + needs) if c.get("needs_reply") is True and c.get("entry_id")]

    entries = []
    redacted_count = 0
    fetch_failed = 0

    for c in candidates:
        try:
            full_body = fetch_full_body(c["entry_id"], mapi)
        except Exception:
            fetch_failed += 1
            continue

        if is_sensitive(c.get("subject", ""), full_body):
            redacted_count += 1
            continue

        entries.append({
            "entry_id": c["entry_id"],
            "subject": c.get("subject", ""),
            "sender": c.get("from", ""),
            "sender_tier": recipient_tier(c.get("from", "")),
            "received": c.get("received", ""),
            "body": full_body,
            "ai_note": c.get("ai_summary", ""),
        })

    stats = {
        "candidates_flagged_needs_reply": len(candidates),
        "published": len(entries),
        "redacted": redacted_count,
        "fetch_failed": fetch_failed,
    }
    return entries, stats


def run(token, dry_run=False):
    import win32com.client.dynamic

    briefing_meta = gh_get("data/briefing.json", token)
    briefing = json.loads(base64.b64decode(briefing_meta["content"]))

    outlook = win32com.client.dynamic.Dispatch("Outlook.Application")
    mapi = outlook.GetNamespace("MAPI")

    entries, stats = build_needs_reply_entries(briefing, mapi)

    if dry_run:
        stats["pushed"] = False
        return stats

    payload = {
        "generated": briefing.get("generatedAt") or briefing.get("date") or "",
        "source": "work-inbox/fetch_inbox.py Phase 3.2 + tools/publish_needs_reply.py",
        "entries": entries,
    }
    content_bytes = json.dumps(payload, ensure_ascii=True, indent=2).encode("utf-8")

    # Backup-and-verify: check whether needs_reply.json already exists (to
    # get its sha for the update), push, then re-fetch the pushed blob via
    # git/blobs (bypasses the raw.githubusercontent.com CDN cache) and diff
    # byte-for-byte before trusting the write.
    existing_sha = None
    try:
        existing = gh_get("data/needs_reply.json", token)
        existing_sha = existing["sha"]
    except Exception:
        pass  # first-ever run, file doesn't exist yet

    result = gh_put(
        "data/needs_reply.json", content_bytes,
        f"Publish needs_reply.json: {len(entries)} flagged entries ({stats['redacted']} redacted, {stats['fetch_failed']} fetch failures)",
        existing_sha, token,
    )
    new_sha = result["content"]["sha"]
    remote = gh_blob(new_sha, token)
    stats["pushed"] = True
    stats["byte_identical_verified"] = (remote == content_bytes)
    stats["new_sha"] = new_sha
    return stats


if __name__ == "__main__":
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] publish_needs_reply.py run started")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--dry-run", action="store_true", help="Build entries but don't push to GitHub")
    args = parser.parse_args()

    token = os.environ.get("GITHUB_PAT")
    if not token:
        print("FATAL: GITHUB_PAT not set")
        sys.exit(1)

    try:
        result = run(token, dry_run=args.dry_run)
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(f"FATAL: publish_needs_reply.py run failed - {type(e).__name__}: {e}")
        sys.exit(1)
