"""
publish_drafted_replies.py
----------------------------
Mirrors Lauren's `agent-commons/pending-email-drafts/drafts.json` into
`work-inbox/data/drafted_replies.json`, so the (public) work-inbox dashboard
can display it without needing any client-side access to agent-commons.

WHY A MIRROR STEP, not a direct client-side cross-repo fetch (correction,
found by testing, not assumed -- 10 Aug 2026):
  The original item-2/item-4 design assumed the dashboard's JS could fetch
  agent-commons/pending-email-drafts/drafts.json the same way it already
  cross-fetches command-centre/data/tasks.json for the CC ticker. That
  precedent only works because command-centre is a PUBLIC repo.
  agent-commons is PRIVATE (confirmed: gh api repos/.../agent-commons ->
  "private": true) -- unauthenticated client-side JS in a public GitHub
  Pages site has no credential to read it, and the existing github-proxy
  Worker (github-proxy.lelitte.co.uk) returns 404 for it (tested live,
  README.md -> 404, vs the same request against work-inbox -> 200), most
  likely because its own PAT doesn't have private-repo read access either.
  Rather than extend that shared proxy's privileges to a private repo, or
  add a new Worker route, this mirrors the exact same pattern already
  proven for needs_reply.json: Drew's own local automation (which holds the
  real GITHUB_PAT) reads the private source, republishes only what's
  already meant to be shown (agent-commons/pending-email-drafts/drafts.json
  is, by design, already redacted/tier-tagged content Kevin is meant to
  see) into work-inbox's own public data/, and the dashboard reads that the
  same way it already reads briefing.json -- same-repo, no cross-origin
  auth problem, zero new infrastructure.

  agent-commons itself is never exposed to any client-side/anonymous
  reader by this design -- only this local script (real GITHUB_PAT,
  running on the admin machine) ever reads it directly.

If agent-commons/pending-email-drafts/drafts.json doesn't exist yet (Lauren
hasn't written anything), this publishes an empty entries list rather than
failing -- the dashboard should render a graceful "no drafts yet" state.

FIX (11 Aug 2026, Drew): two real bugs found while chasing a report of
"mirror ran after the source update but still shows stale data":
  1. The `generated` field on the published payload used
     `datetime.now().isoformat()` -- a timezone-naive LOCAL (BST, UTC+1)
     timestamp -- sitting next to Lauren's timezone-aware UTC timestamps
     in agent-commons. A mirror that actually ran at 12:32 UTC prints
     "13:32", which visually looks *later* than a 13:12 UTC source push
     even though it is 40 minutes *earlier* in absolute time. This was
     the actual explanation for the "stale after a fresh run" report --
     the mirror genuinely hadn't been re-run since the source update, the
     timestamp just made it look like it had. Fixed to emit a proper UTC
     ISO timestamp so this can't be misread again.
  2. normalize_entry()'s optional-passthrough tuple did not include
     `corpus_provenance` -- the object that actually contains
     `content_precedent` (added 11 Aug 2026, agent-commons commit
     eb7af377). Even a byte-fresh read would have silently dropped it.
     `confidence` (including the new "medium" tier) was already a
     whole-dict passthrough and did NOT need a fix -- its staleness in
     the bug report was solely due to #1 above, not a shape-check drop.
"""

import argparse
import base64
import datetime
import json
import os
import sys
import urllib.request
import urllib.error

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from phase_failure_notify import notify_phase_failure

GITHUB_API = "https://api.github.com"
WI_OWNER, WI_REPO = "begb0037admin", "work-inbox"
AC_OWNER, AC_REPO = "begb0037admin", "agent-commons"


def gh_get(owner, repo, path, token):
    req = urllib.request.Request(
        f"{GITHUB_API}/repos/{owner}/{repo}/contents/{path}?ref=main",
        headers={"Authorization": f"token {token}", "Accept": "application/vnd.github+json"},
    )
    with urllib.request.urlopen(req) as r:
        return json.load(r)


def gh_put(owner, repo, path, content_bytes, message, sha, token, branch="main"):
    body = {"message": message, "content": base64.b64encode(content_bytes).decode("ascii"), "branch": branch}
    if sha:
        body["sha"] = sha
    req = urllib.request.Request(
        f"{GITHUB_API}/repos/{owner}/{repo}/contents/{path}",
        data=json.dumps(body).encode("utf-8"),
        headers={"Authorization": f"token {token}", "Accept": "application/vnd.github+json", "Content-Type": "application/json"},
        method="PUT",
    )
    with urllib.request.urlopen(req) as r:
        return json.load(r)


def gh_blob(owner, repo, sha, token):
    req = urllib.request.Request(
        f"{GITHUB_API}/repos/{owner}/{repo}/git/blobs/{sha}",
        headers={"Authorization": f"token {token}", "Accept": "application/vnd.github+json"},
    )
    with urllib.request.urlopen(req) as r:
        return base64.b64decode(json.load(r)["content"])


def normalize_entry(e):
    """Map Lauren's real field names onto the canonical shape the dashboard
    renders, and pass through the richer fields (confidence, inline_flags,
    corpus_provenance) that her actual output includes but the original
    mirror schema didn't account for. Returns None if a truly essential
    field is missing -- logs exactly which one, rather than silently
    dropping."""
    if not isinstance(e, dict):
        return None, "not a dict"

    # source_entry_id was originally required as "core" -- but it only
    # exists for drafts sourced from work-inbox/data/needs_reply.json (a
    # single Outlook EntryID). Drafts 14/15/16 (18 Aug 2026) come from a
    # direct Kevin chat-paste and from Drew's own live-thread retrievals
    # (reply-all threads, no single needs_reply.json EntryID carried into
    # the draft record) -- they were silently dropped by this check even
    # though they're real, complete, reviewable drafts. FIX (18 Aug 2026,
    # Drew): source_entry_id is now optional. When present, it's used as
    # before (identity for tick-dedup + the "Open original" Outlook deep
    # link). When absent, falls back to the draft's own draft_id, which is
    # always unique and always present -- so tick-dedup (mark sent/discard)
    # still works correctly and doesn't collide across entries that all
    # lack a real EntryID. Known, disclosed side effect: for those entries
    # the "Open original" button will still render (dashboard just checks
    # for a non-empty string) but will call openEmail() with a draft_id
    # instead of a real Outlook EntryID, so it won't successfully open the
    # source email. Not fixed here -- would need an app.js-side change
    # (hasSource check) and this fix is intentionally scoped to unblocking
    # visibility only, not a same-night frontend change on top of it.
    core = {"subject", "sender_tier", "draft_text"}
    missing_core = core - e.keys()
    if missing_core:
        return None, f"missing core field(s): {sorted(missing_core)}"

    # Timestamp: Lauren's real output uses "composed_at" -- the original
    # mirror required "drafted_at" and silently dropped every real entry as
    # a result (caught live, 10 Aug 2026, all 4 of Lauren's first real
    # drafts failed this check). Accept either, normalize the OUTPUT key to
    # "drafted_at" so the already-built dashboard template doesn't need a
    # matching rename.
    timestamp = e.get("composed_at") or e.get("drafted_at") or ""
    if not timestamp:
        return None, "missing both composed_at and drafted_at"

    normalized = {
        "source_entry_id": e.get("source_entry_id") or e.get("draft_id") or "",
        "subject": e["subject"],
        "sender_tier": e["sender_tier"],
        "draft_text": e["draft_text"],
        "drafted_at": timestamp,
    }
    # Pass through optional richer fields as-is when present -- these are
    # real content Lauren already produces (confidence level/reason,
    # inline_flags for anything she couldn't verify, corpus_provenance for
    # the content-precedent search results added 11 Aug 2026) that the
    # original dashboard design didn't render at all. Optional: absence
    # doesn't drop the entry, just means the dashboard shows less detail
    # for it. corpus_provenance is where content_precedent actually lives
    # in the source shape -- passed through wholesale so the dashboard sees
    # the exact same structure Lauren wrote, no reshaping in this mirror.
    for optional_field in ("confidence", "inline_flags", "received", "status", "draft_id", "corpus_provenance"):
        if optional_field in e:
            normalized[optional_field] = e[optional_field]

    return normalized, None


def run(token, dry_run=False):
    source_missing = False
    read_error = None
    try:
        source = gh_get(AC_OWNER, AC_REPO, "pending-email-drafts/drafts.json", token)
        drafts = json.loads(base64.b64decode(source["content"]))
        entries = drafts.get("entries", []) if isinstance(drafts, dict) else drafts
    except urllib.error.HTTPError as e:
        entries = []
        if e.code == 404:
            source_missing = True
        else:
            # Any non-404 failure (auth, rate limit, transient 5xx) is
            # logged with its real status/body rather than silently folded
            # into "source_missing" -- a 404 specifically can also mean the
            # token lacks read access to a private repo (GitHub returns 404,
            # not 403, to avoid revealing a private repo's existence), so
            # even the 404 case is worth surfacing explicitly rather than
            # treated as identical to "file genuinely doesn't exist yet."
            try:
                body = e.read().decode("utf-8", errors="replace")
            except Exception:
                body = "<could not read response body>"
            read_error = f"HTTP {e.code}: {body[:500]}"
            source_missing = True

    # Schema check -- normalize + validate each entry, logging exactly why
    # anything gets dropped rather than a silent count. Not a redaction pass
    # (that already happened on Lauren's side per the WATCH item) -- purely
    # a defensive shape check plus field-name normalization on this mirror.
    clean_entries = []
    drop_reasons = []
    for e in entries:
        normalized, reason = normalize_entry(e)
        if normalized is not None:
            clean_entries.append(normalized)
        else:
            drop_reasons.append({"draft_id": e.get("draft_id") if isinstance(e, dict) else None, "reason": reason})

    payload = {
        # Timezone-aware UTC, not datetime.now() (local/BST) -- a naive
        # local timestamp sitting next to agent-commons' UTC timestamps is
        # what made a genuinely-stale mirror LOOK like it had run after a
        # source update it actually predated (11 Aug 2026 incident).
        "generated": datetime.datetime.now(datetime.timezone.utc).isoformat(),
        "source": "agent-commons/pending-email-drafts/drafts.json (mirrored, agent-commons stays private)",
        "source_missing": source_missing,
        "entries": clean_entries,
    }
    content_bytes = json.dumps(payload, ensure_ascii=True, indent=2).encode("utf-8")

    stats = {
        "source_missing": source_missing,
        "read_error": read_error,
        "entries_found": len(entries),
        "entries_published": len(clean_entries),
        "entries_dropped_bad_shape": len(drop_reasons),
        "drop_reasons": drop_reasons,
    }

    if dry_run:
        stats["pushed"] = False
        return stats

    existing_sha = None
    try:
        existing = gh_get(WI_OWNER, WI_REPO, "data/drafted_replies.json", token)
        existing_sha = existing["sha"]
    except urllib.error.HTTPError as e:
        if e.code != 404:
            raise

    result = gh_put(
        WI_OWNER, WI_REPO, "data/drafted_replies.json", content_bytes,
        f"Mirror drafted_replies.json from agent-commons: {len(clean_entries)} entries",
        existing_sha, token,
    )
    new_sha = result["content"]["sha"]
    remote = gh_blob(WI_OWNER, WI_REPO, new_sha, token)
    stats["pushed"] = True
    stats["byte_identical_verified"] = (remote == content_bytes)
    stats["new_sha"] = new_sha
    return stats


if __name__ == "__main__":
    print(f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] publish_drafted_replies.py run started")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    token = os.environ.get("GITHUB_PAT")
    if not token:
        print("FATAL: GITHUB_PAT not set")
        if not args.dry_run:
            notify_phase_failure(
                "Work Inbox Briefing - Drafted Replies publish",
                "GITHUB_PAT not set",
                log_filename="drafted_replies_failure_last.log",
            )
        sys.exit(1)

    try:
        result = run(token, dry_run=args.dry_run)
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(f"FATAL: publish_drafted_replies.py run failed - {type(e).__name__}: {e}")
        # Real desktop toast, not just a log line -- this script fails
        # silently to Kevin otherwise (see phase_failure_notify.py's own
        # docstring / HANDOVER.md Phase 2, 20-21 Aug 2026). Skipped on
        # --dry-run so manual/interactive test runs don't fire a false
        # production alert. Best-effort: never allowed to mask the FATAL
        # already printed above or change this script's own exit code.
        if not args.dry_run:
            notify_phase_failure(
                "Work Inbox Briefing - Drafted Replies publish",
                f"{type(e).__name__}: {e}",
                log_filename="drafted_replies_failure_last.log",
            )
        sys.exit(1)
