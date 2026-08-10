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
"""

import argparse
import base64
import json
import os
import sys
import urllib.request
import urllib.error

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


def run(token, dry_run=False):
    try:
        source = gh_get(AC_OWNER, AC_REPO, "pending-email-drafts/drafts.json", token)
        drafts = json.loads(base64.b64decode(source["content"]))
        entries = drafts.get("entries", []) if isinstance(drafts, dict) else drafts
        source_missing = False
    except urllib.error.HTTPError as e:
        if e.code == 404:
            entries = []
            source_missing = True
        else:
            raise

    # Schema check -- only pass through entries with the fields the panel
    # needs, so a malformed upstream entry can't silently ship as an empty
    # card. Not a redaction pass (that already happened on Lauren's side
    # per the WATCH item) -- purely a defensive shape check on this mirror.
    required = {"source_entry_id", "subject", "sender_tier", "draft_text", "drafted_at"}
    clean_entries = [e for e in entries if isinstance(e, dict) and required.issubset(e.keys())]
    dropped = len(entries) - len(clean_entries)

    payload = {
        "generated": __import__("datetime").datetime.now().isoformat(),
        "source": "agent-commons/pending-email-drafts/drafts.json (mirrored, agent-commons stays private)",
        "source_missing": source_missing,
        "entries": clean_entries,
    }
    content_bytes = json.dumps(payload, ensure_ascii=True, indent=2).encode("utf-8")

    stats = {
        "source_missing": source_missing,
        "entries_found": len(entries),
        "entries_published": len(clean_entries),
        "entries_dropped_bad_shape": dropped,
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
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--dry-run", action="store_true")
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
        print(f"FATAL: publish_drafted_replies.py run failed - {type(e).__name__}: {e}")
        sys.exit(1)
