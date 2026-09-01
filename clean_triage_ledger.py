"""One-time cleanup for data/triage_ledger.json after the MAIL_BACKEND=imap
cutover (see fetch_inbox.py Phase 3.5/3.6 identity fix, drew memory
wi-cc-feed-imap-entryid-fix).

Problem it fixes
----------------
While the live pipeline ran MAIL_BACKEND=imap with the pre-fix code, every
email had entry_id == "".  The `applied` map is keyed
`<mail_key>_<task_id>`, so those writes collapsed to `_<task_id>` -- one
poisoned key per task that then permanently barred that task from ever
receiving another inbox update.  This script removes those poisoned keys
(and only those) so the affected tasks can be updated again.  `promoted`
and `tracked_needs_urgent` are inspected but never modified.

A real key never starts with "_" (a real Outlook EntryID is hex; a real
internet Message-ID is localpart@domain) -- so "key starts with '_'" is an
exact, safe identifier for the poison.

GitHub-only.  Needs $GITHUB_PAT.  Dry-run by default.

  python clean_triage_ledger.py                 # dry run, shows before/after
  python clean_triage_ledger.py --apply         # backup -> verify -> write -> verify

Every run prints a UTC timestamp.
"""
import base64
import json
import os
import sys
import urllib.error
import urllib.request
from datetime import datetime, timezone

REPO = os.environ.get("WI_GITHUB_REPO", "begb0037admin/work-inbox")
PATH = "data/triage_ledger.json"
PAT = os.environ.get("GITHUB_PAT", "")


def _ts():
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _api(method, url, body=None):
    hdrs = {
        "Authorization": "token " + PAT,
        "User-Agent": "clean-triage-ledger",
        "Accept": "application/vnd.github+json",
    }
    data = None
    if body is not None:
        data = json.dumps(body).encode("utf-8")
        hdrs["Content-Type"] = "application/json"
    req = urllib.request.Request(url, data=data, headers=hdrs, method=method)
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read())


def _get_file():
    meta = _api("GET", "https://api.github.com/repos/%s/contents/%s" % (REPO, PATH))
    raw = base64.b64decode(meta["content"])
    return meta["sha"], raw


def main(argv):
    apply = "--apply" in argv
    if not PAT:
        print("FATAL: GITHUB_PAT not set")
        return 2

    print("[%s] clean_triage_ledger  repo=%s  mode=%s" % (_ts(), REPO, "APPLY" if apply else "dry-run"))
    sha, raw = _get_file()
    if not raw:
        print("FATAL: fetched file is empty -- refusing to proceed")
        return 2
    doc = json.loads(raw.decode("utf-8"))

    applied = doc.get("applied", {})
    promoted = doc.get("promoted", {})
    tracked = doc.get("tracked_needs_urgent", {})
    poison = sorted(k for k in applied if k.startswith("_"))

    print("  BEFORE:  applied=%d  promoted=%d  tracked_needs_urgent=%d" % (len(applied), len(promoted), len(tracked)))
    print("  poisoned '_'-prefixed applied keys: %d" % len(poison))
    for k in poison:
        print("     - %s   (was applied %s)" % (k, applied[k]))

    if not poison:
        print("  nothing to clean. exit 0")
        return 0

    cleaned = dict(doc)
    cleaned["applied"] = {k: v for k, v in applied.items() if not k.startswith("_")}
    print("  AFTER :   applied=%d  promoted=%d  tracked_needs_urgent=%d  (removed %d)"
          % (len(cleaned["applied"]), len(promoted), len(tracked), len(poison)))

    # sanity: only the poison keys differ, nothing else touched
    removed = set(applied) - set(cleaned["applied"])
    assert removed == set(poison), removed
    assert cleaned["promoted"] == promoted and cleaned.get("tracked_needs_urgent", {}) == tracked

    if not apply:
        print("[%s] dry run only. re-run with --apply to write (backup + verify included)." % _ts())
        return 0

    # 1. backup current bytes
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    bpath = "data/archive/triage_ledger_%s.json" % stamp
    burl = "https://api.github.com/repos/%s/contents/%s" % (REPO, bpath)
    _api("PUT", burl, {
        "message": "backup: triage_ledger before poison-key cleanup %s" % _ts(),
        "content": base64.b64encode(raw).decode("ascii"),
    })
    print("[%s] backup committed: %s" % (_ts(), bpath))

    # 2. verify backup is byte-identical
    bmeta = _api("GET", burl)
    bback = base64.b64decode(bmeta["content"])
    if bback != raw:
        print("FATAL: backup verify FAILED (byte mismatch). aborting before any write to %s" % PATH)
        return 3
    print("[%s] backup verified byte-identical (%d bytes)" % (_ts(), len(bback)))

    # 3. write cleaned file (optimistic lock on the sha we read)
    new_bytes = (json.dumps(cleaned, indent=1) + "\n").encode("utf-8")
    _api("PUT", "https://api.github.com/repos/%s/contents/%s" % (REPO, PATH), {
        "message": "chore: remove %d poisoned '_'-prefixed applied keys from triage ledger (IMAP entry_id='' era) %s" % (len(poison), _ts()),
        "content": base64.b64encode(new_bytes).decode("ascii"),
        "sha": sha,
    })
    print("[%s] cleaned %s written" % (_ts(), PATH))

    # 4. verify post-write
    _sha2, raw2 = _get_file()
    doc2 = json.loads(raw2.decode("utf-8"))
    still = [k for k in doc2.get("applied", {}) if k.startswith("_")]
    if still:
        print("FATAL: post-write verify FAILED -- poison keys still present: %s" % still)
        return 3
    if doc2.get("promoted", {}) != promoted or doc2.get("tracked_needs_urgent", {}) != tracked:
        print("FATAL: post-write verify FAILED -- promoted/tracked changed unexpectedly")
        return 3
    print("[%s] post-write verify OK: applied=%d, promoted/tracked unchanged. backup at %s"
          % (_ts(), len(doc2.get("applied", {})), bpath))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
