r"""
parity_vs_briefing.py -- work-inbox laptop migration, Phase 5 (mail parity),
SELF-CONTAINED ON THE LAPTOP (no classic Outlook, no COM capture needed).

The strict same-window field-level parity (subject / from_email / is_read /
has_attachments / importance / kevin_is_primary_recipient, received +-120s,
message-id set diff) was already PROVEN on the admin desktop 29 Aug 2026 via
diff_mail_pull.py: "REAL parity issues: 0" (+31 benign X.500->SMTP, +5 read-cap
churn). That needs a COM capture and classic Outlook, which the laptop will not
run.

This script is the ONGOING confidence check Kevin/Lauren run from the laptop
across different inbox states: it pulls the live desktop briefing.json from
GitHub (the COM+claude-p pipeline's output) and checks that a fresh IMAP mail
pull SURFACES THE SAME MESSAGES, attributed the same way. Because briefing.json
is a triaged artifact (no message-id, no is_read/importance/has_attachments per
card, sender is a display name not an address) this is a COVERAGE + attribution
sanity check, not a byte-for-byte field diff -- and briefing.json is a snapshot
from an earlier run, so some drift (new mail since, items filed/read) is normal
and is reported as such, not as a failure.

HOW TO RUN (laptop, as ad-oak\begb0037):
    python parity_vs_briefing.py              # fresh IMAP capture + fetch briefing + diff
    python parity_vs_briefing.py --no-capture # reuse data/parallel/imap_inbox_raw.json
    python parity_vs_briefing.py --history 5  # also diff against the last 5 briefing.json commits

Writes data/parallel/parity_vs_briefing_<ts>.json + prints a summary.
Run once now, then ~once a day for 3-4 days to eyeball parity across states.
Every run prints timestamps. Reads only; pushes / mutates nothing.
"""

import json
import os
import re
import subprocess
import sys
import urllib.error
import urllib.request
from datetime import datetime

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

HERE = os.path.dirname(os.path.abspath(__file__))
PDIR = os.path.join(HERE, "data", "parallel")
IMAP_RAW = os.path.join(PDIR, "imap_inbox_raw.json")
REPO = "begb0037admin/work-inbox"
BRIEFING_PATH = "data/briefing.json"
RECEIVED_TOLERANCE_S = 120

try:
    import diff_mail_pull as _dmp   # pure module; reuse its _tier()
except Exception:
    _dmp = None


def _ts():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def log(msg):
    print(f"[{_ts()}] {msg}", flush=True)


# --------------------------------------------------------------------------- #
#  helpers
# --------------------------------------------------------------------------- #
def _nsub(s):
    return re.sub(r"\s+", " ", (s or "")).strip().lower()


def _nsub_stripped(s):
    """subject with leading Re:/Fw:/Fwd: runs removed -- fallback match key."""
    x = _nsub(s)
    while True:
        m = re.match(r"^(re|fw|fwd|aw|tr)\s*:\s*", x)
        if not m:
            return x
        x = x[m.end():]


def _parse_dt(s):
    """Accept 'YYYY-MM-DD HH:MM:SS', ISO with microseconds and/or tz offset,
    'YYYY-MM-DDTHH:MM:SS'. Return a naive datetime (local wall clock)."""
    s = (s or "").strip()
    if not s:
        return None
    s = s.split(" (")[0].strip()
    # split tz: a trailing '+HH:MM' / '-HH:MM' / 'Z'
    m = re.match(r"^(.*?)(?:Z|[+-]\d{2}:?\d{2})?$", s)
    core = (m.group(1) if m else s).strip()
    core = core.split(".")[0]          # drop fractional seconds
    core = core.replace("T", " ")
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"):
        try:
            return datetime.strptime(core, fmt)
        except Exception:
            pass
    return None


_GH_AUTH = None            # "Bearer <pat>" once _init_gh_auth() has validated it
_GH_AUTH_READY = False


def _init_gh_auth():
    """Validate GITHUB_PAT and build an ASCII-clean Authorization header value.
    A non-ASCII / whitespace-polluted PAT (bad paste, smart quote) is a clean
    exit with a fix hint -- NOT a raw UnicodeEncodeError deep in urllib."""
    global _GH_AUTH, _GH_AUTH_READY
    _GH_AUTH_READY = True
    raw = os.environ.get("GITHUB_PAT", "")
    pat = raw.strip()
    if not pat:
        log("GITHUB_PAT not set -- trying the GitHub API unauthenticated "
            "(works only if the repo is public; low rate limit).")
        return
    if pat != raw:
        log("note: GITHUB_PAT had surrounding whitespace -- stripped it.")
    if not pat.isascii():
        bad = [f"index {i} = U+{ord(c):04X} {c!r}" for i, c in enumerate(pat) if not c.isascii()]
        log("GITHUB_PAT contains NON-ASCII characters -- cannot form an HTTP header:")
        for b in bad[:6]:
            log(f"    {b}")
        log("  Re-set it clean (classic token: 40 chars, ^ghp_[A-Za-z0-9]{36}$):")
        log("    [Environment]::SetEnvironmentVariable('GITHUB_PAT','<paste token>','User')")
        log("  then open a NEW shell and check:")
        log("    $env:GITHUB_PAT.Length   ;   $env:GITHUB_PAT -match '^[\\x21-\\x7E]+$'")
        raise RuntimeError("GITHUB_PAT contains non-ASCII characters -- re-set it (see above)")
    if len(pat) < 20 or " " in pat:
        log(f"warning: GITHUB_PAT looks malformed (len {len(pat)}) -- if the fetch 401s, re-set it.")
    _GH_AUTH = "Bearer " + pat


def _gh_get(url, accept="application/vnd.github+json"):
    """GET api.github.com. Returns parsed JSON, or the raw text body when
    `accept` ends in 'raw'. Any auth/network/HTTP error -> RuntimeError with a
    one-line diagnosis + the URL (never a bare traceback)."""
    if not _GH_AUTH_READY:
        _init_gh_auth()
    req = urllib.request.Request(url, headers={
        "Accept": accept,
        "User-Agent": "work-inbox-parity",
        "X-GitHub-Api-Version": "2022-11-28",
    })
    if _GH_AUTH:
        req.add_header("Authorization", _GH_AUTH)
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            data = r.read()
    except urllib.error.HTTPError as e:
        body = ""
        try:
            body = e.read().decode("utf-8", "replace")
        except Exception:
            pass
        hint = ""
        if e.code in (401, 403):
            hint = " -- GITHUB_PAT missing/expired/insufficient scope (needs repo read)"
        elif e.code == 404:
            hint = " -- repo/path/ref not found, or the PAT can't see a private repo"
        raise RuntimeError(f"HTTP {e.code} {e.reason}{hint}  [{url}]  {body[:200]}")
    except urllib.error.URLError as e:
        raise RuntimeError(f"network error: {e.reason}  [{url}]")
    text = data.decode("utf-8", "replace")
    return text if accept.rstrip().endswith("raw") else json.loads(text)


def fetch_briefings(history):
    """-> [(as_of_dt, sha, briefing_dict), ...] newest first. `history` = how
    many commits of data/briefing.json to pull (>=1)."""
    n = max(1, history)
    commits = _gh_get(
        f"https://api.github.com/repos/{REPO}/commits?path={BRIEFING_PATH}&per_page={n}")
    if not isinstance(commits, list) or not commits:
        raise RuntimeError(f"commits API returned no history for {BRIEFING_PATH} "
                           f"(repo {REPO})")
    out = []
    for c in commits[:n]:
        sha = c["sha"]
        when = c["commit"]["committer"]["date"]            # e.g. 2026-08-28T12:06:41Z
        as_of = _parse_dt(when)
        body = _gh_get(
            f"https://api.github.com/repos/{REPO}/contents/{BRIEFING_PATH}?ref={sha}",
            accept="application/vnd.github.raw")
        doc = json.loads(body)
        out.append((as_of, sha[:9], doc))
    return out


def briefing_messages(doc):
    """Flatten briefing.json mail cards -> comparable pseudo-entries."""
    rows = []
    for tier in ("urgent", "needs", "fyi", "low"):
        for c in doc.get(tier, []) or []:
            rows.append({
                "subject": c.get("subject") or c.get("title") or "",
                "received": c.get("received_raw") or c.get("received") or "",
                "kevin_is_primary_recipient": c.get("kevin_is_primary_recipient"),
                "tier": tier,
                "needs_reply": c.get("needs_reply"),
                "message_count": c.get("messageCount"),
            })
    return rows


def _tier_of(imap_entry):
    if _dmp is not None:
        try:
            return _dmp._tier(imap_entry)
        except Exception:
            pass
    # minimal fallback
    if imap_entry.get("importance") == 2:
        return "urgent"
    return "needs" if not imap_entry.get("is_read", True) else "fyi"


# tiers that mean "this mattered" -- a miss here is a real flag
_HARD_TIERS = {"urgent", "needs"}


def diff_one(as_of, sha, doc, imap):
    bmsgs = briefing_messages(doc)

    # index imap by normalised subject (exact, then prefix-stripped)
    imap_pool = list(imap)
    by_sub = {}
    by_sub_s = {}
    for e in imap_pool:
        by_sub.setdefault(_nsub(e.get("subject")), []).append(e)
        by_sub_s.setdefault(_nsub_stripped(e.get("subject")), []).append(e)

    matched = []            # (bmsg, imap_entry)
    only_briefing = []
    used = set()            # id() of consumed imap entries

    for b in bmsgs:
        bdt = _parse_dt(b["received"])
        cands = by_sub.get(_nsub(b["subject"]), []) or by_sub_s.get(_nsub_stripped(b["subject"]), [])
        best = None
        best_gap = None
        for e in cands:
            if id(e) in used:
                continue
            edt = _parse_dt(e.get("received"))
            if bdt and edt:
                gap = abs((bdt - edt).total_seconds())
                if gap <= RECEIVED_TOLERANCE_S and (best_gap is None or gap < best_gap):
                    best, best_gap = e, gap
            elif best is None:
                best = e  # subject-only match when a date is missing
        if best is not None:
            used.add(id(best))
            matched.append((b, best))
        else:
            only_briefing.append(b)

    only_imap = [e for e in imap_pool if id(e) not in used]

    # classify only_briefing
    ob_hard, ob_soft = [], []
    matched_subs = {_nsub(b["subject"]) for b, _ in matched}
    for b in only_briefing:
        old = (bdt := _parse_dt(b["received"])) and as_of and bdt < as_of
        row = {"subject": b["subject"], "tier": b["tier"], "received": b["received"],
               "note": "pre-dates this briefing (may have been filed/read/deleted since)"
                       if old else "NOT older than the briefing -- unexpected"}
        (ob_hard if b["tier"] in _HARD_TIERS else ob_soft).append(row)

    # classify only_imap
    oi_new, oi_sibling, oi_hard, oi_soft = [], [], [], []
    for e in only_imap:
        edt = _parse_dt(e.get("received"))
        row = {"subject": e.get("subject"), "from_email": e.get("from_email"),
               "is_read": e.get("is_read"), "importance": e.get("importance"),
               "source_folder": e.get("source_folder", ""), "received": e.get("received"),
               "tier": _tier_of(e)}
        if as_of and edt and edt > as_of:
            row["note"] = "arrived AFTER the briefing was generated -- expected"
            oi_new.append(row)
        elif _nsub(e.get("subject")) in matched_subs:
            row["note"] = "same thread as a matched briefing card (grouped) -- not a gap"
            oi_sibling.append(row)
        elif row["tier"] in _HARD_TIERS and not e.get("is_read", True):
            row["note"] = "IMAP surfaced an UNREAD needs/urgent item the briefing did not -- REVIEW"
            oi_hard.append(row)
        else:
            row["note"] = "read / fyi-tier only-in-IMAP -- read-cap boundary churn or drift"
            oi_soft.append(row)

    # matched-pair field checks
    kipr_mismatch, tier_soft = [], []
    for b, e in matched:
        if bool(b.get("kevin_is_primary_recipient")) != bool(e.get("kevin_is_primary_recipient")):
            kipr_mismatch.append({"subject": b["subject"],
                                  "briefing": b.get("kevin_is_primary_recipient"),
                                  "imap": e.get("kevin_is_primary_recipient")})
        bt, it = b["tier"], _tier_of(e)
        if bt != it:
            tier_soft.append({"subject": b["subject"], "briefing_tier": bt, "imap_tier": it})

    # subfolder gap check
    cdr_seen = any("cdr" in (e.get("source_folder", "") or "").lower()
                   or "working group" in (e.get("source_folder", "") or "").lower()
                   for e in imap)

    real_flags = len(ob_hard) + len(oi_hard) + len(kipr_mismatch)
    return {
        "briefing_sha": sha,
        "briefing_as_of": as_of.strftime("%Y-%m-%d %H:%M:%S") if as_of else None,
        "counts": {
            "briefing_cards": len(bmsgs),
            "imap_entries": len(imap),
            "matched": len(matched),
        },
        "REAL_FLAGS": real_flags,
        "only_in_briefing_HARD_needs_urgent": ob_hard,
        "only_in_briefing_soft_fyi_low": ob_soft,
        "only_in_imap_arrived_after": oi_new,
        "only_in_imap_grouped_thread_sibling": oi_sibling,
        "only_in_imap_HARD_unread_needs_urgent": oi_hard,
        "only_in_imap_soft_readcap_or_drift": oi_soft,
        "kevin_is_primary_recipient_mismatch_REAL": kipr_mismatch,
        "derived_tier_differs_SOFT": tier_soft,
        "cdr_subfolder_present_in_imap": cdr_seen,
    }


def folder_diagnostic():
    """Dump NAMESPACE + LIST rows near the 'Bi-monthly CDR/PD working group'
    folder so the '/'-in-name gap can be closed. Reuses imap_mail's auth."""
    try:
        import imap_mail
    except Exception as e:
        return {"error": f"cannot import imap_mail: {e!r}"}
    try:
        token, upn = imap_mail.acquire_token_silent(log=log)
        M = imap_mail._imap_connect(token, upn, log=log)
    except Exception as e:
        return {"error": f"IMAP connect failed: {e!r}"}
    diag = {"upn": upn}
    try:
        try:
            diag["namespace"] = [x.decode("ascii", "replace") if isinstance(x, bytes) else str(x)
                                 for x in (M.namespace()[1] or [])]
        except Exception as e:
            diag["namespace"] = f"namespace() failed: {e!r}"
        typ, data = M.list()
        rows = []
        for row in (data or []):
            if not row:
                continue
            s = row.decode("ascii", "replace") if isinstance(row, bytes) else str(row)
            low = s.lower()
            if "cdr" in low or "working group" in low or "bi-monthly" in low:
                rows.append(s)
        diag["list_rows_matching_cdr"] = rows
        diag["all_inbox_children"] = [
            (row.decode("ascii", "replace") if isinstance(row, bytes) else str(row))
            for row in (data or []) if row and b"INBOX/" in (row if isinstance(row, bytes) else row.encode())
        ][:60]
    finally:
        try:
            M.logout()
        except Exception:
            pass
    return diag


def _print_folder_diag(diag):
    print("\n--- folder diagnostic (for the 'Bi-monthly CDR/PD working group' "
          "'/'-in-name subfolder fix) ---")
    if diag.get("error"):
        print(f"  ERROR: {diag['error']}")
        return
    print(f"  upn: {diag.get('upn')}")
    print(f"  NAMESPACE: {diag.get('namespace')}")
    rows = diag.get("list_rows_matching_cdr") or []
    print(f"  LIST rows matching cdr / working group / bi-monthly ({len(rows)}):")
    for s in rows:
        print(f"       {s}")
    if not rows:
        print("       (none returned by LIST at all -- inspect all_inbox_children:)")
        for s in (diag.get("all_inbox_children") or [])[:40]:
            print(f"       {s}")


# --------------------------------------------------------------------------- #
#  main
# --------------------------------------------------------------------------- #
def main(argv):
    no_capture = "--no-capture" in argv
    history = 1
    for i, a in enumerate(argv):
        if a == "--history" and i + 1 < len(argv):
            try:
                history = int(argv[i + 1])
            except Exception:
                pass

    log("parity_vs_briefing starting")

    if not no_capture:
        log("running a fresh IMAP mail capture (MAIL_BACKEND=imap WI_MAIL_PARALLEL=1) ...")
        env = dict(os.environ)
        env["MAIL_BACKEND"] = "imap"
        env["WI_MAIL_PARALLEL"] = "1"
        r = subprocess.run([sys.executable, os.path.join(HERE, "fetch_inbox.py")],
                           cwd=HERE, env=env)
        if r.returncode != 0:
            log(f"FATAL: capture exited {r.returncode}")
            return 2

    if not os.path.exists(IMAP_RAW):
        log(f"FATAL: {IMAP_RAW} not found -- run without --no-capture first.")
        return 2
    with open(IMAP_RAW, "r", encoding="utf-8") as f:
        imap = json.load(f)
    log(f"IMAP capture: {len(imap)} inbox entries "
        f"({sum(1 for e in imap if not e.get('is_read', True))} unread)")

    os.makedirs(PDIR, exist_ok=True)
    out = os.path.join(PDIR, f"parity_vs_briefing_{datetime.now():%Y%m%d_%H%M%S}.json")

    # ---- FOLDER DIAGNOSTIC FIRST -- the priority output. Runs and prints even
    #      if the briefing fetch below fails (bad PAT / offline). ----
    diag = folder_diagnostic()
    _print_folder_diag(diag)
    # persist immediately so a later failure can't lose it
    with open(out, "w", encoding="utf-8") as f:
        json.dump({"generated": _ts(), "imap_count": len(imap),
                   "reports": [], "folder_diagnostic": diag,
                   "note": "briefing diff not run yet"}, f, indent=2)

    # ---- briefing fetch (graceful) ----
    try:
        briefings = fetch_briefings(history)
    except Exception as e:
        log(f"could not fetch briefing.json from GitHub: {e}")
        log("  The folder diagnostic above is valid. Fix GITHUB_PAT / connectivity "
            "and re-run for the parity diff.")
        with open(out, "w", encoding="utf-8") as f:
            json.dump({"generated": _ts(), "imap_count": len(imap),
                       "reports": [], "folder_diagnostic": diag,
                       "briefing_fetch_error": str(e)}, f, indent=2)
        log(f"wrote {out} (folder diagnostic only)")
        return 2

    reports = []
    for as_of, sha, doc in briefings:
        rep = diff_one(as_of, sha, doc, imap)
        reports.append(rep)

    with open(out, "w", encoding="utf-8") as f:
        json.dump({"generated": _ts(), "imap_count": len(imap),
                   "reports": reports, "folder_diagnostic": diag}, f, indent=2)

    # ---- console summary ----
    for rep in reports:
        c = rep["counts"]
        print(f"\n=== briefing {rep['briefing_sha']} as of {rep['briefing_as_of']} "
              f"vs IMAP now ===")
        print(f"  cards={c['briefing_cards']}  imap={c['imap_entries']}  matched={c['matched']}")
        print(f"  REAL FLAGS: {rep['REAL_FLAGS']}")
        for label in ("only_in_briefing_HARD_needs_urgent",
                      "only_in_imap_HARD_unread_needs_urgent",
                      "kevin_is_primary_recipient_mismatch_REAL"):
            rows = rep[label]
            if rows:
                print(f"  !! {label} ({len(rows)}):")
                for x in rows[:12]:
                    print(f"       - {x.get('subject','')[:90]}  [{x.get('note','')}]")
        exp = (len(rep["only_in_briefing_soft_fyi_low"])
               + len(rep["only_in_imap_arrived_after"])
               + len(rep["only_in_imap_grouped_thread_sibling"])
               + len(rep["only_in_imap_soft_readcap_or_drift"])
               + len(rep["derived_tier_differs_SOFT"]))
        print(f"  expected/soft (drift, read-cap churn, grouped threads, tier-fn diff): {exp}")
        if not rep["cdr_subfolder_present_in_imap"]:
            print("  !! SUBFOLDER GAP: no IMAP entry from 'Bi-monthly CDR/PD working group' "
                  "-- '/' in the name. PHASE 5 BLOCKER -- see the folder diagnostic above / the json.")

    total_real = sum(r["REAL_FLAGS"] for r in reports)
    print(f"\n[{_ts()}] wrote {out}")
    print(f"[{_ts()}] TOTAL REAL FLAGS across {len(reports)} briefing snapshot(s): {total_real}")
    print(f"[{_ts()}] 0 real flags + only soft/expected rows = IMAP pull matches the live "
          f"briefing's mail coverage. Run again tomorrow to see it across a different state.")
    return 0 if total_real == 0 else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
