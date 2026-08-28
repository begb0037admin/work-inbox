"""
diff_mail_pull.py -- parity check between the Outlook-COM mail pull and the
IMAP+OAuth2 mail pull, for the MAIL_BACKEND migration (cautious-change-pace;
see the 17 Aug 2026 regression+revert -- do NOT flip the backend until this is
clean over several scheduled cycles and Kevin has said go).

Usage:
    # in the same scheduled window, capture both:
    set MAIL_BACKEND=com   & set WI_MAIL_PARALLEL=1 & python fetch_inbox.py
    set MAIL_BACKEND=imap  & set WI_MAIL_PARALLEL=1 & python fetch_inbox.py
    # WI_MAIL_PARALLEL=1 makes Phase 1 dump its raw mail lists locally and
    # push / mutate NOTHING. It writes:
    #     data/parallel/com_inbox_raw.json  data/parallel/com_sent_raw.json
    #     data/parallel/imap_inbox_raw.json data/parallel/imap_sent_raw.json
    #
    python diff_mail_pull.py            # diffs whatever is in data/parallel/
    python diff_mail_pull.py A.json B.json   # or diff two explicit inbox dumps

Prints a human-readable report AND writes data/parallel/parity_<ts>.json.
Every run prints a timestamp.
"""

import json
import os
import re
import sys
from datetime import datetime

# The parity subjects contain em-dashes, emoji, etc. -- a cp1252 Windows
# console will UnicodeEncodeError on print(). Force UTF-8, replace on failure.
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

HERE = os.path.dirname(os.path.abspath(__file__))
PDIR = os.path.join(HERE, "data", "parallel")

# received timestamps can legitimately differ by a few seconds
# (Outlook ReceivedTime vs IMAP INTERNALDATE) -- not a parity failure.
RECEIVED_TOLERANCE_SECONDS = 120


def _ts():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _load(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _norm_mid(m):
    return (m.get("message_id") or "").strip().strip("<>").lower()


def _key(m):
    """Cross-backend stable identity. Both backends now capture the internet
    Message-ID (COM via PR_INTERNET_MESSAGE_ID, IMAP from the fetched headers),
    so join on that. Fallback (should be rare): subject + a coarse received
    date -- NOT from_email, since the COM X.500 vs IMAP SMTP difference in that
    field is one of the things the diff is meant to surface."""
    mid = _norm_mid(m)
    if mid:
        return "mid:" + mid
    subj = re.sub(r"\s+", " ", (m.get("subject") or "").strip().lower())
    rcv = (m.get("received") or "")[:10]
    return f"s:{subj}|r:{rcv}"


def _parse_dt(s):
    s = (s or "").split("+")[0].split(" (")[0].strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S"):
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            pass
    return None


# categorise() logic, mirrored here so we can compare the DERIVED tier, not
# just the raw fields -- the tier is what actually reaches the dashboard.
URGENT_SUBJECTS = ["major incident", "priority 1", "p1", "urgent", "critical", "security vulnerab"]
NEEDS_SUBJECTS = ["re:", "fw:", "fwd:", "action", "required", "please", "timeline", "update",
                  "chasing", "waiting", "overdue", "follow", "scoping", "handover", "error",
                  "import", "failed", "issue", "case ", "support"]
FYI_SUBJECTS = ["fyi", "notification", "scheduled", "maintenance", "summary", "workshop",
                "invitation", "invite", "digest", "recap", "newsletter", "annual leave",
                "out of office", "automatic reply", "accepted:", "declined:", "cancelled:"]
LOW_SUBJECTS = ["unsubscribe", "noreply", "no-reply", "do not reply", "automated",
                "github", "pages", "build", "deploy", "run failed", "wisp"]


def _tier(m):
    subj = (m.get("subject") or "").lower()
    sender = (m.get("from_email") or "").lower()
    is_read = m.get("is_read", True)
    imp = m.get("importance", 1)
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


def diff_lists(com, imap, label):
    ck = {_key(m): m for m in com}
    ik = {_key(m): m for m in imap}
    only_com = sorted(set(ck) - set(ik))
    only_imap = sorted(set(ik) - set(ck))
    common = sorted(set(ck) & set(ik))

    field_mismatches = []
    benign = []   # known, expected differences -- reported separately, not "issues"

    def _nsub(s):
        return re.sub(r"\s+", " ", (s or "")).strip().lower()

    for k in common:
        a, b = ck[k], ik[k]
        # subject: compare whitespace-normalised + case-folded (COM keeps
        # trailing spaces; IMAP unfolds CRLF) -- a real difference still shows.
        if _nsub(a.get("subject")) != _nsub(b.get("subject")):
            field_mismatches.append({"key": k, "field": "subject",
                                     "com": a.get("subject"), "imap": b.get("subject")})
        for fld in ("is_read", "has_attachments", "importance",
                    "kevin_is_primary_recipient"):
            if (a.get(fld) or "") != (b.get(fld) or ""):
                field_mismatches.append({"key": k, "field": fld,
                                         "com": a.get(fld), "imap": b.get(fld)})
        # from_email: benign if it's only a case difference, or COM returned an
        # X.500 DN for an internal sender (IMAP gives the real SMTP -- an
        # improvement). A genuinely different address is a real mismatch.
        ce, ie = (a.get("from_email") or ""), (b.get("from_email") or "")
        if ce != ie:
            case_only = ce.lower() == ie.lower()
            x500 = ce.startswith("/") or ce == "" or "=" in ce.split("@")[0]
            (benign if (case_only or x500) else field_mismatches).append(
                {"key": k, "field": "from_email", "com": ce, "imap": ie})
        da, db = _parse_dt(a.get("received")), _parse_dt(b.get("received"))
        if da and db and abs((da - db).total_seconds()) > RECEIVED_TOLERANCE_SECONDS:
            field_mismatches.append({"key": k, "field": "received",
                                     "com": a.get("received"), "imap": b.get("received")})
        if _tier(a) != _tier(b):
            field_mismatches.append({"key": k, "field": "derived_tier",
                                     "com": _tier(a), "imap": _tier(b)})

    # An only-in-one-side row that is UNREAD is a real concern (a message one
    # backend surfaced and the other missed). An only-in-one-side row that is
    # READ is almost always read-cap boundary churn -- COM caps the 30 newest
    # read by ReceivedTime, IMAP by UID order; near the boundary they pick a
    # slightly different 30. Split them so the issue count reflects real gaps.
    only_com_unread = [ck[k].get("subject") for k in only_com if not ck[k].get("is_read")]
    only_imap_unread = [ik[k].get("subject") for k in only_imap if not ik[k].get("is_read")]
    only_com_readcap = [ck[k].get("subject") for k in only_com if ck[k].get("is_read")]
    only_imap_readcap = [ik[k].get("subject") for k in only_imap if ik[k].get("is_read")]

    report = {
        "label": label,
        "counts": {"com": len(com), "imap": len(imap), "common": len(common)},
        "only_in_com_UNREAD": only_com_unread,
        "only_in_imap_UNREAD": only_imap_unread,
        "only_in_com_readcap_churn": only_com_readcap,
        "only_in_imap_readcap_churn": only_imap_readcap,
        "only_in_com": [ck[k].get("subject") for k in only_com],
        "only_in_imap": [ik[k].get("subject") for k in only_imap],
        "field_mismatches": field_mismatches,
        "benign_diffs": benign,
    }
    return report


def _print(rep):
    print(f"\n=== {rep['label']} ===")
    print(f"  counts: COM={rep['counts']['com']}  IMAP={rep['counts']['imap']}  "
          f"common={rep['counts']['common']}")
    if rep.get("only_in_com_UNREAD"):
        print(f"  !! ONLY in COM, UNREAD ({len(rep['only_in_com_UNREAD'])}) -- real gap:")
        for s in rep["only_in_com_UNREAD"]:
            print(f"    - {s}")
    if rep.get("only_in_imap_UNREAD"):
        print(f"  !! ONLY in IMAP, UNREAD ({len(rep['only_in_imap_UNREAD'])}) -- real gap:")
        for s in rep["only_in_imap_UNREAD"]:
            print(f"    - {s}")
    _cc = rep.get("only_in_com_readcap_churn", []) + rep.get("only_in_imap_readcap_churn", [])
    if _cc:
        print(f"  only-in-one-side but READ ({len(_cc)}) -- read-cap boundary churn, not a gap:")
        for s in _cc:
            print(f"    - {s}")
    if rep["field_mismatches"]:
        print(f"  FIELD MISMATCHES ({len(rep['field_mismatches'])}):")
        for fm in rep["field_mismatches"]:
            print(f"    [{fm['field']}] {fm['key'][:70]}")
            print(f"        COM : {fm['com']!r}")
            print(f"        IMAP: {fm['imap']!r}")
    if rep.get("benign_diffs"):
        print(f"  benign/expected diffs ({len(rep['benign_diffs'])}): "
              f"e.g. from_email X.500->SMTP -- not counted as issues")
    _real = (rep.get("only_in_com_UNREAD", []) + rep.get("only_in_imap_UNREAD", [])
             + rep["field_mismatches"])
    if not _real:
        extra = []
        if rep.get("benign_diffs"):
            extra.append(f"{len(rep['benign_diffs'])} benign")
        if _cc:
            extra.append(f"{len(_cc)} read-cap churn")
        print("  PARITY OK" + (f" ({', '.join(extra)})" if extra else ""))


def main(argv):
    print(f"[{_ts()}] diff_mail_pull starting")
    if len(argv) == 3:
        a, b = _load(argv[1]), _load(argv[2])
        rep = diff_lists(a, b, os.path.basename(argv[1]) + " vs " + os.path.basename(argv[2]))
        _print(rep)
        return 0

    need = ["com_inbox_raw.json", "imap_inbox_raw.json",
            "com_sent_raw.json", "imap_sent_raw.json"]
    missing = [n for n in need if not os.path.exists(os.path.join(PDIR, n))]
    if missing:
        print(f"[{_ts()}] missing capture files in {PDIR}: {missing}")
        print("  Run fetch_inbox.py with WI_MAIL_PARALLEL=1 on BOTH backends first.")
        return 2

    reps = [
        diff_lists(_load(os.path.join(PDIR, "com_inbox_raw.json")),
                   _load(os.path.join(PDIR, "imap_inbox_raw.json")), "INBOX pull"),
        diff_lists(_load(os.path.join(PDIR, "com_sent_raw.json")),
                   _load(os.path.join(PDIR, "imap_sent_raw.json")), "SENT pull"),
    ]
    # "Real" issues = field mismatches + UNREAD-only-on-one-side rows.
    # Read-cap boundary churn and benign X.500->SMTP diffs are reported but
    # NOT counted -- they are expected and don't block a cutover decision.
    real_issues = sum(len(r["field_mismatches"])
                      + len(r.get("only_in_com_UNREAD", []))
                      + len(r.get("only_in_imap_UNREAD", [])) for r in reps)
    churn = sum(len(r.get("only_in_com_readcap_churn", []))
                + len(r.get("only_in_imap_readcap_churn", [])) for r in reps)
    benign = sum(len(r.get("benign_diffs", [])) for r in reps)

    # Write the machine-readable result FIRST, so a console-print failure
    # (e.g. an exotic char on a narrow codepage) can never lose it.
    out = os.path.join(PDIR, f"parity_{datetime.now():%Y%m%d_%H%M%S}.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump({"generated": _ts(), "real_issues": real_issues,
                   "readcap_churn": churn, "benign": benign,
                   "reports": reps}, f, indent=2)

    for r in reps:
        try:
            _print(r)
        except Exception as e:
            print(f"  (console print failed: {e} -- see {out})")

    print(f"\n[{_ts()}] wrote {out}")
    print(f"[{_ts()}] REAL parity issues: {real_issues}   "
          f"(+ {churn} read-cap churn, {benign} benign X.500->SMTP -- both expected)")
    print(f"[{_ts()}] 0 real issues = mail pulls match; else read the report above / the json")
    return 0 if real_issues == 0 else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))
