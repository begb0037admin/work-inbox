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
import sys
from datetime import datetime

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


def _key(m):
    """Cross-backend stable identity: prefer the internet Message-ID (both
    backends can carry it once imap_mail is in; COM side would need the same
    header added -- until then fall back to subject+sender+date-ish)."""
    mid = (m.get("message_id") or "").strip().strip("<>")
    if mid:
        return "mid:" + mid
    subj = (m.get("subject") or "").strip().lower()
    frm = (m.get("from_email") or m.get("from") or "").strip().lower()
    rcv = (m.get("received") or "")[:16]
    return f"s:{subj}|f:{frm}|r:{rcv}"


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
    for k in common:
        a, b = ck[k], ik[k]
        for fld in ("subject", "from_email", "is_read", "has_attachments",
                    "importance", "kevin_is_primary_recipient"):
            if (a.get(fld) or "") != (b.get(fld) or ""):
                field_mismatches.append({"key": k, "field": fld,
                                         "com": a.get(fld), "imap": b.get(fld)})
        da, db = _parse_dt(a.get("received")), _parse_dt(b.get("received"))
        if da and db and abs((da - db).total_seconds()) > RECEIVED_TOLERANCE_SECONDS:
            field_mismatches.append({"key": k, "field": "received",
                                     "com": a.get("received"), "imap": b.get("received")})
        if _tier(a) != _tier(b):
            field_mismatches.append({"key": k, "field": "derived_tier",
                                     "com": _tier(a), "imap": _tier(b)})

    report = {
        "label": label,
        "counts": {"com": len(com), "imap": len(imap), "common": len(common)},
        "only_in_com": [ck[k].get("subject") for k in only_com],
        "only_in_imap": [ik[k].get("subject") for k in only_imap],
        "field_mismatches": field_mismatches,
    }
    return report


def _print(rep):
    print(f"\n=== {rep['label']} ===")
    print(f"  counts: COM={rep['counts']['com']}  IMAP={rep['counts']['imap']}  "
          f"common={rep['counts']['common']}")
    if rep["only_in_com"]:
        print(f"  ONLY in COM ({len(rep['only_in_com'])}):")
        for s in rep["only_in_com"]:
            print(f"    - {s}")
    if rep["only_in_imap"]:
        print(f"  ONLY in IMAP ({len(rep['only_in_imap'])}):")
        for s in rep["only_in_imap"]:
            print(f"    - {s}")
    if rep["field_mismatches"]:
        print(f"  FIELD MISMATCHES ({len(rep['field_mismatches'])}):")
        for fm in rep["field_mismatches"]:
            print(f"    [{fm['field']}] {fm['key'][:70]}")
            print(f"        COM : {fm['com']!r}")
            print(f"        IMAP: {fm['imap']!r}")
    if not (rep["only_in_com"] or rep["only_in_imap"] or rep["field_mismatches"]):
        print("  PARITY OK")


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
    for r in reps:
        _print(r)

    out = os.path.join(PDIR, f"parity_{datetime.now():%Y%m%d_%H%M%S}.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump({"generated": _ts(), "reports": reps}, f, indent=2)
    print(f"\n[{_ts()}] wrote {out}")

    total_issues = sum(len(r["only_in_com"]) + len(r["only_in_imap"])
                       + len(r["field_mismatches"]) for r in reps)
    print(f"[{_ts()}] total parity issues: {total_issues}")
    return 0 if total_issues == 0 else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))
