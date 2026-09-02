#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
lane_b_cal_guard.py -- Lane B calendar HALT kill-switch (LANE_B_TEAMS_CAL_DESIGN.md sec.6a)
=========================================================================================

The SECOND HALT layer for Lane B calendar. The first is lane_b_call1.py's
re-contamination guard (HALT if a write / non-allowlisted / non-codex_apps tool
is seen). This one is deliberately asymmetric and stricter: it takes a
`list_events` snapshot immediately BEFORE the Call-1 pull and again immediately
AFTER, and HALTS the whole run on ANY detected calendar change during that
window -- because the calendar blast radius (decline / cancel / RSVP notices
to real attendees) justifies favouring safety over uptime.

Single entry point for the wrapper:

  python lane_b_cal_guard.py --run            # pre-snap -> lane_b_call1 --domain calendar -> post-snap -> diff

  exit 0  = clean. data/lane_b/lane_b_normalised.json is fresh and trustworthy.
  exit 1  = GUARD TRIPPED (a calendar change happened during the read window,
            OR lane_b_call1's own re-contamination guard tripped). The freshly
            written lane_b_normalised.json is QUARANTINED (renamed
            .halted_<ts>) so fetch_inbox.py falls back to "calendar empty".
            data/codex_runs/GUARD_TRIPPED_cal_<ts>.json holds the diff.
            The WRAPPER is responsible for `Disable-ScheduledTask` + a toast.
  exit 2  = usage / environment error.
  exit 3  = a snapshot codex run failed (can't verify -> treated as unsafe: the
            normalised file is quarantined too, no calendar this run).

Also usable in pieces (for tests / manual):
  python lane_b_cal_guard.py --snapshot --out data/codex_runs/cal_baseline_<ts>.json
  python lane_b_cal_guard.py --diff  --pre <baseline.json> --post <after.json>

NOTE (1 Sept 2026): the codex_apps calendar event object has NO
lastModifiedDateTime / last_modified field (confirmed from the real probe), so
the per-event fingerprint diffs on {subject, start, end, response_status, type}
keyed by id. That still catches an add, a drop, a reschedule, an RSVP change,
and a single->cancelled-occurrence flip. Any timestamp-only edit that changed
none of those is not detectable via this surface -- documented residual.
"""

from __future__ import annotations

import argparse
import datetime as _dt
import json
import subprocess
import sys
import time as _time
from pathlib import Path

import lane_b_call1 as lb

REPO_ROOT      = Path(__file__).resolve().parent
LANE_B_DIR     = REPO_ROOT / "data" / "lane_b"
CODEX_RUNS_DIR = REPO_ROOT / "data" / "codex_runs"
NORMALISED     = LANE_B_DIR / "lane_b_normalised.json"

SNAPSHOT_TIMEOUT_S = int(lb.os.environ.get("WI_LANE_B_SNAP_TIMEOUT", "360"))


def _log(m: str) -> None:
    print(f"[{_dt.datetime.now(_dt.timezone.utc).strftime('%Y-%m-%d %H:%M:%SZ')}] lane_b_cal_guard: {m}")


def _snapshot_prompt(win_start_iso: str, win_end_iso: str) -> str:
    return (
        "Using the Microsoft Outlook Calendar app connector, retrieve my calendar events "
        f"between {win_start_iso} and {win_end_iso} from my default calendar and, if it "
        f"exists, the calendar named \"{lb.SHARED_CAL_NAME}\". "
        "Return ONLY the raw connector result as a JSON array of the event objects, with no "
        "summary and no prose. Do not use any other app or tool. Do not create, update, "
        "cancel, delete, move, respond to, or add an attachment to any event. Send nothing. "
        f"{lb.SAFETY_RULE}"
    )


SNAP_RETRIES = max(1, int(lb.os.environ.get("WI_LANE_B_RETRIES", "3")))

# --------------------------------------------------------------------------- #
#  Snapshot normalisation (1 Sept 2026 false-positive fix).
#  Two back-to-back reads of an UNCHANGED calendar were diffing by ~52. Cause:
#   (a) recurring-series OCCURRENCES get a fresh Graph `id` per connector call
#       -> every occurrence looked removed+added;  (b) start/end were compared
#       as raw `.isoformat()` of whatever tz the connector happened to render
#       (UTC one call, a Windows tz label the next) for the SAME instant;
#   (c) response_status / whitespace re-casing.
#  Fix: match on a STABLE natural key (iCalUID, else id, else subject) + the
#  event's start INSTANT canonicalised to UTC; compare only genuinely
#  load-bearing fields, each normalised. Ordering was already irrelevant (the
#  fingerprint is a dict + set diff). The re-contamination guard is untouched.
_CMP_FIELDS = ("subject", "start", "end", "response_status", "all_day")
_STATUS_ALIAS = {"notresponded": "none", "": "none", "not_responded": "none",
                 "tentativelyaccepted": "tentative", "tentatively_accepted": "tentative"}
_WIN_TZ = {
    "gmt standard time": "Europe/London", "w. europe standard time": "Europe/Berlin",
    "central europe standard time": "Europe/Budapest", "romance standard time": "Europe/Paris",
    "greenwich standard time": "Atlantic/Reykjavik", "eastern standard time": "America/New_York",
    "central standard time": "America/Chicago", "pacific standard time": "America/Los_Angeles",
    "india standard time": "Asia/Kolkata", "singapore standard time": "Asia/Singapore",
    "aus eastern standard time": "Australia/Sydney",
}


def _norm_txt(v) -> str:
    return " ".join(str(v or "").split())


def _norm_status(ev: dict) -> str:
    rs = ev.get("response_status") or ev.get("responseStatus") or {}
    v = rs.get("response") if isinstance(rs, dict) else str(rs or "")
    v = (v or "").strip().casefold()
    return _STATUS_ALIAS.get(v, v)


def _tzinfo(name: str):
    key = (name or "").strip().lower()
    if key in ("", "utc", "gmt", "z", "tzid=utc"):
        return _dt.timezone.utc
    ia = _WIN_TZ.get(key) or (name if "/" in (name or "") else None)
    if ia:
        try:
            from zoneinfo import ZoneInfo
            return ZoneInfo(ia)
        except Exception:
            return None
    return None


def _instant_utc(sdt, tz_name: str, all_day: bool) -> str:
    """Canonical UTC ISO-8601 'Z' for a timed event; bare 'YYYY-MM-DD' for an
    all-day event. Applied identically to PRE and POST, so an unchanged event
    matches itself regardless of how the connector rendered the timezone."""
    if sdt is None:
        return ""
    if all_day:
        return sdt.date().isoformat()
    if sdt.tzinfo is not None:
        return sdt.astimezone(_dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    tz = _tzinfo(tz_name)
    if tz is None:
        # unknown Windows zone and no tzdata -- treat the wall clock as UTC.
        # Symmetric across both snapshots; only an intermittent label flip on
        # this exact zone would slip through, and --dry-diff would surface it.
        return sdt.strftime("%Y-%m-%dT%H:%M:%SZ")
    return sdt.replace(tzinfo=tz).astimezone(_dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _natural_key(ev: dict, sdt, tz_name: str, all_day: bool) -> str:
    """Stable across connector calls: iCalUID (shared by a whole series) OR the
    raw id OR the subject, PLUS the start instant -- the instant disambiguates
    the individual occurrences of a recurring series (same iCalUID)."""
    uid = _norm_txt(lb._first(ev, "i_cal_u_id", "iCalUId", "i_cal_uid", " i_cal_uid", "uid", default=""))
    rid = _norm_txt(ev.get("id") or "")
    subj = _norm_txt(ev.get("subject") or ev.get("display_title") or "")
    base = uid or rid or ("subj:" + subj[:80] if subj else "")
    return base + "@" + _instant_utc(sdt, tz_name, all_day)


def _fingerprint_one(ev: dict) -> dict:
    sdt, s_tz = lb._graph_dt_parts(ev.get("start"))
    edt, e_tz = lb._graph_dt_parts(ev.get("end"))
    all_day = bool(lb._is_all_day(ev, sdt, edt))
    return {
        "subject": _norm_txt(ev.get("subject") or ev.get("display_title") or "")[:200],
        "start": _instant_utc(sdt, s_tz, all_day),
        "end": _instant_utc(edt, e_tz, all_day),
        "response_status": _norm_status(ev),
        "all_day": all_day,
        "type": _norm_txt(ev.get("type") or ""),  # stored for the trip payload; NOT compared
    }


def _run_window():
    """The 7-day snapshot window, computed ONCE per guard cycle so PRE and POST
    use byte-identical query params even if the cycle straddles UTC midnight."""
    today = _dt.date.today()
    ws = _dt.datetime(today.year, today.month, today.day, tzinfo=_dt.timezone.utc)
    return ws, ws + _dt.timedelta(days=7)


def take_snapshot(tag: str, window=None) -> tuple[dict | None, dict]:
    """Return (fingerprint_by_id | None, meta).
      dict  -> a verified snapshot (list_events fired).
      None  -> UNAVAILABLE: list_events did not fire across all retries, OR the
               codex run failed. NOT a HALT -- the caller skips Lane B calendar
               this cycle without disabling the task.
    Raises RuntimeError ONLY if the snapshot session's own re-contamination guard
    tripped (a write / non-allowlisted tool was seen) -- that IS a real HALT."""
    ws, we = window if window is not None else _run_window()
    prompt = _snapshot_prompt(ws.strftime("%Y-%m-%dT%H:%M:%SZ"), we.strftime("%Y-%m-%dT%H:%M:%SZ"))

    attempts: list[dict] = []
    for n in range(1, SNAP_RETRIES + 1):
        try:
            events, _raw = lb.run_codex_json(prompt, timeout_s=SNAPSHOT_TIMEOUT_S, tag=f"snap-{tag}#{n}")
        except RuntimeError as e:
            attempts.append({"n": n, "outcome": "codex_failed", "detail": str(e)[:160]})
            if n < SNAP_RETRIES:
                _time.sleep(lb.CALL1_RETRY_BACKOFF_S[min(n - 1, len(lb.CALL1_RETRY_BACKOFF_S) - 1)])
            continue
        tool_calls = lb.extract_tool_calls(events)
        status, detail = lb.guard_recontamination(tool_calls, "calendar")
        if status == "halt":
            raise RuntimeError(f"snapshot re-contamination guard tripped: {detail['unexpected']}")
        if status == "unavailable":
            attempts.append({"n": n, "outcome": "unavailable",
                             "tools": [f"{t['server']}::{t['tool']}" for t in tool_calls]})
            _log(f"snapshot {tag} attempt {n}/{SNAP_RETRIES}: list_events did not fire")
            if n < SNAP_RETRIES:
                _time.sleep(lb.CALL1_RETRY_BACKOFF_S[min(n - 1, len(lb.CALL1_RETRY_BACKOFF_S) - 1)])
            continue

        objs = lb._events_from_results(tool_calls, "calendar", events)
        fp: dict[str, dict] = {}
        for ev in objs:
            if not isinstance(ev, dict):
                continue
            sdt, s_tz = lb._graph_dt_parts(ev.get("start"))
            edt, _e_tz = lb._graph_dt_parts(ev.get("end"))
            all_day = bool(lb._is_all_day(ev, sdt, edt))
            key = _natural_key(ev, sdt, s_tz, all_day)
            if not key.strip("@"):
                continue
            fp[key] = _fingerprint_one(ev)
        attempts.append({"n": n, "outcome": "ok", "count": len(fp)})
        return fp, {"tag": tag, "count": len(fp), "attempts": attempts,
                    "tool_calls": [f"{t['server']}::{t['tool']}" for t in tool_calls]}

    return None, {"tag": tag, "count": 0, "attempts": attempts, "unavailable": True}


def diff_snapshots(pre: dict, post: dict) -> list[dict]:
    """A trip is a genuine calendar change: an event key present on only one
    side (real add/remove, or a reschedule that moved the start instant), or a
    matched key whose subject/start/end/response_status/all_day differs after
    normalisation. `type` and other re-rendered fields are NOT compared."""
    trips: list[dict] = []
    for key in sorted(pre.keys() - post.keys()):
        trips.append({"change": "removed", "key": key, "was": pre[key]})
    for key in sorted(post.keys() - pre.keys()):
        trips.append({"change": "added", "key": key, "now": post[key]})
    for key in sorted(pre.keys() & post.keys()):
        changed = {k: {"was": pre[key].get(k), "now": post[key].get(k)}
                   for k in _CMP_FIELDS
                   if pre[key].get(k) != post[key].get(k)}
        if changed:
            trips.append({"change": "modified", "key": key, "fields": changed})
    return trips


def _quarantine_normalised(ts: str, why: str) -> None:
    if NORMALISED.exists():
        dest = NORMALISED.with_suffix(f".halted_{ts}.json")
        try:
            NORMALISED.rename(dest)
            _log(f"quarantined {NORMALISED.name} -> {dest.name} ({why})")
        except OSError as e:
            _log(f"could not quarantine {NORMALISED.name} ({e}) -- overwriting with a HALT stub")
            NORMALISED.write_text(json.dumps(
                {"calendar": [], "teams": [], "transcripts": [],
                 "meta": {"ts": ts, "lane_b": {"halt": True, "why": why}}}, indent=2), encoding="utf-8")


def _write_trip(ts: str, payload: dict) -> Path:
    CODEX_RUNS_DIR.mkdir(parents=True, exist_ok=True)
    p = CODEX_RUNS_DIR / f"GUARD_TRIPPED_cal_{ts}.json"
    p.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
    return p


def cmd_run() -> int:
    """Exit: 0 clean (calendar verified unchanged; normalised file trustworthy).
             1 PERSISTENT HALT (a real calendar change during the read window,
               or a write-tool seen) -- wrapper Disable-ScheduledTask + toast.
             3 TRANSIENT: connector unavailable / codex failed / can't verify --
               no connector calendar this run, wrapper does NOT disable the task."""
    ts = _dt.datetime.now(_dt.timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    _log(f"--run start ts={ts}")
    _log(f"CODEX_HOME={lb._codex_home()}  account_id={lb._codex_account_id()}")
    CODEX_RUNS_DIR.mkdir(parents=True, exist_ok=True)
    window = _run_window()
    _log(f"read window (pinned for PRE+POST): {window[0].strftime('%Y-%m-%dT%H:%M:%SZ')} .. {window[1].strftime('%Y-%m-%dT%H:%M:%SZ')}")

    # 1. PRE snapshot
    try:
        pre_fp, pre_meta = take_snapshot("pre", window)
    except RuntimeError as e:      # re-contamination during the snapshot session == real HALT
        _log(f"PRE snapshot re-contamination HALT ({e})")
        _quarantine_normalised(ts, f"pre-snapshot re-contamination: {e}")
        _write_trip(ts, {"ts": ts, "phase": "pre", "halt": True, "error": str(e)})
        return 1
    if pre_fp is None:
        _log("PRE snapshot UNAVAILABLE (list_events never fired) -- skipping Lane B calendar "
             "this cycle; NOT disabling the task")
        _quarantine_normalised(ts, "pre-snapshot connector unavailable")
        _write_trip(ts, {"ts": ts, "phase": "pre", "transient": True, "meta": pre_meta})
        return 3
    (CODEX_RUNS_DIR / f"cal_baseline_{ts}.json").write_text(
        json.dumps({"ts": ts, "meta": pre_meta, "fp": pre_fp}, indent=2, ensure_ascii=False), encoding="utf-8")
    _log(f"PRE snapshot: {pre_meta['count']} event(s)")

    # 2. Call-1 (its own re-contamination HALT is inside)
    rc = subprocess.run(
        [sys.executable, str(REPO_ROOT / "lane_b_call1.py"), "--domain", "calendar"],
        cwd=str(REPO_ROOT),
        env={**lb.os.environ, "WI_LANE_B_SKIP_WARMUP": "1"},   # the PRE snapshot already warmed codex
    ).returncode
    _log(f"lane_b_call1.py --domain calendar exit {rc}")
    if rc == 1:
        _log("lane_b_call1 RE-CONTAMINATION guard TRIPPED -- persistent HALT")
        _quarantine_normalised(ts, "lane_b_call1 re-contamination guard tripped")
        _write_trip(ts, {"ts": ts, "phase": "call1", "halt": True, "call1_exit": rc})
        return 1
    if rc != 0:
        _log(f"lane_b_call1 exit {rc} (codex failed / all domains unavailable) -- "
             f"no connector calendar this run; NOT disabling the task")
        _quarantine_normalised(ts, f"lane_b_call1 exit {rc}")
        return 3
    if not NORMALISED.exists():
        _log("lane_b_call1 exit 0 but no lane_b_normalised.json (connector unavailable this cycle) -- "
             "no connector calendar this run; NOT disabling the task")
        return 3

    # 3. POST snapshot
    try:
        post_fp, post_meta = take_snapshot("post", window)
    except RuntimeError as e:
        _log(f"POST snapshot re-contamination HALT ({e})")
        _quarantine_normalised(ts, f"post-snapshot re-contamination: {e}")
        _write_trip(ts, {"ts": ts, "phase": "post", "halt": True, "error": str(e), "pre_meta": pre_meta})
        return 1
    if post_fp is None:
        _log("POST snapshot UNAVAILABLE -- cannot confirm the calendar is unchanged; "
             "quarantining this run's calendar (transient, task stays enabled)")
        _quarantine_normalised(ts, "post-snapshot connector unavailable -- unverified")
        _write_trip(ts, {"ts": ts, "phase": "post", "transient": True, "pre_meta": pre_meta})
        return 3
    _log(f"POST snapshot: {post_meta['count']} event(s)")

    # 4. diff -- ONLY a real change between two verified snapshots is a HALT
    trips = diff_snapshots(pre_fp, post_fp)
    if trips:
        _log(f"GUARD TRIPPED -- {len(trips)} calendar change(s) during the read window -- persistent HALT")
        p = _write_trip(ts, {"ts": ts, "phase": "diff", "halt": True, "trips": trips,
                             "pre_meta": pre_meta, "post_meta": post_meta})
        _quarantine_normalised(ts, f"{len(trips)} calendar change(s) during read window")
        _log(f"wrote {p}. WRAPPER must Disable-ScheduledTask + toast. No auto-resume.")
        return 1

    _log("clean -- calendar unchanged across the read window; lane_b_normalised.json is trustworthy")
    return 0


def cmd_snapshot(out: str) -> int:
    ts = _dt.datetime.now(_dt.timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    try:
        fp, meta = take_snapshot("adhoc")
    except RuntimeError as e:
        _log(f"snapshot re-contamination HALT: {e}")
        return 1
    if fp is None:
        _log(f"snapshot unavailable (list_events never fired): {meta.get('attempts')}")
        return 3
    Path(out).write_text(json.dumps({"ts": ts, "meta": meta, "fp": fp}, indent=2, ensure_ascii=False),
                         encoding="utf-8")
    _log(f"wrote {out} ({meta['count']} events)")
    return 0


def cmd_diff(pre: str, post: str) -> int:
    a = json.loads(Path(pre).read_text(encoding="utf-8")).get("fp", {})
    b = json.loads(Path(post).read_text(encoding="utf-8")).get("fp", {})
    trips = diff_snapshots(a, b)
    print(json.dumps({"trips": trips, "tripped": bool(trips)}, indent=2, ensure_ascii=False))
    return 1 if trips else 0


def cmd_dry_diff() -> int:
    """Two snapshots with NO Call-1 between them; assert zero diff. The cheap
    validator for the normalisation before any full guard cutover retry.
    exit 0 = stable, 1 = residual diff (normalisation still imperfect),
    3 = connector unavailable.
    NOTE (2 Sept 2026): PRE and POST are no longer truly "back-to-back" --
    lane_b_call1.py's run_codex_json() now enforces a minimum quiet gap
    (WI_LANE_B_SNAPSHOT_GAP_S, default 75s) since the last connector touch
    before every codex exec call, including this pair. Evidence: a clean
    manual two-separate-invocations test (natural human-typing gap) vs a
    hard hang on every attempt of the previous zero-gap automated version --
    see HANDOVER.md for the full writeup."""
    ts = _dt.datetime.now(_dt.timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    window = _run_window()
    _log(f"--dry-diff: PRE + POST back-to-back, NO Call-1 (window "
         f"{window[0].strftime('%Y-%m-%dT%H:%M:%SZ')} .. {window[1].strftime('%Y-%m-%dT%H:%M:%SZ')})")
    try:
        a_fp, a_meta = take_snapshot("dry-pre", window)
        b_fp, b_meta = take_snapshot("dry-post", window)
    except RuntimeError as e:
        _log(f"re-contamination guard tripped during --dry-diff: {e}")
        return 1
    if a_fp is None or b_fp is None:
        _log(f"connector unavailable (pre={a_meta.get('count')}, post={b_meta.get('count')}) -- cannot validate")
        return 3
    trips = diff_snapshots(a_fp, b_fp)
    CODEX_RUNS_DIR.mkdir(parents=True, exist_ok=True)
    out = CODEX_RUNS_DIR / f"DRY_DIFF_cal_{ts}.json"
    out.write_text(json.dumps({"ts": ts, "pre_count": a_meta["count"], "post_count": b_meta["count"],
                               "trips": trips}, indent=2, ensure_ascii=False), encoding="utf-8")
    if trips:
        _log(f"--dry-diff NOT CLEAN: {len(trips)} residual diff(s) between two unchanged reads -> {out.name}")
        print(json.dumps(trips[:20], indent=2, ensure_ascii=False))
        return 1
    _log(f"--dry-diff CLEAN: {a_meta['count']} vs {b_meta['count']} events, 0 diffs. "
         f"Normalisation is stable ({out.name}).")
    return 0


def cmd_selftest() -> int:
    """Pure-function checks, no codex. Feeds synthetic connector event dicts
    that mimic the observed re-rendering and asserts benign churn -> 0 trips
    while genuine change -> a trip."""
    fails = []

    def _fp(objs):
        d = {}
        for ev in objs:
            sdt, s_tz = lb._graph_dt_parts(ev.get("start"))
            ad = bool(lb._is_all_day(ev, sdt, lb._graph_dt_parts(ev.get("end"))[0]))
            d[_natural_key(ev, sdt, s_tz, ad)] = _fingerprint_one(ev)
        return d

    def check(name, cond):
        print(("  ok   " if cond else "  FAIL ") + name)
        if not cond:
            fails.append(name)

    # occurrence: same iCalUID, id churns between calls, tz rendered UTC then +01:00 (same instant)
    pre = [{"id": "AAA111", "i_cal_u_id": "UID-weekly-1to1", "subject": "Weekly 1:1",
            "start": {"dateTime": "2026-09-02T08:00:00", "timeZone": "UTC"},
            "end": {"dateTime": "2026-09-02T08:30:00", "timeZone": "UTC"},
            "response_status": {"response": "Accepted"}, "type": "occurrence"}]
    post = [{"id": "BBB222-DIFFERENT", "i_cal_u_id": "UID-weekly-1to1", "subject": "Weekly  1:1 ",
             "start": {"dateTime": "2026-09-02T09:00:00+01:00", "timeZone": "GMT Standard Time"},
             "end": {"dateTime": "2026-09-02T09:30:00+01:00", "timeZone": "GMT Standard Time"},
             "response_status": {"response": "accepted"}, "type": "singleInstance"}]
    check("benign churn (id + tz-label + case + whitespace) -> 0 trips",
          diff_snapshots(_fp(pre), _fp(post)) == [])

    # two occurrences of one series (same UID, different instants) must both survive
    two = [dict(pre[0]), dict(pre[0], start={"dateTime": "2026-09-09T08:00:00", "timeZone": "UTC"},
                end={"dateTime": "2026-09-09T08:30:00", "timeZone": "UTC"})]
    check("two occurrences of one series -> 2 distinct keys", len(_fp(two)) == 2)

    # genuine reschedule
    resched = [dict(pre[0], start={"dateTime": "2026-09-02T15:00:00", "timeZone": "UTC"},
               end={"dateTime": "2026-09-02T15:30:00", "timeZone": "UTC"})]
    check("genuine reschedule -> a trip", len(diff_snapshots(_fp(pre), _fp(resched))) >= 1)

    # genuine subject change on the same instant
    ren = [dict(pre[0], subject="Weekly 1:1 -- CANCELLED cover")]
    tr = diff_snapshots(_fp(pre), _fp(ren))
    check("genuine subject change -> modified trip",
          len(tr) == 1 and tr[0]["change"] == "modified" and "subject" in tr[0]["fields"])

    # genuine add / remove
    check("genuine add -> a trip", len(diff_snapshots(_fp(pre), _fp(pre + ren))) == 1)
    check("genuine remove -> a trip", len(diff_snapshots(_fp(pre + ren), _fp(pre))) == 1)

    # all-day event, date-only both sides
    ad_pre = [{"id": "D1", "i_cal_u_id": "UID-leave", "subject": "A/L", "is_all_day": True,
               "start": {"dateTime": "2026-09-03T00:00:00", "timeZone": "UTC"},
               "end": {"dateTime": "2026-09-04T00:00:00", "timeZone": "UTC"}}]
    ad_post = [dict(ad_pre[0], id="D1-CHURN")]
    check("all-day event, id churn -> 0 trips", diff_snapshots(_fp(ad_pre), _fp(ad_post)) == [])

    print("")
    if fails:
        print("RESULT: %d FAILED" % len(fails))
        return 1
    print("RESULT: all passed")
    return 0


def main(argv: list[str]) -> int:
    ap = argparse.ArgumentParser(description="Lane B calendar HALT kill-switch")
    ap.add_argument("--run", action="store_true", help="pre-snap -> lane_b_call1 calendar -> post-snap -> diff (the wrapper entry point)")
    ap.add_argument("--snapshot", action="store_true")
    ap.add_argument("--out", default=str(CODEX_RUNS_DIR / "cal_snapshot.json"))
    ap.add_argument("--diff", action="store_true")
    ap.add_argument("--dry-diff", dest="dry_diff", action="store_true",
                    help="PRE + POST back-to-back, NO Call-1 -- assert the normalised snapshots are stable")
    ap.add_argument("--selftest", action="store_true", help="pure-function checks, no codex")
    ap.add_argument("--pre")
    ap.add_argument("--post")
    args = ap.parse_args(argv)

    if args.selftest:
        return cmd_selftest()
    if args.dry_diff:
        return cmd_dry_diff()
    if args.run:
        return cmd_run()
    if args.snapshot:
        return cmd_snapshot(args.out)
    if args.diff:
        if not (args.pre and args.post):
            ap.error("--diff needs --pre and --post")
        return cmd_diff(args.pre, args.post)
    ap.error("one of --run / --dry-diff / --selftest / --snapshot / --diff required")


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
