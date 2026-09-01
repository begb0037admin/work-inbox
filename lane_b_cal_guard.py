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

SNAPSHOT_TIMEOUT_S = int(lb.os.environ.get("WI_LANE_B_SNAP_TIMEOUT", "180"))


def _log(m: str) -> None:
    print(f"[{_dt.datetime.now(_dt.timezone.utc).strftime('%Y-%m-%d %H:%M:%SZ')}] lane_b_cal_guard: {m}")


def _snapshot_prompt(win_start_iso: str, win_end_iso: str) -> str:
    return (
        "Using the Microsoft Outlook Calendar app connector, retrieve my calendar events "
        f"between {win_start_iso} and {win_end_iso} from my default calendar and, if it "
        f"exists, the calendar named \"{lb.SHARED_CAL_NAME}\". "
        "Return ONLY the raw connector result as a JSON array of the event objects, with no "
        "summary and no prose. Do not use any other app or tool. Do not create, update, "
        "cancel, delete, move, respond to, or add an attachment to any event. Send nothing."
    )


SNAP_RETRIES = max(1, int(lb.os.environ.get("WI_LANE_B_RETRIES", "3")))


def take_snapshot(tag: str) -> tuple[dict | None, dict]:
    """Return (fingerprint_by_id | None, meta).
      dict  -> a verified snapshot (list_events fired).
      None  -> UNAVAILABLE: list_events did not fire across all retries, OR the
               codex run failed. NOT a HALT -- the caller skips Lane B calendar
               this cycle without disabling the task.
    Raises RuntimeError ONLY if the snapshot session's own re-contamination guard
    tripped (a write / non-allowlisted tool was seen) -- that IS a real HALT."""
    today = _dt.date.today()
    ws = _dt.datetime(today.year, today.month, today.day, tzinfo=_dt.timezone.utc)
    we = ws + _dt.timedelta(days=7)
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
            sdt, _ = lb._graph_dt_parts(ev.get("start"))
            edt, _ = lb._graph_dt_parts(ev.get("end"))
            rs = ev.get("response_status") or {}
            eid = str(ev.get("id") or ev.get("i_cal_u_id") or "")
            if not eid:
                continue
            fp[eid] = {
                "subject": (ev.get("subject") or ev.get("display_title") or "")[:200],
                "start": sdt.isoformat() if sdt else "",
                "end": edt.isoformat() if edt else "",
                "response_status": (rs.get("response") if isinstance(rs, dict) else str(rs or "")),
                "type": ev.get("type") or "",
            }
        attempts.append({"n": n, "outcome": "ok", "count": len(fp)})
        return fp, {"tag": tag, "count": len(fp), "attempts": attempts,
                    "tool_calls": [f"{t['server']}::{t['tool']}" for t in tool_calls]}

    return None, {"tag": tag, "count": 0, "attempts": attempts, "unavailable": True}


def diff_snapshots(pre: dict, post: dict) -> list[dict]:
    trips: list[dict] = []
    for eid in pre.keys() - post.keys():
        trips.append({"change": "removed", "id": eid, "was": pre[eid]})
    for eid in post.keys() - pre.keys():
        trips.append({"change": "added", "id": eid, "now": post[eid]})
    for eid in pre.keys() & post.keys():
        if pre[eid] != post[eid]:
            changed = {k: {"was": pre[eid].get(k), "now": post[eid].get(k)}
                       for k in set(pre[eid]) | set(post[eid])
                       if pre[eid].get(k) != post[eid].get(k)}
            trips.append({"change": "modified", "id": eid, "fields": changed})
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
    CODEX_RUNS_DIR.mkdir(parents=True, exist_ok=True)

    # 1. PRE snapshot
    try:
        pre_fp, pre_meta = take_snapshot("pre")
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
        post_fp, post_meta = take_snapshot("post")
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


def main(argv: list[str]) -> int:
    ap = argparse.ArgumentParser(description="Lane B calendar HALT kill-switch")
    ap.add_argument("--run", action="store_true", help="pre-snap -> lane_b_call1 calendar -> post-snap -> diff (the wrapper entry point)")
    ap.add_argument("--snapshot", action="store_true")
    ap.add_argument("--out", default=str(CODEX_RUNS_DIR / "cal_snapshot.json"))
    ap.add_argument("--diff", action="store_true")
    ap.add_argument("--pre")
    ap.add_argument("--post")
    args = ap.parse_args(argv)

    if args.run:
        return cmd_run()
    if args.snapshot:
        return cmd_snapshot(args.out)
    if args.diff:
        if not (args.pre and args.post):
            ap.error("--diff needs --pre and --post")
        return cmd_diff(args.pre, args.post)
    ap.error("one of --run / --snapshot / --diff required")


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
