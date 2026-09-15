#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
lane_b_cal_guard.py -- Lane B calendar HALT kill-switch (LANE_B_TEAMS_CAL_DESIGN.md sec.6a)
=========================================================================================

REDESIGNED 2 Sept 2026 (Kevin's decision, see HANDOVER.md): this file used to
run TWO HALT layers -- a PRE/POST `list_events` snapshot diff around Call-1
(this file), plus lane_b_call1.py's own re-contamination guard (HALT if a
write / non-allowlisted / non-codex_apps tool is seen). The snapshot-diff
layer required two independent connector reads of the same window to agree
with each other -- which is what actually produced the 1 Sept
52-false-positive bug, the still-undiagnosed all-day-event gap, and a
"second connector call hangs" 360s-timeout pattern that blocked every
`--dry-diff` attempt the week of 1-2 Sept. None of those three incidents was
ever a real write; every one was a symptom of the diff mechanism itself.

**`--run` (the live gate the wrapper calls) now relies SOLELY on
lane_b_call1.py's re-contamination guard** -- it inspects the ACTUAL tool
calls made during the real fetch (including partial output from a
killed/timed-out attempt, since 2 Sept -- see _scan_partial_output_for_writes
in lane_b_call1.py), rather than inferring a write from a before/after
mismatch. It already checks every retry attempt, isn't timing-sensitive, and
needs only one connector read, not two agreeing ones.

The snapshot/diff machinery below (`take_snapshot`, `diff_snapshots`,
`--snapshot`, `--diff`, `--dry-diff`, `--selftest`) is UNCHANGED and STAYS --
it is a useful standalone diagnostic for spot-checking connector-read
normalisation -- but as of 2 Sept it is no longer part of what a live
scheduled run goes through.

Single entry point for the wrapper:

  python lane_b_cal_guard.py --run                            # lane_b_call1 --domain calendar (default); its own re-contamination guard is the gate
  python lane_b_cal_guard.py --run --domain teams              # same gate, Teams-only fetch
  python lane_b_cal_guard.py --run --domain both                # same gate, one call covering both domains

  exit 0  = clean. data/lane_b/lane_b_normalised.json is fresh and trustworthy.
  exit 1  = GUARD TRIPPED -- lane_b_call1's re-contamination guard actually
            observed a write/off-allowlist tool call. The freshly written
            lane_b_normalised.json is QUARANTINED (renamed .halted_<ts>) so
            fetch_inbox.py falls back to "calendar empty".
            data/codex_runs/GUARD_TRIPPED_cal_<ts>.json holds the detail.
            The WRAPPER is responsible for `Disable-ScheduledTask` + a toast.
  exit 2  = usage / environment error.
  exit 3  = the Call-1 codex run failed / connector unavailable this cycle
            (can't verify -> treated as unsafe: no calendar this run, task
            stays enabled, retried next cadence).
  exit 5  = MODEL POLICY VIOLATION -- deterministic model/effort code/config
            bug, not connector flakiness and not a mailbox-safety HALT. The
            wrapper should keep the task enabled but surface the violation for
            investigation rather than retrying it as transient.

Diagnostic-only, NOT part of the live gate (see note above):
  python lane_b_cal_guard.py --snapshot --out data/codex_runs/cal_baseline_<ts>.json
  python lane_b_cal_guard.py --diff  --pre <baseline.json> --post <after.json>
  python lane_b_cal_guard.py --dry-diff
  python lane_b_cal_guard.py --selftest

NOTE (1 Sept 2026): the codex_apps calendar event object has NO
lastModifiedDateTime / last_modified field (confirmed from the real probe), so
the per-event fingerprint diffs on {subject, start, end, response_status, type}
keyed by id. That still catches an add, a drop, a reschedule, an RSVP change,
and a single->cancelled-occurrence flip. Any timestamp-only edit that changed
none of those is not detectable via this surface -- documented residual, and
academic now that this mechanism is diagnostic-only rather than the live gate.
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

# GUARD_RUN_TIMEOUT_S -- added 14 Sep 2026, Bridge Briefing hang investigation
# follow-up (see HANDOVER.md's EDU_PARKED entries for the incident this
# closes). Before this fix, cmd_run()'s subprocess.run() launch of
# lane_b_call1.py carried NO timeout at all -- the only thing that could ever
# stop a hung child was the wrapper .ps1's own scheduled-task
# ExecutionTimeLimit (PT45M), which (a) only fires after the FULL 45-minute
# budget is spent, starving every phase after this one (mail domain,
# fetch_inbox.py's own triage, the dashboard push), and (b) does not
# reliably reach a multi-generation process tree either -- documented live:
# orphaned grandchild python/codex/node processes survived a Task Scheduler
# kill on 14 Sep and needed manual cleanup. This mirrors the exact fix
# already proven one level down, inside lane_b_call1.py's own
# run_codex_json(): a real process-level enforced timeout via
# Popen+wait(timeout=...), and on TimeoutExpired a `taskkill /T /F` against
# the WHOLE process tree (not just the direct child) so an orphaned
# grandchild codex.exe/node.exe can't keep running underneath a defeated
# parent.
#
# COMPUTED, not a hand-guessed constant (a first pass hardcoded 2400s, which
# a Codex review same day correctly flagged as BELOW lane_b_call1.py's own
# documented worst case for `--domain both` -- ~2700-2750s once warm-up +
# EDU_PARKED_RETRIES outer attempts + 2 inner sub-attempts at
# CALL1_TIMEOUT_S + KILL_COOLDOWN_S gaps are counted for both domains. A
# timeout tighter than the thing it's meant to bound would fire on a
# legitimate slow run, not just a genuine hang -- exactly the class of bug
# this whole investigation exists to close, just moved one level up).
# _guard_run_timeout_s() below computes this from lane_b_call1.py's OWN live
# constants so it can never silently go stale if those are retuned later,
# and honestly logs the case where the real worst case would exceed what
# can safely fit under the wrapper's own PT45M (2700s) outer
# ExecutionTimeLimit rather than hiding it. WI_LANE_B_GUARD_RUN_TIMEOUT_S
# still wins outright over the computed value if set, for manual tuning.
_GUARD_HARD_CAP_S = 2500  # stay safely under the wrapper's PT45M (2700s) so
                          # OUR clean tree-kill always wins the race against
                          # Task Scheduler's own less-reliable one.


def _guard_run_timeout_s(domain: str) -> int:
    """Realistic worst-case budget for cmd_run()'s lane_b_call1.py child,
    computed from lane_b_call1.py's own live timeout/retry constants (not a
    hand-guessed number -- see the comment block above this function).

    Per domain, worst case = the parked-mode retry budget for that domain, each
    outer attempt up to 2 inner sub-attempts at CALL1_TIMEOUT_S, separated by a KILL_COOLDOWN_S
    gap (the worst-case gap -- assumes the previous sub-attempt was itself
    killed). Warm-up (CALL1_WARMUP_TIMEOUT_S) is counted ONCE, not once per
    domain, since lane_b_call1.py's own `_WARMED_HOMES` caches it per
    CODEX_HOME for the life of the process -- `--domain both` still only
    warms up once. `--domain both` also pays one extra SNAPSHOT_GAP_S
    between the two domains' calls. +10% on top as a rounding/margin buffer.

    If that computed worst case is ABOVE _GUARD_HARD_CAP_S, this is capped
    there instead and a loud WARNING is logged -- a disclosed tradeoff
    (reliability of a clean kill over giving every possible benefit of the
    doubt to an extreme-worst-case-but-genuinely-legitimate run), not a
    silent bug. If this cap fires often in practice, the real fix is
    revisiting the wrapper's PT45M and/or the retry constants this formula
    is built from, not just raising this number further -- flagged, not
    solved here."""
    override = lb.os.environ.get("WI_LANE_B_GUARD_RUN_TIMEOUT_S", "").strip()
    if override:
        return int(override)
    domains = 2 if domain == "both" else 1
    one_attempt = 2 * lb.CALL1_TIMEOUT_S + lb.KILL_COOLDOWN_S
    if domain == "both":
        # Calendar deliberately has one extra parked-mode outer retry: a fresh
        # 15 Sep personal-only run exhausted both inner attempts, while the
        # scheduled task recovered on its retry path. Teams keeps the original
        # parked budget so a calendar recovery does not silently double every
        # domain's wall-clock allowance.
        per_domain = (lb.EDU_PARKED_CALENDAR_RETRIES * one_attempt
                      + lb.EDU_PARKED_RETRIES * one_attempt)
    else:
        retries = (lb.EDU_PARKED_CALENDAR_RETRIES if domain == "calendar"
                   else lb.EDU_PARKED_RETRIES)
        per_domain = retries * one_attempt
    total = lb.CALL1_WARMUP_TIMEOUT_S + per_domain
    if domains == 2:
        total += lb.SNAPSHOT_GAP_S  # the extra between-domain gap `both` pays that a single domain doesn't
    computed = int(total * 1.10)
    if computed > _GUARD_HARD_CAP_S:
        _log(f"WARNING: computed guard timeout ({computed}s, domain={domain!r}) exceeds the "
             f"{_GUARD_HARD_CAP_S}s hard cap kept under the wrapper's PT45M (2700s) outer "
             f"ExecutionTimeLimit -- capping to {_GUARD_HARD_CAP_S}s so this guard's own clean "
             f"tree-kill always fires before Task Scheduler's own less-reliable one. An extreme-"
             f"worst-case but genuinely legitimate (not hung) run could occasionally be treated as "
             f"transient/retried by this cap -- a disclosed tradeoff, not a silent bug. If this "
             f"fires often, revisit PT45M and/or the retry constants this formula is built from.")
        return _GUARD_HARD_CAP_S
    return computed


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
        except lb.codex_model_policy.ModelPolicyViolation:
            # This is a deterministic model/effort code/config violation, not
            # a transient connector failure. Do not consume SNAP_RETRIES or
            # flatten it into the diagnostic "unavailable" result.
            raise
        except lb.ReContaminationDetected as e:
            # A write was actually observed, even in partial/timed-out output -- never
            # retry this, propagate as the real HALT it is (same as a full-attempt catch below).
            raise RuntimeError(str(e)) from e
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


def cmd_run(domain: str = "calendar") -> int:
    """Exit: 0 clean (Call-1's re-contamination guard saw nothing unexpected;
               normalised file trustworthy).
             1 PERSISTENT HALT (Call-1's re-contamination guard actually
               observed a write/off-allowlist tool call) -- wrapper
               Disable-ScheduledTask + toast.
             3 TRANSIENT: connector unavailable / codex failed / can't verify --
               no connector calendar this run, wrapper does NOT disable the task.

    `domain` -- 'calendar' (default), 'teams', or 'both'. Added 2 Sept 2026 for
    the Lane B Teams wrapper wiring: fetch_inbox.py's own comment on
    TEAMS_BACKEND already documents that this guard (via lane_b_call1's
    re-contamination check, which covers the microsoft_teams.* tool namespace
    same as calendar's) is the SOLE safety mechanism for Teams too -- no
    separate Teams guard exists or is planned. This param is what actually lets
    a live run ask lane_b_call1.py to fetch (and guard) Teams data instead of
    just defaulting to calendar-only every time.

    REDESIGNED 2 Sept 2026 (Kevin's decision): the PRE-snapshot -> POST-snapshot
    diff that used to bookend Call-1 here has been REMOVED from this live gate.
    It required two independent connector reads of the same window to agree
    with each other -- which is what actually produced the 1 Sept
    52-false-positive bug, the still-undiagnosed all-day-event gap, and the
    "second call hangs" 360s-timeout pattern that blocked every --dry-diff
    attempt this week. None of those three incidents were ever a real write --
    every one was a symptom of the diff mechanism itself, not evidence it ever
    caught anything the guard below wouldn't have.

    Call-1's re-contamination guard (lane_b_call1.py: guard_recontamination(),
    now backed by _scan_partial_output_for_writes() so a write logged just
    before a timeout/kill is no longer invisible either) inspects the ACTUAL
    tool calls made during the real fetch -- it doesn't infer a write from a
    before/after mismatch, it sees the write tool get called, directly. It
    already checks every retry attempt, not just the last one, isn't
    timing-sensitive, and needs only ONE connector read, not two agreeing
    ones. It is now the SOLE live safety mechanism for Lane B calendar/Teams.

    take_snapshot() / diff_snapshots() / --dry-diff / --snapshot / --diff /
    --selftest below are UNCHANGED and remain available as standalone
    diagnostic commands (useful for spot-checking normalisation) -- they are
    simply no longer part of what a live scheduled run goes through."""
    ts = _dt.datetime.now(_dt.timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    _log(f"--run start ts={ts} domain={domain}")
    _log(f"CODEX_HOME={lb._codex_home()}  account_id={lb._codex_account_id()}")
    CODEX_RUNS_DIR.mkdir(parents=True, exist_ok=True)

    # `-u` (added 14 Sep 2026, same investigation): the PREVIOUS bare
    # `sys.executable` launch (no `-u`) meant this child's own stdout was
    # block-buffered whenever it wasn't a real terminal -- true here, since
    # this process's own stdout is itself already piped through the wrapper
    # .ps1's `Tee-Object` (python -u lane_b_cal_guard.py ... | Tee-Object).
    # Buffering is decided per-process by each interpreter, not inherited
    # from an already-piped fd, so lane_b_call1.py's own real-time `_log()`
    # lines could sit invisible in memory for the full length of a run and
    # only appear (or never appear, if killed) at exit -- this is the exact
    # "log froze ... zero further log output" symptom the 10 Sep and 14 Sep
    # incidents both recorded, indistinguishable from a genuinely stuck run
    # until now. `-u` forces this child unbuffered, matching every other
    # `python -u ...` invocation the wrapper .ps1 already uses directly.
    #
    # Popen + a real timed wait (see _guard_run_timeout_s() above), not a
    # bare subprocess.run() with no timeout at all: on TimeoutExpired, kill
    # the WHOLE process tree by PID (`taskkill /T /F`) so an orphaned
    # grandchild can't keep the real call running underneath a defeated/
    # reaped parent -- same fix already proven in run_codex_json() and
    # _ensure_warm(). taskkill's own RETURN CODE is checked, not just
    # exceptions from launching it (Codex review finding, 14 Sep 2026: a
    # non-zero taskkill result -- e.g. "process not found", which can mean
    # genuinely already gone OR that the kill itself failed -- must not be
    # silently treated as success; on a non-zero result this also falls back
    # to proc.kill() as a second, best-effort attempt).
    timeout_s = _guard_run_timeout_s(domain)
    _log(f"running: {sys.executable} -u lane_b_call1.py --domain {domain} (enforced timeout {timeout_s}s)")
    proc = subprocess.Popen(
        [sys.executable, "-u", str(REPO_ROOT / "lane_b_call1.py"), "--domain", domain],
        cwd=str(REPO_ROOT),
        env={**lb.os.environ},
    )
    try:
        rc = proc.wait(timeout=timeout_s)
    except subprocess.TimeoutExpired:
        _log(f"TIMEOUT {timeout_s}s hit waiting for lane_b_call1.py (PID {proc.pid}) -- "
             f"killing its whole process tree so an orphaned grandchild codex/node process can't "
             f"keep running underneath a defeated timeout")
        try:
            kill_result = subprocess.run(["taskkill", "/T", "/F", "/PID", str(proc.pid)],
                                         capture_output=True, timeout=15)
            if kill_result.returncode != 0:
                raise RuntimeError(
                    f"taskkill exited {kill_result.returncode}: "
                    f"{(kill_result.stdout or b'').decode('utf-8', 'replace').strip()!r} / "
                    f"{(kill_result.stderr or b'').decode('utf-8', 'replace').strip()!r}"
                )
        except Exception as kill_exc:  # noqa: BLE001 -- non-Windows / taskkill missing / non-zero result / already exited
            _log(f"WARNING: taskkill did not confirm success ({kill_exc}) -- falling back to proc.kill() "
                 f"(direct child only; a grandchild codex/node process may still be running, check manually)")
            try:
                proc.kill()
            except Exception:  # noqa: BLE001
                pass
        try:
            proc.wait(timeout=15)
        except subprocess.TimeoutExpired:
            _log("WARNING: process did not exit within 15s of the taskkill -- check manually")
        # Honest note on what a timeout here does and doesn't tell us (Codex
        # review finding, 14 Sep 2026): this maps to TRANSIENT (exit 3), not
        # a guard HALT (1) -- but that is NOT because a kill here proves no
        # write happened. lane_b_call1.py's own re-contamination guard scans
        # partial output as it's received DURING a call it manages itself
        # (see _scan_partial_output_for_writes() in that file); if THIS
        # outer timeout instead fires while lane_b_call1.py is still mid-
        # call, whatever it had buffered but not yet scanned is lost with
        # it, and this guard cannot positively rule out a write in that
        # window. This exposure already existed before this fix, via the
        # wrapper's own PT45M outer kill (the only backstop previously) --
        # this change makes the HANG itself get caught reliably and cleanly
        # (no orphans), it does not add a new safety gap or close this
        # residual one. Mapping to exit 3 (not disabling the scheduled task
        # on every ordinary slow-connector timeout) is the same tradeoff
        # this file's own design has always made; if evidence ever shows a
        # write slipping through specifically via a killed run, that needs
        # its own fix (e.g. capturing and scanning partial output here too),
        # not just a comment update.
        _log(f"lane_b_call1.py --domain {domain} TIMED OUT after {timeout_s}s and was killed -- "
             f"treating as TRANSIENT (exit 3), not a guard HALT.")
        _quarantine_normalised(ts, f"lane_b_call1 timed out after {timeout_s}s (killed by guard)")
        return 3
    _log(f"lane_b_call1.py --domain {domain} exit {rc}")
    if rc == 1:
        _log("lane_b_call1 RE-CONTAMINATION guard TRIPPED -- persistent HALT")
        _quarantine_normalised(ts, "lane_b_call1 re-contamination guard tripped")
        _write_trip(ts, {"ts": ts, "phase": "call1", "halt": True, "call1_exit": rc, "domain": domain})
        return 1
    if rc == 5:
        # Added 10 Sep 2026, touchpoint-3 Codex review finding: MUST be
        # distinguished from the generic `rc != 0` branch below, which maps
        # everything else (including the pre-existing usage/environment-error
        # exit 2) to this function's own exit 3 ("codex failed / all domains
        # unavailable... NOT disabling the task, retry next cadence"). A
        # ModelPolicyViolation (lane_b_call1.py exit 5) is a deterministic
        # code/config bug, not connector flakiness -- folding it into "3"
        # here would mean the caller (Run Laptop Bridge Briefing.ps1) also
        # cannot tell it apart, and would silently retry it on the normal
        # cadence indefinitely. Propagated as its own distinct 5, not folded.
        _log("lane_b_call1 MODEL POLICY VIOLATION (exit 5) -- code/config bug, not connector "
             "unavailability; not disabling the task (not a security HALT), but this needs "
             "investigation, not just a retry")
        _quarantine_normalised(ts, f"lane_b_call1 MODEL POLICY VIOLATION (exit {rc})")
        return 5
    if rc != 0:
        _log(f"lane_b_call1 exit {rc} (codex failed / all domains unavailable) -- "
             f"no connector data this run; NOT disabling the task")
        _quarantine_normalised(ts, f"lane_b_call1 exit {rc}")
        return 3
    if not NORMALISED.exists():
        _log("lane_b_call1 exit 0 but no lane_b_normalised.json (connector unavailable this cycle) -- "
             "no connector data this run; NOT disabling the task")
        return 3

    _log("clean -- lane_b_call1's re-contamination guard saw nothing unexpected; "
         "lane_b_normalised.json is trustworthy")
    return 0


def cmd_snapshot(out: str) -> int:
    ts = _dt.datetime.now(_dt.timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    try:
        fp, meta = take_snapshot("adhoc")
    except lb.codex_model_policy.ModelPolicyViolation as e:
        _log(f"snapshot MODEL POLICY VIOLATION -- {e}")
        return 5
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
    3 = connector unavailable, 5 = model-policy code/config violation.
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
    except lb.codex_model_policy.ModelPolicyViolation as e:
        _log(f"--dry-diff MODEL POLICY VIOLATION -- {e}")
        return 5
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
    ap.add_argument("--run", action="store_true", help="lane_b_call1 --domain <--domain> ; its own re-contamination guard is the gate (the wrapper entry point)")
    ap.add_argument("--domain", choices=["calendar", "teams", "both"], default="calendar",
                    help="which Lane B domain(s) to fetch+guard on --run (default calendar, back-compat)")
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
        return cmd_run(args.domain)
    if args.snapshot:
        return cmd_snapshot(args.out)
    if args.diff:
        if not (args.pre and args.post):
            ap.error("--diff needs --pre and --post")
        return cmd_diff(args.pre, args.post)
    ap.error("one of --run / --dry-diff / --selftest / --snapshot / --diff required")


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
