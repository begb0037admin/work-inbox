#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
lane_b_call1.py -- Lane B "Call 1": the codex_apps connector dumb-fetch
=====================================================================

Spec: docs/LANE_B_TEAMS_CAL_DESIGN.md (sec.3 prompts, sec.5 allowlists, sec.6c
re-contamination guard, sec.7 per-run log). Corrected 1 Sept 2026 after the
"MAKE-OR-BREAK" false negatives:

  * The connector surface is REAL and works headless on the Edu account
    (`begb0037@ox.ac.uk`). MCP server name = `codex_apps`. Tools are
    namespaced `microsoft_outlook_calendar.*` / `microsoft_teams.*`.
  * `codex_apps` tools are LAZILY surfaced -- a "list every tool" enumeration
    never shows them. So Call 1 asks the model to CALL the named tools and we
    parse the `mcp_tool_call` events out of `codex exec --json` JSONL.
  * The re-contamination guard asserts on the tool calls ACTUALLY OBSERVED in
    the JSONL, not on a manifest turn.

What this does NOT do: summarise, decide, or branch on content. It fetches,
guards, sanitises (via normalise_pull), and writes files. Triage stays
`claude -p`, with no connector, downstream.

Outputs (under data/lane_b/ and data/codex_runs/):
  data/lane_b/lane_b_normalised.json     <- consumed by fetch_inbox.py CAL_BACKEND=connector
  data/lane_b/<ts>_call1_<domain>.jsonl  <- raw codex --json transcript
  data/lane_b/<ts>_lane_b.json           <- per-run log (LANE_B sec.7)
  data/codex_runs/<ts>_sanitiser_hits.json

Usage:
  python lane_b_call1.py --domain calendar
  python lane_b_call1.py --domain teams
  python lane_b_call1.py --domain both
  python lane_b_call1.py --domain calendar --dry-run          # print the prompt, run nothing
  python lane_b_call1.py --domain calendar --from-file probe.jsonl   # parse a captured transcript, no codex run

Exit codes: 0 ok (incl. "connector unavailable this cycle" -> empty + warning);
            1 GUARD TRIPPED (re-contamination / unexpected tool) -- caller should HALT;
            2 usage / environment error;
            3 codex run failed (timeout / non-zero / no parseable output).
"""

from __future__ import annotations

import argparse
import datetime as _dt
import hashlib
import json
import os
import re
import shutil as _shutil
import socket
import subprocess
import sys
import time
from pathlib import Path

try:
    import normalise_pull
except Exception as _e:  # pragma: no cover
    print(f"lane_b_call1: cannot import normalise_pull ({_e})", file=sys.stderr)
    sys.exit(2)

# --------------------------------------------------------------------------- #
REPO_ROOT      = Path(__file__).resolve().parent
LANE_B_DIR     = REPO_ROOT / "data" / "lane_b"
CODEX_RUNS_DIR = REPO_ROOT / "data" / "codex_runs"
NORMALISED_OUT = LANE_B_DIR / "lane_b_normalised.json"

# Recorded ~/.codex/config.toml sha1 baselines, per host (short lowercase
# hostname). WARNING-ONLY -- a mismatch is logged, never a HALT. Set
# WI_CODEX_CONFIG_SHA1 to override for a deliberate re-baseline.
_HOST = socket.gethostname().split(".")[0].lower()
CONFIG_TOML_SHA1_BASELINES = {
    "101l-de013193":   "ba0184e864ffd081069820cc7a6f8f19acf5c845",  # AD-OAK\begb0037, Edu, codex-cli 0.151.0 (1 Sept 2026, first real run)
    "desktop-mjdjm64": "4fd8ef763bf0a8ddad9a138b6679a84fe8536f73",  # admin desktop, Edu (1 Sept 2026)
}
CONFIG_TOML_SHA1_BASELINE = (
    os.environ.get("WI_CODEX_CONFIG_SHA1", "").strip().lower()
    or CONFIG_TOML_SHA1_BASELINES.get(_HOST, "")
)

CODEX_BIN   = os.environ.get("WI_CODEX_BIN", "codex")
CODEX_MODEL = os.environ.get("WI_CODEX_MODEL", "").strip()   # optional -m <model>
# codex-cli 0.151.0 cold-starts SLOW on the Oxford laptop -- attempt 1 of a real
# call was observed taking ~3m37s (1 Sept). Bumped from 240; a one-shot warm-up
# call (see _ensure_warm) absorbs the cold start once per process.
CALL1_TIMEOUT_S        = int(os.environ.get("WI_LANE_B_TIMEOUT", "300"))
CALL1_WARMUP_TIMEOUT_S = int(os.environ.get("WI_LANE_B_WARMUP_TIMEOUT", "360"))
# Headless connector availability FLIPS between runs on the same account
# (confirmed 1 Sept: same laptop/account, calendar fired one run, Teams the
# next). If the expected tool doesn't fire, re-invoke codex a few times before
# giving up on that domain for the cycle. This is NOT a HALT -- it's the
# "connector unavailable this cycle" path.
CALL1_RETRIES = max(1, int(os.environ.get("WI_LANE_B_RETRIES", "3")))
CALL1_RETRY_BACKOFF_S = [5, 12, 20, 30]
_WARMED = False

PRIMARY_CAL_NAME = "Calendar"
SHARED_CAL_NAME  = "People Department - HR Systems"

# --- re-contamination guard (revised 1 Sept 2026 after a false HALT) ---------
# The task-descriptive Call-1 prompt lets the model pick its own read tools, and
# it legitimately picks ones we didn't hard-list (observed: `search_events`
# alongside `list_calendars`/`list_events`). So the guard is now VERB-based, not
# an exact allowlist:
#   * server MUST be `codex_apps`.
#   * tool namespace (before the first '.') MUST be one of the two Lane B
#     connectors -- anything else on codex_apps (`microsoft_outlook_email.*`,
#     `github.*`, `canva.*`, ...) is off-scope -> HALT.
#   * within those namespaces: a READ verb leaf -> allow; a WRITE verb leaf or an
#     unrecognised verb -> HALT (fail closed).
LANE_B_NAMESPACES = {"microsoft_outlook_calendar", "microsoft_teams"}
READ_VERB_RE = re.compile(r"^(list|get|fetch|search|resolve)(_|$)", re.IGNORECASE)
WRITE_VERB_RE = re.compile(
    r"^(send|create|update|delete|remove|add|reply|forward|draft|cancel|respond|"
    r"accept|decline|tentatively|move|mark|set|patch|post|schedule|invite|share|"
    r"rsvp|clear|archive|pin|unpin|hide)(_|$)",
    re.IGNORECASE,
)
# Reference only (no longer the gate): the sec.5 read tool names known on 1 Sept.
CAL_ALLOW = {
    "list_calendars", "list_events", "search_events", "fetch_event", "fetch_events_batch",
    "list_event_instances", "list_recurring_series", "get_mailbox_settings",
}
TEAMS_ALLOW = {
    "list_teams", "list_channels", "list_chats", "resolve_team", "resolve_channel",
    "resolve_chat", "resolve_user", "list_chat_messages", "list_channel_messages",
    "fetch", "search", "get_chat_members", "resolve_scheduled_online_meeting",
    "list_online_meeting_transcripts", "get_online_meeting_transcript_content",
    "list_online_meeting_recordings",
}

EXPECTED_TOOL = {"calendar": "list_events", "teams": "list_chats"}


# --------------------------------------------------------------------------- #
def _utcstamp() -> str:
    return _dt.datetime.now(_dt.timezone.utc).strftime("%Y%m%dT%H%M%SZ")


def _log(msg: str) -> None:
    print(f"[{_dt.datetime.now(_dt.timezone.utc).strftime('%Y-%m-%d %H:%M:%SZ')}] lane_b_call1: {msg}")


def _config_toml_sha1() -> str | None:
    p = Path(os.path.expanduser("~")) / ".codex" / "config.toml"
    if not p.exists():
        return None
    return hashlib.sha1(p.read_bytes()).hexdigest().lower()


# --------------------------------------------------------------------------- #
#  Rigid Call-1 prompts (LANE_B sec.3, adapted to name the codex_apps tools)
# --------------------------------------------------------------------------- #
# NOTE (1 Sept 2026 probe finding): the codex_apps connector tools only load
# when the prompt describes the TASK, NOT when it names a tool imperatively.
# "Call `microsoft_outlook_calendar.list_events` ..." -> the model replies
# "I can't access that in this session" and no tool fires. "Using the Outlook
# Calendar connector, retrieve my events between X and Y ..." -> it works.
# So these prompts are task-descriptive. The "call no other tool / change
# nothing" guardrail clauses are kept -- those are fine.

def build_calendar_prompt(win_start_iso: str, win_end_iso: str) -> str:
    return (
        "Using the Microsoft Outlook Calendar app connector, retrieve my calendar events "
        f"between {win_start_iso} and {win_end_iso} (inclusive), ordered by start time. "
        f"Include events from my default calendar and, if it exists, the calendar named "
        f"\"{SHARED_CAL_NAME}\". Expand any recurring series into its concrete occurrences "
        "within that window. "
        "Return ONLY the raw connector result as a JSON array of the event objects, with no "
        "summary, no interpretation, and no prose. "
        "Do not use any other app or tool. Do not create, update, cancel, delete, move, "
        "respond to, or add an attachment to any event. Do not send any message or email."
    )


def build_teams_prompt(since_iso: str) -> str:
    return (
        "Using the Microsoft Teams app connector, retrieve my 40 most recent chats and, for "
        f"each chat or channel with activity since {since_iso}, the 30 newest messages. "
        "Return ONLY the raw connector results as JSON (the chats list and the messages), "
        "with no summary, no interpretation, and no prose. "
        "Do not use any other app or tool. Do not send or reply to any message, do not create "
        "a chat or channel, and do not touch Planner or tasks."
    )


# --------------------------------------------------------------------------- #
#  codex exec --json
# --------------------------------------------------------------------------- #
def _codex_argv0() -> list[str]:
    """Resolve how to launch codex, robustly on Windows.
    WI_CODEX_BIN may be a bare name ('codex'), a full path to codex.cmd/.exe, or
    a .ps1 (npm global on Kevin's laptop = C:\\Users\\...\\npm\\codex.ps1).
    Python 3.12+ will not run a .cmd/.bat via subprocess without the extension,
    and cannot exec a .ps1 at all -- handle both."""
    cand = CODEX_BIN
    resolved = _shutil.which(cand) or cand
    low = resolved.lower()
    if low.endswith(".ps1"):
        return ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", resolved]
    return [resolved]


def _ensure_warm() -> None:
    """One throwaway `codex exec` per process to absorb the cold-start hang
    (codex-cli 0.151.0 on the Oxford laptop can take 3+ min on the first call).
    Skipped when WI_LANE_B_SKIP_WARMUP=1 (e.g. the guard already warmed the box)."""
    global _WARMED
    if _WARMED or os.environ.get("WI_LANE_B_SKIP_WARMUP", "").strip().lower() in ("1", "true", "yes"):
        _WARMED = True
        return
    _WARMED = True
    _log(f"warming codex (cold start can take ~3+ min; timeout {CALL1_WARMUP_TIMEOUT_S}s)...")
    t0 = time.time()
    try:
        subprocess.run(
            _codex_argv0() + ["exec", "-s", "read-only", "--skip-git-repo-check",
                              "Reply with the single word OK. Use no tools, change nothing."],
            capture_output=True, text=True, timeout=CALL1_WARMUP_TIMEOUT_S,
            cwd=str(REPO_ROOT), env={**os.environ, "PYTHONUTF8": "1"},
        )
        _log(f"codex warm-up done in {time.time() - t0:.0f}s")
    except Exception as e:  # noqa: BLE001
        _log(f"codex warm-up did not complete in {time.time() - t0:.0f}s ({e}) -- continuing")


def run_codex_json(prompt: str, *, timeout_s: int, tag: str) -> tuple[list[dict], str]:
    """Return (parsed_json_objects, raw_stdout). Raises RuntimeError on hard failure."""
    _ensure_warm()
    cmd = _codex_argv0() + ["exec", "-s", "read-only", "--skip-git-repo-check", "--json"]
    if CODEX_MODEL:
        cmd += ["-m", CODEX_MODEL]
    cmd.append(prompt)

    last_raw = ""
    for attempt in (1, 2):
        _log(f"[{tag}] codex exec attempt {attempt}/2 (timeout {timeout_s}s)")
        try:
            proc = subprocess.run(
                cmd, capture_output=True, text=True, timeout=timeout_s,
                cwd=str(REPO_ROOT),
                env={**os.environ, "PYTHONUTF8": "1"},
            )
        except subprocess.TimeoutExpired as te:
            last_raw = (te.stdout or "") if isinstance(te.stdout, str) else ""
            _log(f"[{tag}] timed out after {timeout_s}s (cold-start hang?) -- retrying once" if attempt == 1
                 else f"[{tag}] timed out again")
            continue

        raw = proc.stdout or ""
        last_raw = raw
        objs = _parse_jsonl(raw)
        if objs:
            if proc.returncode != 0:
                _log(f"[{tag}] codex exited {proc.returncode} but produced parseable JSONL -- continuing")
            return objs, raw
        _log(f"[{tag}] no parseable JSONL (exit {proc.returncode}); stderr tail: "
             f"{(proc.stderr or '').strip()[-300:]!r}")
        if attempt == 1:
            continue

    raise RuntimeError(f"[{tag}] codex exec produced no usable JSON output after 2 attempts")


def _parse_jsonl(raw: str) -> list[dict]:
    out: list[dict] = []
    for line in raw.splitlines():
        line = line.strip()
        if not line or line[0] not in "{[":
            continue
        try:
            out.append(json.loads(line))
        except json.JSONDecodeError:
            continue
    return out


# --------------------------------------------------------------------------- #
#  Pull mcp tool calls + results out of the --json event stream.
#  codex's --json schema has shifted across versions, so this is defensive:
#  it recognises several shapes and also deep-scans for {server, tool}.
# --------------------------------------------------------------------------- #
def extract_tool_calls(events: list[dict]) -> list[dict]:
    """
    Return [{server, tool, arguments, result, error, raw}] for every completed
    mcp tool call. Schema (codex exec --json, confirmed 1 Sept 2026):
      {"type":"item.completed","item":{
         "type":"mcp_tool_call","server":"codex_apps",
         "tool":"microsoft_outlook_calendar.list_events","arguments":{...},
         "result":{"content":[{"type":"text","text":"Action completed."}],
                   "structured_content":{"value":[...] , "next_link":null}},
         "error":null,"status":"completed"}}
    The real payload is item.result.structured_content -- calendar: `.value`;
    teams list_chats: `.chats`; message lists: `.messages` / `.value`.
    item.result.content[0].text is just "Action completed." -> ignore.
    """
    found: list[dict] = []
    for ev in events:
        if not isinstance(ev, dict):
            continue
        if ev.get("type") not in ("item.completed", "item.updated"):
            continue
        item = ev.get("item") or {}
        if item.get("type") != "mcp_tool_call" or item.get("status") != "completed":
            continue
        found.append({
            "server": item.get("server") or "",
            "tool": item.get("tool") or "",
            "arguments": item.get("arguments") or {},
            "result": (item.get("result") or {}).get("structured_content"),
            "error": item.get("error"),
            "raw": item,
        })

    # de-dupe on (server, tool, args); keep the one that carries a result
    dedup: dict[tuple, dict] = {}
    for tc in found:
        key = (tc["server"], tc["tool"],
               json.dumps(tc["arguments"], sort_keys=True, default=str))
        if key not in dedup or (dedup[key]["result"] is None and tc["result"] is not None):
            dedup[key] = tc
    return list(dedup.values())


def final_assistant_text(events: list[dict]) -> str:
    """Best-effort: the model's last textual message (the 'return only the JSON array' output)."""
    texts: list[str] = []

    def _walk(node):
        if isinstance(node, dict):
            t = node.get("type") or node.get("role")
            if t in ("agent_message", "assistant", "message", "item.completed", "agent_message_delta"):
                for k in ("text", "message", "content", "delta"):
                    v = node.get(k)
                    if isinstance(v, str):
                        texts.append(v)
                    elif isinstance(v, list):
                        for seg in v:
                            if isinstance(seg, dict) and isinstance(seg.get("text"), str):
                                texts.append(seg["text"])
            for v in node.values():
                _walk(v)
        elif isinstance(node, list):
            for v in node:
                _walk(v)

    for ev in events:
        _walk(ev)
    return "\n".join(texts[-6:]).strip()


def _json_array_from_text(text: str) -> list | None:
    if not text:
        return None
    # strip code fences
    text = re.sub(r"```(?:json)?", "", text)
    depth = 0
    start = None
    for i, ch in enumerate(text):
        if ch == "[":
            if depth == 0:
                start = i
            depth += 1
        elif ch == "]":
            depth -= 1
            if depth == 0 and start is not None:
                frag = text[start:i + 1]
                try:
                    val = json.loads(frag)
                    if isinstance(val, list):
                        return val
                except json.JSONDecodeError:
                    pass
                start = None
    return None


# --------------------------------------------------------------------------- #
#  Re-contamination guard (LANE_B sec.6c, revised: assert on observed calls)
# --------------------------------------------------------------------------- #
def guard_recontamination(tool_calls: list[dict], domain: str) -> tuple[str, dict]:
    """Verb-based, not an exact allowlist. Returns ('ok'|'halt'|'unavailable', detail).
    HALT on: server != codex_apps; a tool namespace outside the two Lane B
    connectors; a write-verb leaf; an unrecognised-verb leaf (fail closed)."""
    seen: list[str] = []
    unexpected: list[str] = []
    for tc in tool_calls:
        srv, tool = tc["server"], tc["tool"]
        seen.append(f"{srv}::{tool}")
        if srv != "codex_apps":
            unexpected.append(f"{srv}::{tool} (server != codex_apps)")
            continue
        ns, _dot, leaf = tool.partition(".")
        if not _dot:
            unexpected.append(f"{tool} (no namespace)")
            continue
        if ns not in LANE_B_NAMESPACES:
            unexpected.append(f"{tool} (off-scope connector namespace '{ns}')")
            continue
        if WRITE_VERB_RE.match(leaf):
            unexpected.append(f"{tool} (write verb)")
            continue
        if not READ_VERB_RE.match(leaf):
            unexpected.append(f"{tool} (unrecognised verb -- add to READ_VERB_RE if it is a read)")
            continue
        # codex_apps + Lane B namespace + read verb -> allowed
    detail = {"seen": sorted(set(seen)), "unexpected": sorted(set(unexpected))}
    if unexpected:
        return "halt", detail
    if not any(tc["tool"].split(".")[-1] == EXPECTED_TOOL[domain] for tc in tool_calls):
        return "unavailable", detail
    return "ok", detail


# --------------------------------------------------------------------------- #
#  Map connector event/message objects -> normalise_pull raw shape
# --------------------------------------------------------------------------- #
def _first(d: dict, *keys, default=""):
    for k in keys:
        if k in d and d[k] not in (None, ""):
            return d[k]
    return default


_FRAC_RE = re.compile(r"\.\d+")


def _graph_dt_parts(v):
    """Graph datetime is {'dateTime': '2026-09-01T00:00:00.0000000', 'timeZone': 'UTC'}.
    Returns (naive_datetime | None, tz_name). Strips the 7-digit fractional part
    (datetime.fromisoformat rejects >6 digits)."""
    if isinstance(v, dict):
        s = _first(v, "dateTime", "date_time", "datetime", "value", default="")
        tz = _first(v, "timeZone", "time_zone", default="UTC")
    else:
        s, tz = (v or ""), "UTC"
    if not s:
        return None, tz
    s = _FRAC_RE.sub("", str(s).strip()).replace("Z", "")
    try:
        return _dt.datetime.fromisoformat(s), tz
    except ValueError:
        try:
            return _dt.datetime.fromisoformat(s[:19]), tz
        except ValueError:
            return None, tz


def _is_all_day(ev: dict, sdt, edt) -> bool:
    flag = ev.get("is_all_day", ev.get("isAllDay", ev.get("all_day")))
    if isinstance(flag, bool):
        return flag
    # heuristic: both ends at exactly midnight and the span is a whole number of days
    if sdt and edt and sdt.time() == _dt.time(0) and edt.time() == _dt.time(0):
        span = (edt - sdt)
        return span.days >= 1 and span.seconds == 0
    return False


def _to_pipeline_start(sdt, tz_name: str, all_day: bool) -> str:
    """Produce an ISO string fetch_inbox.py can datetime.fromisoformat().
    Timed events: convert UTC -> machine-local, emit with local offset (matches
    the COM path's str(item.Start) local wall time). All-day: keep the date at
    naive midnight, no tz shift (a UTC->BST shift could bump the day)."""
    if sdt is None:
        return ""
    if all_day:
        return sdt.replace(hour=0, minute=0, second=0, microsecond=0).isoformat()
    tzn = (tz_name or "UTC").strip().lower()
    src = _dt.timezone.utc if tzn in ("utc", "gmt", "z", "") else _dt.timezone.utc
    return sdt.replace(tzinfo=src).astimezone().isoformat()


def calendar_events_to_raw(events: list) -> list[dict]:
    out: list[dict] = []
    for ev in events:
        if not isinstance(ev, dict):
            continue
        org_ea = ((ev.get("organizer") or {}).get("emailAddress")
                  or (ev.get("organizer") or {}).get("email_address") or {})
        loc = ev.get("location") or {}
        rs = ev.get("response_status") or ev.get("responseStatus") or {}
        online = ev.get("online_meeting") or ev.get("onlineMeeting") or {}

        sdt, s_tz = _graph_dt_parts(ev.get("start"))
        edt, _e_tz = _graph_dt_parts(ev.get("end"))
        all_day = _is_all_day(ev, sdt, edt)

        out.append({
            "calendar_name": _first(ev, "calendar_name", "calendarName", default=""),
            "id": str(_first(ev, "id", "i_cal_u_id", "iCalUId", default="")),
            "subject": _first(ev, "subject", "display_title", "title", default=""),
            "start": _to_pipeline_start(sdt, s_tz, all_day),
            "end": _to_pipeline_start(edt, _e_tz, all_day),
            "is_all_day": all_day,
            "location": _first(loc, "displayName", "display_name", default="") if isinstance(loc, dict) else str(loc or ""),
            "organizer_name": _first(org_ea, "name", default="") or _first(ev, "organizer_name", default=""),
            "organizer_email": _first(org_ea, "address", "email", default="") or _first(ev, "organizer_email", default=""),
            "is_cancelled": bool(_first(ev, "is_cancelled", "isCancelled", default=False)),
            "response_status": _first(rs, "response", default="") if isinstance(rs, dict) else str(rs or ""),
            "series_master_id": str(ev.get("series_master_id") or ev.get("seriesMasterId") or ""),
            "has_online_meeting": bool(
                ev.get("has_online_meeting", ev.get("isOnlineMeeting",
                    bool(online) or bool(_first(ev, "online_meeting_join_url", default=""))))
            ),
            "online_meeting_join_url": _first(online, "join_url", "joinUrl", default="") or _first(ev, "online_meeting_join_url", default=""),
            "body_preview": (_first(ev, "bodyPreview", "body_preview", default="")
                             or (_first(ev.get("body") or {}, "content", default="")[:300]
                                 if isinstance(ev.get("body"), dict) else "")),
        })
    return out


def teams_messages_to_raw(messages: list) -> list[dict]:
    out: list[dict] = []
    for m in messages:
        if not isinstance(m, dict):
            continue
        frm = m.get("from") or m.get("sender") or {}
        frm_user = (frm.get("user") if isinstance(frm, dict) else {}) or frm
        out.append({
            "kind": _first(m, "kind", default="chat"),
            "container_name": _first(m, "container_name", "chat_name", "channel_name", "conversation_name", default=""),
            "message_id": str(_first(m, "message_id", "id", default="")),
            "from_name": _first(frm_user, "display_name", "displayName", "name", default="") or _first(m, "from_name", default=""),
            "from_email": _first(frm_user, "email", "mail", "userPrincipalName", default="") or _first(m, "from_email", default=""),
            "created": _first(m, "created", "created_date_time", "createdDateTime", default=""),
            "is_from_me": bool(_first(m, "is_from_me", "from_me", default=False)),
            "has_attachments": bool(m.get("attachments") or _first(m, "has_attachments", default=False)),
            "body_preview": _first(m, "body_preview", "bodyPreview", default="") or (
                _first(m.get("body") or {}, "content", default="")[:400] if isinstance(m.get("body"), dict) else ""
            ),
        })
    return out


def _events_from_results(tool_calls: list[dict], domain: str, events: list[dict]) -> list:
    """Pull the arrays out of item.result.structured_content. Never uses the
    model's final agent_message (per the 1 Sept probe: it's unreliable / summary).
    The `_json_array_from_text` fallback is retained only for a --from-file
    transcript that predates the structured_content schema."""
    data_tools = ({"list_events", "search_events", "list_event_instances", "fetch_events_batch"}
                  if domain == "calendar"
                  else {"list_chats", "list_chat_messages", "list_channel_messages",
                        "list_channels", "search"})
    keys = (("value", "events", "items") if domain == "calendar"
            else ("chats", "messages", "value", "items"))
    collected: list = []
    for tc in tool_calls:
        if tc["tool"].split(".")[-1] not in data_tools:
            continue
        if tc.get("error"):
            _log(f"[{domain}] tool {tc['tool']} returned error: {tc['error']!r} -- skipped")
            continue
        res = tc.get("result")
        if isinstance(res, str):
            try:
                res = json.loads(res)
            except json.JSONDecodeError:
                res = None
        if isinstance(res, dict):
            for k in keys:
                if isinstance(res.get(k), list):
                    collected.extend(res[k])
                    break
        elif isinstance(res, list):
            collected.extend(res)
    if collected:
        return collected
    arr = _json_array_from_text(final_assistant_text(events))
    if arr:
        _log(f"[{domain}] WARNING: no structured_content array found; fell back to final-message JSON ({len(arr)} items)")
    return arr or []


# --------------------------------------------------------------------------- #
def run_domain(domain: str, events: list[dict], *, window_days: int) -> dict:
    """Returns a result dict for this domain: {status, guard, counts, raw_items}."""
    tool_calls = extract_tool_calls(events)
    status, guard_detail = guard_recontamination(tool_calls, domain)
    _log(f"[{domain}] tool calls observed: {guard_detail['seen'] or '(none)'}")
    if guard_detail["unexpected"]:
        _log(f"[{domain}] UNEXPECTED tools: {guard_detail['unexpected']}")

    raw_items: list = []
    if status == "ok":
        objs = _events_from_results(tool_calls, domain, events)
        raw_items = (calendar_events_to_raw(objs) if domain == "calendar"
                     else teams_messages_to_raw(objs))
        _log(f"[{domain}] extracted {len(raw_items)} item(s)")
    elif status == "unavailable":
        _log(f"[{domain}] connector did not return {EXPECTED_TOOL[domain]} -- treating as unavailable this cycle (empty, no HALT)")

    return {
        "domain": domain,
        "status": status,
        "guard": guard_detail,
        "tool_calls": [f"{t['server']}::{t['tool']}" for t in tool_calls],
        "count": len(raw_items),
        "raw_items": raw_items,
    }


def fetch_domain(domain: str, prompt: str, *, window_days: int, ts: str, retries: int) -> dict:
    """Run codex exec for one domain, retrying while the expected connector tool
    does not fire. Returns a run_domain-shaped dict plus an `attempts` list.
    Never raises. Terminal statuses: 'ok', 'halt' (re-contamination),
    'unavailable' (expected tool never fired across all attempts),
    'codex_failed' (every attempt's codex run failed)."""
    attempts: list[dict] = []
    result: dict | None = None
    for n in range(1, retries + 1):
        try:
            events, raw = run_codex_json(prompt, timeout_s=CALL1_TIMEOUT_S, tag=f"{domain}#{n}")
        except RuntimeError as e:
            attempts.append({"n": n, "outcome": "codex_failed", "detail": str(e)[:200]})
            _log(f"[{domain}] attempt {n}/{retries}: codex run failed -- {e}")
            if n < retries:
                time.sleep(CALL1_RETRY_BACKOFF_S[min(n - 1, len(CALL1_RETRY_BACKOFF_S) - 1)])
            continue
        try:
            (LANE_B_DIR / f"{ts}_call1_{domain}_a{n}.jsonl").write_text(raw, encoding="utf-8")
        except OSError:
            pass
        result = run_domain(domain, events, window_days=window_days)
        attempts.append({"n": n, "outcome": result["status"],
                         "tools": result["tool_calls"], "count": result["count"]})
        if result["status"] in ("ok", "halt"):
            break
        _log(f"[{domain}] attempt {n}/{retries}: {EXPECTED_TOOL[domain]} did not fire "
             f"(connector unavailable)" + (" -- retrying" if n < retries else " -- giving up this cycle"))
        if n < retries:
            time.sleep(CALL1_RETRY_BACKOFF_S[min(n - 1, len(CALL1_RETRY_BACKOFF_S) - 1)])

    if result is None:
        result = {"domain": domain, "status": "codex_failed",
                  "guard": {"seen": [], "unexpected": []},
                  "tool_calls": [], "count": 0, "raw_items": []}
    result["attempts"] = attempts
    _log(f"[{domain}] final status={result['status']} after {len(attempts)} attempt(s)")
    return result


# --------------------------------------------------------------------------- #
def main(argv: list[str]) -> int:
    ap = argparse.ArgumentParser(description="Lane B Call-1 codex_apps connector fetch")
    ap.add_argument("--domain", choices=["calendar", "teams", "both"], default="calendar")
    ap.add_argument("--window-days", type=int, default=7)
    ap.add_argument("--teams-lookback-h", type=int, default=72)
    ap.add_argument("--dry-run", action="store_true", help="print the prompt(s), run nothing")
    ap.add_argument("--from-file", help="parse this pre-captured codex --json transcript instead of running codex")
    ap.add_argument("--out", default=str(NORMALISED_OUT))
    args = ap.parse_args(argv)

    ts = _utcstamp()
    _log(f"start domain={args.domain} window_days={args.window_days} ts={ts}")
    LANE_B_DIR.mkdir(parents=True, exist_ok=True)
    CODEX_RUNS_DIR.mkdir(parents=True, exist_ok=True)

    today = _dt.date.today()
    win_start = _dt.datetime(today.year, today.month, today.day, tzinfo=_dt.timezone.utc)
    win_end = win_start + _dt.timedelta(days=args.window_days)
    win_start_iso = win_start.strftime("%Y-%m-%dT%H:%M:%SZ")
    win_end_iso = win_end.strftime("%Y-%m-%dT%H:%M:%SZ")
    since_iso = (_dt.datetime.now(_dt.timezone.utc) - _dt.timedelta(hours=args.teams_lookback_h)).strftime("%Y-%m-%dT%H:%M:%SZ")

    domains = ["calendar", "teams"] if args.domain == "both" else [args.domain]
    prompts = {
        "calendar": build_calendar_prompt(win_start_iso, win_end_iso),
        "teams": build_teams_prompt(since_iso),
    }

    if args.dry_run:
        for d in domains:
            print(f"\n===== {d} prompt =====\n{prompts[d]}\n")
        return 0

    sha_before = _config_toml_sha1()
    if sha_before and not args.from_file:
        if CONFIG_TOML_SHA1_BASELINE and sha_before != CONFIG_TOML_SHA1_BASELINE:
            _log(f"WARNING (never a HALT): ~/.codex/config.toml sha1 {sha_before} != recorded "
                 f"baseline {CONFIG_TOML_SHA1_BASELINE} for host {_HOST} "
                 f"(set WI_CODEX_CONFIG_SHA1 to re-baseline)")
        elif not CONFIG_TOML_SHA1_BASELINE:
            _log(f"note: no recorded ~/.codex/config.toml baseline for host {_HOST}; "
                 f"observed sha1 {sha_before} -- add it to CONFIG_TOML_SHA1_BASELINES")

    per_domain: dict[str, dict] = {}
    overall_rc = 0
    for d in domains:
        if args.from_file:
            events = _parse_jsonl(Path(args.from_file).read_text(encoding="utf-8", errors="replace"))
            if not events:
                _log(f"[{d}] --from-file produced no parseable JSONL")
                return 3
            per_domain[d] = run_domain(d, events, window_days=args.window_days)
            per_domain[d]["attempts"] = [{"n": 1, "outcome": per_domain[d]["status"], "from_file": True}]
        else:
            per_domain[d] = fetch_domain(d, prompts[d], window_days=args.window_days,
                                         ts=ts, retries=CALL1_RETRIES)
        if per_domain[d]["status"] == "halt":
            overall_rc = 1

    sha_after = _config_toml_sha1()
    any_ok = any(r["status"] == "ok" for r in per_domain.values())

    # --- assemble the normalise_pull raw shape, sanitise, write ---
    raw_lane_b = {
        "calendar": per_domain.get("calendar", {}).get("raw_items", []),
        "teams": per_domain.get("teams", {}).get("raw_items", []),
        "transcripts": [],
    }
    normalised, hits = normalise_pull.normalise(raw_lane_b, ts=ts)

    # carry Lane B provenance + guard status into meta so fetch_inbox.py can gate
    _dom_keys = ("status", "count", "tool_calls", "guard", "attempts")
    normalised["meta"]["lane_b"] = {
        "ts": ts,
        "domains": {d: {k: per_domain[d].get(k) for k in _dom_keys} for d in per_domain},
        "config_toml_sha1_before": sha_before,
        "config_toml_sha1_after": sha_after,
        "config_toml_sha1_match": (sha_before == sha_after) if (sha_before and sha_after) else None,
        "halt": overall_rc == 1,
    }

    if overall_rc == 1:
        # RE-CONTAMINATION GUARD TRIPPED: write the marker, do NOT overwrite a good normalised file
        trip = CODEX_RUNS_DIR / f"GUARD_TRIPPED_{ts}.json"
        trip.write_text(json.dumps(normalised["meta"]["lane_b"], indent=2), encoding="utf-8")
        _log(f"GUARD TRIPPED -- wrote {trip}. NOT updating {args.out}. Caller must HALT.")
        _write_run_log(ts, args, per_domain, sha_before, sha_after, len(hits), halted=True)
        return 1

    if not any_ok:
        # every requested domain was 'unavailable' / 'codex_failed' this cycle
        # (headless connector flakiness). Do NOT overwrite a previous good file --
        # fetch_inbox.py will use it until it ages past WI_LANE_B_MAX_AGE_H, then
        # degrade to empty+warning. This is the documented flaky path, not an error.
        _log(f"no domain returned data this cycle "
             f"({ {d: per_domain[d]['status'] for d in per_domain} }); "
             f"leaving any existing {Path(args.out).name} in place (last-good until it ages out).")
        _write_run_log(ts, args, per_domain, sha_before, sha_after, len(hits), halted=False)
        return 0

    Path(args.out).write_text(json.dumps(normalised, indent=2, ensure_ascii=False), encoding="utf-8")
    (CODEX_RUNS_DIR / f"{ts}_sanitiser_hits.json").write_text(
        json.dumps({"ts": ts, "hits": hits}, indent=2, ensure_ascii=False), encoding="utf-8")
    _write_run_log(ts, args, per_domain, sha_before, sha_after, len(hits), halted=False)

    counts = normalised["meta"]["counts"]
    _log(f"wrote {args.out} calendar={counts['calendar']} teams={counts['teams']} "
         f"sanitiser_hits={len(hits)} "
         f"status={ {d: per_domain[d]['status'] for d in per_domain} }")
    return overall_rc


def _write_run_log(ts, args, per_domain, sha_before, sha_after, n_hits, *, halted):
    log = {
        "ts": ts,
        "domains_requested": args.domain,
        "window_days": args.window_days,
        "per_domain": {d: {k: per_domain[d].get(k)
                           for k in ("status", "count", "tool_calls", "guard", "attempts")}
                       for d in per_domain},
        "sanitiser_hits": n_hits,
        "config_toml_sha1_before": sha_before,
        "config_toml_sha1_after": sha_after,
        "config_toml_sha1_match": (sha_before == sha_after) if (sha_before and sha_after) else None,
        "halted": halted,
        "codex_bin": CODEX_BIN,
    }
    (LANE_B_DIR / f"{ts}_lane_b.json").write_text(json.dumps(log, indent=2, ensure_ascii=False), encoding="utf-8")


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
