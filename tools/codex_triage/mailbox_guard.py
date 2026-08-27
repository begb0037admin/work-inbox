#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
mailbox_guard.py -- post-run Outlook COM delta-sweep KILL-SWITCH for the Codex
parallel AI-triage run (work-inbox).  Build Step 1 of docs/OPTION1_KILLSWITCH.md.

WHY THIS EXISTS
---------------
Kevin has explicitly accepted the residual write-risk of running the Codex
AI-triage phases on his EXISTING ChatGPT account with the Microsoft Outlook /
Calendar / Teams connectors still attached.  Every preventive write-gate control
was tested 26-27 Aug 2026 and FAILED (see
docs/CODEX_CONNECTOR_MIGRATION_RESEARCH.md Section 9).  That acceptance is
CONTINGENT on this kill-switch:

  A mandatory before/after Outlook COM snapshot around every Codex run that
  HARD-DISABLES the scheduled task, ALERTS Kevin, and writes a timestamped
  INCIDENT RECORD the moment ANY mailbox mutation is detected that the run
  itself did not intend.

The Codex parallel run's ONLY intended output is local files
(docs/codex_*.json + data/codex_runs/*).  It performs NO mailbox write of any
kind.  So the set of "expected" mailbox deltas is EMPTY -- any delta trips.

This is DETECTION, not prevention.  It cannot stop a write in progress; it stops
the NEXT run and makes sure Kevin knows within one cycle.

SCOPE OF THE SWEEP  (Outlook COM, Kevin's primary mailbox)
---------------------------------------------------------
Covered:
  * Top-level Inbox -- newest INBOX_MAX within INBOX_LOOKBACK_DAYS, PLUS every
    message currently carrying a category or a flag regardless of age (so a
    category landed on an old message is still caught).
  * The five named Inbox subfolder trees the live pipeline reads
    (Senior Management / Bi-monthly CDR/PD working group / H&S / Team /
    Projects), one level deep.
  * Per message: Categories, FlagStatus, FlagRequest, UnRead, Importance and
    parent folder path (a move shows as disappearance + reappearance).
  * Sent Items count + newest-5 identities  (a Codex send raises the count).
  * Drafts count + newest-5 identities.
  * Default Calendar + the "People Department - HR Systems" shared calendar,
    working window [today-CAL_BACK_DAYS, today+CAL_FWD_DAYS]: per event
    Start/End/Subject/Location hash, plus adds / removes (a Codex
    create_event / update_event shows here).
NOT covered -- documented residual gap:
  * Microsoft Teams.  Teams is not reachable through Outlook COM.  A Codex
    Teams chat.ReadWrite misfire would NOT be seen by this sweep.  See
    docs/OPTION1_KILLSWITCH.md "Teams gap" for the accepted mitigation.

STALE-CACHE GOTCHA (reconfirmed 27 Aug 2026)
--------------------------------------------
A COM Categories/flag read taken within seconds of a Graph-side write returns a
false unchanged value.  Before the AFTER snapshot this script forces
Namespace.SyncObjects .Start() on every sync object and then sleeps
--settle-seconds (default 60) so a real write has propagated back to the local
store before it is read.  This only ever risks a FALSE NEGATIVE (missed write),
never a false trip, so erring long is safe.

CLI
---
  snapshot  --out FILE [--label before|after] [--settle-seconds N] [--no-settle]
  diff      --before FILE --after FILE [--out FILE]
  guard     --before FILE --after FILE --task "Work Inbox Codex Parallel"
                                        [--sensitivity strict|writes-only]
                                        [--dry-run]
  prove     [--settle-seconds N] [--keep-artifacts]
  clear-flag

Exit codes:  0 = clean / no delta;  2 = delta found (guard tripped);
             3 = refused (task name is the live pipeline);  4 = COM / usage error.

Every path prints an ISO-8601 timestamped line (standing requirement).
Stdlib + pywin32 only.
"""

import argparse
import ctypes
import datetime as _dt
import hashlib
import json
import os
import subprocess
import sys
import time

# --------------------------------------------------------------------------- #
#  Constants
# --------------------------------------------------------------------------- #

HERE = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.abspath(os.path.join(HERE, "..", ".."))
RUNS_DIR = os.path.join(REPO_ROOT, "data", "codex_runs")
TRIPPED_FLAG = os.path.join(RUNS_DIR, "GUARD_TRIPPED.flag")

# Reuse the live pipeline's notification mechanism verbatim (fetch_inbox.py
# _notify_phase_failure -> Show-TaskNotification.ps1 / BurntToast).
NOTIFY_SCRIPT_PATH = r"D:\OneDrive - lelitte.com\Desktop\Show-TaskNotification.ps1"

# The one task name this script must NEVER disable -- the live briefing pipeline.
LIVE_PIPELINE_TASK = "work inbox briefing"

# Named Inbox subfolder trees the live Phase 1c sweep reads (fetch_inbox.py).
SUBFOLDER_TREES = [
    "Senior Management",
    "Bi-monthly CDR/PD working group",
    "H&S",
    "Team",
    "Projects",
]

INBOX_LOOKBACK_DAYS = 30
INBOX_MAX = 250
CAL_BACK_DAYS = 1
CAL_FWD_DAYS = 8
DEFAULT_SETTLE_SECONDS = 60

# Outlook default folder ids
OL_INBOX = 6
OL_SENT = 5
OL_DRAFTS = 16
OL_CALENDAR = 9

# Senders considered safe to touch in `prove` mode (automated, non-work).
PROVE_SAFE_DOMAINS = (
    "distrokid.com", "github.com", "linkedin.com", "spotify.com",
    "e.linkedin.com", "notifications.google.com", "youtube.com",
    "medium.com", "substack.com", "meetup.com", "eventbrite.com",
)


# --------------------------------------------------------------------------- #
#  Small helpers
# --------------------------------------------------------------------------- #

def now_iso():
    return _dt.datetime.now().strftime("%Y-%m-%dT%H:%M:%S")


def log(msg):
    print(f"[{now_iso()}] mailbox_guard: {msg}", flush=True)


def sha12(text):
    return hashlib.sha1((text or "").encode("utf-8", "replace")).hexdigest()[:12]


def _com_str(v):
    return "" if v is None else str(v)


def ensure_runs_dir():
    os.makedirs(RUNS_DIR, exist_ok=True)


# --------------------------------------------------------------------------- #
#  Outlook COM
# --------------------------------------------------------------------------- #

def connect_mapi(max_attempts=3, wait_s=20):
    import pythoncom
    import win32com.client
    last = None
    for attempt in range(1, max_attempts + 1):
        try:
            pythoncom.CoInitialize()
            app = win32com.client.Dispatch("Outlook.Application")
            ns = app.GetNamespace("MAPI")
            _ = ns.GetDefaultFolder(OL_INBOX).Name  # force a real call
            log(f"Outlook COM connected on attempt {attempt}/{max_attempts}.")
            return ns
        except Exception as e:  # noqa: BLE001
            last = e
            log(f"Outlook COM attempt {attempt}/{max_attempts} failed: {e}")
            if attempt < max_attempts:
                time.sleep(wait_s)
    raise RuntimeError(f"Outlook COM connection failed after {max_attempts} attempts: {last}")


def force_sync(ns):
    """Trigger send/receive on every configured sync object so a Graph-side
    write propagates back to the local store before the AFTER snapshot reads."""
    try:
        sos = ns.SyncObjects
        n = sos.Count
        for i in range(1, n + 1):
            try:
                sos.Item(i).Start()
            except Exception:  # noqa: BLE001
                pass
        log(f"force_sync: started {n} sync object(s).")
    except Exception as e:  # noqa: BLE001
        log(f"force_sync: skipped ({e}).")


def _msg_record(item):
    """Mutable-state fingerprint for a mail item, keyed later by EntryID."""
    try:
        parent_path = _com_str(getattr(item.Parent, "FolderPath", ""))
    except Exception:  # noqa: BLE001
        parent_path = ""
    cats = _com_str(getattr(item, "Categories", ""))
    try:
        flag_status = int(getattr(item, "FlagStatus", 0) or 0)
    except Exception:  # noqa: BLE001
        flag_status = 0
    flag_request = _com_str(getattr(item, "FlagRequest", ""))
    try:
        unread = bool(getattr(item, "UnRead", False))
    except Exception:  # noqa: BLE001
        unread = False
    try:
        importance = int(getattr(item, "Importance", 1) or 1)
    except Exception:  # noqa: BLE001
        importance = 1
    subject_sha1 = sha12(_com_str(getattr(item, "Subject", "")))
    mutable = f"{cats}|{flag_status}|{flag_request}|{unread}|{importance}|{parent_path}"
    return {
        "subject_sha1": subject_sha1,
        "categories": cats,
        "flag_status": flag_status,
        "flag_request": flag_request,
        "unread": unread,
        "importance": importance,
        "folder_path": parent_path,
        "state_hash": sha12(mutable),
    }


def _iter_folder_items(folder):
    try:
        items = folder.Items
    except Exception:  # noqa: BLE001
        return
    try:
        items.Sort("[ReceivedTime]", True)
    except Exception:  # noqa: BLE001
        pass
    for it in items:
        yield it


def _scan_inbox(ns, records, cutoff):
    inbox = ns.GetDefaultFolder(OL_INBOX)
    seen = 0
    kept = 0
    for it in _iter_folder_items(inbox):
        seen += 1
        try:
            recv = it.ReceivedTime
            recv_dt = _dt.datetime(recv.year, recv.month, recv.day, recv.hour, recv.minute, recv.second)
        except Exception:  # noqa: BLE001
            recv_dt = None
        try:
            eid = it.EntryID
        except Exception:  # noqa: BLE001
            continue
        in_window = recv_dt is not None and recv_dt >= cutoff
        has_marking = bool(_com_str(getattr(it, "Categories", "")).strip()) or int(getattr(it, "FlagStatus", 0) or 0) != 0
        if not in_window and not has_marking:
            # Past the lookback window AND not marked -- older items are
            # ordered after this point, so once we are clearly past the window
            # with the cap already met we can stop.
            if kept >= INBOX_MAX:
                break
            continue
        if kept >= INBOX_MAX and not has_marking:
            break
        records[eid] = _msg_record(it)
        kept += 1
    log(f"_scan_inbox: {kept} tracked of {seen} scanned (top-level Inbox).")


def _scan_subfolders(ns, records):
    try:
        inbox = ns.GetDefaultFolder(OL_INBOX)
        subs = inbox.Folders
    except Exception as e:  # noqa: BLE001
        log(f"_scan_subfolders: cannot open Inbox.Folders ({e}) -- skipped.")
        return
    wanted = {n.lower() for n in SUBFOLDER_TREES}
    added = 0
    for i in range(1, subs.Count + 1):
        try:
            sub = subs.Item(i)
        except Exception:  # noqa: BLE001
            continue
        if _com_str(sub.Name).lower() not in wanted:
            continue
        for it in _iter_folder_items(sub):
            try:
                eid = it.EntryID
            except Exception:  # noqa: BLE001
                continue
            records[eid] = _msg_record(it)
            added += 1
    log(f"_scan_subfolders: {added} tracked across named trees.")


def _folder_summary(folder, label):
    try:
        items = folder.Items
        items.Sort("[ReceivedTime]", True)
        count = items.Count
    except Exception as e:  # noqa: BLE001
        log(f"_folder_summary({label}): failed ({e}).")
        return {"count": None, "newest": []}
    newest = []
    n = 0
    for it in items:
        if n >= 5:
            break
        try:
            newest.append({
                "subject_sha1": sha12(_com_str(getattr(it, "Subject", ""))),
                "entry_id_sha1": sha12(_com_str(getattr(it, "EntryID", ""))),
            })
            n += 1
        except Exception:  # noqa: BLE001
            continue
    return {"count": count, "newest": newest}


def _scan_calendar(ns, records):
    today = _dt.date.today()
    lo = today - _dt.timedelta(days=CAL_BACK_DAYS)
    hi = today + _dt.timedelta(days=CAL_FWD_DAYS)

    def scan(folder, tag):
        try:
            items = folder.Items
            items.IncludeRecurrences = True
            items.Sort("[Start]")
        except Exception as e:  # noqa: BLE001
            log(f"_scan_calendar({tag}): failed ({e}).")
            return 0
        added = 0
        for it in items:
            try:
                st = it.Start
                st_d = _dt.date(st.year, st.month, st.day)
            except Exception:  # noqa: BLE001
                continue
            if st_d > hi:
                break
            if st_d < lo:
                continue
            try:
                eid = it.EntryID
            except Exception:  # noqa: BLE001
                continue
            subj = sha12(_com_str(getattr(it, "Subject", "")))
            loc = sha12(_com_str(getattr(it, "Location", "")))
            start_s = _com_str(getattr(it, "Start", ""))
            end_s = _com_str(getattr(it, "End", ""))
            mutable = f"{subj}|{loc}|{start_s}|{end_s}|{tag}"
            records[f"{tag}:{eid}"] = {
                "calendar": tag,
                "subject_sha1": subj,
                "location_sha1": loc,
                "start": start_s,
                "end": end_s,
                "state_hash": sha12(mutable),
            }
            added += 1
        log(f"_scan_calendar({tag}): {added} events in window.")
        return added

    scan(ns.GetDefaultFolder(OL_CALENDAR), "primary")
    try:
        kevin_store = None
        for st in ns.Folders:
            if _com_str(st.Name) == "kevin.lelitte@admin.ox.ac.uk":
                kevin_store = st
                break
        if kevin_store is not None:
            hr_cal = kevin_store.Folders("Calendar").Folders("People Department - HR Systems")
            scan(hr_cal, "hr_systems_shared")
        else:
            log("_scan_calendar: kevin.lelitte store not found -- shared HR calendar skipped.")
    except Exception as e:  # noqa: BLE001
        log(f"_scan_calendar: shared HR calendar skipped ({e}).")


def take_snapshot(label, settle_seconds, do_settle):
    ns = connect_mapi()
    if do_settle:
        force_sync(ns)
        log(f"settling {settle_seconds}s for write propagation before reading...")
        time.sleep(max(0, int(settle_seconds)))
    messages = {}
    _scan_inbox(ns, messages, _dt.datetime.now() - _dt.timedelta(days=INBOX_LOOKBACK_DAYS))
    _scan_subfolders(ns, messages)
    calendar = {}
    _scan_calendar(ns, calendar)
    sent = _folder_summary(ns.GetDefaultFolder(OL_SENT), "Sent Items")
    drafts = _folder_summary(ns.GetDefaultFolder(OL_DRAFTS), "Drafts")
    snap = {
        "schema": "mailbox_guard/snapshot/v1",
        "label": label,
        "timestamp": now_iso(),
        "settled_seconds": int(settle_seconds) if do_settle else 0,
        "counts": {
            "messages_tracked": len(messages),
            "calendar_events_tracked": len(calendar),
            "sent_items": sent["count"],
            "drafts": drafts["count"],
        },
        "messages": messages,
        "calendar": calendar,
        "sent": sent,
        "drafts": drafts,
        "teams_covered": False,
        "notes": [
            "Teams is NOT covered -- not reachable via Outlook COM. See docs/OPTION1_KILLSWITCH.md.",
        ],
    }
    return snap


# --------------------------------------------------------------------------- #
#  Diff
# --------------------------------------------------------------------------- #

# Delta types that Kevin's own passive email-reading on another device can
# legitimately produce mid-run.  `writes-only` sensitivity ignores exactly
# these; a headless read-only Codex run cannot cause them.  `strict` (default)
# ignores nothing.
USER_PLAUSIBLE_TYPES = {"read_state_changed"}


def diff_snapshots(before, after):
    deltas = []
    b_msgs = before.get("messages", {})
    a_msgs = after.get("messages", {})

    for eid, b in b_msgs.items():
        a = a_msgs.get(eid)
        if a is None:
            deltas.append({
                "type": "message_disappeared",
                "entry_id_sha1": sha12(eid),
                "subject_sha1": b.get("subject_sha1"),
                "before": {"folder_path": b.get("folder_path")},
                "after": None,
                "severity": "critical",
                "note": "Tracked message no longer present at its EntryID -- moved or deleted.",
            })
            continue
        if a.get("state_hash") == b.get("state_hash"):
            continue
        for field, dtype, sev in (
            ("categories", "categories_changed", "critical"),
            ("flag_status", "flag_changed", "critical"),
            ("flag_request", "flag_changed", "critical"),
            ("importance", "importance_changed", "critical"),
            ("unread", "read_state_changed", "warn"),
            ("folder_path", "message_moved", "critical"),
        ):
            if b.get(field) != a.get(field):
                deltas.append({
                    "type": dtype,
                    "entry_id_sha1": sha12(eid),
                    "subject_sha1": b.get("subject_sha1"),
                    "field": field,
                    "before": b.get(field),
                    "after": a.get(field),
                    "severity": sev,
                })

    # Sent Items -- a rise is a possible outbound send by Codex.
    b_sent = (before.get("sent") or {}).get("count")
    a_sent = (after.get("sent") or {}).get("count")
    if b_sent is not None and a_sent is not None and a_sent > b_sent:
        deltas.append({
            "type": "sent_items_increased",
            "before": b_sent, "after": a_sent,
            "severity": "critical",
            "note": "Sent Items count rose during the run -- possible outbound send.",
        })

    # Drafts -- any change (Codex create_draft / send leaves a trace).
    b_dr = (before.get("drafts") or {}).get("count")
    a_dr = (after.get("drafts") or {}).get("count")
    if b_dr is not None and a_dr is not None and a_dr != b_dr:
        deltas.append({
            "type": "drafts_changed",
            "before": b_dr, "after": a_dr,
            "severity": "critical",
            "note": "Drafts count changed during the run.",
        })

    # Calendar
    b_cal = before.get("calendar", {})
    a_cal = after.get("calendar", {})
    for key, b in b_cal.items():
        a = a_cal.get(key)
        if a is None:
            deltas.append({
                "type": "calendar_item_disappeared",
                "calendar": b.get("calendar"),
                "subject_sha1": b.get("subject_sha1"),
                "before": {"start": b.get("start"), "end": b.get("end")},
                "severity": "critical",
            })
        elif a.get("state_hash") != b.get("state_hash"):
            deltas.append({
                "type": "calendar_item_changed",
                "calendar": b.get("calendar"),
                "subject_sha1": b.get("subject_sha1"),
                "before": {"start": b.get("start"), "end": b.get("end"), "location_sha1": b.get("location_sha1")},
                "after": {"start": a.get("start"), "end": a.get("end"), "location_sha1": a.get("location_sha1")},
                "severity": "critical",
            })
    for key, a in a_cal.items():
        if key not in b_cal:
            deltas.append({
                "type": "calendar_item_added",
                "calendar": a.get("calendar"),
                "subject_sha1": a.get("subject_sha1"),
                "after": {"start": a.get("start"), "end": a.get("end")},
                "severity": "critical",
                "note": "New calendar event appeared during the run -- possible create_event.",
            })

    return deltas


def classify(deltas, sensitivity):
    """Return (tripping_deltas, ignored_deltas) for the chosen sensitivity."""
    if sensitivity == "writes-only":
        trip = [d for d in deltas if d.get("type") not in USER_PLAUSIBLE_TYPES]
        ign = [d for d in deltas if d.get("type") in USER_PLAUSIBLE_TYPES]
        return trip, ign
    return list(deltas), []


# --------------------------------------------------------------------------- #
#  Enforce  (the kill-switch actions)
# --------------------------------------------------------------------------- #

def _disable_task(task_name, dry_run):
    if task_name.strip().lower() == LIVE_PIPELINE_TASK:
        log(f"REFUSING to disable '{task_name}' -- that is the live briefing pipeline.")
        return {"action": "refused", "reason": "live pipeline task", "task": task_name}
    if dry_run:
        log(f"[dry-run] would run: schtasks /Change /TN \"{task_name}\" /DISABLE")
        return {"action": "dry-run", "task": task_name}
    try:
        p = subprocess.run(
            ["schtasks", "/Change", "/TN", task_name, "/DISABLE"],
            capture_output=True, text=True, timeout=30,
        )
        out = (p.stdout or "").strip()
        err = (p.stderr or "").strip()
        if p.returncode == 0:
            log(f"scheduled task '{task_name}' DISABLED.")
            return {"action": "disabled", "task": task_name, "stdout": out}
        # Task not present is expected during Step-3 manual runs (Step 6 job
        # not yet created) -- still a real trip, just nothing to disable.
        log(f"schtasks returned {p.returncode} for '{task_name}': {err or out}")
        return {"action": "task_absent_or_error", "task": task_name,
                "returncode": p.returncode, "stderr": err, "stdout": out}
    except Exception as e:  # noqa: BLE001
        log(f"schtasks call failed: {e}")
        return {"action": "error", "task": task_name, "error": str(e)}


def _alert_kevin(detail_text, task_label):
    ensure_runs_dir()
    detail_path = os.path.join(RUNS_DIR, f"guard_incident_{_dt.datetime.now().strftime('%Y%m%d_%H%M%S')}.detail.log")
    try:
        with open(detail_path, "w", encoding="utf-8") as f:
            f.write(f"[{now_iso()}] {task_label}\n{detail_text}\n")
    except Exception as e:  # noqa: BLE001
        log(f"could not write alert detail file: {e}")
        detail_path = None
    toast = {"attempted": False, "ok": False}
    try:
        if os.path.exists(NOTIFY_SCRIPT_PATH) and detail_path:
            toast["attempted"] = True
            p = subprocess.run(
                ["powershell.exe", "-NoProfile", "-WindowStyle", "Hidden", "-File", NOTIFY_SCRIPT_PATH,
                 "-Status", "Failure", "-TaskName", task_label, "-LogPath", detail_path],
                timeout=25, capture_output=True, text=True,
            )
            toast["ok"] = (p.returncode == 0)
            toast["returncode"] = p.returncode
            log(f"BurntToast alert invoked (rc={p.returncode}).")
        else:
            log(f"toast alert skipped -- notify script not found at {NOTIFY_SCRIPT_PATH}")
    except Exception as e:  # noqa: BLE001
        log(f"toast alert failed: {e}")
        toast["error"] = str(e)
    # Console banner -- always, even if toast unavailable.
    _banner("MAILBOX GUARD TRIPPED", detail_text)
    return {"detail_path": detail_path, "toast": toast}


def _banner(title, body):
    line = "!" * 72
    print(f"\n{line}\n[{now_iso()}] {title}\n{line}\n{body}\n{line}\n", flush=True)


def write_incident_record(before, after, deltas, ignored, enforcement, sensitivity, task_name):
    ensure_runs_dir()
    ts = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    path = os.path.join(RUNS_DIR, f"guard_incident_{ts}.json")
    rec = {
        "schema": "mailbox_guard/incident/v1",
        "timestamp": now_iso(),
        "sensitivity": sensitivity,
        "task_name": task_name,
        "before_snapshot": {"label": before.get("label"), "timestamp": before.get("timestamp"),
                            "counts": before.get("counts")},
        "after_snapshot": {"label": after.get("label"), "timestamp": after.get("timestamp"),
                           "counts": after.get("counts")},
        "tripping_deltas": deltas,
        "ignored_deltas": ignored,
        "enforcement": enforcement,
        "teams_gap_note": "Teams writes are NOT covered by this sweep (no Outlook COM path).",
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(rec, f, indent=2, ensure_ascii=True, sort_keys=True)
    with open(TRIPPED_FLAG, "w", encoding="utf-8") as f:
        f.write(f"{now_iso()} guard tripped -- see {os.path.basename(path)}. "
                f"Clear with: python tools/codex_triage/mailbox_guard.py clear-flag\n")
    log(f"incident record written: {path}")
    log(f"sentinel written: {TRIPPED_FLAG} (blocks the wrapper until cleared)")
    return path


def run_guard(before, after, task_name, sensitivity, dry_run):
    deltas = diff_snapshots(before, after)
    trip, ignored = classify(deltas, sensitivity)
    if not trip:
        if ignored:
            log(f"{len(ignored)} user-plausible delta(s) seen and ignored under "
                f"'{sensitivity}' sensitivity; NO write-type delta. Guard CLEAN.")
        else:
            log("no mailbox delta detected. Guard CLEAN.")
        return 0
    summary_lines = [f"{len(trip)} mailbox delta(s) detected that the Codex run "
                     f"did not intend (sensitivity={sensitivity}):"]
    for d in trip:
        summary_lines.append(f"  - {d.get('type')} [{d.get('severity')}] "
                             f"field={d.get('field','-')} "
                             f"before={d.get('before')!r} after={d.get('after')!r} "
                             f"subj={d.get('subject_sha1','-')}")
    detail_text = "\n".join(summary_lines)
    enforcement = {"disable_task": _disable_task(task_name, dry_run)}
    enforcement["alert"] = _alert_kevin(detail_text, f"Work Inbox Codex Parallel -- MAILBOX GUARD TRIPPED")
    incident_path = write_incident_record(before, after, trip, ignored, enforcement,
                                          sensitivity, task_name)
    enforcement["incident_record"] = incident_path
    log("GUARD TRIPPED -- scheduled task disable attempted, Kevin alerted, incident recorded.")
    return 2


# --------------------------------------------------------------------------- #
#  prove -- end-to-end proof-of-fire with a synthetic COM injection
# --------------------------------------------------------------------------- #

def _find_disposable_message(ns):
    """Newest already-read Inbox message from a known automated sender with no
    category set.  Returns (entry_id, sender, subject_sha1) or None."""
    inbox = ns.GetDefaultFolder(OL_INBOX)
    items = inbox.Items
    try:
        items.Sort("[ReceivedTime]", True)
    except Exception:  # noqa: BLE001
        pass
    n = 0
    for it in items:
        n += 1
        if n > 400:
            break
        try:
            if _com_str(getattr(it, "Categories", "")).strip():
                continue
            if int(getattr(it, "FlagStatus", 0) or 0) != 0:
                continue
            if bool(getattr(it, "UnRead", False)):
                continue  # do not disturb unread state
            sender = _com_str(getattr(it, "SenderEmailAddress", "")).lower()
            dom = sender.split("@")[-1] if "@" in sender else ""
            if dom and any(dom == d or dom.endswith("." + d) for d in PROVE_SAFE_DOMAINS):
                return it.EntryID, sender, sha12(_com_str(getattr(it, "Subject", "")))
        except Exception:  # noqa: BLE001
            continue
    return None


def _read_categories_settled(ns, entry_id, settle_seconds):
    force_sync(ns)
    time.sleep(max(0, int(settle_seconds)))
    it = ns.GetItemFromID(entry_id)
    return _com_str(getattr(it, "Categories", ""))


def _residue_sweep(ns, marker):
    """Restrict('[Categories] <> \\'\\'') sweep -- count anything carrying our marker."""
    inbox = ns.GetDefaultFolder(OL_INBOX)
    try:
        restricted = inbox.Items.Restrict("[Categories] <> ''")
    except Exception as e:  # noqa: BLE001
        log(f"residue sweep Restrict failed ({e}) -- falling back to full scan.")
        restricted = inbox.Items
    hits = 0
    for it in restricted:
        try:
            if marker in _com_str(getattr(it, "Categories", "")):
                hits += 1
        except Exception:  # noqa: BLE001
            continue
    return hits


def cmd_prove(args):
    ensure_runs_dir()
    settle = args.settle_seconds
    marker = f"Drew-guard-selftest-{_dt.datetime.now().strftime('%Y%m%d-%H%M%S')}"
    dummy_task = "Drew Guard Selftest Dummy"
    result = {"marker": marker, "steps": [], "pass": False}

    def step(name, ok, extra=None):
        entry = {"step": name, "ok": bool(ok), "at": now_iso()}
        if extra is not None:
            entry["detail"] = extra
        result["steps"].append(entry)
        log(f"PROVE step {'OK ' if ok else 'FAIL'} -- {name}"
            + (f" :: {extra}" if extra is not None else ""))
        return ok

    ns = connect_mapi()

    # 1. Pick a disposable message.
    pick = _find_disposable_message(ns)
    if not pick:
        step("select disposable message", False, "no safe automated-sender message with empty categories found")
        _emit_prove_result(result)
        return 4
    entry_id, sender, subj_sha1 = pick
    step("select disposable message", True, {"sender": sender, "subject_sha1": subj_sha1})

    # 2. BEFORE snapshot.
    before = take_snapshot("before-selftest", settle, do_settle=False)
    before_path = os.path.join(RUNS_DIR, "selftest_before.json")
    with open(before_path, "w", encoding="utf-8") as f:
        json.dump(before, f, indent=1, sort_keys=True)
    step("before snapshot", entry_id in before.get("messages", {}),
         {"messages_tracked": before["counts"]["messages_tracked"]})

    # 3. Create a real dummy scheduled task so the disable path is genuinely exercised.
    made_task = False
    try:
        p = subprocess.run(
            ["schtasks", "/Create", "/TN", dummy_task, "/TR", "cmd /c exit",
             "/SC", "ONCE", "/ST", "23:59", "/F"],
            capture_output=True, text=True, timeout=30,
        )
        made_task = (p.returncode == 0)
        step("create dummy scheduled task", made_task, (p.stdout or p.stderr or "").strip())
    except Exception as e:  # noqa: BLE001
        step("create dummy scheduled task", False, str(e))

    # 4. Inject the synthetic category via COM.
    injected = False
    try:
        it = ns.GetItemFromID(entry_id)
        orig = _com_str(getattr(it, "Categories", ""))
        it.Categories = marker
        it.Save()
        injected = True
        step("inject synthetic category via COM", True, {"was": orig, "now": marker})
    except Exception as e:  # noqa: BLE001
        step("inject synthetic category via COM", False, str(e))
        _cleanup_prove(ns, entry_id, "", dummy_task, made_task, marker)
        _emit_prove_result(result)
        return 4

    # 5. AFTER snapshot (settled).
    after = take_snapshot("after-selftest", settle, do_settle=True)
    after_path = os.path.join(RUNS_DIR, "selftest_after.json")
    with open(after_path, "w", encoding="utf-8") as f:
        json.dump(after, f, indent=1, sort_keys=True)

    # 6. Diff must show exactly the one categories_changed on our EntryID.
    deltas = diff_snapshots(before, after)
    cat_deltas = [d for d in deltas if d.get("type") == "categories_changed"
                  and d.get("entry_id_sha1") == sha12(entry_id)]
    step("diff detects the synthetic write", len(cat_deltas) == 1,
         {"total_deltas": len(deltas), "matching": len(cat_deltas),
          "delta": cat_deltas[0] if cat_deltas else None})

    # 7. Enforce -- disable dummy task + alert + incident record.
    enforcement = {"disable_task": _disable_task(dummy_task, dry_run=False)}
    enforcement["alert"] = _alert_kevin("SELF-TEST -- synthetic injection, not a real incident.\n"
                                        + "\n".join(f"{d['type']} {d.get('before')!r} -> {d.get('after')!r}"
                                                    for d in cat_deltas),
                                        "Work Inbox Codex Parallel -- MAILBOX GUARD SELF-TEST")
    incident_path = write_incident_record(before, after, cat_deltas, [], enforcement,
                                          "strict", dummy_task)
    step("enforce: disable task action", enforcement["disable_task"].get("action") == "disabled",
         enforcement["disable_task"])
    step("enforce: alert detail file written",
         bool(enforcement["alert"].get("detail_path")
              and os.path.exists(enforcement["alert"]["detail_path"])))
    step("enforce: incident record + sentinel written",
         os.path.exists(incident_path) and os.path.exists(TRIPPED_FLAG))

    # 8. Verify the dummy task really shows Disabled.
    disabled_ok = False
    try:
        q = subprocess.run(["schtasks", "/Query", "/TN", dummy_task, "/FO", "LIST", "/V"],
                           capture_output=True, text=True, timeout=30)
        disabled_ok = "Disabled" in (q.stdout or "")
        step("verify dummy task state == Disabled", disabled_ok,
             next((ln.strip() for ln in (q.stdout or "").splitlines()
                   if "Scheduled Task State" in ln or "Status" in ln), ""))
    except Exception as e:  # noqa: BLE001
        step("verify dummy task state == Disabled", False, str(e))

    # 9. Remediate the synthetic change + confirm clean.
    remediated_val = _remediate_prove(ns, entry_id, orig, settle)
    step("remediate synthetic category (COM, settled re-read)", remediated_val == orig,
         {"expected": orig, "read_back": remediated_val})
    residue = _residue_sweep(ns, "Drew-guard-selftest")
    step("Restrict sweep: 0 self-test residue", residue == 0, {"residue_hits": residue})

    # 10. Tidy up test scaffolding.
    _cleanup_prove(ns, entry_id, orig, dummy_task, made_task, marker, delete_task=True)
    if not args.keep_artifacts:
        for p in (before_path, after_path, incident_path,
                  enforcement["alert"].get("detail_path"), TRIPPED_FLAG):
            try:
                if p and os.path.exists(p):
                    os.remove(p)
            except Exception:  # noqa: BLE001
                pass
        log("self-test artifacts removed (use --keep-artifacts to retain).")
    else:
        log("self-test artifacts kept under data/codex_runs/.")

    result["pass"] = all(s["ok"] for s in result["steps"])
    _emit_prove_result(result)
    return 0 if result["pass"] else 2


def _remediate_prove(ns, entry_id, orig, settle):
    try:
        it = ns.GetItemFromID(entry_id)
        it.Categories = orig
        it.Save()
    except Exception as e:  # noqa: BLE001
        log(f"remediation write failed: {e}")
        return "<remediation-write-failed>"
    return _read_categories_settled(ns, entry_id, settle)


def _cleanup_prove(ns, entry_id, orig, dummy_task, made_task, marker, delete_task=False):
    try:
        it = ns.GetItemFromID(entry_id)
        if marker in _com_str(getattr(it, "Categories", "")):
            it.Categories = orig
            it.Save()
            log("cleanup: reverted lingering self-test category.")
    except Exception as e:  # noqa: BLE001
        log(f"cleanup: category revert check failed: {e}")
    if made_task and delete_task:
        try:
            subprocess.run(["schtasks", "/Delete", "/TN", dummy_task, "/F"],
                           capture_output=True, text=True, timeout=30)
            log(f"cleanup: dummy task '{dummy_task}' deleted.")
        except Exception as e:  # noqa: BLE001
            log(f"cleanup: could not delete dummy task: {e}")


def _emit_prove_result(result):
    ensure_runs_dir()
    path = os.path.join(RUNS_DIR, f"selftest_result_{_dt.datetime.now().strftime('%Y%m%d_%H%M%S')}.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(result, f, indent=2, sort_keys=True)
    verdict = "PASS" if result.get("pass") else "FAIL"
    _banner(f"MAILBOX GUARD SELF-TEST: {verdict}",
            "\n".join(f"  [{'x' if s['ok'] else ' '}] {s['step']}" for s in result["steps"]))
    log(f"self-test result written: {path}")


# --------------------------------------------------------------------------- #
#  argparse
# --------------------------------------------------------------------------- #

def main(argv=None):
    ap = argparse.ArgumentParser(description="Outlook COM delta-sweep kill-switch for the Codex parallel run.")
    sub = ap.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("snapshot", help="capture a mailbox state snapshot")
    s.add_argument("--out", required=True)
    s.add_argument("--label", default="snapshot")
    s.add_argument("--settle-seconds", type=int, default=DEFAULT_SETTLE_SECONDS)
    s.add_argument("--no-settle", action="store_true",
                   help="skip SyncObjects + settle sleep (use for the BEFORE snapshot)")

    d = sub.add_parser("diff", help="diff two snapshots, print deltas, no enforcement")
    d.add_argument("--before", required=True)
    d.add_argument("--after", required=True)
    d.add_argument("--out")

    g = sub.add_parser("guard", help="diff + enforce: disable task, alert, incident record on any delta")
    g.add_argument("--before", required=True)
    g.add_argument("--after", required=True)
    g.add_argument("--task", default="Work Inbox Codex Parallel")
    g.add_argument("--sensitivity", choices=["strict", "writes-only"], default="strict")
    g.add_argument("--dry-run", action="store_true")

    p = sub.add_parser("prove", help="end-to-end proof-of-fire with a synthetic COM injection")
    p.add_argument("--settle-seconds", type=int, default=DEFAULT_SETTLE_SECONDS)
    p.add_argument("--keep-artifacts", action="store_true")

    sub.add_parser("clear-flag", help="clear the GUARD_TRIPPED sentinel after review")

    args = ap.parse_args(argv)

    if args.cmd == "snapshot":
        snap = take_snapshot(args.label, args.settle_seconds, do_settle=not args.no_settle)
        with open(args.out, "w", encoding="utf-8") as f:
            json.dump(snap, f, indent=1, ensure_ascii=True, sort_keys=True)
        log(f"snapshot '{args.label}' written to {args.out} "
            f"({snap['counts']['messages_tracked']} msgs, "
            f"{snap['counts']['calendar_events_tracked']} events, "
            f"sent={snap['counts']['sent_items']}, drafts={snap['counts']['drafts']}).")
        return 0

    if args.cmd == "diff":
        before = json.load(open(args.before, encoding="utf-8"))
        after = json.load(open(args.after, encoding="utf-8"))
        deltas = diff_snapshots(before, after)
        if args.out:
            with open(args.out, "w", encoding="utf-8") as f:
                json.dump(deltas, f, indent=2, sort_keys=True)
        log(f"{len(deltas)} delta(s).")
        for dl in deltas:
            print(json.dumps(dl, sort_keys=True))
        return 2 if deltas else 0

    if args.cmd == "guard":
        before = json.load(open(args.before, encoding="utf-8"))
        after = json.load(open(args.after, encoding="utf-8"))
        return run_guard(before, after, args.task, args.sensitivity, args.dry_run)

    if args.cmd == "prove":
        return cmd_prove(args)

    if args.cmd == "clear-flag":
        if os.path.exists(TRIPPED_FLAG):
            os.remove(TRIPPED_FLAG)
            log("GUARD_TRIPPED sentinel cleared.")
        else:
            log("no GUARD_TRIPPED sentinel present.")
        return 0

    return 1


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        sys.exit(130)
