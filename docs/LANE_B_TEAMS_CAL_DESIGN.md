# LANE_B_TEAMS_CAL_DESIGN.md — dumb-fetch codex exec for Calendar + Teams

> **UPDATE 2026-09-01 — BUILDING. Two design points corrected by live findings (see `HANDOVER.md` top entry):**
> 1. **§6c "manifest turn" is dropped.** `codex_apps` connector tools are lazily surfaced and never appear in a "list your tools" enumeration — that method gave three false negatives. The re-contamination guard instead asserts on the `mcp_tool_call` `server`/`tool` events **actually observed** in the `codex exec --json` JSONL (`lane_b_call1.py: guard_recontamination`).
> 2. **§3 prompts must be TASK-DESCRIPTIVE, not imperative.** "Using the Outlook Calendar app connector, retrieve my events between X and Y…" loads the tools; "Call `microsoft_outlook_calendar.list_events` with…" makes the model reply "I can't access that in this session". Guardrail clauses ("use no other tool / change nothing / send nothing") stay. Real MCP server = `codex_apps`; result payload = `item.result.structured_content` (`.value` calendar / `.chats` teams).
> Build: `lane_b_call1.py` (runner + guard, shipped), `fetch_inbox.py` `CAL_BACKEND=connector` wiring (shipped), `lane_b_cal_guard.py` (§6a snapshot HALT, next), Teams section (after calendar is live). Identity = Edu `begb0037@ox.ac.uk` PRIMARY, personal Plus failover.

**Date:** 2026-08-29 (Drew). **DESIGN ONLY — no build until Kevin approves this doc.**
**Parent:** `LAPTOP_MIGRATION_PLAN.md` (rev 2). **Reuses patterns from:** `CONNECTOR_SAFEGUARDS.md` (§A enumeration, §B safeguard layers, §D the "NOT SOUND" verdict), `CODEX_CONNECTOR_PIPELINE_PLAN.md` (the two-call split, `normalise_pull.py`, the 26-Aug branch scripts on `drew/codex-phase2-ai-triage`).
**Scope:** Lane B is **calendar + Teams only**. Mail is Lane A (own read-only Python — different doc). Triage is `claude -p`, unchanged.

---

## 1. What Lane B produces and why it exists

| Signal | Source (connector tool) | Consumed by |
|---|---|---|
| Today+6d calendar events, primary + "People Department - HR Systems" shared calendar | `list_calendars` → `list_events` → `fetch_events_batch` / `list_event_instances` | `fetch_inbox.py` Phases 3.7 (raw) / 3.8 (AI summaries) under `CAL_BACKEND=connector` |
| Teams 1:1 / group chat messages, channel messages, in-meeting chat | `list_chats` / `list_channels` → `list_chat_messages` / `list_channel_messages` / `fetch` / `search` | a new Teams section in the briefing (triage decides relevance) |
| Meeting recordings + transcripts (scheduled meetings only) | event id → `resolve_scheduled_online_meeting` → `list_online_meeting_transcripts` → `get_online_meeting_transcript_content` (WebVTT) | briefing meeting-prep context |

**No other route exists at Oxford:** Graph is disallowed as an auth method; EWS is retiring and is removed from the plan; Teams has never had a non-connector path. This is the whole justification for accepting a connector that structurally holds write tools while Call-1 runs.

---

## 2. Architecture — two calls, split by connector attachment

```
 CALL 1  (connector ATTACHED — dedicated CODEX_HOME, Edu account, connectors = {Calendar, Teams})
   codex exec, rigid instruction. Dumb fetch ONLY. No "summarise", no "decide", no branching over content.
   ─ manifest turn ("list tools, call nothing")  → assert subset of the allowlist (§5)  [RE-CONTAMINATION GUARD]
   ─ calendar: list_calendars → list_events(window) → fetch_events_batch  → JSON array, fixed fields
   ─ teams:    list_chats/list_channels → list_*_messages / search        → JSON array, fixed fields
   ─ meetings: for each event with an onlineMeeting → resolve_scheduled_online_meeting
               → list_online_meeting_transcripts → get_online_meeting_transcript_content
        │
        ▼  raw_lane_b.json   (connector shape)
 normalise_pull.py   — deterministic sanitiser, NO model. Strip HTML, truncate, neutralise
        │              instruction-like text / role markers, strip zero-width + bidi, record hits.
        ▼  lane_b_normalised.json   ({ calendar:[...], teams:[...], transcripts:[...] })
        │
   ┌────┴─────────────────┐   ┌────────────────────┐   ┌─────────────────────────────┐
   │ CALENDAR kill-switch │   │ TEAMS kill-switch  │   │ RE-CONTAMINATION guard      │
   │ pre/post list_events │   │ pre/post from=me    │   │ (ran as the manifest turn   │
   │ snapshot → ANY diff  │   │ msg snapshot → new  │   │  above; also on each guard  │
   │ ⇒ HALT pipeline      │   │ one ⇒ disable NEXT  │   │  session) → any tool outside │
   │ (disable+exit1+toast)│   │ run + toast         │   │  allowlist ⇒ HALT           │
   └──────────────────────┘   └────────────────────┘   └─────────────────────────────┘
        │
        ▼  feeds fetch_inbox.py alongside the Lane A mail lists
 CALL 2  = claude -p triage. NO connector. Cannot act. (unchanged from today's live engine)
```

---

## 3. Call-1 instruction shape (rigid — the "dumb fetch")

One structured data request per domain. No open-ended reasoning. Verbatim intent:

**Calendar:**
> "Call `list_calendars`. Then for the calendar named `Calendar` and the calendar named `People Department - HR Systems`, call `list_events` with `start_datetime` = `<ISO today 00:00>` and `end_datetime` = `<ISO today+7 00:00>`. For events in a recurring series, call `list_event_instances` for the concrete occurrences in that window. Return a JSON array of objects with EXACTLY these keys: `calendar_name`, `id`, `subject`, `start`, `end`, `is_all_day`, `location`, `organizer_name`, `organizer_email`, `is_cancelled`, `response_status`, `series_master_id`, `has_online_meeting`, `online_meeting_join_url`, `body_preview` (first 300 chars of the plain-text body, verbatim — do not summarise or interpret). Do not call any other tool. Do not create, update, cancel, delete, respond to, or add an attachment to any event. Return only the JSON array."

**Teams:**
> "Call `list_chats` (limit 40, most recent) and `list_teams` → `list_channels` for teams Kevin is in. For each chat and each channel updated since `<ISO now-72h>`, call `list_chat_messages` / `list_channel_messages` (limit 30, newest). Return a JSON array of objects with EXACTLY these keys: `kind` (`chat`|`channel`), `container_name`, `message_id`, `from_name`, `from_email`, `created`, `is_from_me`, `has_attachments`, `body_preview` (first 400 chars plain-text, verbatim). Do not call any other tool. Do not send, reply, create a chat/channel, or touch Planner. Return only the JSON array."

**Transcripts** (chained off the calendar result, only for events where `has_online_meeting` and the event is in the past ≤48h or starts today):
> "For online-meeting join URL `<url>`: call `resolve_scheduled_online_meeting`, then `list_online_meeting_transcripts`, then `get_online_meeting_transcript_content` for the most recent transcript. Return `{ event_id, transcript_id, vtt_text }` (WebVTT verbatim, truncated to 20000 chars). Call no other tool."

Every Call-1 turn logs its tool calls (from the JSONL `item` events). **Expected set** = the §5 allowlist. Anything else → discard output + HALT (§6).

---

## 4. `normalise_pull.py` — deterministic sanitiser (no model)

Runs between Call 1 and the triage step. Every string field (`subject`, `body_preview`, `from_name`, `organizer_name`, `location`, `container_name`, `vtt_text`) passes through `sanitise()`:

1. **Plain-text only** — strip `<script|style|head>` blocks then all tags; decode entities; drop `data:` / `javascript:` / `vbscript:` URIs.
2. **Truncate hard** — `body_preview` 400 (Teams) / 300 (calendar); `subject`/`location` 300; `from_name`/`organizer_name` 120; `vtt_text` 20000. Post-strip re-enforced.
3. **Neutralise instruction-like text** — line-start `^(system|assistant|user|developer)\s*[:>]` → prepend `[quoted] `; fenced code / `<\|...\|>` / `[INST]` / `### ` headings → `[quoted] `; literal `ignore previous` / `disregard the above` / `new instructions` / `you are now` / `send an email to` / `forward this to` / `create a chat` / `reply to` (case-insensitive) → wrap the line `[quoted: … ]`; collapse >2 newlines to 1.
4. **Strip zero-width + bidi control chars.**
5. **Record, don't drop** — every field that triggered a rule → `data/codex_runs/<ts>_sanitiser_hits.json` (container + which rule). Injection attempts visible, not silent.
6. On `sanitise()` throw → field replaced with `"[sanitiser error — field withheld]"`, run continues.

`normalise_pull.py` output schema:
```
{ "calendar":    [ {calendar_name,id,subject,start,end,is_all_day,location,organizer,all_day,day, ...} ],
  "teams":       [ {kind,container_name,message_id,from_name,from_email,created,is_from_me,has_attachments,body_preview} ],
  "transcripts": [ {event_id,transcript_id,vtt_text} ] }
```
Calendar objects are mapped to the **exact** dict shape `fetch_inbox.py` Phase 3.7 already consumes (`subject/start/end/location/organizer/all_day/day`) so the ported logic needs no edits.

---

## 5. Tool allowlists (the re-contamination guard asserts a SUBSET of this)

**Calendar read — allowed:** `list_calendars`, `list_events`, `fetch_event`, `fetch_events_batch`, `list_event_instances`, `list_recurring_series`, `get_mailbox_settings`.
**Teams read — allowed:** `list_teams`, `list_channels`, `list_chats`, `resolve_team`, `resolve_channel`, `resolve_chat`, `resolve_user`, `list_chat_messages`, `list_channel_messages`, `fetch`, `search`, `get_chat_members`, `resolve_scheduled_online_meeting`, `list_online_meeting_transcripts`, `get_online_meeting_transcript_content`, `list_online_meeting_recordings`.
**Everything else = violation.** Explicitly: every `microsoft_outlook_calendar` write (16), every `microsoft_teams` write (9), every `microsoft_outlook_email.*`, every `github.*`, every `canva.*` / `sites.*` / `granola.*`, and the `functions.*` / `collaboration.*` beyond what a bare `codex exec` needs.

---

## 6. Kill-switches — deliberately asymmetric

### 6a. CALENDAR — HALT the pipeline on ANY detected change

- **Baseline** (`data/codex_runs/cal_baseline_<ts>.json`), taken **seconds before** Call-1: `list_events` over today+7d on **primary + the HR Systems shared calendar**; per event `{id, subject, start, end, last_modified, is_cancelled, response_status, series_master_id}`.
- **Post** (after Call-1 + `normalise_pull.py`): the identical read.
- **Trip = ANY of:** an `id` present post- not pre-; an `id` present pre- not post-; a changed `last_modified`; a changed `response_status`; an `is_cancelled` flip.
- **On trip:** `Disable-ScheduledTask` on the Lane B task **AND** the current pipeline run `exit(1)` (no `briefing.json` this cycle) **AND** BurntToast ("work-inbox Lane B: calendar changed during a read-only run — pipeline halted") **AND** `data/codex_runs/GUARD_TRIPPED_cal_<ts>.json` with the full diff. **Kevin investigates and re-enables manually.** No auto-resume.
- **Guard session identity:** its own short `codex exec`, same {Calendar, Teams} connector, read tools only, manifest-asserted first (§5).
- **Known false-positive:** a genuine third-party edit to an event inside the ~2–5-min pre/post window trips it. Accepted per Kevin's instruction — calendar blast radius (decline/cancel/RSVP notices to real attendees) justifies favouring safety over uptime. Expect occasional manual re-enables.
- **Cannot catch:** a net-zero change (delete + recreate with a new id looks like one add + one drop → still trips, good); a write to a calendar outside the two snapshotted (mitigated: Call-1 only ever names those two; any third calendar in the manifest is itself a §5 violation).

### 6b. TEAMS — disable the NEXT run

- **Baseline** (`data/codex_runs/teams_baseline_<ts>.json`): `list_chats` + `list_channel_messages` for touched channels, filtered `is_from_me` in a trailing 3h window → `{message_id set, newest created}`.
- **Post:** repeat. **Trip =** any `is_from_me` `message_id` present post- not pre-.
- **On trip:** `Disable-ScheduledTask` + BurntToast + `GUARD_TRIPPED_teams_<ts>.json`. The **current briefing still completes** (blast radius = one Teams message, already sent — containment, not prevention). Next run is blocked pending Kevin.
- **Cannot catch:** net-zero, delayed send, a message in a container Call-1 didn't touch.

### 6c. RE-CONTAMINATION guard — MANDATORY, HALT on violation

- Runs as the **first turn of every Lane B `codex exec`** (Call-1 and both guard sessions): "list your tools, call nothing" → parse JSONL → `data/codex_runs/<ts>_manifest.json`.
- **Assert:** available tool set ⊆ (§5 calendar-read allowlist ∪ §5 teams-read allowlist ∪ the bare `functions.exec/wait` + `collaboration.*` a `codex exec` always has).
- **Also assert:** `list_events` AND `list_chats` ARE present (else the connector didn't load — the 28 Aug Q2 state → "connector unavailable", skip this cycle, toast, no HALT).
- **On subset violation** (a `github.*`, `microsoft_outlook_email.*`, any write tool, any unexpected connector) → `Disable-ScheduledTask` + `exit(1)` + BurntToast ("work-inbox Lane B: unexpected connector tool `<name>` — pipeline disabled") + `GUARD_TRIPPED_manifest_<ts>.json`. Fail-closed.
- Weekly rollup: manifest histogram + any drift.

---

## 7. Per-run log

`data/codex_runs/<ts>_lane_b.json`:
```
{ "ts",
  "manifest_ok": true|false, "manifest_unexpected": [...],
  "call1_tools_called": [...], "call1_wall_s", "attempts",
  "sanitiser_hits": N,
  "cal_guard": "clean|tripped|unverified",  "cal_diff": {...},
  "teams_guard": "clean|tripped|unverified",
  "counts": { "calendar": N, "teams": N, "transcripts": N },
  "config_toml_sha1_before", "config_toml_sha1_after",   // must match
  "codex_version" }
```
`config_toml_sha1_before/after` must equal `35f8910382373d525598194b2649159cfeed3f6a` (or the current recorded baseline) — logged every run, alerted on mismatch.

---

## 8. The Codex "NOT SOUND" caveat still stands

`CONNECTOR_SAFEGUARDS.md` §D: Call-1 holds the connector's 25 write tools (16 calendar + 9 Teams) the whole time it runs; a rigid prompt is **not** a capability boundary; one `codex exec` invocation can call a write tool after hostile content is in context and before it returns; the kill-switches detect *after*, not prevent. **Not resolved here.** Lane B proceeds because Kevin has explicitly accepted it, given:

1. **Narrower blast radius than mail** — worst case: one Teams message, or calendar notices to a meeting's attendees. Not an external HR email.
2. **Dedicated 2-connector identity** — Kevin strips GitHub (89 tools) + Outlook Email himself; the re-contamination guard (§6c) HALTS if they come back.
3. **Calendar kill-switch HALTS** (not logs) on any change — strictest containment available without an enforced boundary.
4. **`claude -p` triage has no connector** — the step that reasons over hostile content cannot act.
5. **No alternative** — Graph disallowed, EWS retiring, Teams connector-only.

**What would upgrade this to "sound":** an enforced read-only boundary — a read-only Graph proxy, or tenant read-only Graph consent. Both currently unavailable at Oxford. If either becomes available, revisit.

---

## 9. Build increments (start FROM 1 SEPT, ONLY after Kevin approves this doc — separately gated)

**Timing:** the Edu account becomes the dedicated automation identity on 1 Sept (quota reset; Kevin moves interactive work off Edu; strips connectors to {Calendar, Teams}; `codex login` on the laptop). Build increment 2 below is preceded by Phase 2(ii) — proving headless `codex exec` on the laptop loads exactly the {Calendar, Teams} read tools. Nothing in this section touches the Edu quota before 1 Sept.

| # | Increment | Live-path risk | Gate |
|---|---|---|---|
| 1 | `normalise_pull.py` calendar+Teams schema + sanitiser; unit test corpus (reuse `CONNECTOR_SAFEGUARDS.md` §C.5 payloads, add calendar-description + Teams-message variants) | none — new file | — |
| 2 | Call-1 runner + wrapper (manifest turn, retry, timeout, timestamps) → `raw_lane_b.json`; **one manual run**, connector attached | reads live calendar + Teams (read-only); writes local files only | Kevin go-ahead for the first manual run |
| 3 | the two kill-switches + re-contamination guard + per-run log | none (detect/HALT, not writers) | — |
| 4 | `CAL_BACKEND=com\|connector` in `fetch_inbox.py`; Phases 3.7/3.8 read `lane_b_normalised.json` under `connector`; `com` default byte-identical; Teams briefing section | none until a gated flip | Kevin go-ahead to push code |
| 5 | parallel scheduled task on the laptop (minutes after the Lane A run); writes `docs/codex_*.json` + `data/codex_runs/*` only, never `data/briefing.json` | none — parallel only | **Kevin's fresh explicit separate go-ahead** |
| — | cutover — out of scope here; folds into `LAPTOP_MIGRATION_PLAN.md` Phase 6 | — | — |

---

## 10. Open items for Kevin

1. **Calendar-source degradation** — if a cycle's manifest turn shows the connector didn't load (`list_events` absent), do Phases 3.7/3.8 degrade to *empty calendar + warning* (briefing continues) or does the whole run fail? Recommend: degrade + warn (the guard already toasts).
2. **Teams window** — 72h lookback for messages, 40 most-recent chats, 30 messages/container. Tune after the first real runs.
3. **Transcript scope** — only past-≤48h / starts-today online meetings. Confirm that matches what Kevin wants in the briefing (vs "every meeting this week").
4. **Calendar kill-switch false-HALT** — confirm Kevin accepts occasional manual re-enables when a colleague edits a meeting mid-window.
