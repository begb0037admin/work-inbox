# Codex Connector Pipeline — the settled plan

**Status:** 2026-08-29 (Drew). **Direction is decided** — this is the execution plan, not a scoping doc.
**Supersedes:** `OPTION3_BUILD_PLAN.md` (connector-*free* + COM pull — the opposite shape; now historical) and the connector half of `PHASE2_BRIEF.md`.
**Prime design constraint (Kevin, 29 Aug):** Kevin does **not** normally run classic Outlook. The whole point of the connector route is that **no desktop app is open — everything is cloud-side.** The end state has **zero dependency on classic Outlook**, including for calendar and including the kill-switch. Anything in this plan that would need Outlook running is either reworked to a connector/Graph read or explicitly flagged as a blocker.
**Reads alongside:** `EMAIL_AUTOMATION_SECURITY_MITIGATIONS.md` (the layered write-gate model — the security spine of this plan), `CODEX_CONNECTOR_MIGRATION_RESEARCH.md` Section 9 (history), `codex_phase2_run_20260826/` on branch `drew/codex-phase2-ai-triage` (the 26 Aug dry-run machinery being reused).

---

## 1. Decision & why

**Kevin decided (28 Aug): use the ChatGPT M365 connector for the mail/calendar/Teams pull, and move the AI triage onto Codex.**

- **It is a funding choice, not a compliance one.** Claude is allowed at Oxford. But `claude -p` (the live triage engine since 27 Aug) runs on Kevin's **personal** Claude subscription. Oxford pays for **ChatGPT Edu**, whose entitlement **covers Codex CLI / programmatic use** (confirmed). Moving the pipeline to Codex puts the running cost on Oxford's paid entitlement instead of Kevin's pocket. (This is *not* the old "£36/mo Anthropic API" saving — that ended at the 27 Aug `claude -p` cutover. The saving now is Kevin's personal Claude usage.)
- Kevin also has a personal **ChatGPT Plus** account (a fallback identity option).
- The connector is the **only** route that also reaches **Microsoft Teams** — lots of valuable Teams messages are currently invisible to work-inbox — and shared calendar, in one place.
- The connector's write tools (`send_email`, `reply_to_email`, `forward_email`, …) **cannot be reliably denied** in headless `codex exec` (every local + account-side control failed 26–27 Aug; codex-cli 0.149.1 has no `--allowed-tools`). So the write risk is handled by the **layered mitigation model**, not by trying to gate the tools.

**`claude -p` stays live** as the fallback until this Codex path is proven at parity. Stopping it = no briefings.

---

## 2. Architecture — two Codex calls, split by connector attachment

```
   ┌─ CALL 1: connector ATTACHED ──────────────┐
   │  codex exec, Outlook Email + Calendar     │   ← dumb fetch only
   │  (+ Teams, later). Rigid instruction:     │
   │  "list messages / events / chats in this  │
   │   date range, return exactly these        │
   │   fields." No reasoning over body text.   │
   └──────────────┬───────────────────────────┘
                  │  raw_pull.json  (connector shape)
                  ▼
        normalise_pull.py   ← schema-normalise + Layer-2 sanitise
                  │           (strip HTML, truncate bodies, neutralise
                  │            injection-prone patterns) BEFORE anything reasons
                  ▼
        pull_normalised.json
                  │
        categorise_and_stage.py   ← VERBATIM port of fetch_inbox.py
                  │                 categorise()/badge_for()/make_card()
                  │                 + Phase 3.3/3.3b/3.9 deterministic Python
        build_granola_context.py  ← VERBATIM port of Phase 3.7b
        build_call2_brief.py      ← assembles the brief; verbatim system prompts
                  ▼
   ┌─ CALL 2: NO connector attached ───────────┐
   │  codex exec, ZERO microsoft_* tools in    │   ← pure language/reasoning
   │  the session. Five judgement phases over  │      over sanitised text only.
   │  the staged text. Cannot send/write —     │      No write path exists here.
   │  there is no tool to do it with.          │
   └──────────────┬───────────────────────────┘
                  │  call2_judgement_output.json
                  ▼
        output writers → docs/codex_briefing.json
                         docs/codex_suggestions.json
                         docs/codex_triage_ledger.json   (separate dedup namespace)
                  │
                  ▼  (parallel-run: DIFF against the live data/briefing.json;
                      never feeds the dashboard, never pushed, no CC/ledger write)
```

**The security spine (from `EMAIL_AUTOMATION_SECURITY_MITIGATIONS.md`):**

| Layer | Where in this pipeline |
|---|---|
| 1 — connector = dumb fetch, no reasoning over hostile text | Call 1's rigid instruction |
| 2 — sanitise / truncate bodies before any AI step | `normalise_pull.py` |
| 3 — connector/account read-only setting | belt-and-braces on Call 1's account; re-verify against current CLI |
| 4 — draft-only system instruction | Call 2 brief (already says "do not attempt any write action of any kind") |
| 5 — human review of every draft | existing Lauren drafting loop — nothing auto-sent, ever |
| 6 — post-run **connector/Graph** Sent-folder delta-check | kill-switch backstop — **reworked to be COM-free** (see §6a). Snapshots the Sent folder (message-ids + count) via a read-only connector call before the run and again after; any new item in the window → disable the scheduled task + BurntToast. Same blind spots as the old COM guard (Teams sends, on-behalf/shared Sent, net-zero count), same "detect ~minutes after delivery, not prevent" limit — it is a backstop, never a gate. |

---

## 3. What is reused (mostly built already, 26 Aug dry run)

On branch `drew/codex-phase2-ai-triage`, `tools/codex_triage/`:

| File | Role | Change needed |
|---|---|---|
| `prompts/brief_inbox.txt` / `brief_sent.txt` / `brief_calendar.txt` | Call 1 dumb-fetch instructions | tighten to Layer-1 rigidity; add explicit "return field X" lists; add a Teams variant later |
| `categorise_and_stage.py` | deterministic split + stages all Call-2 inputs | update field names to the normalised schema (§4); it currently reads connector-era `call_inbox_result.json` |
| `build_granola_context.py` | Phase 3.7b Granola port | point at `pull_normalised.json` instead of `real_briefing.json` |
| `build_call2_brief.py` | assembles Call 2 brief, verbatim system prompts | field-name refresh only; the 5 verbatim system prompts stay byte-identical to `fetch_inbox.py` |
| (main) `tools/codex_triage/mailbox_guard.py` | Layer-6 kill-switch, **COM** delta-sweep, proof-fired 26 Aug | **DO NOT wire as-is** — its COM sweep reintroduces the "classic Outlook must be running" dependency this route exists to remove. Rework to a connector/Graph Sent-folder read (§6a). Keep the disable+toast mechanism; replace the data source. |
| `codex_phase2_run_20260826/PARALLEL_RUN_QUALITY_GATE_DESIGN.md` | the false-demotion + fidelity gate | bring onto the new branch; **drop** the "missing importance" section — the connector `fetch_message` full-detail call returns `importance` (gate design §B Step 1), and COM is no longer the pull |

Verbatim system prompts (do not paraphrase): `EMAIL_SUMMARY_SYSTEM`, `TRIAGE_SYSTEM`, `SUMMARY_SYSTEM`, `CAL_SUM_SYSTEM`, Phase-2 context `SYSTEM` — copied from `fetch_inbox.py`, only "Anthropic API" → "you".

**Genuinely new:** `normalise_pull.py` (schema + Layer-2 sanitiser), the Call-1 runner/wrapper, output writers + `codex_triage_ledger.json`, the parallel diff harness, and (last, separately gated) the parallel scheduled task.

---

## 4. Normalised pull schema (`pull_normalised.json`)

One shape, backend-agnostic, so `categorise_and_stage.py` doesn't care whether the source was the connector or a COM dump used for testing. Field names match `fetch_inbox.py`'s current Phase-1 dict so the ported `categorise()` needs no edits:

```
{ "inbox": [ { "subject", "from", "from_email", "received",      # "YYYY-MM-DD HH:MM:SS"
               "is_read", "has_attachments", "importance",        # 0/1/2
               "message_id",                                      # internet Message-ID = the cross-backend key
               "kevin_is_primary_recipient",
               "body_preview" (sanitised, <=150 chars),
               "source_folder" (subfolder items) } ... ],
  "sent":  [ { "subject", "to", "sent", "body_preview" (sanitised, <=100), "message_id" } ... ],
  "calendar": [ { "subject", "start", "end", "location", "organizer", "all_day", "day" } ... ],
  "teams": [ ... ]  # PHASE 2 of the build only — see §6
}
```

`normalise_pull.py` accepts either (a) the connector Call-1 output, or (b) a `WI_MAIL_PARALLEL` COM dump (`data/parallel/com_inbox_raw.json` + `com_sent_raw.json`). **(b) is a developer convenience only** — it lets the deterministic staging (categorise/badge/stage) be smoke-tested without a live connector. It is **not** the parity method (§6b): Kevin does not run classic Outlook, so parity must not depend on him producing COM captures.

**Layer-2 sanitisation in `normalise_pull.py`:** strip HTML tags/entities; collapse whitespace; truncate to the preview length; drop/neutralise `data:`/`javascript:` URIs; flatten anything that looks like an instruction block or role marker in body text (`^\s*(system|assistant|user)\s*:`, fenced blocks, `<\|...\|>`); never pass a raw untruncated body downstream.

---

## 5. Identity & connector attachment (Kevin-blocked — Build Step 1)

The connector-loading behaviour is **currently unknown and must be established with Kevin** (no `codex login` without telling him):

- **Q2 finding, 28 Aug:** a plain headless `codex exec -s read-only` on this machine right now loads **zero** connector tools (only `functions.exec/wait` + `collaboration.*`) — a change from 26–27 Aug when `set_message_categories` was callable. Cause undetermined (connector auth expired / ChatGPT app-server bridge `cua_repl` disabled / 27 Aug residual state).
- So Build Step 1 is: **determine how to get the Outlook Email + Calendar read tools into a `codex exec` session** (which account — Oxford ChatGPT Edu vs personal Plus; which `CODEX_HOME`; whatever login/bridge is needed), **and** confirm a second invocation/`CODEX_HOME` loads **none**. Record `~/.codex/config.toml` sha1 before/after. Report the two tool lists. Nothing downstream runs live until this passes.
- Recommended split (to be confirmed): a dedicated automation `CODEX_HOME` (e.g. `C:\CodexAutomation\.codex`) signed into whichever account carries the Oxford Microsoft connectors, used **only** for Call 1; Call 2 runs with `--disable apps` (proven 28 Aug to strip the entire Apps surface) or from a connector-free `CODEX_HOME`, so it structurally cannot touch the mailbox.

---

## 6. Build increments (report after each; do not run to completion silently)

| # | Increment | Live-path risk | Gate |
|---|---|---|---|
| 1 | `normalise_pull.py` (schema + Layer-2 sanitiser) + refreshed `categorise_and_stage.py`, both reading the normalised schema; smoke-tested against a `WI_MAIL_PARALLEL` COM dump (dev convenience only, §6b) | none — new files, nothing invokes them | — |
| 2 | refreshed `build_granola_context.py` + `build_call2_brief.py` against the normalised schema; assemble a Call-2 brief from a dump-derived stage and eyeball it | none | — |
| 3 | **Kevin-blocked:** Build Step 1 (§5) — connector tool-list verification, both directions. Needs Kevin's account choice + any login | none until it passes | Kevin |
| 4 | Call-1 runner + wrapper (warm-up, retry, timeout, timestamped) → produces `raw_pull.json` (mail + calendar); run once by hand | reads the live mailbox/calendar via connector (read-only, dumb-fetch); writes local files only | Kevin go-ahead for the first manual run |
| 5 | output writers + `codex_triage_ledger.json` + the parity harness (§6b, against `data/briefing.json` GitHub history — **no manual COM capture**); one full manual dry-run end-to-end | none live — `docs/codex_*.json` only, never `data/`, never pushed | Kevin review of the dry-run diff |
| 6 | **COM-free kill-switch** (§6a) — connector Sent-folder delta-check + the existing disable+toast mechanism | none (detect-and-disable, not a writer) | — |
| 7 | **Teams ingestion** — Teams connector on Call 1, `teams` branch in `normalise_pull.py`, Teams card/section. **Deliberately after** mail+calendar parity is proven | none live (parallel) | Kevin — separate phase |
| 8 | parallel scheduled task (new task, minutes after each live `fetch_inbox.py` run; writes only `docs/codex_*.json` + `data/codex_runs/*`) | none — parallel only, no cutover | **Kevin's fresh explicit separate go-ahead** |
| — | **Cutover** — out of scope of this plan entirely. Separate decision after the parallel window, judged against the quality gate, with its own restore point and go-ahead. Triggers the §9 retirement list. | — | — |

### 6a. COM-free kill-switch (Layer 6, reworked)

The 26 Aug `mailbox_guard.py` does an Outlook **COM** delta-sweep of Sent/Drafts/categories. That reintroduces "classic Outlook must be running" — the exact dependency this route removes. **Do not wire it as-is.** Rework:

- **Pre-run:** a read-only connector call lists the Sent Items folder for a small trailing window (e.g. last 2h) → snapshot `{message_id set, count}` to `data/codex_runs/sent_baseline_<ts>.json`.
- **Post-run:** same call again. **Any `message_id` present post- but not pre-**, or a count increase → the run produced an unexplained send → `Disable-ScheduledTask` on the parallel task + BurntToast (reuse `mailbox_guard.py`'s existing disable+notify code, swap only the data source).
- **Feasibility:** the connector's `list_messages` can be folder-scoped to Sent — this is a Call-1-style dumb read, no COM, no desktop app. **Feasible.** If a live check ever shows the connector cannot fold-scope to Sent, or cannot be called read-only in that session, **stop and flag it — do not fall back to a COM guard.**
- **Known blind spots (documented, unchanged from the COM version):** on-behalf / shared-mailbox Sent, Teams messages, a net-zero count change, and detection lag ≈ run duration + sweep (minutes, not prevention). It is a backstop, not a gate — Layers 1–2 + human-review (Layer 5) do the real work.

### 6b. Parity validation — against `data/briefing.json` history, not manual COM runs

Kevin does not keep classic Outlook open, so parity **must not** require him to produce side-by-side COM captures.

- The live scheduled `fetch_inbox.py` still runs on `claude -p` throughout and pushes `data/briefing.json` to GitHub **whenever classic Outlook happens to be up** (via the `Classic Outlook Keepalive` watchdog). Each such push is a timestamped baseline in the repo's commit history.
- The parallel Codex run writes `docs/codex_briefing.json`. The parity harness pulls the **nearest-preceding** `data/briefing.json` from GitHub history (Contents API + commit list, with a cache-buster) and diffs field-by-field: context paragraph (spot-check), per-card tier / `needs_reply` / `no_action_needed`, task suggestions, calendar-prep. Match cards on normalised subject + received-to-the-minute.
- Windows where no live baseline exists (Outlook was down) simply have no comparison that cycle — acceptable; the judgement is over the whole window (target ≈ 40+ compared runs), not any single one.
- This also means: **the more the connector path proves out, the less it matters that the COM baseline is intermittent** — the baseline is only needed during validation, and only on the cycles it exists.

---

## 7. Hard gates

- **No cutover.** No `.bat` / scheduled-task change, no `main` default-behaviour change, without Kevin's fresh explicit go-ahead **for that specific step**.
- `claude -p` stays the live triage engine throughout.
- Parallel-run build only: writes nothing live, pushes nothing, mutates no Command-Centre task / no `data/triage_ledger.json`. Output is `docs/codex_*.json` + `data/codex_runs/*` exclusively.
- `~/.codex/config.toml` sha1 recorded before **and** after every session that touches Codex; must match, or the change is explained and Kevin-approved. Baseline at time of writing: `35f8910382373d525598194b2649159cfeed3f6a`.
- No `codex login` / account switch / `[apps]` edit without telling Kevin first.
- Every run prints a timestamp (standing requirement).
- The kill-switch is detect-and-disable only — a backstop, never a gate; it does not make the write path safe on its own.
- **No COM dependency may be added anywhere in this pipeline** — not in the pull, not in calendar, not in the kill-switch. If a step seems to need Outlook running, stop and flag it.

---

## 8. Retire on cutover (COM-free end state — do NOT remove anything yet)

Once the connector path is proven at parity and Kevin authorises cutover, these go — the target is **zero classic-Outlook dependency, no desktop app open**:

| Retire | Replaced by | Notes |
|---|---|---|
| `Classic Outlook Keepalive` scheduled task (+ `Ensure-ClassicOutlook.ps1`, `Run Classic Outlook Keepalive Hidden.vbs`, `Register-/Unregister-ClassicOutlookKeepalive.ps1`) | nothing — no longer needed | WS1 work from 28 Aug. Unregister command is in its own script. |
| `Ensure-ClassicOutlook.ps1` preflight block in `Run Inbox Briefing.bat` / `Run Draft Diff Capture.bat` | nothing | timestamped `.bat` backups exist |
| `fetch_inbox.py` Phase 1 Outlook COM pull (inbox / VIP / subfolder / Sent) | Call 1 connector pull → `normalise_pull.py` | the `MAIL_BACKEND=com` path and `connect_to_outlook()` |
| `fetch_inbox.py` Phases 3.7 / 3.8 COM calendar pull (primary + "People Department - HR Systems" shared calendar) | Call 1 connector **calendar** pull | the connector reads calendar directly — this is the piece that finally lets calendar go COM-free |
| `open_email.py` + `openmail://` protocol handler + `js/app.js openEmail()` COM path | OWA deep-link opener (`web_link` on `outlook.office.com`, reusing command-centre's `openEmailWeb` validation) | the connector / Graph message carries a `webLink`; store it, open in a new tab |
| `MAIL_BACKEND` flag + `imap_mail.py` (parked IMAP path) | delete or leave dormant | superseded by the connector decision; parked, see `MAIL_BACKEND_MIGRATION_PLAN.md` |
| `mailbox_guard.py` COM sweep code | the §6a connector Sent-folder delta-check | keep the file, replace the data source |
| the 6 `claude -p` calls in `fetch_inbox.py` | Call 2 Codex judgement | `AI_BACKEND=claude_code` path retires; `api` path can stay as an emergency fallback |

Everything downstream of the briefing JSON (dashboard, `command-centre` sync, `data/ticks.json`, Lauren's drafting loop) is **unchanged** by cutover — it consumes `briefing.json` / `inbox_suggestions.json` regardless of what produced them.

---

## 9. Supersession note for a cold session

If you're reading `OPTION3_BUILD_PLAN.md`: that plan is the **opposite** shape (connector-free, COM pull) and is historical. This file is current. If you're reading `PHASE2_BRIEF.md`: its connector-pull half is superseded here; its hard constraints (Codex writes local files only, orchestrator pushes; separate dedup ledger keyed on the message id; no sends/drafts/calendar writes) still hold.
