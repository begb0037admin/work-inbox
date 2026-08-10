# work-inbox — Living Handover Document

**Last updated:** 2026-08-10 - draft_final_diff_capture.py added (forward-going half of the style-learning corpus, issue #3), sent_corpus_pull.py refactored to share redaction logic via new style_corpus_common.py; real baseline established against live Outlook Drafts (Drew).
**Status:** Active — pipeline fully working. Live at https://wi.lelitte.co.uk/ | https://begb0037admin.github.io/work-inbox/.

---

## NEXT SESSION — START HERE

### 1. Granola calendar context — CLOSED 2026-07-04 ✅

**DO NOT reopen.** Do not refactor, retune, or alter Phase 3.7b or Phase 3.8.

**Root cause (fixed):** `fetch_inbox.py` only read `detail["summary"]`. Granola note detail responses return usable content in `summary_text` and `summary_markdown`.

**Production fix (commits `7bc621f`, `cf6ca85`, `48e57ea`):**
- `fetch_inbox.py` now falls back to `summary_text` / `summary_markdown`.
- Granola context passed into Phase 3.8 increased to 1500 characters.
- Phase 3.8 asks for 2-3 concise prep sentences with a 900 token response budget.
- Title matching behaviour deliberately unchanged.
- No debug logging, forced matches, phase-skip flags, or dry-run mode in production.

**Future proposals (separate phases only):**
- A first-class DRY_RUN mode for safer diagnostics may be proposed later.
- Any title matching changes require a separate approved phase.

---

## Session 2026-08-10 — sent_corpus_pull.py built (Drew) — not yet run against real data

**Scope:** `begb0037admin/agent-commons` issue #3 (cross-agent email/Teams style-learning pipeline), item 3/4 — bulk-ingest Kevin's own Sent items as the initial style corpus, via Graph API originally, redirected mid-task to reusing work-inbox's proven Outlook COM access instead.

**What was added:** `tools/sent_corpus_pull.py` — a new, separate script, NOT a change to `fetch_inbox.py` or the live 6x/day pipeline. Reuses the identical COM connection pattern already in production (`win32com.client.dynamic.Dispatch("Outlook.Application")` → `GetNamespace("MAPI")` → `GetDefaultFolder(5)` for Sent Mail), but pulls full body text over a month-chunked historical `[SentOn]` window (existing `fetch_inbox.py` Sent read is 7-day/100-char-preview only, feeding ephemeral AI-triage context — never persisted, and not touched by this addition).

**Redaction pass (automated, per Kevin's decision):** keyword/pattern-based, 4 categories (`health`, `bereavement`, `hr_case`, `absence`) — any match anywhere in subject+body excludes the whole message from the corpus. Redaction ledger records only `entry_id`/date/category/known-name-flag, never matched text. Tested against 13 synthetic cases (all 4 categories + 2 negative controls, including a "leave" false-positive check) — 13/13 passed. Chunking logic separately verified for gaps/overlaps across a year boundary.

**Not done yet:** no real pull has been run against live Outlook — this session's environment didn't have Outlook running, and starting it to pull real historical mail was treated as past the "build and report" checkpoint, needing Kevin's explicit go-ahead first. Proposed durable output location: `begb0037admin/agent-commons` `corpus/sent-items/` (scaffold README pushed there, no real corpus data yet).

**Update, same session — real dry run against live Outlook (Kevin started Outlook Classic mid-session):** ran in `--stats-only` mode (aggregate counts only, nothing written to disk) against the real last-90-days Sent folder. First run: `total_seen: 740` vs `clean_count(327) + redacted_count(76) = 403` — 337 items (45%) silently disappearing through a bare `except: continue`. Root cause: Sent Items also holds meeting requests/responses/cancellations (COM `Class` 53/54/55/56/57), which lack mail-style `Body`/`To` and threw an unhandled `AttributeError` indistinguishable from a real bug. Fixed by explicitly filtering to `Class == 43` (`olMail`) up front instead of relying on exception shape. Re-run fully reconciled: 403 = 327 clean + 76 redacted (health 60, hr_case 16, bereavement 3, absence 2), zero unexpected errors on real mail items. ~19% of real Sent Mail over 90 days matched a redaction category.

Full writeup and open questions (recipient-PII in the `to` field, redaction being pattern-based not NLP): `begb0037admin/drew` `memory/sent-items-corpus-investigation.md` and `begb0037admin/agent-commons` issue #3 comments. Cross-agent Outlook COM gotcha (Sent Items non-mail Classes) also logged to `begb0037admin/agent-commons` `memory/index.json`.

**Still not done, on purpose:** no content written to disk yet (all runs stats-only), nothing pushed to `agent-commons/corpus/sent-items/` beyond the design-doc README. Next: real (non-stats-only) pull to local staging, spot-check locally, then push only the reviewed redacted corpus.json.

---

## Session 2026-08-10 (continued) -- draft_final_diff_capture.py built, real baseline established (Drew)

**Scope:** `begb0037admin/agent-commons` issue #3, forward-going half of the corpus approach (item 3) -- capture principal's draft-to-final edits over time, not just the one-time Sent-items backfill.

**Feasibility investigated first (read-only structural probe, no content read/stored):** Outlook's `EntryID` is NOT a safe key to correlate a Drafts-folder item with its eventual Sent-folder counterpart -- sending mints a new MAPI entry. `ConversationID` is: present on 40/40 sampled items in both Drafts (103 total) and Sent Items (1585 total).

**Built:**
- `tools/style_corpus_common.py` -- redaction classifier (health/bereavement/hr_case/absence), `recipient_tier` mapping, and the `OL_MAIL_CLASS` non-mail-item filter, factored out of `sent_corpus_pull.py` now that a second script needs the identical logic.
- `tools/sent_corpus_pull.py` -- refactored to import the shared module instead of duplicating it. Re-ran the original 13-case synthetic redaction suite + chunking test against the refactor -- zero regression.
- `tools/draft_final_diff_capture.py` -- periodic snapshot-and-correlate (not an event-driven listener -- considered `Application.ItemSend` for perfect fidelity, rejected for v1 since it needs a persistently-running process, a different architecture from every other script here). Snapshots Drafts each run, diffs against the previous run's local-only ledger to find vanished drafts, correlates against Sent Items by `ConversationID` within a 72h window (earliest match wins, no fallback guessing), applies the same whole-pair redaction exclusion as Sent-items (either side sensitive excludes both), computes `recipient_tier`, classifies `edit_type`/`note` via claude-haiku-4-5 on the redacted pair (confirmed OK with Kevin, same model `fetch_inbox.py` already uses).

**Verification:** correlation logic -- 5 mocked-Outlook cases (window bounds, non-mail filtering, multiple-candidate tiebreak, no-match), 5/5 pass. Whole-pair redaction gate -- confirmed both directions (draft-only and final-only sensitivity both correctly exclude the pair) with the real classifier. `edit_type` classification -- 5 synthetic pairs against the real API, 5/5 valid enum, 4/5 exact intended match (1 legitimately ambiguous test case, not a classifier bug). **Real baseline run against live Outlook:** 96 drafts tracked into the local-only ledger (`C:/Users/admin/Documents/CorpusStaging/draft_watch/ledger.json`, confirmed outside any git working tree), 0 vanished/0 pairs -- expected and correct for a first run, not a bug (the mechanism is inherently forward-looking).

**Not done, on purpose:** no diff pairs exist yet (need a real send to happen between two runs), nothing pushed to `agent-commons/corpus/draft-final-diffs/`. Not yet wired into Task Scheduler -- holding for confirmation given it makes live Anthropic API calls per pair on an unattended schedule.

**Also this session:** Teams draft-staging design moved from proposal to concrete (surface confirmed as work-inbox by Kevin) -- new "Pending Teams Replies" panel, data cross-fetched from `agent-commons/pending-teams-drafts/drafts.json` (mirrors the existing CC-ticker cross-repo-fetch pattern; preserves the standing rule that Lauren never writes into work-inbox directly), reusing the existing `workInbox_ticks_v1` Cloudflare-Worker-synced tick mechanism for "mark as sent" rather than building new write-back infra. Design only -- still blocked on the separately-deferred Teams read-access question, not built. Full detail: `begb0037admin/agent-commons` issue #3.

---

## Fix list

3. **Drag reorder animation** — No visual feedback during drag. Cards need to visually shift in real time as Kevin drags — placeholder in the DOM during `dragover`.

4. **Phase 3.8 calendar-summary mismatch on days starting with an all-day event** — investigated by Drew 2026-08-04, root cause confirmed, NOT fixed (Phase 3.8 is closed — needs Kevin to explicitly reopen it before any code change). See "Session 2026-08-04" below for full detail and full writeup in `begb0037admin/drew` repo, `memory/calendar-summary-offset-bug.md`.

---

## Session 2026-08-04 — Calendar-summary offset bug investigated, NOT fixed (Drew)

**Scope:** Flagged during unrelated meeting-records work — live `data/briefing.json` `calTomorrow` items had AI-generated `summary` text describing a *different* meeting than the one it was attached to. Investigated whether this is a Python index bug in `fetch_inbox.py` Phase 3.8.

**Confirmed against live `data/briefing.json` (Wednesday 5 August briefing), pulled fresh with a cache-buster:**
- `calToday` (9 items, first item is a real 09:30 meeting, no preceding all-day item): zero mismatches, every summary correctly self-referential.
- `calTomorrow` (9 items, idx 0 is an all-day event — "Simon out of the office - funeral"): idx 1, 2, 3 each carry the summary content that rightfully belongs to the *next* item (idx 1 shows idx 2's title, idx 2 shows idx 3's title, idx 3 shows idx 4's actual topic while idx 4 itself is left with no summary). idx 6 and 7 are correctly self-referential — the mismatch does not persist for the whole day.

**Root cause (confirmed, not guessed):** the Python index bookkeeping in Phase 3.8 (`_cal_for_summary` construction and the `target[item["idx"]] = ...` write-back, ~line 1168 onward) is correct — re-read line by line, positions match. The mismatch correlates exactly with whether the day's `idx` sequence fed to the model starts at 0 or not. `calToday` starts at idx 0 (no shift). `calTomorrow`'s first non-all-day item carries idx 1 (idx 0 was filtered out as an all-day event) — and the model's own generation of the `"day_idx"` JSON response keys (`tomorrow_1`, `tomorrow_2`, ...) mismatches the content it writes for the first few real items before self-correcting by idx 6. This is a **prompt/model reliability issue** (claude-haiku-4-5, the model Phase 3.8 is locked to, appears to fall back to counting output position from 0 rather than reliably echoing the literal `idx` value whenever that value doesn't start at 0) — not a deterministic code bug.

**Proposed fix (not applied):** decouple the index shown to the model from the index used for write-back — renumber the `idx` sent to the model to always start at 0 within `_cal_for_summary` (sequential by array position), and keep the original `cal_today_items`/`cal_tomorrow_items` position in a separate field never exposed to the model, used only for the write-back. Small, contained diff, same file.

**Why this was not pushed:** Phase 3.8 is marked closed in this file and in `CLAUDE.md` — "do not modify without Kevin explicitly opening a new approved phase." No message in the task that surfaced this constituted that explicit reopening, so this was investigated and root-caused only, not fixed. Full writeup: `begb0037admin/drew` repo, `memory/calendar-summary-offset-bug.md`.

---

## Session 2026-08-02 — CC ticker done-task filtering fix (Drew's first task)

**Scope:** Fix bug where the "Command Centre Focus" sidebar ticker (`loadCcTicker()` in `js/app.js`) counted ALL Command Centre tasks per tier, including tasks marked `done: true`, so it disagreed with Command Centre's own "Daily Focus" tile, which correctly counts only open (`!t.done`) tasks per tier.

**Confirmed against live `command-centre/data/tasks.json` before the fix:** 39 tasks total, 13 marked done. All-tasks counts (old, wrong) vs. open-only counts (new, correct — matches CC's own tile):
| Tier | All (old) | Open only (new/CC) |
|---|---|---|
| Today | 10 | 5 |
| Tomorrow | 6 | 4 |
| Week | 13 | 8 |
| Parked | 10 | 9 |

**Root cause:** none of the four `tasks.filter(t=>t.tier===...)` calls in `loadCcTicker()` excluded `t.done`. Command Centre's own `js/app.js` `renderBoard()` does `tasks.filter(t=>t.tier==='today'&&!t.done)` — work-inbox's ticker didn't match that.

**Fix (commit `8582608`):** added `const openTasks=tasks.filter(t=>!t.done);` right after the tasks array is built, and switched all four tier-count filters (`cc-today-count`, `cc-tmrw-count`, `cc-week-count`, `cc-parked-count`) plus the age-based stats (`ages`, `stalled`, `oldest`, `avg`, `twoWeeks` — i.e. `cc-stalled`, `cc-oldest`, `cc-avg`, `cc-twoweeks`) to run over `openTasks` instead of the full `tasks` array. Judgement call, not explicitly requested by Kevin: extended the fix to the age stats too, for consistency — a completed-but-old task shouldn't be able to drag "Oldest task"/"Avg age" up or count toward "stalled"/"2+ weeks old", since those exist to flag *open* work going stale. Worth Kevin's explicit confirmation if this reads wrong once he's looking at real numbers.

**Verification:** pushed via GitHub Contents API PUT, base SHA `cbf52b72a7b84ceed9df287ceb9d5436d55ccc09` → new SHA `d07b6c97b65d5b9496466d011dbd4fa2071f1f55`. Re-fetched the pushed blob via `git/blobs/{sha}` (not `raw.githubusercontent.com`, which caches) and diffed byte-for-byte against the intended patched file — exact match. Re-fetched live `command-centre/data/tasks.json` after the fix and simulated the new `loadCcTicker()` logic: Today 5 / Tomorrow 4 / Week 8 / Parked 9 — matches Command Centre's own `renderBoard()` tier-count logic (`tier===x && !done`) exactly on the same live data. `node --check` confirmed the pushed file is syntactically valid JS.

**Diff scope confirmed minimal:** the only changes in the file are inside `loadCcTicker()` — one added line (`openTasks`) and eight filter calls switched from `tasks` to `openTasks`. Nothing else in `js/app.js` touched.

**Also flagged to Kevin, not actioned this session (his call to prioritize):**
- **Command Centre "Save failed — HTTP 502" + no way to exit inbox-suggestions view** — worker-side 502 from `cc-tasks-writer.kevinlelitte.workers.dev` on `persistTasks()`, cause not yet confirmed (possibly inbox auto-promotion write frequency, possibly cold start). Separately, `showView('inbox')` in command-centre's `js/app.js` has no matching `showView('board')` control in the markup — clicking "From your inbox" strands the user on the Inbox Suggestions view with no way back except F5. Kevin previously asked to hold other command-centre work until this is resolved — likely next in line.
- **Work Inbox "Command Centre Focus" ticker redesign (6-across / drop Parked, widen to show Urgent + Needs response)** — dropped entirely by Kevin mid-session, 2026-08-02: the two tiles (CC's own Daily Focus vs. work-inbox's cross-reference into CC) are supposed to show different things by design; the redesign ask was based on a misunderstanding, not a real gap. Nothing from that thread was ever pushed to GitHub — it only existed as unapproved scratchpad edits (`wi_app_fresh.js`, `wi_index_fresh.html`, `wi_styles_fresh.css`) — so there is nothing to revert on the live site. Do not resurrect this thread without Kevin raising it again.

---

## Session 2026-07-04 — Absence tomorrow-detection fix (commit `3aab85c`)

**Scope:** `fetch_inbox.py` absence detection extended to surface tomorrow's leave in the sidebar absences panel. Weekend-aware labelling added.

**What changed:**
- Absence detection block replaced with version that scans both today and next working day.
- Today's absences on weekends/Sundays show `"(next week)"` suffix — avoids "today" implying a working day when today is Saturday.
- Absences starting on `tomorrow` (= `next_workday(today)`) labelled `"(tomorrow)"` on Mon–Thu, `"(next week)"` on Fri/Sat/Sun.
- Shared `_extract_absence_name()` helper removes duplication from the name-stripping logic.
- No duplicate checking needed: date logic naturally prevents double-listing the same person.

**Kevin approval:** "Yep, approved."

---

## Session 2026-07-04 - Pipeline hardening review follow-ups

**Scope:** Apply quick review follow-ups after Granola rollout.

**What changed:**
- `fetch_inbox.py`: Added a shared GitHub API timeout for script GitHub reads/writes.
- `fetch_inbox.py`: Made Phase 3.6 task action append idempotent by skipping exact duplicate action text.
- `fetch_inbox.py`: Renamed the Granola comment to Phase 3.7b to reduce diagnostic ambiguity; behaviour unchanged.
- `js/app.js`: Added HTML escaping for calendar times, titles, organisers, and summaries before rendering.

**Remaining non-blocking improvement:** A first-class DRY_RUN mode would still make future diagnostics safer because Phase 3.6, Phase 4, and Phase 5 can write to GitHub.

---

## Session 2026-07-04 - Granola calendar context fix (CLOSED — do not reopen)

**Scope:** Fix Phase 3.7 Granola context and improve Phase 3.8 meeting prep summaries.

**What changed:**
- `fetch_inbox.py`: Granola note detail extraction now falls back from `summary` to `summary_text` / `summary_markdown`.
- `fetch_inbox.py`: Granola context passed into Phase 3.8 increased from 500 to 1500 characters.
- `fetch_inbox.py`: Phase 3.8 now asks for 2-3 concise prep sentences and has a 900 token response budget.

**Validation:** Local debug smoke test confirmed `FA Team Daily Catchup` matched `FA Team Catch-up - 03/07`; dashboard smoke test used `Company 90 - Status Update` and confirmed the calendar summary display works.

**Not included:** No title matching changes, no forced debug matches, no diagnostic logging spam, no phase skip flags, and no `fetch_inbox_debug.py` changes in production.

---

## Session 2026-07-03 — Calendar scroll (approved, pushed to main)

**Scope:** Replace expand/collapse toggle on Today and Tomorrow calendar columns with independent vertical scrolling. Keep fixed height (260px), same size and position.

**What changed:**
- **`css/styles.css`** (commit `dc3544b`): Removed expand/collapse styles (`.cal-col-body` with `overflow:hidden`, `.cal-expand-footer`, `.cal-expand-btn`). Added scroll styles — `.cal-col-body { max-height: 260px; overflow-y: auto; overflow-x: hidden }` with 4px webkit scrollbar (`#d1d9e6` thumb, hover `#94a3b8`).
- **`js/app.js`** (commit `6589384`): `renderBlock()` inside `renderCalPanel()` — return statement no longer includes `cal-expand-footer` div. `toggleCalExpand()` function removed entirely. Both Today (`calBodyToday`) and Tomorrow (`calBodyTom`) columns now scroll independently via the same `renderBlock` code path.

**Kevin approval:** "perfect, approved ensure that it's on both columns today and tomorrow."

---

## Session 2026-07-03 — Granola 0-matches investigation (superseded — see CLOSED phase above)

**Scope:** Diagnosing why Phase 3.7 Granola fetch returns 10 notes but matches 0 calendar items.

**Resolution:** Fixed 2026-07-04. Root cause was `summary_text`/`summary_markdown` fallback missing. See CLOSED phase entry above.

---

## Session 2026-07-04 — Crest rule propagation

No code changes to work-inbox this session. Cross-repo maintenance only.

- **Crest audit completed** — all dashboards inspected for Oxford crest usage:
  - work-inbox: external file `images/oxford-crest.jpg` — intact ✅
  - hris-launcher: base64 JPEG `<img class="sidebar-crest">` — intact ✅
  - command-centre: base64 JPEG `<img class="sb-crest">` — intact ✅
  - hr-fa-knowledge-base: base64 JPEG `<img class="crest">` — intact ✅
  - hris-dashboard: emoji 🎓 (no image) — N/A
  - ag-flexpoints: no crest — N/A
- **Hard rule propagated** — added to CLAUDE.md for hris-launcher, command-centre, hr-fa-knowledge-base.

---

## Session 2026-07-02 (end) — small fixes pushed to main

- **`ctx-strip` label restored** — `setupCtxTicker()` was missing `<div class="ctx-label">Briefing context</div>`. Added back. Commit `fb178b5`.
- **Badge position fixed** — NEW/UPDATED badges moved from inside `.card-ph-title` to `.card-ph-actions` (right side, next to CC→). Commit `2d39b9e`. Confirmed working.
- **OSM IT Services URL** — sidebar link updated to `https://oxford.saasiteu.com/Modules/SelfService/#home`. Commit `e4cc1fd`.

---

## Session 2026-07-02 (continued) — calendar panel corrections

Commits pushed to main: `af12dff` (equal 3-col, July+August, AI summaries), `1da688d` (combined mini-cals into one card, narrowed calendar column).

### What changed
- **`css/styles.css`**: `.main-cal-panel` grid changed to `7fr 7fr 4fr` — Today and Tomorrow take equal wider columns; mini-cal column is narrower (≈22% of row).
- **`js/app.js`**: `renderMiniCal(monthOffset)` now returns inner content only (no wrapping block). Both months rendered inside a single `.main-cal-block` with a `.mini-cal-divider` `<hr>` between them. AI summaries (`c.summary`) shown on Today/Tomorrow entries as `.main-cal-summary` divs.

---

## Session 2026-07-02 — v5 design corrections (commit `12ff90d`)

- **Removed** email address from sidebar
- **Links updated**: 6 approved links, all now populated
- **Cards redesigned**: flat `.card-ph` design (drag handle, circle done button, title + sub, email + CC→ icons, NEW/UPDATED badges on right)
- **Layout corrected**: left col = Today + Tomorrow, right col = Week + Parked
- **Oxford crest**: restored as external file `images/oxford-crest.jpg` — NEVER embed as base64, NEVER delete, NEVER change the `src` attribute

---

## Architecture

| Component | Description |
|-----------|-------------|
| `fetch_inbox.py` | Outlook COM via pywin32. Pulls inbox → Anthropic triage (claude-haiku-4-5) → pushes `data/briefing.json` to GitHub via Contents API |
| `index.html` | Shell — HTML structure only. Loads `css/styles.css` → `js/app.js`. No framework, no build step. |
| `css/styles.css` | All styles. |
| `js/app.js` | All JS — briefing render, cal panel, ctx ticker, CC ticker, drag-and-drop, tick sync, archive, live clock. |
| `open_email.py` | Registered `openmail://` protocol handler — opens exact email in classic Outlook via EntryID COM |

---

## Current State

### Working
- fetch_inbox.py — all phases confirmed working
- **Granola calendar context (Phase 3.7b + 3.8)** — COMPLETE. Matching via keyword overlap; summary extracted from `summary_text`/`summary_markdown`. Do not modify.
- **Absence detection** — today's leave + tomorrow's leave (weekend-aware labelling). Commit `3aab85c`.
- Task Scheduler — `WorkInbox-0900` / `WorkInbox-1200` / `WorkInbox-1500` (Mon–Fri)
- Dashboard loads live briefing.json on load, falls back to localStorage archive
- Oxford navy sidebar — crest (external `images/oxford-crest.jpg`), branding, live clock, filter, CC ticker, absences, all 6 links populated
- 3-column calendar panel (Today `7fr` | Tomorrow `7fr` | July+August mini-cals in one card `4fr`)
- **Calendar columns scroll independently** — Today and Tomorrow each have `max-height: 260px; overflow-y: auto` with 4px scrollbar. Expand/collapse removed.
- Rotating context strip with "Briefing context" label, dot nav
- 2×2 priority grid with tier filter — flat `.card-ph` design, NEW/UPDATED badges on right
- CC ticker reads live from CC tasks.json every 60s
- drag-and-drop, tick sync, archive, show done, openmail:// all working
- Multi-machine setup complete (begb0037.AD-OAK)

### Known issues (fix next session)
- Drag reorder has no visual animation
- Phase 3.8 calendar-summary mismatch on days starting with an all-day event (see Fix list item 4 and Session 2026-08-04 above) — root cause confirmed, fix scoped, awaiting Kevin's explicit reopening of Phase 3.8

---

## localStorage Keys

| Key | Purpose |
|-----|--------|
| `workInbox_briefings_v1` | Archive of past briefing JSON objects, keyed by date string |
| `workInbox_today_v1` | Key of the currently displayed briefing |
| `workInbox_ticks_v1` | Tick (done) state for all cards |
| `workInbox_priOverrides_v1` | Per-card section overrides for priority drag-and-drop |
| `workInbox_priOrder_v1` | Per-section sort order for priority cards |
| `workInbox_customPri_v1` | Email cards manually dragged into priority sections |

---

## Technical Notes

**index.html edits:** always use binary `atob()`/`btoa()` — NEVER `TextEncoder` on file content (re-encodes em-dash bytes).

**Priority drag-and-drop sections:** `pt` (today), `ptom` (tomorrow), `pw` (week), `pfyi` (parked/FYI), `ur` (urgent overlay), `nr` (needs overlay).

---

## File Locations

| File | Location |
|------|---------|
| Repo | github.com/begb0037admin/work-inbox |
| Proxy | github-proxy.lelitte.co.uk/work-inbox/ |
| Dashboard (primary) | wi.lelitte.co.uk |
| Dashboard (GitHub Pages) | begb0037admin.github.io/work-inbox/ |
| Styles | `css/styles.css` |
| JS | `js/app.js` |
| Script | `fetch_inbox.py` |
| Opener | `open_email.py` |
| Briefing | `data/briefing.json` |
| Local | `C:\Users\admin\Documents\Claude\Projects\work-inbox\` |
| Scheduler recovery | `create_inbox_tasks.bat` in repo root — run as Administrator |

---

## Standing Rules
- Never commit tokens or raw data
- All GitHub writes via Contents API (PAT from `GITHUB_PAT` env var)
- `index.html` edits: always use binary `atob()`/`btoa()` — NEVER `TextEncoder`
- Desktop bat: always download fresh via PowerShell — never rename an existing file
- Every raw.githubusercontent.com fetch MUST include `?t=<timestamp>` cache-buster
- **NEVER touch `images/oxford-crest.jpg` or the `<img class="sidebar-crest">` src attribute** — external file only, never base64
- **Phase 3.7b and Phase 3.8 are closed** — do not modify without Kevin explicitly opening a new approved phase
