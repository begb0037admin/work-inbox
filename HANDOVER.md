# work-inbox — Living Handover Document

**Last updated:** 2026-07-04 - Absence tomorrow-detection with weekend-aware labelling applied.
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

## Fix list

3. **Drag reorder animation** — No visual feedback during drag. Cards need to visually shift in real time as Kevin drags — placeholder in the DOM during `dragover`.

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
