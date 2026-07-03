# work-inbox — Living Handover Document

**Last updated:** 2026-07-03 — end of session. All three outstanding fix items resolved.
**Status:** Active — pipeline fully working. Live at https://wi.lelitte.co.uk/ | https://begb0037admin.github.io/work-inbox/.

---

## NEXT SESSION — Fix list

No outstanding fix items. All three items from the previous fix list are resolved.

---

## Session 2026-07-03 — Three fix items completed (branch `claude/handover-md-review-fm0r44`, PR #27)

### Item 3 — Drag animation fixed (commits `a4e8d648`)
- Root cause: v5 redesign renamed priority card HTML class from `pri-card` to `card-ph`, but all drag selectors in `js/app.js` and `css/styles.css` still referenced `.pri-card`.
- **`js/app.js`**: Updated selectors in `priDragStart`, `priDragEnd`, `priCardDragOver`, `priCardDrop`, `priZoneDrop` — all `.pri-card` references replaced with `.card-ph`.
- **`css/styles.css`**: Updated three rules — `.pri-card.pri-dragging`, `.pri-card.pri-drop-before`, `.pri-card.pri-drop-after` → `.card-ph.*`.
- No logic or appearance changes — pure selector correction.

### Item 2 — AI calendar summaries (no changes needed)
- Verified that Phase 3.8 (`ai_calendar_summaries`) was already fully implemented in `fetch_inbox.py`. The previous fix list entry was outdated. No code changes required.

### Item 1 — Tomorrow absences now surfaced (commit `678e5f89`)
- Root cause: `fetch_inbox.py` absence detection loop only checked `item_start <= today < item_end`. Team members going on leave tomorrow were invisible today.
- Added `elif item_start == tomorrow and item_end > tomorrow` branch — adds `" (from tomorrow)"` suffix to the name in `absence_set`.
- `tomorrow` is already computed via `next_workday(today)` (skips weekends). The calendar window already covers tomorrow so tomorrow's events were already in the `calendar` list — no additional fetch needed.

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
- **Hard rule propagated** — added to CLAUDE.md for hris-launcher, command-centre, hr-fa-knowledge-base (matching the rule already in work-inbox CLAUDE.md). Visual verification on live dashboards is Kevin's to confirm — no grey squares expected based on source inspection.
- **Reminder trigger** (set 2026-07-03, trig_01U3X2rbvpD1aR7V32DYW7N4) deleted — session resumed manually.

---

## Session 2026-07-03 (aborted — GitHub MCP disconnected)

GitHub MCP server disconnected before handover could be pushed. No code changes this session. Reminder trigger set for 2026-07-04 07:00 UTC. Two new items identified (crest rule propagation, grey square verification) — both addressed in 2026-07-04 session above.

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
- **Links updated**: 6 approved links, all now populated (OSM IT Services = `https://oxford.saasiteu.com/Modules/SelfService/#home`)
- **Cards redesigned**: flat `.card-ph` design (drag handle, circle done button, title + sub, email + CC→ icons, NEW/UPDATED badges on right)
- **Layout corrected**: left col = Today + Tomorrow, right col = Week + Parked
- **Oxford crest**: restored as external file `images/oxford-crest.jpg` — NEVER embed as base64, NEVER delete, NEVER change the `src` attribute

---

## Session 2026-07-02 — v5 full redesign (PR #24)

Three files changed: `index.html`, `css/styles.css`, `js/app.js`. Approved by Kevin, cherry-picked to main throughout session.

**`js/app.js`** — functions added: `renderCalPanel`, `setupCtxTicker`, `_renderCtx`, `_jumpCtx`, `updateInboxWidget`, `applyFilter`, `clearSel`, `clickStat`, `loadCcTicker`. `renderBriefing()` updated. `loadTasksWidget()` replaced by `loadCcTicker()`.

All existing mechanics preserved: drag/drop, openmail://, tick sync, archive, show done, priority overrides.

---

## Session 2026-06-29 — Task Scheduler re-established

- Re-established via `create_inbox_tasks.bat` (repo root). Run as Administrator.
- Schedule: `WorkInbox-0900` / `WorkInbox-1200` / `WorkInbox-1500` Mon–Fri.

---

## Session 2026-06-28 — Live clock added to sidebar

- `updateWiClock()` + `setInterval` added to `js/app.js`. Merged to main via PR #20.

---

## Session 2026-06-27 — File split + Cloudflare custom domain

- 65KB monolithic `index.html` split into `index.html` (shell) + `css/styles.css` + `js/app.js`.
- `wi.lelitte.co.uk` live via Cloudflare. Auto-deploys from main.

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
- Task Scheduler — `WorkInbox-0900` / `WorkInbox-1200` / `WorkInbox-1500` (Mon–Fri). Recovery: run `create_inbox_tasks.bat` as Administrator.
- Dashboard loads live briefing.json on load, falls back to localStorage archive
- Oxford navy sidebar — crest (external `images/oxford-crest.jpg`), branding, live clock, filter, CC ticker, inbox widget, absences, all 6 links populated
- 3-column calendar panel (Today `7fr` | Tomorrow `7fr` | July+August mini-cals in one card `4fr`)
- Rotating context strip with "Briefing context" label, dot nav
- 2×2 priority grid with tier filter — flat `.card-ph` design, NEW/UPDATED badges on right
- CC ticker reads live from CC tasks.json every 60s
- drag-and-drop, tick sync, archive, show done, openmail:// all working
- Multi-machine setup complete (begb0037.AD-OAK)
- Priority card drag animation — `.card-ph` selectors correct in `js/app.js` and `css/styles.css`
- Tomorrow absences surfaced in `fetch_inbox.py` with `" (from tomorrow)"` suffix
- AI calendar summaries (Phase 3.8) — confirmed fully implemented

### Known issues
- None outstanding.

### Critical Note — Desktop Bat File
- **Never rename** — always download fresh via PowerShell:
  `Invoke-WebRequest -UseBasicParsing "https://raw.githubusercontent.com/begb0037admin/work-inbox/main/Run_Inbox_Briefing.bat" -OutFile "$env:USERPROFILE\Desktop\Run Inbox Briefing.bat"`

---

## localStorage Keys

| Key | Purpose |
|-----|---------|
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
|------|----------|
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
