# work-inbox — Living Handover Document

**Last updated:** 2026-07-02 — v5 design corrections: flat cards, sidebar links, email removal. Pushed to main (commit `12ff90d`).
**Status:** Active — pipeline fully working. Live at https://wi.lelitte.co.uk/ | https://begb0037admin.github.io/work-inbox/.

---

## Session 2026-07-02 — v5 design corrections (commit `12ff90d`)

Three files changed: `index.html`, `css/styles.css`, `js/app.js`. Pushed directly to main.

### What changed
- **Removed** email address (`kevin.lelitte@admin.ox.ac.uk`) from sidebar `<div class="user-block">`
- **Links updated**: sidebar now has all 6 approved links (Command Centre, HRIS Dashboard, HRIS Launcher, Access Group Support, FA Knowledge Base, OSM IT Services). OSM IT Services URL is `#` placeholder — Kevin to provide real URL.
- **Cards redesigned**: priority cards switched from old expandable `.pri-card` design to approved flat `.card-ph` design (drag handle, circle done button, title + action sub-text, small email + CC→ icon buttons)
- **Layout corrected**: inbox grid is stacked left/right — left col = Today + Tomorrow, right col = Week + Parked
- **CSS added**: `.card-ph`, `.card-drag`, `.card-done-btn`, `.card-ph-body`, `.card-ph-title`, `.card-ph-sub`, `.card-ph-actions`, `.card-icon`, `.card-icon-cc`

### Pending
- OSM IT Services URL: placeholder `#` — Kevin to provide
- Cloudflare cache purge for wi.lelitte.co.uk (or use begb0037admin.github.io/work-inbox/ to verify)

---

## Session 2026-07-02 — v5 full redesign (PR #24)

Three files changed: `index.html`, `css/styles.css`, `js/app.js`. Approved by Kevin, awaiting merge.

### What changed

**`index.html`**
- Real Oxford crest (base64 JPEG) in sidebar
- Filter dropdown — All tiers / Today / Tomorrow / This Week / Parked
- Command Centre Focus ticker — reads live CC `tasks.json` via github-proxy; shows Today / Tomorrow / Week / Parked counts + Stalled in Today / Oldest / Avg age / 2+ weeks age stats
- "From your inbox" widget with urgent badge count
- `calPanel` div added before `contextBar` in main area

**`css/styles.css`** — new v5 styles appended:
- `.wi-clock-time` — tabular-nums HH:MM:SS
- `.filter-select` — chevron dropdown
- `.inbox-ticker` / `.ticker-stat` / `.ticker-num` — CC ticker; red/gold/green counts; `.selected` highlight on click
- `.inbox-widget` / `.inbox-widget-badge`
- `.main-cal-panel` — 3-col grid (today | tomorrow | 195px mini-month)
- `.ctx-strip` — pale blue `#eff6ff` with Oxford navy bold text, `ctxFlipIn` animation
- `.ctx-dot.active`, `.sec-count`

**`js/app.js`** — new functions:
- `renderCalPanel(data)` → `#calPanel` (3-col calendar, past/next highlighting)
- `setupCtxTicker(context)` / `_renderCtx()` / `_jumpCtx(i)` — rotating sentences, 4500ms, hover-pause, dot nav
- `updateInboxWidget(data)` — urgent badge + total count
- `applyFilter(val)` — shows/hides `sec-{today,tomorrow,week,parked}-wrap`; collapses to 1-col
- `clearSel()` / `clickStat(tier)` — ticker stat click toggles filter + selected state
- `loadCcTicker()` — fetches CC tasks.json; today/tomorrow/week/parked counts + stalled/oldest/avg/2-weeks age; runs on load + every 60s

- `renderBriefing()` updated — calls new functions; `sec-*-wrap` IDs; `sec-count` spans; "FYI / Parked" label; `absEl` null-guard
- `loadTasksWidget()` replaced by `loadCcTicker()` — old function read CC counts into a removed `#tasksWidget` element; new function targets the sidebar ticker IDs

All existing mechanics preserved: drag/drop, openmail://, tick sync, archive, show done, priority overrides.

---

## Session 2026-06-29 — Task Scheduler re-established

- **Task Scheduler (WorkInbox-Briefing) was lost** — all tasks were missing. Cause unknown.
- Re-established via `create_inbox_tasks.bat` (committed to repo root). Must be run as Administrator.
- New schedule: **3 tasks at 09:00 / 12:00 / 15:00 Mon–Fri** (reduced from 6 runs/day).
- Tasks created: `WorkInbox-0900`, `WorkInbox-1200`, `WorkInbox-1500`
- Command used: `schtasks /create` with `/sc weekly /d MON,TUE,WED,THU,FRI /rl highest`
- Bat file target: `C:\Users\admin\Desktop\Run Inbox Briefing.bat`
- **If Task Scheduler is ever lost again:** download `create_inbox_tasks.bat` from repo root, right-click → Run as administrator.

---

## Session 2026-06-28 — Live clock added to sidebar

- Added `<div class="wi-clock-time" id="wi-clock-time">` to the `date-block` in `index.html`
- Added `updateWiClock()` + `setInterval(updateWiClock, 1000)` to `js/app.js`
- Approved by Kevin 2026-06-28. Merged to main via PR #20.

---

## Session 2026-06-27 — File split + Cloudflare custom domain

**File split:** 65KB monolithic `index.html` split into 3 files.

| File | SHA | Size |
|---|---|---|
| `index.html` (shell) | `94272f1b` | 13,957 chars |
| `css/styles.css` | `1b1d1f60` | 18,481 chars |
| `js/app.js` | `cc8a0a9d` | 32,615 chars |

**Cloudflare:** `wi.lelitte.co.uk` live and confirmed loading. Auto-deploys from main.
**CORS:** `cc-tasks-writer` updated — `https://wi.lelitte.co.uk` added.

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
- fetch_inbox.py — all phases confirmed working (calendar sort, IncludeRecurrences, max_tokens=4096 for summaries)
- Task Scheduler — `WorkInbox-0900` / `WorkInbox-1200` / `WorkInbox-1500` (Mon–Fri). Recovery: run `create_inbox_tasks.bat` as Administrator.
- Dashboard loads live briefing.json from GitHub on load, falls back to localStorage archive
- Oxford navy sidebar — crest, branding, live clock, filter, CC ticker, inbox widget, absences, links
- 3-column calendar panel (today | tomorrow | mini-month)
- Rotating context strip (pale blue, Oxford navy text, dot nav)
- 2×2 priority grid with tier filter
- CC ticker reads live from CC tasks.json every 60s
- All drag-and-drop, tick sync, archive, show done preserved
- openmail:// click-through — confirmed working
- Multi-machine setup complete (begb0037.AD-OAK)

### Critical Note — Desktop Bat File
- **Never rename an existing bat file** — always download fresh via PowerShell:
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
