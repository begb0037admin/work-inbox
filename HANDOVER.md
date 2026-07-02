# work-inbox — Living Handover Document

**Last updated:** 2026-07-02 — end of session. Three fixes outstanding for next session (see below).
**Status:** Active — pipeline fully working. Live at https://wi.lelitte.co.uk/ | https://begb0037admin.github.io/work-inbox/.

---

## NEXT SESSION — Fix list (priority order)

1. **Badge position broken** — Attempted to move NEW/UPDATED badges from inside `.card-ph-title` to `.card-ph-actions` (right side, next to CC→). Change was pushed (`js/app.js`, commit `2d39b9e`) but did not work on live site. Investigate and fix. The badge is rendered as `theBadge` in `renderPriorityCards()` — currently placed inside `<div class="card-ph-actions">` but not displaying correctly there. Revert or fix.

2. **Drag reorder animation** — Cards currently show no visual feedback during drag. Kevin needs: as you drag a card down, the card below should flip up above it in real time; as you drag up, the card above should move down. Requires rewriting drag handlers to insert a live placeholder element into the DOM during `dragover`, shifting surrounding cards visually. This is a meaningful piece of work — plan before coding.

3. **OSM IT Services URL** — sidebar link is `#` placeholder. Kevin to provide the real URL.

---

## Session 2026-07-02 (end) — small fixes pushed to main

- **`ctx-strip` label restored** — `setupCtxTicker()` was missing `<div class="ctx-label">Briefing context</div>`. Added back. Commit `fb178b5`.
- **Badge move attempted** — `theBadge` moved from `.card-ph-title` to `.card-ph-actions` in `renderPriorityCards()`. Commit `2d39b9e`. **Did not work on live site — needs fixing next session.**

---

## Session 2026-07-02 (continued) — calendar panel corrections

Commits pushed to main: `af12dff` (equal 3-col, July+August, AI summaries), `1da688d` (combined mini-cals into one card, narrowed calendar column).

### What changed
- **`css/styles.css`**: `.main-cal-panel` grid changed to `7fr 7fr 4fr` — Today and Tomorrow take equal wider columns; mini-cal column is narrower (≈22% of row).
- **`js/app.js`**: `renderMiniCal(monthOffset)` now returns inner content only (no wrapping block). Both months rendered inside a single `.main-cal-block` with a `.mini-cal-divider` `<hr>` between them. AI summaries (`c.summary`) shown on Today/Tomorrow entries as `.main-cal-summary` divs.

---

## Session 2026-07-02 — v5 design corrections (commit `12ff90d`)

- **Removed** email address from sidebar
- **Links updated**: 6 approved links (OSM IT Services = `#` placeholder — Kevin to provide URL)
- **Cards redesigned**: flat `.card-ph` design (drag handle, circle done button, title + sub, email + CC→ icons)
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
- Oxford navy sidebar — crest (external `images/oxford-crest.jpg`), branding, live clock, filter, CC ticker, inbox widget, absences, links
- 3-column calendar panel (Today `7fr` | Tomorrow `7fr` | July+August mini-cals in one card `4fr`)
- Rotating context strip with "Briefing context" label, dot nav
- 2×2 priority grid with tier filter — flat `.card-ph` design
- CC ticker reads live from CC tasks.json every 60s
- drag-and-drop, tick sync, archive, show done, openmail:// all working
- Multi-machine setup complete (begb0037.AD-OAK)

### Known issues (fix next session)
- NEW/UPDATED badge position broken — see fix list above
- Drag reorder has no visual animation — see fix list above

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
