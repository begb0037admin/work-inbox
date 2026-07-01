# work-inbox — Living Handover Document

**Last updated:** 2026-07-01 (Hope session, failover — Kevin unavailable, see command-centre AGENT_MODEL.md Section 6) — Daily Focus / cross-dashboard sidebar redesign: mockups approved, not yet implemented
**Status:** Active — pipeline fully working. Task Scheduler confirmed working 2026-06-30.

---

## Session 2026-07-01 (Hope, failover) — Daily Focus / cross-dashboard sidebar redesign — MOCKUPS APPROVED, AWAITING GO-AHEAD TO IMPLEMENT

**Full spec lives in command-centre repo:** `command-centre/docs/DAILY_FOCUS_CROSS_DASHBOARD_REDESIGN.md` — read this before touching any sidebar/main-area code for this feature.
**Approved mockup (final, reference artefact, not live code):** `docs/mockups/wi-full-v5.html` (this repo), `command-centre/docs/mockups/cc-full-v5.html`.

**One-line summary:** Both dashboards move to one shared sidebar template. Work Inbox's Daily Focus widget becomes cross-dashboard — it will show Command Centre's tier/stalled-task data instead of (or in addition to) its own inbox stats; Command Centre's Daily Focus widget will show Work Inbox data. Absences (already exists here, `#absencesSidebar` / `js/app.js:445-447`) stays as its own widget and gets added to Command Centre's sidebar too, sourced from this repo's `data/briefing.json`. Work Inbox also gains a new "From your inbox" badge widget (mirroring the one that already exists on Command Centre, sourced from `command-centre/data/inbox_suggestions.json`). Main-area layout also changes on both dashboards (separate, larger piece of work — see spec doc Section 5).

**Status: NOT IMPLEMENTED.** No changes have been made to `index.html` or `js/app.js`. **Do not begin implementation without Kevin's (or an authorised failover operator's) explicit go-ahead.**

**Next action:** wait for approval, then implement per the spec doc in command-centre.

---

## Session 2026-06-30 (evening) — CC→ button window reuse; hashchange listener in CC

**CC→ button window reuse fix (this repo — js/app.js):**
- Problem: `window.open(url, 'command-centre')` opened a new CC tab on every click, even when CC was already open.
- Root cause: named-window reuse only works for tabs opened via `window.open` — a tab the user navigated to manually has no window name.
- Fix: added module-level `_ccWindow` variable. `openCC(id)` checks `_ccWindow && !_ccWindow.closed`; if true, navigates in place via `_ccWindow.location.href` + `.focus()`; otherwise opens new tab and stores reference. Button onclick updated to call `openCC(id)`.
- Backup before write: `Archive/app_js_backup_20260630_2110.js` (commit `2d9833d5`)
- js/app.js SHA after fix: `44867d67`. Confirmed working by Kevin.

**hashchange listener (command-centre/js/app.js):**
- Problem: CC card not selected/animated when CC→ button used while CC is already open.
- Root cause: CC reads `window.location.hash` once on page load. Navigating an already-open tab via `_ccWindow.location.href` fires `hashchange` in the CC tab — but CC had no listener.
- Fix: `window.addEventListener('hashchange', function(){ var id=window.location.hash.replace('#',''); if(id) goToCard(id); });` appended to CC `js/app.js`.
- CC backup before write: `Archive/app_js_backup_20260630_2115.js` (commit `8e3d84dd`)
- CC `js/app.js` new SHA: `12e0b1be` (commit `8b3c880`). Confirmed working by Kevin.

---

## Session 2026-06-30 — fetch_inbox.py fixes

### Phase 3.7 AI summaries (Unterminated string fix)
- **Problem:** `max_tokens=1000` too low for 25 tasks — JSON response truncated mid-string, causing `json.JSONDecodeError: Unterminated string`.
- **Fix:** `max_tokens` raised to `4096`. Input payload reduced: `description[:300]` (was `[:500]`), `actions[-5:]` (was `[-10:]`).
- **Confirmed:** Phase 3.7 done — 24 summaries generated.

### Calendar sort fix
- **Problem:** Calendar items displayed in Outlook's arbitrary iteration order (e.g. 15:00 shown above 14:00).
- **Fix:** `build_cal_items()` now sorts items by `start` ascending before rendering. Applies to calToday, calTomorrow, and calFull.

### Recurring meetings fix
- **Problem:** Recurring meetings (e.g. HR Systems team weekly) not appearing in calendar — raw `Items` iteration only returns the master record whose original start date is outside the lookback window.
- **Fix:** `_cal_items.IncludeRecurrences = True` + `_cal_items.Sort("[Start]")` + `break` when `t.date() > week_end`. Expands recurring instances within the window without runaway iteration.
- **Note:** The CLAUDE.md constraint "do NOT use Restrict() with IncludeRecurrences" still stands — the fix uses direct sorted iteration, not Restrict().

### Task Scheduler / bat file
- **Problem:** Scheduled tasks were running an old bat file on the desktop (BOM, wrong path `C:\Users\admin\work-inbox`, wrong wording "Updating inbox briefing..."). File had been renamed from a pre-existing version rather than freshly downloaded.
- **Fix:** Fresh download via PowerShell: `Invoke-WebRequest -UseBasicParsing "https://raw.githubusercontent.com/begb0037admin/work-inbox/main/Run_Inbox_Briefing.bat" -OutFile "$env:USERPROFILE\Desktop\Run Inbox Briefing.bat"`
- **Confirmed:** All three scheduled tasks (09:00, 12:00, 15:00) confirmed working.
- **Rule:** Never rename an existing bat file — always download fresh from GitHub via PowerShell to avoid BOM and stale content.

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

- Added `<div class="wi-clock-time" id="wi-clock-time">` to the `date-block` in `index.html`, between the "Today" label and `sidebarDate` div
- Added `updateWiClock()` function + `setInterval(updateWiClock, 1000)` to `js/app.js` — writes live HH:MM:SS to `wi-clock-time` and live day/date to `sidebarDate`, ticking every second
- Removed static `document.getElementById('sidebarDate').textContent=data.date` from `renderBriefing()` — date is now driven by the live clock rather than the briefing JSON
- Archive backup committed before write: `Archive/index_backup_20260628_0600.html` (full original index.html content)
- Approved by Kevin 2026-06-28. Merged to main via PR #20 (squash commit `6cd84f0e`).
- Clock moved from `begb0037admin/command-centre` — see CC HANDOVER.md for corresponding removal.

---

## Session 2026-06-27 — File split + Cloudflare custom domain

**File split:** 65KB monolithic `index.html` split into 3 files via atomic Git Tree API commit.

| File | SHA | Size |
|---|---|---|
| `index.html` (shell) | `94272f1b` | 13,957 chars |
| `css/styles.css` | `1b1d1f60` | 18,481 chars |
| `js/app.js` | `cc8a0a9d` | 32,615 chars |

**Archive backup:** `Archive/index_backup_20260627.html` — SHA `9f168db4` (matches pre-split live file).

**Atomic commit:** `450d5d62`

**Cloudflare deployment:** `wi.lelitte.co.uk` — live and confirmed loading. Cloudflare Worker connected to `begb0037admin/work-inbox` (main branch), auto-deploys on push.

**CORS:** `cc-tasks-writer` updated — `https://wi.lelitte.co.uk` added to `CORS_ORIGINS`.

**CC link updated:** Tasks widget "Open command centre" link updated from GitHub Pages URL to `https://cc.lelitte.co.uk/`.

**Unicode chars preserved:** 22 in JS, 7 in CSS, 3 in HTML — all handled correctly via Python binary base64 encoding. Mojibake constraint maintained.

**Note:** Dashboard shows stale briefing on Mac — expected. `fetch_inbox.py` is Windows-only (Outlook COM). Dashboard display works on any browser.

## Session 2026-06-25 — Command Centre updated from HR Systems Managers Meeting 24/06

**Note: Command Centre and Work Inbox work in tandem. Changes to tasks.json in CC are reflected here; inbox triage automation (Phase 3.6) writes back to CC. Both HANDOVERs must be kept in sync.**

No changes to the work-inbox pipeline this session. The session focused on reviewing Granola notes from the HR Systems Managers Meeting (24 June 2026) and applying the resulting task updates to `begb0037admin/command-centre/data/tasks.json`.

**Key items from the meeting that may surface in inbox or future briefings:**

| Topic | Status | What to watch for |
|---|---|---|
| REF attributes via ESS | Active — scoping meeting to be set by Nathan | Invite from Nathan (Kevin, Simon, Michelle); avoid Thu 25 Jun |
| SHSMS Evaluation (DTP1334) | Active — follow-up by 2 July | Evaluator briefing invites (14 Jul); scoring 65/30/5 confirmed |
| Sickness Absence open bug | Active — Kevin investigating | Emails from Simon Burford re pay impact |
| WFM / GLAM Rostering | Active — resolution meeting to be set | Calendar invite: Kevin, Simon, Marie, Michelle, Julie |
| Holiday Records (case 69001638) | Active — scheduling to be agreed | Access Group consultant contact (within 3 weeks of 24 Jun) |
| DPIA Stage 7 | Urgent — 30 Jun deadline | Chase from Kevin to Marie re signature |
| HR Reporting SSO Migration | Active — VS2022 licensing blocker | No emails expected until licensing resolved |
| Volunteering Hours | Active — Kevin to forward thread to Marie | Matt Thomas thread to forward |

**Follow-up cleanup pass (same session):** Several session corrections did not apply cleanly on the initial write (likely overwritten by inbox automation between writes). A second pass applied the following:

| Change | Detail |
|---|---|
| t1781099880461 deleted | 10 Jun Holiday Records duplicate removed — 11 Jun task (t1781204987882) is the definitive record with full Managers Meeting context |
| t028 deleted | UKVI Skilled Workers — Kevin confirmed not taking forward |
| t015 updated | Web pages / URL linking — E-Form save bug (case 69001555) confirmed fixed by Maura McGlynn (AG) across Production, OXU, OXZ; URL updates now unblocked |
| t025 title fixed | Corrected to "Support cover w/c 6 July" (was incorrectly showing 1 July) |
| t003 closed | Flex Points Plan marked done |
| t022 closed | Pop Notes to Chris — Cority PUG minutes sent to Christopher Sanders, resolved |

**tasks.json state after cleanup pass:** 35 active tasks. Commit `b0326d0` (25 Jun 2026).

---

## Session 2026-06-25 — Roadmap Cleared

All outstanding roadmap items confirmed resolved by Kevin:

| Item | Status |
|---|---|
| KNOWN_ABSENCES for Marie and James removed from fetch_inbox.py | ✅ Done |
| Missing delete rule for surveys.theaccessgroup.com added | ✅ Done |
| AG FlexPoints repo created, Pages enabled, env vars set, Playwright installed, weekly run scheduled | ✅ Done |
| Task Scheduler bat updated to auto-pull fetch_inbox.py before running | ✅ Done |
| openmail:// click-through reliability — monitored, stable | ✅ Done |
| Multi-machine setup (begb0037.AD-OAK) | ✅ Done |

---

## Session 2026-06-24 — Inbox Organisation Complete

Full inbox organisation completed based on analysis of 2,742 inbox + 377 sent emails (Dec 2025 – Jun 2026).

| Item | Status | Detail |
|------|--------|--------|
| 19 folders created inside Inbox | ✅ Done | Access Group (My Cases, Team Cases), PeopleXD System, Reports, Team (Michael, Asta, James), H&S (Cority, DSE·IRIS·Risk Base), Projects (DTP1092, DTP1334, ePloy), Reference (HR Broadcast, ICT Mailing Lists, Bodleian & Sector), _Archive |
| 12 auto-delete rules | ✅ Done | Created via Python COM. Del- prefix. Teams, GitHub, Access Group marketing, New Vacancy Notification, Cority status, DistroKid, Anthropic/Claude, Descript, Accessplanit, Skype voicemail, MetaCompliance, Annual Leave system |
| 13 auto-file rules | ✅ Done | Created by Cowork via Outlook Rules Wizard UI (Python COM blocked by Oxford Exchange for MoveToFolder). Team James/Michael/Asta, Reports ITSRVXT/PeopleXD Reports, PeopleXD System, Cority, HR Broadcast, ICT subject/senders, Bodleian & Sector, Access My Cases/Team Cases |
| Run Rules Now | ✅ Done | Inbox 362 → 170 (192 emails filed) |
| data/inbox_export.json | ✅ Deleted | Commit cacbcfe |
| surveys.theaccessgroup.com delete rule | ✅ Done | Added manually via Rules Wizard |

Reference document: `docs/INBOX_ORGANISATION.md`
Cowork brief: `docs/COWORK_BRIEF_INBOX_RULES.md`

---

## Architecture

| Component | Description |
|-----------|-------------|
| fetch_inbox.py | Outlook COM via pywin32. Pulls inbox → Anthropic triage (claude-haiku-4-5) → Python post-processing → pushes data/briefing.json to GitHub via Contents API |
| index.html | Shell — HTML structure only. Loads css/styles.css → js/app.js. No framework, no build step. |
| css/styles.css | All styles extracted from former monolithic index.html. |
| js/app.js | All JS — briefing render, drag-and-drop, tick sync, archive, tasks widget, live sidebar clock, CC→ button window reuse (`_ccWindow`). |
| open_email.py | Registered openmail:// protocol handler — strips prefix and trailing slash, opens exact email in classic Outlook via EntryID COM |

---

## Current State (fully working as of 2026-06-30)

### Working
- fetch_inbox.py — all phases confirmed working
- Phase 3.7 AI summaries — `max_tokens=4096`, confirmed generating summaries for all priority tasks
- Calendar pull — `IncludeRecurrences=True` + `Sort([Start])` + break at week_end; captures recurring meetings (e.g. HR Systems team weekly). Do NOT use `Restrict()` with `IncludeRecurrences`.
- Calendar sort — `build_cal_items()` sorts by start time ascending; confirmed correct order
- Python post-processing of calendar items — KNOWN_ABSENCES list cleared (Marie and James both returned)
- Absence dedup — partial-name matching prevents duplicates when KNOWN_ABSENCE_DATES and calendar-detected names overlap
- Email received date on tiles — safe pywintypes.datetime parsing → "9 Jun" format, right-aligned on card
- Email sort — newest-first within each category (Urgent, Needs Response, FYI, Low)
- index.html — ALL garbled Unicode fully resolved (48 characters total across two passes)
- Font sizes — increased throughout to match Command Centre scale
- open_email.py — openmail:// protocol registered, confirmed working
- Task Scheduler — 3 tasks: `WorkInbox-0900`, `WorkInbox-1200`, `WorkInbox-1500` (Mon–Fri). Confirmed working 2026-06-30. Recovery: run `create_inbox_tasks.bat` from repo root as Administrator.
- Dashboard loads live briefing.json from GitHub on load, falls back to localStorage archive
- Oxford navy sidebar (#002147, 340px) with crest, branding, calendar, absences
- **Live clock in sidebar** — `wi-clock-time` shows ticking HH:MM:SS above the live day/date in the Today block
- Time-of-day greeting (Good morning/afternoon/evening, Kevin) — UK timezone
- Archive panel — past briefings by date, Load arrow to restore
- Yellow accent bar for Needs Response section
- Context bar — 5-7 sentence specific briefing
- Card click-through — whole tile clickable via openmail://, hover shadow
- Tick to hide — cards and priority rows fade then hide after 1500ms
- Show Done / Hide Done — boolean flag, fully working
- Absences — white bullet list, text justified
- Fuzzy EntryID matching — fallback if AI subject slightly differs
- Badge CSS normalised — NEW (green) and UPDATED (blue) badges identical to Command Centre
- "Open email" buttons — emoji removed from all instances
- "Priority actions — next week" section (pnw drop zone) — sits between "this week" and "urgent"; teal dot accent
- Drag-and-drop fully re-enabled and extended — email cards draggable into any priority section; custom priority items persisted in `workInbox_customPri_v1`
- Multi-machine setup complete (begb0037.AD-OAK)
- **CC→ button window reuse (2026-06-30):** `_ccWindow` variable stores CC window reference; navigates existing tab in-place rather than opening a new one every click.

### Local Path
Folder: `C:\Users\admin\Documents\Claude\Projects\work-inbox`
Task Scheduler, Registry (openmail://), and desktop bat all updated 2026-06-09.

### Critical Note — Desktop Bat File
- Bat file target: `C:\Users\admin\Desktop\Run Inbox Briefing.bat`
- **Never rename an existing bat file** — always download fresh from GitHub via PowerShell:
  `Invoke-WebRequest -UseBasicParsing "https://raw.githubusercontent.com/begb0037admin/work-inbox/main/Run_Inbox_Briefing.bat" -OutFile "$env:USERPROFILE\Desktop\Run Inbox Briefing.bat"`
- Renaming an existing file only changes the name, never the content — stale content (BOM, wrong path) will persist.
- Bat auto-pulls latest fetch_inbox.py from raw.githubusercontent.com (with `?t=` cache-buster) before every run — no git dependency.

---

## localStorage Keys (index.html)

| Key | Purpose |
|-----|--------|
| `workInbox_briefings_v1` | Archive of past briefing JSON objects, keyed by date string |
| `workInbox_today_v1` | Key of the currently displayed briefing |
| `workInbox_ticks_v1` | Tick (done) state for all cards |
| `workInbox_priOverrides_v1` | Per-card section overrides for priority drag-and-drop |
| `workInbox_priOrder_v1` | Per-section sort order for priority cards |
| `workInbox_customPri_v1` | Email cards manually dragged into priority sections (persisted across refreshes) |

---

## Technical Notes — index.html

### Mojibake Fix (critical for future edits)
index.html was triple-encoded (UTF-8 → Latin-1 misread x3), producing 12-byte sequences per intended character.
Two-pass fix applied 2026-06-09:
- Pass 1: em-dash (U+2014) and down-triangle (U+25BC) — 37 instances
- Pass 2: en-dash (U+2013), rightwards arrow (U+2192), curly quotes (U+201C/D), checkmark (U+2713), white circle (U+25CB), up-triangle (U+25B2), warning sign (U+26A0) — 11 instances
Total: 48 instances fixed. Zero triple-encoding patterns remain.

**Rule: Any future edit to index.html MUST use the binary atob()/btoa() approach — NEVER TextEncoder on the file content.** TextEncoder will re-encode the multi-byte Unicode chars and re-garble them.

### CSS Override Pattern
Font size increases were added as CSS overrides at the end of the style block (before `</style>`). Safest pattern for future edits — existing rules untouched, cascade handles priority.

### Priority Drag-and-Drop Architecture
- Five named sections: `pt` (today), `pw` (this week), `pnw` (next week), `ur` (urgent overlay), `nr` (needs overlay)
- Sections always render — content not gated by briefing data
- Two drag types handled: priority cards (`_priDragState`) and email cards (`_emailDragData`)
- Email card drop → `_addEmailCardToPriority()` → stores in `workInbox_customPri_v1` + sets override
- `applyPriOverrides(data)` merges AI items (prioritiesToday/prioritiesWeek) + custom items, then applies overrides and order

---

## Prompt Standards (fetch_inbox.py)

### Context field
5-7 sentences. Must include: full names and exact return dates of every absent colleague; specific projects/systems/cases blocked; most time-critical deadline with exact date; emails waiting 48hrs+; one thing Kevin should open first. No generalisations. No GitHub PAT/CI/workflow mentions.

### Subject field
Copy exact email subject verbatim. Fuzzy matching fallback in Python if slight drift occurs.

---

## File Locations

| File | Location |
|------|----------|
| Repo | github.com/begb0037admin/work-inbox |
| Proxy | github-proxy.lelitte.co.uk/work-inbox/ |
| Dashboard (GitHub Pages) | begb0037admin.github.io/work-inbox/ |
| Dashboard (primary) | wi.lelitte.co.uk |
| Styles | work-inbox/css/styles.css |
| JS | work-inbox/js/app.js |
| Script | work-inbox/fetch_inbox.py |
| Opener | work-inbox/open_email.py |
| Briefing | work-inbox/data/briefing.json |
| Local | C:\Users\admin\Documents\Claude\Projects\work-inbox\ |
| Registry | HKCU:\Software\Classes\openmail (points to new path) |
| Scheduler | WorkInbox-0900 / WorkInbox-1200 / WorkInbox-1500 (confirmed working 2026-06-30) |
| Scheduler recovery | `create_inbox_tasks.bat` in repo root — run as Administrator |
| Desktop bat | `C:\Users\admin\Desktop\Run Inbox Briefing.bat` — download fresh via PowerShell, never rename |

---

## AG FlexPoints — see its own repo
AG FlexPoints has its own repo: **begb0037admin/AG-FlexPoints** with its own HANDOVER.md.

---

## Roadmap

No outstanding items as of 2026-06-30.

---

## Standing Rules
- Never commit tokens or raw data
- All GitHub writes via Contents API (PAT from GITHUB_PAT env var)
- Every change committed immediately
- Seat A never references local disk — all reads via GitHub proxy
- Local machine must stay in sync: bat auto-pulls fetch_inbox.py from raw.githubusercontent.com before every run
- index.html edits: always use binary atob()/btoa() — NEVER TextEncoder on file content
- Desktop bat: always download fresh via PowerShell — never rename an existing file
