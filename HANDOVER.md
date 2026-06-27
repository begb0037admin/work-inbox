# work-inbox — Living Handover Document

**Last updated:** 2026-06-27 (file split complete — wi.lelitte.co.uk live)
**Status:** Active — pipeline fully working, roadmap clear.

---

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
|---|---|---|
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
| js/app.js | All JS — briefing render, drag-and-drop, tick sync, archive, tasks widget. Both script blocks merged. |
| open_email.py | Registered openmail:// protocol handler — strips prefix and trailing slash, opens exact email in classic Outlook via EntryID COM |

---

## Current State (fully working as of 2026-06-25)

### Working
- fetch_inbox.py — all three phases confirmed working
- Calendar pull — direct iteration with 30-day lookback + 6-day forward window (confirmed calendar:27 items); Restrict() approach abandoned
- Python post-processing of calendar items — KNOWN_ABSENCES list cleared (Marie and James both returned)
- Absence dedup — partial-name matching prevents duplicates when KNOWN_ABSENCE_DATES and calendar-detected names overlap
- Email received date on tiles — safe pywintypes.datetime parsing → "9 Jun" format, right-aligned on card
- Email sort — newest-first within each category (Urgent, Needs Response, FYI, Low)
- index.html — ALL garbled Unicode fully resolved (48 characters total across two passes)
- Font sizes — increased throughout to match Command Centre scale
- open_email.py — openmail:// protocol registered, confirmed working
- Task Scheduler — WorkInbox-Briefing runs at 7am/9am/11am/1pm/3pm/5pm Mon-Fri; bat auto-pulls fetch_inbox.py before running
- Dashboard loads live briefing.json from GitHub on load, falls back to localStorage archive
- Oxford navy sidebar (#002147, 340px) with crest, branding, calendar, absences
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
- "Priority actions – next week" section (pnw drop zone) — sits between "this week" and "urgent"; teal dot accent
- Drag-and-drop fully re-enabled and extended — email cards draggable into any priority section; custom priority items persisted in `workInbox_customPri_v1`
- Multi-machine setup complete (begb0037.AD-OAK)

### Local Path
Folder: `C:\Users\admin\Documents\Claude\Projects\work-inbox`
Task Scheduler, Registry (openmail://), and desktop bat all updated 2026-06-09.

### Critical Note — Local Script
- Task Scheduler runs local fetch_inbox.py — must stay in sync with GitHub
- Bat auto-pulls latest fetch_inbox.py from raw.githubusercontent.com (with `?t=` cache-buster) before every run — no git dependency

---

## localStorage Keys (index.html)

| Key | Purpose |
|-----|---------|
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
Font size increases were added as CSS overrides at the end of the style block (before </style>). Safest pattern for future edits — existing rules untouched, cascade handles priority.

### Priority Drag-and-Drop Architecture
- Five named sections: `pt` (today), `pw` (this week), `pnw` (next week), `ur` (urgent overlay), `nr` (needs overlay)
- Sections always render — not gated by briefing data
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
| Scheduler | WorkInbox-Briefing (Task Scheduler) |

---

## AG FlexPoints — see its own repo
AG FlexPoints has its own repo: **begb0037admin/AG-FlexPoints** with its own HANDOVER.md.

---

## Roadmap

All items resolved as of 2026-06-25. File split and custom domain completed 2026-06-27. No outstanding work.

---

## Standing Rules
- Never commit tokens or raw data
- All GitHub writes via Contents API (PAT from GITHUB_PAT env var)
- Every change committed immediately
- Seat A never references local disk — all reads via GitHub proxy
- Local machine must stay in sync: bat auto-pulls fetch_inbox.py from raw.githubusercontent.com before every run
- index.html edits: always use binary atob()/btoa() — never TextEncoder on file content
