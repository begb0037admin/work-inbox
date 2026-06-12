# work-inbox — Living Handover Document

**Last updated:** 2026-06-11 (evening)
**Status:** Active — pipeline working but scheduled runs unreliable (see Known Issues).

---

## Architecture

| Component | Description |
|-----------|-------------|
| fetch_inbox.py | Outlook COM via pywin32. Pulls inbox → Anthropic triage (claude-haiku-4-5) → Python post-processing → pushes data/briefing.json to GitHub via Contents API |
| index.html | Static GitHub Pages dashboard at begb0037admin.github.io/work-inbox/ |
| open_email.py | Registered openmail:// protocol handler — strips prefix and trailing slash, opens exact email in classic Outlook via EntryID COM |

---

## Current State (fully working as of 2026-06-09 late evening)

### Working
- fetch_inbox.py — all three phases confirmed working
- Calendar pull — direct iteration with 30-day lookback + 6-day forward window (confirmed calendar:27 items); Restrict() approach abandoned
- Python post-processing of calendar items — KNOWN_ABSENCES list rewrites sub/alert for Marie Cooksey and James Salas Guillen with hardcoded specific text, bypassing AI entirely for those fields
- Absence dedup — partial-name matching prevents duplicates when KNOWN_ABSENCE_DATES and calendar-detected names overlap (e.g. "Marie" + "Marie Cooksey")
- Email received date on tiles — safe pywintypes.datetime parsing → "9 Jun" format, right-aligned on card
- Email sort — newest-first within each category (Urgent, Needs Response, FYI, Low)
- index.html — ALL garbled Unicode fully resolved (48 characters total across two passes):
  - Pass 1: 37 instances — em-dash (—) and triangle (▼) in section headings
  - Pass 2: 11 instances — en-dash (–), rightwards arrow (→), curly quotes (" "), checkmark (✓), white circle bullet (○), up-triangle (▲), warning sign (⚠)
  - Zero triple-encoding patterns remaining. briefing.json confirmed clean (no encoding issues).
- Font sizes — increased throughout to match Command Centre scale: sidebar section headers 13px, calendar titles 15px, absence list 14px, section labels 14px, card titles 16px, user/date values 15px
- open_email.py — openmail:// protocol registered, confirmed working
- Task Scheduler — WorkInbox-Briefing runs at 7am/9am/11am/1pm/3pm/5pm Mon-Fri
- Dashboard loads live briefing.json from GitHub on open, falls back to localStorage archive
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

### Local Path
Folder: `C:\Users\admin\Documents\Claude\Projects\work-inbox`
Task Scheduler, Registry (openmail://), and desktop bat all updated 2026-06-09.

### Critical Note — Local Script
- Task Scheduler runs local fetch_inbox.py — must stay in sync with GitHub
- If calendar items revert to vague text, run: git fetch origin && git checkout origin/main -- fetch_inbox.py
- Then run: python fetch_inbox.py

### Known Issues / Next Session
- **GITHUB_PAT expired 2026-06-11 (~after 1pm):** root cause of the missed 3pm/5pm runs — all GitHub pushes return 401 Unauthorized. Outlook pull and Anthropic triage unaffected. Fix: generate a new PAT (must cover work-inbox AND command-centre repos), update the GITHUB_PAT user env var on the admin machine. If tick sync stops working, the Cloudflare Worker cc-tasks-writer may hold the same PAT and need updating too.
- **Scheduled 7am run missing (2026-06-11):** machine likely asleep/locked. Check Task Scheduler → WorkInbox-Briefing → History; enable "Wake the computer to run this task" on the Conditions tab.
- **Local git repo unreliable (2026-06-11):** `git fetch` on the admin machine triggered gc which hung on "Deletion of directory '.git/objects/00' failed" — likely OneDrive/antivirus locking `.git` (repo lives under Documents). Both bat files failed as a result. Fix applied: `Run_Inbox_Briefing.bat` no longer uses git — it downloads fetch_inbox.py directly from raw.githubusercontent.com (with `?t=` cache-buster) then runs it. The push to GitHub happens inside fetch_inbox.py via the Contents API, so git is not needed at all for a refresh. Consider excluding the project folder from OneDrive backup, and re-download the bat to the local machine (the old local copy still uses git).
- Task Scheduler bat file should auto-pull latest fetch_inbox.py before running (use the same no-git download approach as Run_Inbox_Briefing.bat)
- James Salas Guillen not appearing in calendar (no Outlook calendar block) — post-processing only fires if AI includes him; monitor until he returns 18 Jun
- Multi-machine setup not yet done (work machine begb0037.AD-OAK)

---

## KNOWN_ABSENCES — Hardcoded Calendar Context

These are hardcoded in fetch_inbox.py and rewrite AI output for known absent colleagues:

**Marie Cooksey** (leave 8–13 June)
- sub: "Marie is on leave 8-13 June. Any items requiring her approval or sign-off must wait until she returns. Kevin and Chris are covering H&S support queue and OSM escalations."
- alert: "Marie unavailable all week - action DTP1092 comments and volunteer reporting queries independently"

**James Salas Guillen** (leave until 18 June)
- sub: "James is on leave until 18 June. DSE/Cardinus archiving, SBS users in feed and applicant data work all on hold. Handover document received Fri 6 Jun."
- alert: "James away until 18 June - Kevin and Chris covering OSM tickets and H&S support queue"

**Action:** Update/remove these entries when absences end.

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
| Dashboard | begb0037admin.github.io/work-inbox/ |
| Script | work-inbox/fetch_inbox.py |
| Opener | work-inbox/open_email.py |
| Briefing | work-inbox/data/briefing.json |
| Local | C:\Users\admin\Documents\Claude\Projects\work-inbox\ |
| Registry | HKCU:\Software\Classes\openmail (points to new path) |
| Scheduler | WorkInbox-Briefing (Task Scheduler) |

---

## Session 2026-06-12 — AG FlexPoints scaffolded

New sister project built in `ag-flexpoints/` on branch `claude/determined-allen-zanfr0`
(PR open): recurring report on Access Group FlexPoints, destined for its own repo
**begb0037admin/AG-FlexPoints** (could not be created from the cloud session —
GitHub App token is scoped to work-inbox only; Kevin must create the repo and
grant the Claude GitHub App access to it).

- `ag-flexpoints/fetch_flexpoints.py` — Playwright login to accessgroup.my.site.com
  (ACCESS_PORTAL_USER / ACCESS_PORTAL_PASSWORD Windows env vars — same
  username/password flow as the hr-fa-knowledge-base scrape; lift its login code
  if the generic selectors miss), claude-haiku-4-5 parses pages → full report JSON
  → pushed via Contents API. Fallback: parse newest page saved into `source/`.
- `ag-flexpoints/index.html` — hybrid dashboard: Access portal layout (summary
  tiles, Quote Received, Requested, Booked, Transaction History) in work-inbox
  style (navy sidebar, context bar). CSV + PDF/print export.
- `ag-flexpoints/data/flexpoints.json` — seeded with real data from Kevin's
  2026-06-12 portal screenshot (2,935 pts available, ALL expiring 2026-06-29;
  quote 69001638 for 2,800 pts awaiting approval).
- Portal is unreachable from cloud agents (403) — local script is the only bridge.

## Roadmap

| Priority | Task |
|----------|------|
| 0 | AG FlexPoints: create repo begb0037admin/AG-FlexPoints, move ag-flexpoints/ contents to its root, enable Pages, set ACCESS_PORTAL_USER/ACCESS_PORTAL_PASSWORD env vars, pip install playwright, schedule weekly run |
| 1 | Update Task Scheduler bat to auto-pull fetch_inbox.py before running |
| 2 | Add James Salas Guillen to Outlook calendar so he appears without AI inference |
| 3 | Monitor card openmail:// click-through reliability |
| 4 | Multi-machine — replicate setup on work machine (begb0037.AD-OAK) |
| 5 | Update KNOWN_ABSENCES when Marie returns 13 Jun and James returns 18 Jun |

---

## Standing Rules
- Never commit tokens or raw data
- All GitHub writes via Contents API (PAT from GITHUB_PAT env var)
- Every change committed immediately
- Seat A never references local disk — all reads via GitHub proxy
- Local machine must stay in sync: git fetch origin && git checkout origin/main -- fetch_inbox.py
- index.html edits: always use binary atob()/btoa() — never TextEncoder on file content
