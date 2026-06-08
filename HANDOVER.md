# work-inbox — Living Handover Document

**Last updated:** 2026-06-08 (evening session 2)
**Status:** Active — fully working. Task Scheduler recreated with git pull baked in. Test run result pending verification.

---

## Architecture

| Component | Description |
|-----------|-------------|
| fetch_inbox.py | Outlook COM via pywin32. Pulls inbox → Anthropic triage (claude-haiku-4-5) → Python post-processing → pushes data/briefing.json to GitHub via Contents API |
| index.html | Static GitHub Pages dashboard at begb0037admin.github.io/work-inbox/ |
| open_email.py | Registered openmail:// protocol handler — strips prefix and trailing slash, opens exact email in classic Outlook via EntryID COM |

---

## Current State (as of 2026-06-08 evening session 2)

### Working
- fetch_inbox.py — all three phases confirmed working
- Python post-processing of calendar items — KNOWN_ABSENCES list rewrites sub/alert for Marie Cooksey and James Salas Guillen with hardcoded specific text, bypassing AI entirely for those fields
- open_email.py — openmail:// protocol registered, confirmed working
- Task Scheduler — WorkInbox-Briefing runs at 7am/9am/11am/1pm/3pm/5pm Mon-Fri
  - Recreated 2026-06-08 with correct path and git pull baked in:
  - `cmd.exe /c cd /d "C:\Users\admin\Documents\Claude\Projects\work-inbox" && git fetch origin && git checkout origin/main -- fetch_inbox.py && C:\Python314\python.exe fetch_inbox.py`
  - Previous task was silently failing with error -2147024629 (ERROR_DIRECTORY) — was pointing at old dead path `C:\Users\admin\work-inbox` since folder moved months ago
  - Test run initiated via `schtasks /run /tn "WorkInbox-Briefing"` — result not yet verified
- Dashboard loads live briefing.json from GitHub on open, falls back to localStorage archive
- Oxford navy sidebar (#002147, 340px) with crest, branding, calendar, absences
- Time-of-day greeting (Good morning/afternoon/evening, Kevin) — UK timezone
- Archive panel — past briefings by date, Load arrow to restore
- Yellow accent bar for Needs Response section
- Context bar — 15px font, 5-7 sentence specific briefing
- Calendar items — Python hardcodes specific sub/alert for known absent colleagues
- Card click-through — whole tile clickable via openmail://, hover shadow
- Tick to hide — cards and priority rows fade then hide after 1500ms
- Show Done / Hide Done — boolean flag, fully working
- Absences — white bullet list, text justified
- Fuzzy EntryID matching — fallback if AI subject slightly differs

### Local Path
Folder: `C:\Users\admin\Documents\Claude\Projects\work-inbox`
Task Scheduler, Registry (openmail://), and desktop bat all updated.

### Critical Note — Local Script
- Task Scheduler now auto-pulls latest fetch_inbox.py before running (baked into task command)
- If calendar items revert to vague text, run manually: `git fetch origin && git checkout origin/main -- fetch_inbox.py && python fetch_inbox.py`

### Known Issues / Pending
- Task Scheduler test run result not yet verified — check: `schtasks /query /tn "WorkInbox-Briefing" /fo LIST /v | Select-String "Last Run Time|Last Result|Status"`
  - If Last Result = 0: working. If non-zero: diagnose error code.
- James Salas Guillen not appearing in calendar (no Outlook calendar block) — post-processing only fires if AI includes him; monitor
- Multi-machine setup not yet done (work machine begb0037.AD-OAK)
- HANDOVER.md and CLAUDE.md were not updated during the session that recreated Task Scheduler — this commit is that update

---

## KNOWN_ABSENCES — Hardcoded Calendar Context

These are hardcoded in fetch_inbox.py and rewrite AI output for known absent colleagues:

**Marie Cooksey**
- sub: "Marie is on leave 8-13 June. Any items requiring her approval or sign-off must wait until she returns. Kevin and Chris are covering H&S support queue and OSM escalations."
- alert: "Marie unavailable all week - action DTP1092 comments and volunteer reporting queries independently"

**James Salas Guillen**
- sub: "James is on leave until 18 June. DSE/Cardinus archiving, SBS users in feed and applicant data work all on hold. Handover document received Fri 6 Jun."
- alert: "James away until 18 June - Kevin and Chris covering OSM tickets and H&S support queue"

Update these when absences change.

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
| Backup (working) | work-inbox/fetch_inbox_backup_2026-06-09.py |
| Backup (broken Phase 1 only) | work-inbox/fetch_inbox_backup_2026-06-08.py |
| Local | C:\Users\admin\Documents\Claude\Projects\work-inbox\ |
| Registry | HKCU:\Software\Classes\openmail (points to current path) |
| Scheduler | WorkInbox-Briefing (Task Scheduler, Run As: DESKTOP-MJDJM64\admin) |

---

## Roadmap

| Priority | Task |
|----------|------|
| 1 | Verify Task Scheduler test run result (Last Result = 0 expected) |
| 2 | Unactioned thread logic -- cross-reference sent items, 30-day lookback for no-reply threads |
| 3 | Subfolder scanning -- prereq for inbox sort (emails filed to subfolders vanish from dashboard without this) |
| 4 | Inbox sort -- blocked on subfolder scanning |
| 5 | Add James Salas Guillen to Outlook calendar so he appears without AI inference |
| 6 | Monitor card openmail:// click-through reliability |
| 7 | Multi-machine -- replicate setup on work machine (begb0037.AD-OAK) |
| 8 | Update command-centre ROADMAP.md |
| ~~COMPLETE~~ | ~~Task Scheduler auto-pull -- git pull now baked into task command directly~~ |

---

## Standing Rules
- Never commit tokens or raw data
- All GitHub writes via Contents API (PAT from GITHUB_PAT env var)
- Every change committed immediately
- Seat A never references local disk -- all reads via GitHub proxy or API
- Local machine must stay in sync: git fetch origin && git checkout origin/main -- fetch_inbox.py
- Task Scheduler requires elevated PowerShell (Run as Administrator) to Register or Unregister tasks
- Unicode em dashes in repo files silently break PowerShell .Replace() -- use split-on-section-headers approach for doc edits
- No em dashes in PowerShell commit messages -- plain hyphens only
