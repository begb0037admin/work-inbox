# work-inbox — Living Handover Document

**Last updated:** 2026-06-08 (session 2 -- complete)
**Status:** Active -- fully working. Task Scheduler fixed and tested. All session 2 changes committed.

---

## Architecture

| Component | Description |
|-----------|-------------|
| fetch_inbox.py | Outlook COM via pywin32. Pulls inbox → Anthropic triage (claude-haiku-4-5) → Python post-processing → pushes data/briefing.json to GitHub via Contents API |
| index.html | Static GitHub Pages dashboard at begb0037admin.github.io/work-inbox/ |
| open_email.py | Registered openmail:// protocol handler — strips prefix and trailing slash, opens exact email in classic Outlook via EntryID COM |

---

## Current State (fully working as of 2026-06-08 late evening)

### Working
- fetch_inbox.py — all three phases confirmed working
- Python post-processing of calendar items — KNOWN_ABSENCES list rewrites sub/alert for Marie Cooksey and James Salas Guillen with hardcoded specific text, bypassing AI entirely for those fields
- open_email.py — openmail:// protocol registered, confirmed working
- Task Scheduler — WorkInbox-Briefing runs at 7am/9am/11am/1pm/3pm/5pm Mon-Fri
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

### Local Path (updated 2026-06-09)
Folder moved to: `C:\Users\admin\Documents\Claude\Projects\work-inbox`
Task Scheduler, Registry (openmail://), and desktop bat all updated by Seat C.

### Critical Note — Local Script
- Task Scheduler runs local fetch_inbox.py — must stay in sync with GitHub
- If calendar items revert to vague text, run: git fetch origin && git checkout origin/main -- fetch_inbox.py
- Then run: python fetch_inbox.py

### Known Issues / Next Session
- Task Scheduler bat file should auto-pull latest fetch_inbox.py before running
- James Salas Guillen not appearing in calendar (no Outlook calendar block) — post-processing only fires if AI includes him; monitor
- Multi-machine setup not yet done (work machine begb0037.AD-OAK)

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
| Local | C:\Users\admin\Documents\Claude\Projects\work-inbox\ |
| Registry | HKCU:\Software\Classes\openmail (points to new path) |
| Scheduler | WorkInbox-Briefing (Task Scheduler) |

---

## Roadmap

### Completed -- 2026-06-08

**3-phase script restore**
fetch_inbox.py in GitHub contained Phase 1 only -- it pulled from Outlook and stopped with no API triage, no briefing.json output, and no GitHub push. Dashboard was showing stale data on every scheduled run because the working 3-phase script existed only as a local backup. Fixed by restoring the full script from fetch_inbox_backup_2026-06-09.py. Broken version preserved as fetch_inbox_backup_2026-06-08.py for rollback if needed.

**Sort fix -- unread newest-first**
Emails were sorted unread-first (correct) but oldest-first within each group (wrong). Emails arriving this morning were landing at the bottom of the unread section. Fixed: both unread and read groups now sort newest-first, so the most recent emails always appear at the top of their group.

**Restrict filter date format**
Outlook COM Restrict filter requires 12-hour US locale date format. The script was passing a mixed 24-hour/AM-PM format that could cause silent filter failure on afternoon runs, falling back to a full inbox scan. Fixed to %m/%d/%Y %I:%M %p.

**Task Scheduler -- broken path fixed, auto-pull baked in**
Task Scheduler was pointing at C:\Users\admin\work-inbox -- a path that no longer exists after the folder was moved to Documents\Claude\Projects\work-inbox. Every scheduled run had been failing with error -2147024629 (ERROR_DIRECTORY -- invalid directory). Dashboard was showing stale data all day as a result. Fixed: task deleted and recreated via elevated PowerShell with correct working directory and git auto-pull baked directly into the task command. GitHub is now always the source of truth before every scheduled execution. Test run confirmed: Last Result 0 at 16:20:32 on 2026-06-08. Task command: cmd.exe /c cd /d "C:\Users\admin\Documents\Claude\Projects\work-inbox" && git fetch origin && git checkout origin/main -- fetch_inbox.py && C:\Python314\python.exe fetch_inbox.py. Note: recreating the task requires elevated PowerShell (Run as Administrator) -- standard session gives Access Denied.

### Outstanding

**1. Unactioned thread logic** -- next priority
Currently the dashboard uses read/unread status to decide what to surface. Read does not mean actioned -- an email may have been opened and never replied to. The fix is to cross-reference sent items (already pulled by the script) to identify threads with no reply from Kevin. Those unactioned threads get a 30-day lookback instead of 7 days so nothing drops off the dashboard just because it is older than a week. The Conor O'Brien / Access Group email from 21 May 2026 is a real example of this failure. Affects: fetch_inbox.py Phase 1 inbox pull logic and the API triage prompt.

**2. Inbox sort -- DO NOT RUN until item 3 is complete**
A one-off task to file the entire inbox into named folders: Simon, Marie, Projects, and others. A Chrome brief has already been written and is ready to run. Hard blocker: the dashboard currently reads only the default inbox folder. If the sort runs before item 3 is done, any email moved to a subfolder vanishes from the dashboard silently with no error message. Affects: fetch_inbox.py Phase 1, dashboard coverage.

**3. Subfolder scanning -- prerequisite for item 2**
Update fetch_inbox.py to scan named subfolders in addition to the main inbox. Required before the inbox sort runs. Approximately 15 lines of additional code in Phase 1. Affects: fetch_inbox.py Phase 1 only.

**4. James Salas Guillen calendar**
James is absent and appears on the dashboard only because his details are hardcoded into fetch_inbox.py. This requires manual script updates whenever his absence changes. Root cause: no Outlook calendar block exists for his leave. Fix: add him to the Outlook calendar so the script picks him up automatically the same way it handles Marie Cooksey. Affects: fetch_inbox.py KNOWN_ABSENCES block and Outlook calendar.

**5. Click-through reliability monitoring**
The openmail:// protocol uses a unique EntryID to open the exact email in Outlook when a dashboard card is clicked. A fuzzy subject-matching fallback handles cases where the AI slightly rephrases a subject. Not yet tested across all email types -- long subjects, special characters, CC threads, calendar invites. No action required now -- flag any click-through failures as they occur and report the exact subject string that caused the mismatch. Affects: open_email.py, EntryID injection in fetch_inbox.py.

**6. Work machine replication**
The full setup -- Task Scheduler, openmail:// registry entry, Python environment, local script folder -- exists only on the admin machine (C:\Users\admin). The work machine (begb0037.AD-OAK) has none of it. If working from the office, scheduled runs do not happen. Requires replicating the complete setup on the work machine. Affects: all components, new machine.

**7. command-centre ROADMAP.md**
The command-centre repo is the master tracker across all HR systems work. Its ROADMAP.md has not been updated to reflect work-inbox project growth. Low priority -- nothing breaks -- but the command-centre view of the project landscape is stale. Affects: command-centre repo only.

---
## Standing Rules
- Never commit tokens or raw data
- All GitHub writes via Contents API (PAT from GITHUB_PAT env var)
- Every change committed immediately
- Seat A never references local disk — all reads via GitHub proxy
- Local machine must stay in sync: git fetch origin && git checkout origin/main -- fetch_inbox.py
