# work-inbox — Living Handover Document

**Last updated:** 2026-06-08
**Status:** Active — dashboard fully functional. Session ended 2026-06-08 evening.

---

## Architecture

| Component | Description |
|-----------|-------------|
| fetch_inbox.py | Outlook COM via pywin32. Pulls inbox → Anthropic triage (claude-haiku-4-5) → pushes data/briefing.json to GitHub via Contents API |
| index.html | Static GitHub Pages dashboard at begb0037admin.github.io/work-inbox/ |
| open_email.py | Registered openmail:// protocol handler — strips prefix and trailing slash, opens exact email in classic Outlook via EntryID COM |

---

## Current State (as of 2026-06-08 evening, commit 097a9e4)

### Working
- fetch_inbox.py — all three phases confirmed working (Outlook pull, Anthropic triage, GitHub push)
- open_email.py — openmail:// protocol registered, strips trailing slash, opens exact email in Outlook
- Task Scheduler — WorkInbox-Briefing runs at 7am/9am/11am/1pm/3pm/5pm Mon-Fri
- Dashboard loads live briefing.json from GitHub on open, falls back to localStorage archive
- Oxford navy sidebar (#002147, 340px) with crest, branding, calendar, absences
- Time-of-day greeting (Good morning/afternoon/evening, Kevin) — UK timezone
- Archive panel — past briefings by date, Load arrow to restore
- Yellow accent bar for Needs Response section
- Context bar — 15px font, richer 5-7 sentence specific briefing (names, dates, cases, no PAT/CI mentions)
- Calendar items — specific sub and alert text with correct/wrong examples baked into prompt
- AI cross-references OOO emails and handover emails to infer absences not in calendar
- Card click-through — whole tile clickable via openmail://, hover shadow, checkbox isolated
- Tick to hide — ticking a card fades then hides after 1500ms
- Priority actions — tick fades then hides after 1500ms
- Show Done button — reveals all hidden ticked items for untick or reference
- Absences — single bullet list block in white text, text justified (not individual pills)
- Fuzzy EntryID matching — fallback if AI subject slightly differs from inbox subject
- Verbatim subject prompt — AI instructed to copy exact email subject character for character

### Known Issues / Next Session
- Calendar items still occasionally too vague despite prompt — monitor over next few runs
- Card openmail:// click-through dependent on entry_id being present in briefing.json — verify on next run
- Multi-machine setup not yet done (work machine begb0037.AD-OAK)

---

## Prompt Standards (fetch_inbox.py)

### Context field
5-7 sentences. Must include: full names and exact return dates of every absent colleague; specific projects/systems/cases blocked; most time-critical deadline with exact date; emails waiting 48hrs+; one thing Kevin should open first. No generalisations. No GitHub PAT/CI/workflow mentions.

### Calendar items
- time: "All day" for all-day events, never date ranges
- title: "Event Type - Full Name" e.g. "Annual Leave - Marie Cooksey"
- sub: specific — exact dates, what is blocked, who is covering. No vague text.
- alert: name specific projects/actions affected. No generic text like "Colleague absent" or "Team member absent".
- Cross-reference OOO emails and handover emails to infer absences not in calendar blocks
- Correct/wrong examples baked into prompt to enforce standard

### Subject field
Copy exact email subject verbatim — character for character. Do not paraphrase. Fuzzy matching fallback in Python if slight drift occurs.

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
| Local | C:\Users\admin\work-inbox\ |
| Registry | HKCU:\Software\Classes\openmail |
| Scheduler | WorkInbox-Briefing (Task Scheduler) |

---

## Roadmap

| Priority | Task |
|----------|------|
| 1 | Monitor calendar item quality over next few morning runs — refine prompt if still vague |
| 2 | Verify card openmail:// click-through working reliably with fuzzy matching |
| 3 | Multi-machine — replicate setup on work machine (begb0037.AD-OAK) |
| 4 | Update command-centre ROADMAP.md |
| 5 | Show Done — verify priority rows restore correctly on untick |

---

## Standing Rules
- Never commit tokens or raw data
- fetch_inbox.py and open_email.py are correct — do not touch without explicit instruction
- All GitHub writes via Contents API (PAT from GITHUB_PAT env var) — never git push from Cowork
- Every change committed immediately — no batching
- Seat A never references local disk — all reads via GitHub proxy
