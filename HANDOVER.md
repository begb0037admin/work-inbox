# work-inbox — Living Handover Document

**Last updated:** 2026-06-08
**Status:** Active development — dashboard visual rebuild in progress.

---

## Architecture

| Component | Description |
|-----------|-------------|
| fetch_inbox.py | Outlook COM via pywin32. Pulls inbox → Anthropic triage (claude-haiku-4-5) → pushes data/briefing.json to GitHub via Contents API |
| index.html | Static GitHub Pages dashboard at begb0037admin.github.io/work-inbox/ |
| open_email.py | Registered openmail:// protocol handler — strips prefix and trailing slash, opens exact email in classic Outlook via EntryID COM |

---

## Current State (as of 2026-06-08, commit fa505b6)

### Working
- fetch_inbox.py — all three phases (Outlook pull, Anthropic triage, GitHub push) confirmed working
- open_email.py — openmail:// protocol registered, strips trailing slash, opens exact email in Outlook
- Task Scheduler — WorkInbox-Briefing runs at 7am/9am/11am/1pm/3pm/5pm Mon-Fri
- Dashboard loads live briefing.json from GitHub on open, falls back to localStorage
- Oxford navy sidebar (#002147, 340px) with crest, branding, calendar, absences
- Time-of-day greeting (Good morning/afternoon/evening, Kevin) — UK timezone
- Archive panel — past briefings by date, Load arrow to restore
- Yellow accent for Needs Response section
- Context bar — richer 5-7 sentence specific briefing (names, dates, cases)
- Calendar items — specific sub and alert text (no generic phrases)

### Broken / In Progress (next session)
- Card click-through to Outlook — whole tile should be clickable via openmail://, hover shadow effect. entry_id present in briefing.json but not wired in current render
- Tick to hide — ticking a card should fade then hide after 1500ms. Currently only adds .done class, no hide
- Show Done button — should appear in header, reveals all hidden ticked items for untick or reference. Currently missing entirely

---

## Prompt Standards (fetch_inbox.py)

### Context field
5-7 sentences. Must include: full names and exact return dates of every absent colleague; specific projects/systems/cases blocked; most time-critical deadline with exact date; emails waiting 48hrs+; one thing Kevin should open first. No generalisations. No GitHub PAT/CI/workflow mentions.

### Calendar items
- time: "All day" for all-day events, never date ranges
- title: "Event Type - Full Name" e.g. "Annual Leave - Marie Cooksey"
- sub: specific — exact dates, what is blocked, who is covering. No vague text.
- alert: name specific projects/actions affected. No generic text like "Colleague absent".

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
| 1 | Reinstate card click-through (openmail://, whole tile, hover shadow) |
| 2 | Reinstate tick → fade → hide after 1500ms |
| 3 | Reinstate Show Done button — reveals hidden items for untick/reference |
| 4 | Multi-machine — replicate setup on work machine (begb0037.AD-OAK) |
| 5 | Update command-centre ROADMAP.md |

---

## Standing Rules
- Never commit tokens or raw data
- fetch_inbox.py and open_email.py are correct — do not touch without explicit instruction
- All GitHub writes via Contents API (PAT from GITHUB_PAT env var) — never git push from Cowork
- Every change committed immediately — no batching
