# CHAT_PROMPT.md — work-inbox

> This file documents how to interact with the work-inbox system.
> The system is fully automated — no manual email pasting required.

## How It Works

The dashboard runs automatically. `fetch_inbox.py` is scheduled via Windows Task Scheduler to run at 7am, 9am, 11am, 1pm, 3pm, 5pm Monday-Friday on the admin machine.

Each run:
1. Pulls the 50 most recent emails from Outlook inbox + sent items + today/tomorrow calendar
2. Sends to Anthropic API (claude-haiku-4-5) for triage
3. Pushes structured `briefing.json` to GitHub
4. Dashboard at begb0037admin.github.io/work-inbox/ fetches and renders it

## Manual Refresh

To trigger a manual refresh outside the schedule:

**Option A — Desktop shortcut:**
Double-click `Run Inbox Briefing.bat` on the desktop.

**Option B — PowerShell:**
```powershell
cd C:\Users\admin\Documents\Claude\Projects\work-inbox
git fetch origin
git checkout origin/main -- fetch_inbox.py
python fetch_inbox.py
```

## Starting a New Session to Work on This Project

Tell Claude:
```
Read CLAUDE.md and HANDOVER.md from the work-inbox repo, then tell me the current state.
```

Claude will read both files via the GitHub proxy and report back without asking Kevin for a recap.

## Dashboard Features
- Time-of-day greeting (Good morning/afternoon/evening) — UK timezone
- Amber context bar — 5-7 sentence specific briefing with names, dates, actions
- Urgent / Needs Response / FYI / Low Priority card sections
- Click any card to open the exact email in Outlook (openmail:// protocol)
- Tick a card to mark done — fades and hides after 1.5 seconds
- Show Done / Hide Done button to reveal/re-hide ticked items
- Archive panel — past briefings by date
- Last refreshed timestamp — bottom right, updates every script run
- Sidebar — calendar today/tomorrow, absences bullet list

## Kevin's Communication Style (for AI triage reference)
| Rule | Detail |
|------|--------|
| Opening | Always `Hi [Name],` |
| Tone | Professional, warm, contractions fine |
| Length | 5 sentences or fewer unless complexity warrants |
| Sign-off | `Best,` then `Kevin` |
| Never | Commit to a date/time without a calendar check |
| Never | "Dear", "Kind regards", "Please don't hesitate" |

## VIP Contacts
Marie Cooksey and Simon Burford are always Urgent or Needs Response — never lower priority.
