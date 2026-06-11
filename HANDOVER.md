# work-inbox — Living Handover Document

**Last updated:** 2026-06-11
**Status:** Active — fully working. 2026-06-11 session: mass-tick incident investigated, state restored, root-cause fixes applied.

---

## Incident 2026-06-11 — all priority actions marked done

**Symptom:** Morning of 11 June, dashboard showed the 10 June briefing with every priority action / urgent / needs item ticked and struck through. Kevin did not tick them.

**Root cause (from git history of data/ticks.json):** Not a refresh, rollover, or key-mismatch bug — the ticks were real writes synced via the Cloudflare Worker between 00:17 and 01:19 BST on 11 June, in strict top-to-bottom index order in rapid batches (e.g. urgent_0–9 inside one 1.5s-debounce window, then needs_0–7, stopping at 8 of 40 needs items). Mechanism: ticking a card hid it **instantly**, collapsing the list so the next checkbox slid under the pointer — repeated clicks at one screen position (stuck/faulty mouse button, phantom touch, or similar runaway input on a machine with the dashboard open) tick an entire section. The header said "Wednesday 10 June" because no 11 June briefing had been generated yet; the stale date was a red herring.

**Recovery:** data/ticks.json restored to the last user-intended state (commit 336c5e8, everything unticked). Incident keys kept explicitly `false` so the merge-on-load overrides stale `true` values in any browser's localStorage. No data was deleted — full tick history remains in git.

**Fixes applied (index.html):**
1. **Stable tick keys** — ticks now keyed by content hash (email entry_id, or title for priority tasks) instead of list position, with legacy positional keys still honoured as a read fallback. Briefing regeneration/reordering can no longer re-attach done state to the wrong items.
2. **Burst guard** — 5+ ticks within 4 seconds prompts "keep going?" before continuing; blocks runaway input.
3. **Delayed hide (700ms)** — ticked cards no longer vanish instantly, so a repeated click toggles the *same* item rather than cascading down the list.
4. **Remote tick load merges instead of replaces** localStorage (remote wins per key) — machine-local history is no longer destroyed on load.
5. **Destructive "safety reset" removed** — the old init() path wiped the *entire* tick store (all days) if every inbox item was ticked. Now: nothing is deleted; the dashboard auto-shows done items with a toast.
6. **Stale-briefing banner** — if the displayed briefing is not from today, a yellow banner says so.

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
- Task Scheduler bat file should auto-pull latest fetch_inbox.py before running
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

## Roadmap

| Priority | Task |
|----------|------|
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
