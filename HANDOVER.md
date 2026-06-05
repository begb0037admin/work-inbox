# work-inbox — Living Handover Document
**Last updated:** 2026-06-05
**Status:** Phase 1 build in progress

## Architecture
Kevin's daily inbox briefing system. Replaces Claude in Chrome workflow entirely.

| Component | Description |
|-----------|-------------|
| `fetch_inbox.py` | Python script. Microsoft Graph API via device code auth. Pulls inbox, sent items, calendar → `data/inbox_raw.json` |
| `index.html` | Dashboard UI. Loads JSON, renders triage view |
| Seat A (Claude chat) | Receives raw JSON, triages, produces dashboard JSON |
| Kevin | Pastes dashboard JSON into import box |

## Auth
- Client ID: `04b07795-8ddb-461a-bbee-02f9e1bf7b46` (Azure CLI pre-registered)
- Flow: Device code auth (first run = interactive login; subsequent = silent via cached refresh token)
- Token cache: `msal_token_cache.json` — **gitignored, never commit**
- Oxford Graph Explorer confirmed 200 OK — delegated access enabled

## Run instructions
```bash
python fetch_inbox.py
```
First run: follow device code prompt in terminal (visit URL, enter code).
Subsequent runs: silent — no interaction needed.
Output: `data/inbox_raw.json` (gitignored)

## Files
| File | Status | Notes |
|------|--------|-------|
| `fetch_inbox.py` | ✅ Created | Graph API pull script |
| `index.html` | ✅ Bug fixed | Resizer IIFE moved outside init() |
| `.gitignore` | ✅ Created | Covers token + data output |
| `HANDOVER.md` | ✅ Created | This file |
| `CHAT_PROMPT.md` | ⚠️ Partially obsolete | Review before next session |
| `CHROME_PROMPT.md` | ❌ Obsolete | Chrome workflow abandoned |
| `data/inbox_raw.json` | Gitignored | Script output |
| `msal_token_cache.json` | Gitignored | Auth token cache |

## Roadmap
| Phase | Description | Status |
|-------|-------------|--------|
| 1 | fetch_inbox.py + index.html fix + HANDOVER.md | ✅ Complete |
| 2 | Seat A triage prompt updated — consumes raw JSON, produces dashboard JSON | ⬜ Pending |
| 3 | Explore piping script output directly into Claude chat | ⬜ Pending |
| 4 | Windows Task Scheduler for pre-fetch before Kevin sits down | ⬜ Pending |

## Standing rules
- Never commit `msal_token_cache.json` or `data/inbox_raw.json`
- Seat C updates this file at end of every work session
- CHROME_PROMPT.md is obsolete — do not use
