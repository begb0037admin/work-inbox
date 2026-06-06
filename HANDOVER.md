# work-inbox â Living Handover Document

**Last updated:** 2026-06-05
**Status:** Phase 1 complete. Phase 2 not started.

## Architecture

| Component | Description |
|-----------|-------------|
| etch_inbox.py | Outlook COM via pywin32. Pulls inbox ? data/inbox_raw.json |
| index.html | Dashboard UI |
| Seat A (Claude API) | Phase 2: script calls API directly for triage |

## What was tried and abandoned

| Approach | Reason abandoned |
|----------|-----------------|
| Claude in Chrome | Too slow, too many tokens |
| Microsoft Graph API (Azure CLI client ID) | Oxford blocks â AADSTS65002 |
| Microsoft Graph API (Graph Explorer client ID) | Oxford blocks â AADSTS1001010 |
| Microsoft Graph API (Office client ID) | Oxford blocks â AADSTS65002 |
| Microsoft Graph PowerShell module | Oxford blocks â needs admin consent |

## Current approach: Outlook COM

- Library: pywin32 (win32com.client)
- No OAuth. No tokens. Outlook handles auth.
- Pulls: Inbox (138 messages, last 7 days)
- Sent: folder empty. Calendar: no events today/tomorrow.

## Phase 2 decision

Script calls Claude API directly:
pull inbox ? call Claude API ? triage ? write dashboard JSON ? open dashboard.
One command. Zero manual steps.
Requires: Anthropic API key stored as environment variable.

## Roadmap

| Phase | Description | Status |
|-------|-------------|--------|
| 1 | fetch_inbox.py via Outlook COM | Complete |
| 2 | Script calls Claude API directly | Pending |
| 3 | Windows Task Scheduler pre-fetch | Pending |

## Standing rules

- Never commit tokens or raw data
- Local: https://github.com/begb0037admin/work-inbox
- Repo: github.com/begb0037admin/work-inbox
