# CLAUDE.md — AG-FlexPoints
> AI bootstrap entry point. Read this first, then HANDOVER.md (current state, open issues). Keep under 200 lines.

## Identity
- **Project:** AG FlexPoints — recurring Access Group FlexPoints report
- **Purpose:** Tracks FlexPoints balance, expiry risk, booked services and the
  service catalogue from the Access Customer Success Portal. Local script →
  Anthropic triage → GitHub → live dashboard. Sister project to work-inbox.
- **Owner:** Kevin Lelitte, Manager/Director HR Systems, University of Oxford
- **Status:** Scaffolded — awaiting first live run
- **Repo:** https://github.com/begb0037admin/AG-FlexPoints
- **Source portal:** https://accessgroup.my.site.com/Support/s/our-services

## Architecture
| Component | Description |
|---|---|
| `fetch_flexpoints.py` | Runs on admin Windows machine. Phase 1: Playwright login to Access portal (env-var credentials) pulls FlexPoints + Our Services pages; fallback reads newest saved page from `source/`. Phase 2: claude-haiku-4-5 → structured report JSON. Phase 3: daily snapshot into history. Phase 4: push both files via GitHub Contents API. |
| `index.html` | Single-file GitHub Pages dashboard — hybrid of the Access FlexPoints portal layout and work-inbox style. Oxford navy sidebar, summary tiles, Quote Received (yellow needs-response accent), Requested, Booked, Transaction History. Export buttons: CSV + PDF/print. Fetches `data/flexpoints.json` with `?t=` cache-buster, falls back to relative path until the repo's raw URL is live. |
| `data/flexpoints.json` | Latest report (generated each run). Currently seeded with real data from Kevin's 2026-06-12 portal screenshot. |
| `data/flexpoints_history.json` | Balance snapshots (one/day, ~104 kept) for the trend chart. |
| `source/` | Gitignored drop-zone for manually saved portal pages (fallback input). |

## Key Constraints
- Portal requires Kevin's Access login — `ACCESS_PORTAL_USER` /
  `ACCESS_PORTAL_PASSWORD` Windows User env vars. Never in any file. Cloud
  agents cannot reach the portal (403) — local script is the only bridge.
- Model locked to claude-haiku-4-5 (consistent with work-inbox).
- Recurring schedule via Task Scheduler — weekly Mon 08:00 suggested.
- Points roll over 12 months; expiry risk is the report's main job.
- AI must never invent balances — nulls when the page doesn't show a figure.
- Always pull `fetch_flexpoints.py` from GitHub before running.
- Every raw fetch needs `?t=<timestamp>` cache-buster.

## What Was Tried and Abandoned
| Approach | Reason |
|---|---|
| Cloud-side fetch of portal/public Access pages | 403 — portal needs customer login; public site blocks automated fetchers |

## Hard Rules
- Never commit credentials, raw portal pages, or API keys.
- GitHub is the only working surface (locally-run scripts excepted; they pull
  latest from GitHub before every run).
