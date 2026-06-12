# AG FlexPoints

Recurring report on Access Group FlexPoints — balance, expiry risk, booked
services and the services catalogue from the
[Access Customer Success Portal](https://accessgroup.my.site.com/Support/s/our-services).

Same architecture as work-inbox: a script on the admin Windows machine pulls
the data (the portal sits behind Kevin's Access login, so nothing cloud-side
can reach it), Anthropic API structures it, the result is pushed to GitHub and
rendered on a live GitHub Pages dashboard.

## What FlexPoints are (context for the report)
- Service credits from The Access Group, allocated based on contract value.
- Roll over for 12 months — unused points expire, which is the main thing this
  report watches for.
- Spent on training, configuration, consultancy and supplementary
  implementation services, booked by the organisation's **Points Guardian**
  via the Customer Success Portal catalogues.

## Pipeline
| Phase | What happens |
|---|---|
| 1 | Playwright logs in to `accessgroup.my.site.com/Support` and pulls the rendered FlexPoints + Our Services pages. Fallback: newest file saved into `source/` from a logged-in browser. |
| 2 | Anthropic API (claude-haiku-4-5) parses page text → structured JSON mirroring the full portal page: summary tiles (available/requested/booked/expiring 30-60-90/awaiting), Quote Received, Requested, Booked, full Transaction History, plus AI highlights and a context paragraph. |
| 3 | Snapshot appended to `data/flexpoints_history.json` (one per day, ~2 years kept) for trend tracking. |
| 4 | Report + history pushed to GitHub via Contents API → dashboard picks them up. |

The dashboard is a hybrid of the Access portal layout and the work-inbox
style (Oxford navy sidebar, context bar, action-accented sections) and can
**export the full report** as CSV (all tables + summary + highlights) or
PDF (print stylesheet).

## One-time setup (admin Windows machine)
1. Create the GitHub repo **`begb0037admin/AG-FlexPoints`** (public), enable
   GitHub Pages on `main`, and move this folder's contents to the repo root.
2. Windows **User** environment variables (never in any file):
   - `ACCESS_PORTAL_USER` — Access portal username
   - `ACCESS_PORTAL_PASSWORD` — Access portal password
   - `ANTHROPIC_API_KEY`, `GITHUB_PAT` — already set for work-inbox
3. `pip install anthropic playwright` then `playwright install chromium`
4. Test run: `Run_FlexPoints_Report.bat`
5. Task Scheduler: weekly, **Monday 08:00** (suggested) — recurring cadence is
   whatever the trigger is set to.

If the Playwright login selectors don't match Access's login form, lift the
working login from the `hr-fa-knowledge-base` scrape script (same
username/password flow) into `get_portal_text()` in `fetch_flexpoints.py` —
the rest of the pipeline is independent of how Phase 1 gets the page text.

## Hard rules (inherited from work-inbox)
- Never commit credentials, API keys, or raw portal pages (`source/` is
  gitignored).
- Always pull `fetch_flexpoints.py` from GitHub before running.
- Every raw.githubusercontent.com fetch includes a `?t=<timestamp>`
  cache-buster.
