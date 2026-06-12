# HANDOVER — AG FlexPoints
> Living handover for this project. Hope: read this first.

Session closed 2026-06-12 at Kevin's usage limit; this is the live state.

### What AG FlexPoints is
Sister project to work-inbox: recurring report on Access Group FlexPoints
(service credits) for manager reporting. Repo **begb0037admin/AG-FlexPoints**
(live dashboard: https://begb0037admin.github.io/AG-FlexPoints/). Development
happens in **work-inbox `ag-flexpoints/`** (cloud sessions are scoped to
work-inbox only — writes to AG-FlexPoints are denied at session level), and
Kevin syncs files across with this block (cache-buster mandatory):
```powershell
cd C:\Users\admin\AG-FlexPoints
$t = Get-Date -UFormat %s
Invoke-WebRequest "https://raw.githubusercontent.com/begb0037admin/work-inbox/main/ag-flexpoints/index.html?t=$t" -OutFile index.html
Invoke-WebRequest "https://raw.githubusercontent.com/begb0037admin/work-inbox/main/ag-flexpoints/fetch_flexpoints.py?t=$t" -OutFile fetch_flexpoints.py
Invoke-WebRequest "https://raw.githubusercontent.com/begb0037admin/work-inbox/main/ag-flexpoints/HANDOVER.md?t=$t" -OutFile HANDOVER.md
git add -A; git commit -m "sync from work-inbox"; git push origin main
```

### Shipped this session (PRs #4–#11, all merged to work-inbox main)
- Dashboard = Access portal format (Kevin's explicit template choice) with the
  Oxford navy sidebar kept as hybrid. Clickable summary tiles (Awaiting
  Approval opens the single item directly when its link is captured),
  portal-style tables, AI overview + action highlights.
- PDF export = portal replica ("Access Group" header + red ring, dark
  FlexPoints banner, 3 paginated sections). Print needs "Background graphics"
  ticked, "Headers and footers" unticked.
- CSV export with full detail columns + links.
- Per-item links: script harvests ref→URL from rendered portal pages.
- Item detail for manager reports: script visits each item's detail page
  (≤25/run) → Service/Type/Product/Agreed Scope/Description/Outcomes/
  Requester/Assigned-to. FlexPoints Guardians extracted (Simon Burford, Marie
  Cooksey, Kevin) → sidebar + PDF header.
- Smoke test: Node DOM-harness, ALL PASS (no real browser available in cloud —
  network blocks cdn.playwright.dev; Claude-in-Chrome not drivable from cloud).

### THE ONE OPEN ISSUE — portal login (blocks live data, links, details)
`fetch_flexpoints.py` Phase 1 fails: "could not find username field" on
https://accessgroup.my.site.com/Support/s/login/ (headless Chromium, selectors
written blind; portal 403s all cloud fetches). Already hardened: searches all
iframes, wide selector set, on failure dumps `source/_login_debug.png` +
`.html` (gitignored). **Next action: get Kevin to run
`.\Run_FlexPoints_Report.bat` and send the debug PNG, then fix the selectors
from evidence.** Alternative quick win: Kevin saves the FlexPoints page from
his logged-in browser (Ctrl+S → "Webpage, Complete") into
`C:\Users\admin\AG-FlexPoints\source\` and re-runs the bat — the fallback
parses saved HTML including hrefs, which unlocks real per-item links
immediately. Also consider lifting the working login from the
hr-fa-knowledge-base scrape (same username/password flow) — that repo needs
adding to the session scope.
- Env vars confirmed working on admin machine: ACCESS_PORTAL_USER/_PASSWORD.
- `playwright` CLI not on PATH → use `python -m playwright install chromium`.

### Remaining tasks
1. Fix portal login (above) → first live pull → links/details/guardians real.
2. Task Scheduler job (weekly Mon 08:00) — not yet created.
3. Seed data is from Kevin's 2026-06-12 screenshots; live pull supersedes it.
4. URGENT for Kevin (business, not code): all 2,935 pts expire 2026-06-29;
   quote 69001638 (2,800 pts) sits in Awaiting Approval.
5. Kevin's GITHUB_PAT was exposed in a chat screenshot — advised to rotate.
6. Squash merges desync `claude/*` branches — reset branch to main after each
   merge (done at session end; branch currently clean on main).
