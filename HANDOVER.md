# work-inbox — Living Handover Document

**Last updated:** 2026-07-03 — Granola 0-matches bug. Debug lines pushed (commit `2026f36`). Awaiting Kevin's script output to identify title mismatch.
**Status:** Active — pipeline fully working. Live at https://wi.lelitte.co.uk/ | https://begb0037admin.github.io/work-inbox/.

---

## NEXT SESSION — START HERE

### 1. Granola 0-matches bug — ACTIVE (highest priority)

**Problem:** Phase 3.7 fetches 10 Granola notes successfully but matches 0 of them to the 11 calendar candidates. `_granola_context` stays empty. Phase 3.8 runs blind — AI summaries for calendar items have no prior-meeting context.

**Confirmed so far:**
- API call works: response shape is `{"notes": [...], "hasMore": bool, "cursor": "..."}`, 10 notes returned
- Calendar candidates exist: 11 non-all-day items built from `cal_today_items` + `cal_tomorrow_items`
- 0 keyword matches between note titles and calendar item titles

**Root cause: unknown.** Most likely one of:
- Granola note titles use a different format than Outlook calendar titles (e.g. Granola: `"FA Team Catch-up — 03/07"` vs Outlook: `"FA Team Daily Catchup"` — note `Daily` vs nothing, and hyphen vs em-dash)
- Granola notes from the last 10 days are all from meetings not on today's/tomorrow's calendar
- A subtle normalisation mismatch in `_granola_keywords()`

**What's been done:** Two diagnostic debug lines pushed to `fetch_inbox.py` main (commit `2026f36`). Now in the Granola block, immediately after `_cal_candidates` is built:
```python
print(f"Phase 3.7 debug - note titles: {[n.get('title','') for n in _g_notes]}")
print(f"Phase 3.7 debug - cal candidates: {[c['title'] for c in _cal_candidates]}")
```

**Next step:** Kevin runs the script. Script output will show both lists. Compare them — find the title mismatch. Then fix `_granola_keywords()` or the matching logic to bridge the gap. Once the fix is confirmed working (context > 0 meetings), remove all three debug print lines in one clean commit.

**The `_granola_keywords()` function (current):**
```python
def _granola_keywords(title):
    t = re.sub(r'\b\d{1,2}/\d{2}\b', '', title)   # remove DD/MM dates
    t = re.sub(r'\b\d{4}\b', '', t)                # remove years
    t = re.sub(r'[—\-&]', ' ', t)                  # dashes and ampersands to spaces
    t = re.sub(r'[^\w\s]', '', t)                  # strip remaining punctuation
    return set(w.lower() for w in t.split() if len(w) >= 2)
```

**Matching logic:** score = keyword intersection count; match requires score ≥ 1; best-scoring note wins per calendar item.

---

## Fix list (after Granola is resolved)

2. **Absences not showing tomorrow's absences** — Sidebar absences panel is blank even when a team member is on leave tomorrow. The `fetch_inbox.py` triage needs to detect upcoming absences from calendar events and include them in `briefing.json` `absences[]` with forward notice (at least 1 day prior).

3. **AI calendar summaries are too generic** — Phase 3.8 runs blind without Granola context (blocked by item 1 above). Once Granola matching is fixed, summaries should use prior-meeting notes as primary source and surface carry-forwards, open actions, live decisions.

4. **Drag reorder animation** — No visual feedback during drag. Cards need to visually shift in real time as Kevin drags — placeholder in the DOM during `dragover`.

---

## Session 2026-07-03 — Granola 0-matches investigation

**Scope:** Diagnosing why Phase 3.7 Granola fetch returns 10 notes but matches 0 calendar items.

**What happened:**
- First debug line added and pushed (earlier commit, SHA `6bc0941`): `print(f"Phase 3.7 debug - API keys: ..., notes: {len(_g_notes)}")`.
- Kevin ran the script. Output confirmed: `API keys: ['notes', 'hasMore', 'cursor'], notes: 10` and `Granola context for 0 meetings`. API works; shape correct; 0 matches.
- Root cause unresolved — need to see the actual titles on both sides.
- Two additional debug lines pushed (commit `2026f36`): print note titles list and cal candidates titles list.
- Session ended before Kevin could run the updated script.

**Current `fetch_inbox.py` SHA on main:** `5dd6f684ba69c959e32d84c8ed248e142b83dfb4`

---

## Session 2026-07-04 — Crest rule propagation

No code changes to work-inbox this session. Cross-repo maintenance only.

- **Crest audit completed** — all dashboards inspected for Oxford crest usage:
  - work-inbox: external file `images/oxford-crest.jpg` — intact ✅
  - hris-launcher: base64 JPEG `<img class="sidebar-crest">` — intact ✅
  - command-centre: base64 JPEG `<img class="sb-crest">` — intact ✅
  - hr-fa-knowledge-base: base64 JPEG `<img class="crest">` — intact ✅
  - hris-dashboard: emoji 🎓 (no image) — N/A
  - ag-flexpoints: no crest — N/A
- **Hard rule propagated** — added to CLAUDE.md for hris-launcher, command-centre, hr-fa-knowledge-base.

---

## Session 2026-07-03 (aborted — GitHub MCP disconnected)

GitHub MCP server disconnected before handover could be pushed. No code changes. Reminder trigger set for 2026-07-04 07:00 UTC. Two new items identified (crest rule propagation, grey square verification) — both addressed in 2026-07-04 session above.

---

## Session 2026-07-02 (end) — small fixes pushed to main

- **`ctx-strip` label restored** — `setupCtxTicker()` was missing `<div class="ctx-label">Briefing context</div>`. Added back. Commit `fb178b5`.
- **Badge position fixed** — NEW/UPDATED badges moved from inside `.card-ph-title` to `.card-ph-actions` (right side, next to CC→). Commit `2d39b9e`. Confirmed working.
- **OSM IT Services URL** — sidebar link updated to `https://oxford.saasiteu.com/Modules/SelfService/#home`. Commit `e4cc1fd`.

---

## Session 2026-07-02 (continued) — calendar panel corrections

Commits pushed to main: `af12dff` (equal 3-col, July+August, AI summaries), `1da688d` (combined mini-cals into one card, narrowed calendar column).

### What changed
- **`css/styles.css`**: `.main-cal-panel` grid changed to `7fr 7fr 4fr` — Today and Tomorrow take equal wider columns; mini-cal column is narrower (≈22% of row).
- **`js/app.js`**: `renderMiniCal(monthOffset)` now returns inner content only (no wrapping block). Both months rendered inside a single `.main-cal-block` with a `.mini-cal-divider` `<hr>` between them. AI summaries (`c.summary`) shown on Today/Tomorrow entries as `.main-cal-summary` divs.

---

## Session 2026-07-02 — v5 design corrections (commit `12ff90d`)

- **Removed** email address from sidebar
- **Links updated**: 6 approved links, all now populated
- **Cards redesigned**: flat `.card-ph` design (drag handle, circle done button, title + sub, email + CC→ icons, NEW/UPDATED badges on right)
- **Layout corrected**: left col = Today + Tomorrow, right col = Week + Parked
- **Oxford crest**: restored as external file `images/oxford-crest.jpg` — NEVER embed as base64, NEVER delete, NEVER change the `src` attribute

---

## Architecture

| Component | Description |
|-----------|-------------|
| `fetch_inbox.py` | Outlook COM via pywin32. Pulls inbox → Anthropic triage (claude-haiku-4-5) → pushes `data/briefing.json` to GitHub via Contents API |
| `index.html` | Shell — HTML structure only. Loads `css/styles.css` → `js/app.js`. No framework, no build step. |
| `css/styles.css` | All styles. |
| `js/app.js` | All JS — briefing render, cal panel, ctx ticker, CC ticker, drag-and-drop, tick sync, archive, live clock. |
| `open_email.py` | Registered `openmail://` protocol handler — opens exact email in classic Outlook via EntryID COM |

---

## Current State

### Working
- fetch_inbox.py — all phases confirmed working (Phase 3.7 Granola matching broken — see fix list)
- Task Scheduler — `WorkInbox-0900` / `WorkInbox-1200` / `WorkInbox-1500` (Mon–Fri)
- Dashboard loads live briefing.json on load, falls back to localStorage archive
- Oxford navy sidebar — crest (external `images/oxford-crest.jpg`), branding, live clock, filter, CC ticker, inbox widget, absences, all 6 links populated
- 3-column calendar panel (Today `7fr` | Tomorrow `7fr` | July+August mini-cals in one card `4fr`)
- Rotating context strip with "Briefing context" label, dot nav
- 2×2 priority grid with tier filter — flat `.card-ph` design, NEW/UPDATED badges on right
- CC ticker reads live from CC tasks.json every 60s
- drag-and-drop, tick sync, archive, show done, openmail:// all working
- Multi-machine setup complete (begb0037.AD-OAK)

### Known issues (fix next session)
- **Granola 0-matches** — see NEXT SESSION above (highest priority)
- Absences not showing tomorrow's leave
- AI calendar summaries too generic (blocked by Granola fix)
- Drag reorder has no visual animation

---

## localStorage Keys

| Key | Purpose |
|-----|---------|
| `workInbox_briefings_v1` | Archive of past briefing JSON objects, keyed by date string |
| `workInbox_today_v1` | Key of the currently displayed briefing |
| `workInbox_ticks_v1` | Tick (done) state for all cards |
| `workInbox_priOverrides_v1` | Per-card section overrides for priority drag-and-drop |
| `workInbox_priOrder_v1` | Per-section sort order for priority cards |
| `workInbox_customPri_v1` | Email cards manually dragged into priority sections |

---

## Technical Notes

**index.html edits:** always use binary `atob()`/`btoa()` — NEVER `TextEncoder` on file content (re-encodes em-dash bytes).

**Priority drag-and-drop sections:** `pt` (today), `ptom` (tomorrow), `pw` (week), `pfyi` (parked/FYI), `ur` (urgent overlay), `nr` (needs overlay).

---

## File Locations

| File | Location |
|------|---------|
| Repo | github.com/begb0037admin/work-inbox |
| Proxy | github-proxy.lelitte.co.uk/work-inbox/ |
| Dashboard (primary) | wi.lelitte.co.uk |
| Dashboard (GitHub Pages) | begb0037admin.github.io/work-inbox/ |
| Styles | `css/styles.css` |
| JS | `js/app.js` |
| Script | `fetch_inbox.py` |
| Opener | `open_email.py` |
| Briefing | `data/briefing.json` |
| Local | `C:\Users\admin\Documents\Claude\Projects\work-inbox\` |
| Scheduler recovery | `create_inbox_tasks.bat` in repo root — run as Administrator |

---

## Standing Rules
- Never commit tokens or raw data
- All GitHub writes via Contents API (PAT from `GITHUB_PAT` env var)
- `index.html` edits: always use binary `atob()`/`btoa()` — NEVER `TextEncoder`
- Desktop bat: always download fresh via PowerShell — never rename an existing file
- Every raw.githubusercontent.com fetch MUST include `?t=<timestamp>` cache-buster
- **NEVER touch `images/oxford-crest.jpg` or the `<img class="sidebar-crest">` src attribute** — external file only, never base64
