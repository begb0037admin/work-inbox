# Codex brief — work-inbox: amber "unavailable this run" banner (Kevin-approved, 24 Sep)

Branch `drew/wi-connector-amber-banner` off main. Work only in C:/Users/admin/github/work-inbox.
Only touch `fetch_inbox.py`, `js/app.js`, `HANDOVER.md`. Don't read other repos, `data/`,
`Archive/`, `js/vendor/`. Keep UTF-8 intact. Commit locally; no push; short final message.

Problem: on 24 Sep every connector call failed (usage limit) yet the dashboard showed the green
"● Up to date" banner. Rule: **never green when data failed.**

## 1. Pipeline — `fetch_inbox.py`
Write a new top-level field into briefing.json: `"connector_status": {"mail_inbox": s,
"mail_sent": s, "calendar": s, "teams": s}` where `s` is the per-domain status string from
`data/lane_b/lane_b_normalised.json` as the loaders already read it (`"ok"`, `"unavailable"`,
`"halt"`, …), or `"n/a"` when that backend isn't `connector` / not configured. Find where
`_load_lane_b_mail()` / calendar / Teams loaders read the per-domain status and capture it into
module-level variables (default `"n/a"`), then add the field next to `"mail_truncation_risk"`
(~line 4570). Must never raise: any error → leave that domain `"n/a"`. `python -m py_compile`.

## 2. Dashboard — the status banner (`js/app.js`, ~1100-1140)
- Failed domains = those whose status is not `ok`/`n/a` in `data.connector_status`; ALSO, when
  `connector_status` is absent (older briefings, e.g. the current one), fetch
  `data/laptop_status/briefing_status.json` (same raw.githubusercontent base + cache-buster as the
  other data files, fetched once alongside the briefing) and treat `lane_b_domains.calendar` /
  `.teams` with `status` not `ok` as failed — only if its `ts` is within 2 hours of the
  briefing's refresh time (so a stale status file can't misreport).
- When the run is up to date BUT any domain failed: amber banner (background `#b45309`, white
  text): "⚠ Mail/calendar unavailable this run — <names> didn't load, so this briefing may be
  missing items. Last ran <refreshed_at>." Names in plain English: "inbox mail", "sent mail",
  "calendar", "Teams" joined with commas/"and". Keep the truncation note line if set.
- Green only when up to date AND nothing failed. The red stale banner keeps priority over amber
  (if stale, red; append the failed-domains sentence to it too).
`node --check js/app.js`. Top-of-HANDOVER entry.
