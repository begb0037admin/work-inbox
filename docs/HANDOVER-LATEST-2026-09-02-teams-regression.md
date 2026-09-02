# READ THIS FIRST -- Handover -- 2 September 2026, ~21:20, live regression found

Both connectors are cut over on the live scheduled task, but are silently delivering ZERO data. Full details below. If a newer entry exists at the top of `HANDOVER.md`, prefer that -- this file is a fast-checkpoint fallback pushed under low usage budget.

## What actually happened tonight, in order
1. Calendar (`#31`) merged to `main`, live-validated clean earlier tonight.
2. Teams (`#33`) merged to `main` at 19:23:59 (`f2da964`), after a clean combined shadow-run.
3. The live `Work Inbox Bridge Briefing` scheduled task on `101L-DE013193` / `AD-OAK\begb0037` was cut over to **both** connectors together, per Kevin's standing decision (confirmed via `schtasks /query /tn "Work Inbox Bridge Briefing" /xml`):
   ```
   -File "C:\Users\begb0037.AD-OAK\work-inbox\Run Laptop Bridge Briefing.ps1" -CalBackend connector -TeamsBackend connector
   ```
4. The task fired for real at **20:58:18**, `LastTaskResult 0` (reported success).

## THE PROBLEM -- found live over RDP, 2 Sept ~21:15-21:20
The 20:58 run's own summary (`data\lane_b\20260902T195821Z_lane_b.json`) shows **both domains actually failed**, silently, while the wrapper still exited 0:

```json
"per_domain": {
  "calendar": {"status":"codex_failed","count":0,"attempts":[{"n":1,"identity":"primary","outcome":"codex_failed","detail":"[calendar#primary1] codex exec produced no usable JSON output after 1 attempt(s) (CODEX_HOME=C:\\WorkInboxAI\\codex-laneb)"}],"served_by":null},
  "teams":    {"status":"codex_failed","count":0,"attempts":[{"n":1,"identity":"primary","outcome":"codex_failed","detail":"[teams#primary1] codex exec produced no usable JSON output after 1 attempt(s) (CODEX_HOME=C:\\WorkInboxAI\\codex-laneb)"}],"served_by":null}
}
```

Three distinct bugs identified from the raw evidence:

1. **Automatic failover never fired.** Only ONE attempt is logged per domain, `identity: "primary"`. The whole point of tonight's earlier failover build was that a `codex_failed` primary outcome should always trigger a `failover` (personal-account) attempt. It did not. Leading hypothesis: the live wrapper calls `lane_b_call1.py --domain both`, and the combined/both-domain code path may not be wired to the same failover orchestration (`fetch_domain()`) that was built and proven earlier tonight in the single-domain (`--domain teams`, `--domain calendar`) tests.

2. **Teams connector actually worked, but the run is classified as a hard failure anyway.** Raw transcript `data\lane_b\20260902T195643Z_call1_teams_primary_a1.jsonl` shows the model successfully called `microsoft_teams.list_chats` and got back 49 real chats. But the model's final `agent_message` only returned `{"chats":[...49 entries...],"messages":[]}` -- it listed the chats and stopped, never calling a message-fetch tool to pull actual message content from any of them. That's valid, well-formed JSON -- but the guard/parser layer treats it as "no usable JSON output" / `codex_failed` rather than e.g. "ok, 0 items" or a distinct "incomplete" status. Likely caused by the same tight `PRIMARY_TIMEOUT_S=290s` / `PRIMARY_MAX_ATTEMPTS=1` budget cut made earlier tonight (speed request) -- may simply not be enough turns/time for Teams to both list chats AND pull messages from each one.

3. **Calendar produced literally no raw `.jsonl` output at all** for this run (no matching file near 19:56-19:58Z) -- a different, more severe failure than Teams' "incomplete but present" output. Root cause not yet found; needs a look at whatever wrapper-level logs exist (`logs\`, `data\codex_runs\`) for that exact window.

## Net effect right now
The live scheduled task is technically "cut over" and reports success, but delivers **zero real calendar/Teams content** on every run. Mail (IMAP) is unaffected and still working normally. This is silent -- nothing alerts on it -- so it needs fixing before this can be trusted as actually live.

## Status: diagnosis in progress
Dispatched **Drew** (background agent, visible in Fleet) at ~21:22 to:
- Read current `lane_b_call1.py` / `fetch_inbox.py` on `main` and find why the `--domain both` path skips failover.
- Find/fix why an empty-but-valid Teams response is classified as `codex_failed`.
- Root-cause calendar's total silent failure for this run.
- Propose/implement a fix on a new branch + draft PR -- explicitly told NOT to touch the live scheduled task, NOT to merge, NOT to run anything against live accounts without Kevin's explicit go-ahead.
- Report back with root cause(s), fix status, and exactly what needs sign-off.

**Not yet known (check Drew's follow-up report, a fresh top entry in `HANDOVER.md`, or ask the coordinator):** whether the fix is in, whether it's been proven, whether Kevin has approved re-testing/re-cutover.

## What NOT to re-litigate
- No permanent COM fallback for calendar -- connector is mandatory, standing decision.
- Snapshot-diff mechanism deliberately dropped -- re-contamination guard is the sole live safety layer, don't reintroduce diff-based checks.
- Edu-primary / personal-automatic-failover is the correct architecture (this session's bug is that failover isn't *firing*, not that the architecture is wrong).
- Both calendar and Teams go live together, not staged -- already done, don't split them back out.
- Live scheduled task's mail (IMAP) pull path is fine and untouched by any of this.
