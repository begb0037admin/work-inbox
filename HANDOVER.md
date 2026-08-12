# work-inbox — Living Handover Document















**Last updated:** 2026-08-12 - "Marie K: Non-working day" day-view leak fixed and verified live (SECOND occurrence of the keyword-list-gap failure class, after bare-"AL" on 10 Aug). Checkpoint per agent-commons SESSION_PROTOCOL.md. Closed.







**Status:** Active — pipeline fully working. Live at https://wi.lelitte.co.uk/ | https://begb0037admin.github.io/work-inbox/.















---















## Session 2026-08-12 (continued) — "Marie K: Non-working day" day-view leak fixed, verified live; SECOND occurrence of this failure class (Drew)

**Scope:** Kevin spotted "Marie K: Non-working day" showing in the Tomorrow/Friday day-view calendar columns. He wants leave/absence entries excluded from the day-view columns entirely (he already has annual leave on the sidebar Absences panel) — same standing decision as the 10 Aug bare-"AL" fix.

**Root cause:** `_DAY_VIEW_EXCLUDE_KEYWORDS` (and its sidebar counterparts `ABSENCE_KEYWORDS`/`ABSENCE_NOISE`) had `"annual leave", "a/l", "on leave", "out of office", "ooo", "holiday", "away", "sick leave"` plus the bare-`AL` regex, but no entry for "non-working day" — a real, recurring phrasing Marie King's leave bookings on the "People Department - HR Systems" calendar use (confirmed live via Outlook COM: `Marie K: Non-working day`, real recurring all-day entries going back to Nov 2024, including 13 and 14 Aug 2026 — exactly Kevin's reported Tomorrow/Friday columns). This is the **second** occurrence of this exact failure class — a real leave-phrasing variant the keyword list hadn't seen yet, not a new kind of bug. Worth recognizing fast if it happens a third time.

**Fix, `fetch_inbox.py`:** added `"non-working day"` and `"non working day"` (hyphen and space variant, defensively — only the hyphenated form was found live) to all three keyword lists: `_DAY_VIEW_EXCLUDE_KEYWORDS` (excludes from day-view), `ABSENCE_KEYWORDS` (triggers sidebar Absences detection), and `ABSENCE_NOISE` (used by `_clean_absence_name()` to strip the phrase out of the display name during name-cleaning/splitting). All three needed the update, not just the first — read the actual code before assuming symmetry, per the brief: `ABSENCE_KEYWORDS`/`ABSENCE_NOISE` are for the sidebar panel Kevin explicitly wants this entry to keep appearing on, so this is a case where both lists needed the SAME new term added (unlike a hypothetical case where a day-view-only or sidebar-only term would need asymmetric treatment) — confirmed by tracing `_clean_absence_name()`'s split-on-`":"` logic by hand against the real subject "Marie K: Non-working day", which only produces the correct "Marie K" fallback name when "non-working day" is also in `ABSENCE_NOISE`.

**Verified against real live data, twice independently (real Outlook COM pull, real GitHub push, no local-file assumptions):**
- Ran `python fetch_inbox.py` directly (uncommitted local fix) twice against live Outlook. Both real production runs pushed successfully (commits `276cca48` and `a1289fad`).
- Pulled each pushed commit's actual `data/briefing.json` content back from the GitHub API (not the local working copy — confirmed as a real gotcha this session, see below) and checked directly: `calToday`/`calTomorrow`/`calDay2`/`calDay3` all show zero "Marie" matches in both runs; sidebar `absences` correctly shows `"Marie King - off tomorrow, returns Friday 14 August"` — excluded from day-view, still present on the sidebar, exactly what Kevin wants.
- Cross-checked against the pre-fix archived briefing (`data/archive/briefing_20260812_090349.json`, from the 09:00 scheduled run, before this fix): confirms "Marie K: Non-working day" genuinely was present in `calTomorrow`/`calDay2` before the fix, and genuinely absent after — a real before/after comparison, not just "the new code looks right."
- Isolated unit-style check on the exact literal subject string pulled live from Outlook (`'Marie K: Non-working day'`, confirmed plain ASCII hyphen, no unicode lookalike): pre-fix keyword list → not excluded (the bug); post-fix keyword list → excluded. Matches the live production result.

**A real verification gotcha hit and resolved this session, worth flagging for next time:** `fetch_inbox.py` never writes `data/briefing.json` to local disk — Phase 4 only pushes via the GitHub Contents API (`PUT`). Checking the local working-copy file after running the script directly (rather than via the Desktop `.bat`, which does a fresh `git checkout` afterward) shows stale pre-run data and will look like the fix isn't taking effect even when it is. Always re-pull from `raw.githubusercontent.com` (with a cache-buster) or `gh api repos/.../contents/...` after a direct `python fetch_inbox.py` run, not the local file.

**Proposed, not built — flagging per "propose before non-trivial engineering":** since this is the second keyword-list gap in three days, checked whether Outlook's own calendar metadata could supplement subject-keyword matching. Real live comparison (`BusyStatus`, `Categories`, `Sensitivity`, `AllDayEvent` pulled directly via COM for both the People Dept - HR Systems calendar and Kevin's own calendar, same window): 13 of 14 real leave/absence entries on the People Dept calendar are `BusyStatus=0` (Free); every regular meeting checked on Kevin's own calendar is `BusyStatus=1` or `2` (Tentative/Busy) — a genuinely strong correlation. Not a clean signal on its own though: one real entry ("Julie annual leave", 13 Aug, non-all-day) is booked `BusyStatus=2` (Busy) despite being genuine leave, and `Categories`/`Sensitivity` showed no useful pattern at all (mostly empty/Normal across both leave and non-leave items). Proposal, if Kevin wants it: add `BusyStatus == 0` (Free) as an *additional* OR condition alongside the keyword lists (not a replacement — would still miss the one Busy-booked outlier found live), which would have caught "non-working day" the first time it appeared without needing a keyword-list update at all. Real engineering work (touches the Phase 1 calendar pull to capture `BusyStatus`, plus both detection paths, plus testing against a longer real history than this session's 3-day window) — not started, needs Kevin's go-ahead first.

**Not done, on purpose:** the BusyStatus proposal above (needs a decision); no attempt to hunt for further undiscovered leave-phrasing variants beyond what's confirmed live.

---

## Session 2026-08-12 — Run-start timestamp logging: verified live end to end, closed (Drew)

**Scope:** Per the new estate-wide `agent-commons/SESSION_PROTOCOL.md` (mandatory from 12 Aug 2026), checked the actual live state of the timestamp-logging work Kevin approved earlier this session — did not trust a prior in-chat "done" claim, verified against GitHub commit history and live Desktop/log files directly.

**Code — confirmed pushed:** commit `b74a794` (2026-08-12T07:38:40Z) — "Add run-start timestamp to every console/log-producing script." Touches `fetch_inbox.py` (+19/-11, adds a `log()` helper used on every Phase-boundary print plus the Outlook COM retry lines — the exact lines involved in the 11 Aug incident that prompted this) and all four `tools/*.py` scripts (`draft_final_diff_capture.py`, `publish_drafted_replies.py`, `publish_needs_reply.py`, `sent_corpus_pull.py`), each gaining a one-line `print(f"[...] <script> run started")` as the literal first statement under `if __name__ == "__main__":`.

**Desktop .bat launchers — confirmed edited, not tracked in git (Desktop-only files):** both `Run Inbox Briefing.bat` and `Run Draft Diff Capture.bat` on `D:\OneDrive - lelitte.com\Desktop\` got a 3-line timestamp block inserted immediately after `title`, before any other work:
```bat
for /f "delims=" %%I in ('powershell -NoProfile -Command "Get-Date -Format \"yyyy-MM-dd HH:mm:ss\""') do set "RUN_TS=%%I"
echo Run started: %RUN_TS%
```
Confirmed via `diff` against the pre-edit backups both scripts' own convention already produced (`Run Inbox Briefing.bat.backup-20260812-083908`, `Run Draft Diff Capture.bat.backup-20260812-083908`). This echo is console/title-only (prints before the `python -u ... | Tee-Object` redirection starts), so it's visible on an interactive double-click run but not captured in the `*_last_run.log` files — the log-file self-dating guarantee comes from the Python `log()`/print lines below, not this echo.

**Live verification — not assumed, checked directly:**
- `fetch_inbox.py`: real 09:00 scheduled "Work Inbox Briefing" run (post-dating the 08:38 BST code push) — `inbox_briefing_last_run.log` shows `[2026-08-12 09:00:08] fetch_inbox.py run started` plus timestamped Phase 1/2/3/3.2/3.5/3.7/4 boundary lines. Task completed `LastTaskResult 0`.
- `tools/publish_needs_reply.py`: same 09:00 run's downstream chain — `needs_reply_last_run.log` shows `[2026-08-12 09:04:00] publish_needs_reply.py run started`.
- `tools/publish_drafted_replies.py`: same chain — `drafted_replies_last_run.log` shows `[2026-08-12 09:04:07] publish_drafted_replies.py run started`.
- `tools/draft_final_diff_capture.py`: not naturally due until 09:30 and its last log predated the code push (06:30, stale). Manually triggered `Start-ScheduledTask -TaskName 'Draft Diff Capture'` (same hidden VBS-wrapper `/update` path Task Scheduler itself uses, so this is a genuine exercise of the real automated path, not a synthetic test) and polled to completion — `LastTaskResult 0`, `draft_diff_capture_last_run.log` now shows `[2026-08-12 09:12:24] draft_final_diff_capture.py run started`.
- `tools/sent_corpus_pull.py`: has no scheduled task (manual-only, requires explicit `--start`/`--end`). Local copy in `tools/` was stale (pre-dated the code push); backed it up (`sent_corpus_pull.py.backup-20260812-091146`) and re-pulled the current version from GitHub, then ran it directly with its own built-in `--stats-only` safe/read-only dry-run flag (`--start 2026-08-11 --end 2026-08-12 --stats-only`, writes nothing to disk, no GitHub push) — console output opened with `[2026-08-12 09:12:19] sent_corpus_pull.py run started` followed by the stats JSON.

All five files (`fetch_inbox.py` + 4 `tools/*.py`) and both Desktop `.bat` launchers now confirmed live and verified with real timestamped output, not just pushed code. Nothing remains open on this item.

**Next action:** none — closed. If a future incident needs this data, `*_last_run.log` files under `C:\Users\admin\Documents\Claude\Projects\work-inbox\` (and `\tools\`) are the first place to check; each now opens with a `[YYYY-MM-DD HH:MM:SS] <script> run started` line.

---

## Session 2026-08-11 (continued) — Hang bug FIXED and verified live; issue #3 closed (Drew)

**Scope:** Kevin approved applying the proposed fix from the previous session entry immediately, without waiting further on Codex's review of the issue #3 comment.

**Applied:** backed up `Run Inbox Briefing.bat` first (`Run Inbox Briefing.bat.backup-20260811-205756`, pre-edit SHA256 `21E42234...` recorded), then inserted the two-line guard into `:run_script`, exactly as proposed and exactly matching `Run Draft Diff Capture.bat`'s existing proven pattern:
```bat
if /I "%~1"=="/run" exit /b %RUN_EXIT%
if /I "%~1"=="/update" exit /b %RUN_EXIT%
```
placed immediately after the `if "%RUN_EXIT%"=="0" (...) else (...)` block and before the final `choice /c MQ` prompt. No other lines touched.

**Regression check on the manual/interactive path — verified directly, not assumed:** built an isolated smoke-test harness (a copy of the actual post-edit file with only the real Outlook COM/python call stubbed to an instant fake success, avoiding a real 4-minute run or network side effects for this specific check).
- No-argument invocation, chose `[R]`, ran to completion: still hit `Press M for the menu, or Q to quit:` exactly as before — confirms `%~1` stays empty across `goto run_script`, so the guard correctly does nothing for a manual double-click run. `Q` exited cleanly, code 0.
- `/update` invocation, no stdin available: went straight to `update_script` as before, ran to completion, and this time **skipped the prompt entirely** — exited immediately, code 0, no blocking. Mirror-image proof the guard fires on the real Task Scheduler invocation path.

**Real scheduled-task run — clean exit, no hang:** triggered `Start-ScheduledTask -TaskName "Work Inbox Briefing"` at 20:59:20 BST, polled every 15s. **Exited at 21:05:06 BST — ~5m46s total, `LastTaskResult: 0`.** No forced kill, first clean exit on this task after 6 consecutive hung runs today (06:00/12:00/15:00/18:00 scheduled + 17:17/19:35 manual, all `LastTaskResult 267014`). Real work confirmed genuine via the run log and live GitHub commits matching in timing: `f94a9bf`/`3f6c98f`/`96ee79d`/`5c29d1e`/`b17a2a0` (backup, briefing update, suggestions, needs_reply publish, drafted_replies mirror — the full chain).

**Toast notification — confirmed fired:** BurntToast's per-AppId registry counter (`PeriodicNotificationCount`) read 10 immediately before the trigger, 11 immediately after — a clean +1 tied to this run. This is the first Work Inbox Briefing run where that's been true; every earlier hang left the counter unchanged because the VBS wrapper's notification call (which sits after `objShell.Run` returns from the batch) was never reached before the forced kill tore down the job.

**Final confirmation posted to issue #3:** https://github.com/begb0037admin/agent-commons/issues/3#issuecomment-5258344987

**Closes:** the Work Inbox Briefing process-exit hang, first surfaced when the task was switched to run fully hidden, root-caused and handed to Codex earlier the same day (no action taken), then fixed directly per Kevin's approval after re-confirming Codex hadn't acted.

---

## Session 2026-08-11 (continued) — Hang bug NOT fixed by Codex; independently re-confirmed live, proposed fix drafted for review (Drew)

**Scope:** Kevin reported Codex had "finished fixing" the Work Inbox Briefing hang bug (task completes real work but the instance never exits, force-killed at `ExecutionTimeLimit=PT15M`, `LastTaskResult 267014`). Asked for independent verification before trusting that claim, same discipline as everything else that day.

**Verification, not assumption:**
- Read `D:\OneDrive - lelitte.com\Desktop\Run Inbox Briefing.bat` live — the `/run`/`/update` early-exit guard that `Run Draft Diff Capture.bat` already has (added 10 Aug specifically to prevent landing on the interactive `choice /c MQ` prompt) is still absent from `Run Inbox Briefing.bat`.
- Read the full `agent-commons` issue #3 thread (1486 lines via `gh issue view 3 --comments`) end to end. The last comment on the issue is Drew's own "HANDOFF TO CODEX" brief from earlier the same day — explicitly marked "Do not fix — investigation/fix is being handed to Codex. This comment is a handoff brief only." **There is no reply from Codex anywhere in the thread.** No fix was ever applied.
- Triggered a real manual run of the "Work Inbox Briefing" scheduled task (`Start-ScheduledTask`, 19:35:27) and monitored `Get-ScheduledTask`/`Get-ScheduledTaskInfo` every 20s in the background. Real work visibly completed early in the run log, but the task instance stayed `State=Running` for the full 15-minute window and was force-killed at 19:50:34 — `LastTaskResult: 267014`, an exact re-reproduction with zero code changes in between. Confirms the bug is still live as of this session, not fixed.

**Per Kevin's decision:** since Codex didn't act on the handoff, Drew drafted the exact proposed fix directly (mirroring `Run Draft Diff Capture.bat`'s proven guard verbatim) and posted it as a review-request comment on issue #3, asking Codex specifically to check the logic and confirm the interactive/manual-run path (no argument, double-click) still shows the menu correctly. **Comment:** https://github.com/begb0037admin/agent-commons/issues/3#issuecomment-5257829522

**Proposed change (not applied — live file untouched):** insert two lines into `:run_script`, immediately after the `if "%RUN_EXIT%"=="0" (...) else (...)` block and before the final `choice /c MQ /n /m "Press M for the menu, or Q to quit: "` line:
```bat
if /I "%~1"=="/run" exit /b %RUN_EXIT%
if /I "%~1"=="/update" exit /b %RUN_EXIT%
```
Task Scheduler invokes the batch as `/update` via the VBS wrapper (`Run Inbox Briefing Hidden.vbs`) with no console attached — currently execution falls through unconditionally to the interactive prompt, which blocks forever until the forced kill, before the VBS's own exit-code-passthrough and BurntToast notification call can ever run (`objShell.Run` for the batch never returns). The guard exits immediately with the real pipeline exit code once real work is done, matching the pattern already proven working in `Run Draft Diff Capture.bat` since 10 Aug.

**Not done, on purpose:** fix not applied to `Run Inbox Briefing.bat`; not pushed anywhere. Waiting on Codex's review comment on issue #3, then Kevin's explicit approval, before anyone implements.

---

## Session 2026-08-11 — Outlook COM connection retry (commit `3bd0649`, Drew)

**Scope:** Kevin reported `fetch_inbox.py` had failed twice in one day with a hard exit-1 at the first Outlook COM call, each time confirmed transient by manual retry succeeding minutes later. Priority fix, not just a diagnosis.

**Root cause:** `mapi.GetDefaultFolder(6)` (line 250 before this fix) intermittently raises `pywintypes.com_error (-2147418111, 'Call was rejected by callee.', None, None)` when Outlook's COM automation layer is momentarily busy (mid-sync, a dialog open, etc.). Confirmed from the real `inbox_briefing_last_run.log` in `C:\Users\admin\Documents\Claude\Projects\work-inbox\` — traceback pointed at exactly this line, both times.

**What changed (`fetch_inbox.py`):**
- New `connect_to_outlook(max_attempts=3, retry_wait_seconds=45)` wraps `Dispatch("Outlook.Application")` + `GetNamespace("MAPI")` + the first `GetDefaultFolder(6)` call (the exact call site of both real failures). On `pywintypes.com_error`, logs the attempt, waits 45s, retries — up to 3 total attempts — then re-raises (hard exit 1) only once exhausted.
- The first inbox loop (Phase 1's main pull) now reuses the folder handle `connect_to_outlook()` already opened, instead of a second unretried `GetDefaultFolder(6)` call.
- Deliberately scoped to this initial connection step only — no retry logic added anywhere else in the script, so a genuine error deeper in Phase 1+ still fails immediately instead of being masked.

**Verification:**
- Full live run against real Outlook (`python fetch_inbox.py` in the up-to-date clone) completed end-to-end in the normal ~3.5 min, exit code 0, all phases through Phase 5 completed and pushed to GitHub (commits `857f7b9`/`fbe9e86`). Phase 1 connected on the first attempt with no retry log lines — confirms zero added latency on the normal path.
- Confirmed the pushed script downloads cleanly via the real production path (`raw.githubusercontent.com` with cache-buster) and still contains the `^def build_fallback_context` marker the Desktop batch script's download-validation step checks for.
- Could **not** force-trigger the real busy-callee condition live to prove the retry path fires against genuine Outlook — noted as an honest limitation. Instead verified the exact shipped `connect_to_outlook()` control flow via a mocked-`pywintypes.com_error` test harness (4 scenarios: fails twice then succeeds, fails once then succeeds, fails all 3 and re-raises, and a clean zero-failure run) — all four behaved correctly, including confirming the exhausted-retries path still re-raises rather than swallowing the error.

**Not touched:** hris-dashboard, SAASIT, SSO/MFA (explicitly out of scope) and no other COM call sites in the script.

---

## Session 2026-08-11 (continued) — Draft Diff Capture rescheduled off Work Inbox Briefing's collision times (Drew)

**Scope:** Kevin asked whether Work Inbox Briefing and Draft Diff Capture (`tools/draft_final_diff_capture.py`, hourly 7am-7pm Mon-Fri) could safely run concurrently, since their schedules collided at 9am/12pm/3pm/6pm — both open Outlook COM connections at the same trigger moment.

**Investigation (real data, not assumption):**
- Pulled the actual Windows Task Scheduler Operational event log (`Microsoft-Windows-TaskScheduler/Operational`), not just the two tasks' trigger definitions. Confirmed both tasks' action processes launch within ~15ms of each other at every collision trigger (09:00, 12:00, 15:00, 18:00).
- Today's (11 Aug) real outcomes, using presence/absence of the Phase 4 `data/archive/briefing_*.json` file as the success proxy (the per-run log gets overwritten): 06:00 (no collision) succeeded; 09:00 (collision) succeeded; **12:00 (collision) failed — no archive written**; **15:00 (collision) failed — confirmed via log content, same `com_error` at line 250**. Both of today's two known failures landed on exact collision moments; the one non-collision trigger didn't fail.
- Technical basis: `Outlook.Application` is served by one running `OUTLOOK.EXE` as a single-threaded apartment — every calling process shares that one STA message pump, with no per-caller isolation. `RPC_E_CALL_REJECTED` ("Call was rejected by callee") is COM's standard STA reentrancy-protection response, not a fluke. `tools/draft_final_diff_capture.py` has the same unguarded `Dispatch`/`GetNamespace`/`GetDefaultFolder` pattern fetch_inbox.py had before this session's retry fix — it just hadn't been unlucky yet (0 failures in ~30 runs).
- Also surfaced, flagged separately as a distinct issue (not fixed this session): both tasks have `ExecutionTimeLimit=PT15M`; Work Inbox Briefing hit that forced kill on 3 of 4 checked triggers today (06:00, 12:00, 15:00), including the 06:00 run which had already completed all its real work successfully (archived 06:03:14) yet Task Scheduler didn't register it as finished until the 15-minute timeout — a process-exit hang somewhere in the `cmd.exe → powershell → python(COM)` chain, independent of this collision.

**Kevin's approved fix, implemented:** Changed **Draft Diff Capture only** from hourly (`StartBoundary=07:00`, `Interval=PT1H`, `Duration=PT12H`) to 5 fixed weekly triggers at **06:30, 09:30, 12:30, 15:30, 18:30 Mon-Fri** — each 30 minutes after Work Inbox Briefing's own times, via `Set-ScheduledTask -TaskName "Draft Diff Capture" -Trigger $triggers`. Work Inbox Briefing's own schedule (6/9/12/15/18) was explicitly not touched, and the two scripts were not merged. Kevin accepted the tradeoff (5x/day diff-pair capture instead of 13x/day) since ConversationID correlation means nothing is lost, only delayed to the next run.

**Verified live, not assumed:** re-ran `Get-ScheduledTask`/`schtasks /query /xml` after the change — confirms exactly 5 triggers (06:30/09:30/12:30/15:30/18:30, `DaysOfWeek=62`=Mon-Fri, no leftover hourly repetition), Action/Principal/Settings (`ExecutionTimeLimit=PT15M`, `MultipleInstancesPolicy=IgnoreNew`) all unchanged, `NextRunTime` correctly showing the next of the new fixed times, and Work Inbox Briefing's own 5 triggers confirmed byte-for-byte unchanged.

**Not done this session (flagged for later, not requested yet):** applying the same connect-with-retry pattern to `draft_final_diff_capture.py`; investigating the `ExecutionTimeLimit`/process-exit-hang finding.

---

## Session 2026-08-11 (continued) — Draft Diff Capture's missed-trigger catch-up disabled (Drew)

**Scope:** Kevin asked a follow-up architecture question after the schedule stagger above: if the machine is off at a trigger time and turns on later, could Windows Task Scheduler's `StartWhenAvailable` catch-up mechanism fire both tasks' missed triggers at once on wake, recreating the exact Outlook COM collision just fixed — just triggered by machine-on time instead of the clock?

**Investigation (real data):**
- Confirmed both tasks had `StartWhenAvailable=true` via live XML export.
- Found direct historical proof in the Windows Task Scheduler Operational event log: on 10 Aug, the machine booted at 06:41:08 (missing Work Inbox Briefing's 06:00 trigger). Task Scheduler's next catch-up check didn't run until 07:50:22 — and at that exact second, **17 separate tasks caught up together, "Work Inbox Briefing" among them.** This proves Windows batches all eligible missed-trigger catch-ups into one simultaneous launch, with no spacing or randomization — confirming the risk was real, not just plausible.
- No `RandomDelay` configured on either task's triggers, so nothing today would have broken up a simultaneous catch-up if both Work Inbox Briefing's and Draft Diff Capture's triggers were missed on the same day.

**Kevin's approved fix, implemented:** Disabled `StartWhenAvailable` on **Draft Diff Capture only**; left Work Inbox Briefing's catch-up enabled. Rationale: a skipped Draft Diff Capture catch-up costs nothing real (same zero-data-loss logic already accepted for the schedule stagger — ConversationID correlation picks up any pending pair on the next real run), while Work Inbox Briefing's catch-up still has real value (recovering a fully-missed morning briefing rather than waiting up to 3 hours for the next slot). Disabling only one side is sufficient — the collision requires both tasks to catch up together, so removing either side's ability to catch up removes the risk entirely.

```powershell
$task = Get-ScheduledTask -TaskName "Draft Diff Capture"
$settings = $task.Settings
$settings.StartWhenAvailable = $false
Set-ScheduledTask -TaskName "Draft Diff Capture" -Settings $settings
```

**Verified live afterward, two independent methods:**
- `Get-ScheduledTask` CIM object: Draft Diff Capture `StartWhenAvailable = False`; Work Inbox Briefing `StartWhenAvailable = True`.
- Raw XML export (`schtasks /query /xml`): Draft Diff Capture's `<Settings>` block now omits `<StartWhenAvailable>` entirely (Task Scheduler's schema only serializes this element when `true` — its absence is the correct signature of `false`, not a query failure, confirmed by cross-checking against the CIM read). Work Inbox Briefing's XML still explicitly shows `<StartWhenAvailable>true</StartWhenAvailable>`.
- All other Draft Diff Capture settings/triggers/action/principal confirmed unchanged: `ExecutionTimeLimit=PT15M`, `MultipleInstancesPolicy=IgnoreNew`, `RestartOnFailure` (Count=2/Interval=PT5M), the 5 triggers (06:30/09:30/12:30/15:30/18:30 Mon-Fri), action (`wscript.exe` + hidden VBS wrapper), principal (`RunLevel=Limited`, `UserId=admin`).

---

## Session 2026-08-11 (continued) — Both tasks run fully hidden, with success/failure desktop notifications (Drew)

**Scope:** Kevin's screenshot showed "Work Inbox Briefing" popping up a visible interactive terminal ("Press M for the menu, or Q to quit") when Task Scheduler fires it. Wanted the window gone entirely, matching the already-hidden "Draft Diff Capture" pattern, but with a lightweight notification (not silence) so he still knows a run happened, and definitely knows if one failed.

### Hidden window
Confirmed the real working mechanism by reading `Run Draft Diff Capture Hidden.vbs` directly rather than assuming: Task Scheduler's own "Hidden" task property does NOT suppress the console window (it only hides the task definition from the Task Scheduler UI); `WScript.Shell.Run(cmd, 0, True)` is the actual mechanism that gives a genuinely invisible window. Created `Run Inbox Briefing Hidden.vbs` (Desktop) on the same pattern and repointed Work Inbox Briefing's Task Scheduler action at it (`wscript.exe "Run Inbox Briefing Hidden.vbs"`), leaving triggers/settings/principal untouched. Also found and fixed a real bug in the existing Draft Diff Capture VBS while there: it never captured `objShell.Run`'s return value, so `LastTaskResult` always read 0 regardless of real success/failure — both wrappers now capture the exit code and propagate it via `WScript.Quit`.

### Desktop notifications — long verification story, told straight
First built a plain WinForms popup (`Show-TaskNotification.ps1`), reusing Echo's own `EchoShowIndicator.ps1` pattern (non-activating `WS_EX_NOACTIVATE` window), reasoning raw WinRT toast calls fail silently without a registered AppUserModelID. Two real, hard-won findings from direct testing, both now documented in the script/wrapper comments:
1. **Fire-and-forget children get killed by Task Scheduler's own Job Object** the instant the wrapped action process exits — confirmed by a process staying alive (correct session, no errors) but never rendering. Fixed by making the notification launch synchronous.
2. **The non-activating popup style never reliably surfaced** from a background/Task-Scheduler-launched process, confirmed via real screenshot captures across many isolated variants (border style, taskbar visibility, manual vs. CenterScreen position, `.Activate()` vs. Win32 `SetForegroundWindow`) — none of the small (440x88-150px) variants rendered visibly, while an identical large/maximized window did. Root cause not fully pinned down.

Given the size/rendering rabbit hole and that Kevin's actual intent (surfaced mid-session) was a genuine Windows toast — non-blocking, appears bottom-right, settles into Action Center, explicitly not a WinForms panel or blocking dialog — **switched to the BurntToast PowerShell module** (`Install-Module BurntToast -Scope CurrentUser`, v1.1.0). BurntToast registers its own AppUserModelID automatically, which is exactly what raw toast calls are missing.

**Verified, with an honest limitation stated plainly:** BurntToast's own per-AppId notification counter in the registry (`HKCU:\...\Notifications\Settings\{AppId}\...\powershell.exe`, `PeriodicNotificationCount`) was confirmed to increment by exactly 1 for each real trigger checked — a real Draft Diff Capture run (3→4) and a direct failure-path test (4→5) — with zero entries in the notification script's own fallback error log either time. This proves the OS is genuinely generating and queuing each notification, tied one-to-one to real events, not just that the API call "didn't error." **What was not conclusively caught: a live on-screen screenshot of the toast actually rendering** — every attempt (multiple timings, multiple durations) missed it, the same way the small WinForms popup was never caught either, despite the process-level evidence being solid both times. This may be a screenshot-timing/environment artifact specific to this verification method rather than a real display failure — Kevin seeing it appear during normal day-to-day use is the real confirmation still needed, and worth a quick "did you see it?" check after the next few real runs.

**Failure-path content confirmed correct**, independent of the visibility question: tested `Show-TaskNotification.ps1 -Status Failure` against a synthetic log containing the real `pywintypes.com_error` text from earlier today's incident — the extracted detail text was the exact error line, not a generic "something went wrong," matching Kevin's explicit requirement.

### A pre-existing issue this surfaced, not caused — and it broke the notification for this task
Triggering the real Work Inbox Briefing task end-to-end (17:17:03) showed the actual Python pipeline (Phase 1-5, archived 17:20:45) plus both chained downstream publishers (needs_reply 17:20:58, drafted_replies 17:21:04) all completing normally within the usual ~4 minutes — but the Task Scheduler task instance itself stayed "Running" for another ~11 minutes, until Task Scheduler force-killed it at the 15-minute `ExecutionTimeLimit` (`LastTaskResult 267014`, the exact same forced-termination code seen earlier today at the 06:00/12:00/15:00 triggers). This is the **same pre-existing process-exit hang already flagged earlier this session**, confirmed again here, not introduced today — Draft Diff Capture uses the identical VBS-wrapping pattern and consistently completes cleanly in ~15-20 seconds with no hang, so this is specific to Work Inbox Briefing's own longer chained-script execution, not the hidden-window/notification changes.

**Concrete, confirmed consequence for this specific run: the notification never fired.** BurntToast's registry notification counter was 5 before this run and still 5 afterward — no increment, no fallback error logged either, and no lingering `wscript.exe`/notification `powershell.exe` process (everything in the job was killed at the 15-minute mark). This is because the notification call sits inside the VBS *after* `objShell.Run(...)` returns from the wrapped batch — and since that call never returned (the hang), execution never reached the notification step at all before Task Scheduler tore down the whole job. **So today's fix, as it stands, is incomplete for Work Inbox Briefing specifically: the window is confirmed hidden, but "you'll be told if it ran/failed" is not reliable while this hang persists** — on any run that hits it (which was most of today's checked triggers), there is no notification and no way to distinguish "still legitimately working" from "hung" without checking Task Scheduler directly. The hang itself needs its own investigation before the notification promise is genuinely met for this task; not done as part of this session, flagged here for prioritization.

### One more thing to disclose plainly
Mid-session, while cleaning up a diagnostic test window, an overly broad `taskkill /F /IM powershell.exe /FI "STATUS eq RUNNING"` command force-killed 5 unrelated, long-running `powershell.exe` processes (PIDs 28272, 7664, 13880, 9636, 24516, all running since 10 Aug) that were not the diagnostic target — almost certainly the underlying shells behind other active terminal sessions/agents visible on this shared machine at the time. This was a genuine mistake (an overly broad filter, not a scoped or deliberate cleanup) and is disclosed here in full rather than omitted.

---

## NEXT SESSION — START HERE















### 1. Granola calendar context — CLOSED 2026-07-04 ✅















**DO NOT reopen.** Do not refactor, retune, or alter Phase 3.7b or Phase 3.8.















**Root cause (fixed):** `fetch_inbox.py` only read `detail["summary"]`. Granola note detail responses return usable content in `summary_text` and `summary_markdown`.















**Production fix (commits `7bc621f`, `cf6ca85`, `48e57ea`):**







- `fetch_inbox.py` now falls back to `summary_text` / `summary_markdown`.







- Granola context passed into Phase 3.8 increased to 1500 characters.







- Phase 3.8 asks for 2-3 concise prep sentences with a 900 token response budget.







- Title matching behaviour deliberately unchanged.







- No debug logging, forced matches, phase-skip flags, or dry-run mode in production.















**Future proposals (separate phases only):**







- A first-class DRY_RUN mode for safer diagnostics may be proposed later.







- Any title matching changes require a separate approved phase.















---















## Session 2026-08-10 (continued) — Calendar tab: 4-day rolling window + 4-month mini-cal, leave excluded from day-view, offset bug fixed (Drew)















**Scope:** Kevin's explicit request, same session as the needs_reply staleness-cutoff and 3-tab dashboard work above: "I have the annual leave on the sidebar so I don't actually need the annual leave to display in my calendar... let's just go with four days: today, tomorrow, day after that, and day after that... add August, September, October, November to the calendars on the right-hand side." This explicitly reopens Phase 3.8 (previously marked closed 2026-07-04 — see above — do not treat this note as a general invitation to touch it again beyond what's described here).















**`fetch_inbox.py` changes:**







- `day2`/`day3` computed via `next_workday(tomorrow)` / `next_workday(day2)` — same weekend-skipping semantics `tomorrow` already used, so a Thursday's day2/day3 are Monday/Tuesday, not a blank Saturday/Sunday.







- Leave/absence items excluded from all 4 day-view columns via `_DAY_VIEW_EXCLUDE_KEYWORDS` / `_is_leave_item()` (duplicates the existing `ABSENCE_KEYWORDS` term list used for the sidebar Absences panel rather than restructuring the file to share one constant — keep both in sync if either changes).







- New `cal_day2_items` / `cal_day3_items`, output as `calDay2` / `calDay3` in the briefing JSON alongside the existing `calToday` / `calTomorrow`.







- `calendar_summary_count()` / `weak_calendar_summary_count()` (the same-day-update safety gate in `validate_briefing_update()`) extended to check all 4 keys, not just the original 2.







- **Also fixed while in this code, since extending to 4 columns would have doubled its surface area:** the previously-documented calendar-summary index-offset bug (root-caused 2026-08-04, `begb0037admin/drew` `memory/calendar-summary-offset-bug.md`) — `enumerate()` was applied before filtering out all-day items, so a non-all-day item's index could start above 0 whenever an all-day item preceded it, and claude-haiku-4-5 was sometimes observed echoing output-position instead of the literal idx in that case, silently misattributing a summary to the wrong meeting. Fixed via a new shared `_non_all_day_candidates()` helper that produces both a model-facing sequential `idx` (always starts at 0) and a write-back-only `real_idx` (the item's true position in the day's list); Phase 3.7b (Granola) and Phase 3.8 both now consume the same `_all_day_candidates` list instead of each building their own.







- Preservation logic (`preserve_existing_calendar_summaries`) extended to cover `calDay2`/`calDay3`.















**Frontend (`js/app.js`, `css/styles.css` — `index.html` untouched, it's just a container div):**







- `renderCalPanel()` rewritten: `renderBlock()` now takes an explicit `bodyId` param and is called 4 times (today/tomorrow/day2/day3, DOM ids `calBodyToday`/`calBodyTom`/`calBodyDay2`/`calBodyDay3`). Day2/day3 headers show just the weekday name + date (e.g. "Wednesday 12 August"), not a "Today —"/"Tomorrow —"-style prefix, matching how Kevin described them.







- `renderMiniCal()` now takes a `mtgDates` array (real `Date` objects for whichever of the 4 day-view columns have at least one item) so "has-meeting" dots work across all 4 rendered months, not just the first two hardcoded ones. Called with offsets 0-3 → 4 months, rolling with whatever month "today" is in (currently August–November 2026).







- `.main-cal-panel` restructured from a 3-column `7fr 7fr 4fr` grid (which couldn't fit 4 day-columns + 4 months) into two full-width rows — `.main-cal-days-row` (4 equal columns) and `.main-cal-months-row` (4 equal columns), each still `display:grid;grid-template-columns:repeat(4,1fr)`.







- Confirmed `renderMainCal()` (a separate, older function, ~line 285) is genuinely dead/unused code before touching anything — not edited.















**Verification:**







- `python -m py_compile` on the backend, `node --check` on the frontend.







- Real production run: `D:\OneDrive - lelitte.com\Desktop\Run Inbox Briefing.bat /update` — exit code 0, "Phase 3.8 done - 12 calendar summaries generated", "Phase 3.8 preservation - reused 8 existing same-day calendar summaries", needs_reply and drafted_replies publishers both succeeded with `byte_identical_verified: true`.







- Pulled the live `data/briefing.json` after that run and confirmed `calDay2`/`calDay3` present and populated (5 and 6 items that day), no leave-keyword titles leaked into any of the 4 day-view columns except one gap (see below), and every freshly-generated (non-preserved) Phase 3.8 summary in the brand-new `calDay2`/`calDay3` columns correctly named its own meeting — no cross-contamination, confirming the offset-bug fix works on fresh data.







- Node DOM-stub test (same pattern as the Drafted Replies / tabs work, harness at `begb0037admin/drew` scratchpad, not committed) against the real edited `renderCalPanel()`: confirmed 4 day-columns, 4 correctly-named months, correct "has-meeting" dots, and correct Friday→"Next Week"-labeled-Monday weekend-boundary chaining for day2/day3.







- Live-browser screenshot of `https://begb0037admin.github.io/work-inbox/` Calendar tab after pushing matched the test output exactly.















**Known gap found during verification — FIXED same session, Kevin's explicit follow-up ("yes fix it - i dont want it to show"):** the leave-exclusion keyword list (and the pre-existing sidebar `ABSENCE_KEYWORDS` list it mirrors) matched `"a/l"` (with slash) but not the bare `"AL"` abbreviation. A real live entry, "Michael - AL", leaked through both the day-view exclusion and the sidebar Absences panel. Live Outlook check (bounded to the same date window `fetch_inbox.py` itself uses, not an unbounded scan — an earlier unbounded attempt over-ran and had to be killed) found this wasn't a one-off: two separate "Michael - AL" all-day entries exist on the "People Department - HR Systems" calendar (7 Aug and 10 Aug 2026), confirming the naming convention recurs.







**Fix:** added `_BARE_AL_RE = re.compile(r"al", re.IGNORECASE)` — standalone-word matching, not a plain substring, specifically because a substring match on bare "al" would false-positive constantly (inside "annual", "practical", "Sal", "Alan", "Alison", "Malcolm", "Salary", etc.). Verified against 12 real/adversarial cases (all passed) before touching production code. Wired in as an additional OR condition in `_is_leave_item()` (day-view exclusion) and the sidebar absence-detection loop's keyword check, plus a targeted `_BARE_AL_RE.sub(" ", ...)` step inside `_clean_absence_name()` so "Michael - AL" cleans to "Michael" rather than the literal "Michael - Al" (real names containing "al" as a substring, e.g. "Alan Smith", are provably untouched — verified with a standalone test before pushing).







**Verified against real production data, same run:** re-ran `Run Inbox Briefing.bat /update` (exit 0) and pulled the live `data/briefing.json` — "Michael - AL" no longer appears in `calToday`, and the sidebar `absences` list now correctly includes `"Michael O'Sullivan - off today, returns Tuesday 11 August"` (using the calendar item's real Organizer field, not the cleaned subject, since Organizer was a genuine person name here). Confirmed live in-browser too — screenshot of the Calendar tab and sidebar both matched.







---







## Session 2026-08-10 (continued again) — Calendar column height + Drafted Replies card style (Drew)







**Scope:** Two small follow-up UI requests from Kevin right after the 4-day calendar work above, both in `css/styles.css` only.







- `.cal-col-body` scroll cap raised from `260px` (tuned for the old 3-column layout, where day-columns sat beside a fixed-height mini-cal) to `560px` — Kevin: "we have a scroll bar but we have quite a lot of real estate beneath... make them longer so I have less to scroll." Now that the mini-cal moved to its own full-width row below (see above), the day-columns had no sibling height constraint and real spare page space was going unused. Still capped, not removed, so one exceptionally busy day doesn't blow out the page layout.



- `.dr-card` (Drafted Replies panel cards) — removed the `border-left:3px solid var(--purple)` accent bar so drafted-reply cards use a plain 1px border all round, matching every other card style on the dashboard (`.card-ph`, `.main-cal-block`, etc.) instead of standing out with a colour bar.







Both verified live in-browser after pushing (hard-reload + screenshot): taller day-columns show more of today's schedule without scrolling, and the Drafted Replies cards now have a plain border with no purple bar.



---



## Session 2026-08-11 — needs_reply staleness cutoff revised 60 -> 30 days (Drew)

**Scope:** Kevin's final word on the last open parameter of the needs_reply precision fix (agent-commons issue #3 step-3 brief). The fix itself -- capturing the To-vs-CC signal (`kevin_is_primary_recipient`), computing message age (`age_days`), passing both into the Phase 3.2 AI classification prompt as explicit signals, and a deterministic hard override that can only ever flip `needs_reply` from true to false (never the reverse) for anything past the cutoff -- was already fully built and live from earlier the same day (10 Aug 2026, see the "1 two months" confirmation earlier in this doc). Kevin's cutoff choice changed from 60 days ("two months") to 30 days.

**Change:** `STALENESS_CUTOFF_DAYS` in `fetch_inbox.py`'s Phase 3.2, `60` -> `30`. Nothing else needed changing -- `_kevin_is_primary_recipient()`, `KEVIN_EMAIL`, the `age_days` computation, and the `EMAIL_SUMMARY_SYSTEM` prompt instructions to the model were all already in place and unaffected by this threshold change.

**Verified against real production data:** re-ran `Run Inbox Briefing.bat /update` (exit 0) -- log line "2 flagged needs_reply (1 overridden false for being older than 30 days)", versus the prior 60-day runs earlier the same day which consistently showed "0 overridden" (no email happened to fall in the 30-60 day gap until the cutoff tightened). Pulled the live `data/needs_reply.json` and confirmed both surviving entries are genuinely recent (4 Aug and 27 Jul, i.e. 7 and 15 days old respectively as of 11 Aug) -- well inside the 30-day window, confirming the override is doing real work, not just present in the code.

---

## Session 2026-08-10 (continued again) — Calendar CC link now deep-links to the matching Command Centre task (Drew)



**Scope:** Kevin's explicit follow-up, same session: "whe i click on the cc on one of the schedules it take me to command centre but not the item - it should high[light] the item so i can drill dowwn into the email if required - one links to the other."



**Root cause:** the Calendar tab's per-meeting "CC →" link was always a bare `href="https://cc.lelitte.co.uk"` with no task id at all -- calendar meetings (raw Outlook data) never carried any Command Centre task reference. This is different from the Priorities tab's CC buttons, which already deep-link correctly via `#${p.id}` since priority cards ARE sourced directly from Command Centre's own `tasks.json` (confirmed by reading `command-centre/js/app.js` directly: on load it reads `window.location.hash`, looks up `document.getElementById('card-'+hash)`, scrolls to it, and adds a `deep-linked-<tier>` highlight class -- this mechanism already existed and works, it just had nothing to link to from the calendar side).



**Fix, `fetch_inbox.py`:** new `_match_cc_task_id()` -- for each calendar meeting, looks for an exact (case-insensitive) match between the meeting's title and a not-done Command Centre task's `emailRef` field. Confirmed live against real `tasks.json` that several tasks carry the verbatim meeting title in `emailRef` (e.g. "Sickness Absence Survey working group", "Confidential - OH Consultation"). Deliberately did NOT also match against `task.source` (which often names a meeting too, e.g. "HR Systems Managers Meeting 24/06") -- `source` carries a trailing date but no way to tell which week's occurrence of a *recurring* meeting it refers to, so matching against it risked deep-linking to a stale prior occurrence's task. If more than one not-done task shares the identical `emailRef`, no link is attached rather than guessing. Matched items get a new `ccTaskId` field.



**Fix, `js/app.js`:** the CC link now renders as `href="https://cc.lelitte.co.uk/#${c.ccTaskId}"` when `c.ccTaskId` is present, and is omitted entirely otherwise -- a link that goes nowhere useful is worse than no link, per Kevin's complaint.



**Verified:** `python -m py_compile` + `node --check`. Matching logic unit-tested against the real live `tasks.json` and 13 real calendar meeting titles seen this session -- exactly the 2 genuine matches came back ("Sickness Absence Survey working group" -> `t2608071200560`, "Confidential - OH Consultation" -> `t2608071501072`), zero false positives on the other 11. Node DOM-stub test of the real edited `renderCalPanel()` confirmed the matched item gets the deep-link href and the unmatched item gets no CC link at all (not the old generic homepage link). Real production run (`Run Inbox Briefing.bat /update`, exit 0) confirmed `ccTaskId` correctly attached to the same two live items in `data/briefing.json`, and a live-browser screenshot confirmed the CC link now shows on only those two meetings on the actual dashboard.



**Not verified live end-to-end (Chrome extension disconnected mid-session):** did not get a live click-through confirming the deep-link actually scrolls to and highlights the task on the Command Centre page itself. High confidence this works -- it's the exact same hash format and exact same `command-centre/js/app.js` mechanism the Priorities tab's CC buttons already use successfully -- but flagging honestly that this specific last step was verified by direct code inspection + matching test output, not a live click. Worth a quick manual click-check next session if Kevin hasn't already confirmed it works.















---















## Session 2026-08-10 — sent_corpus_pull.py built (Drew) — not yet run against real data















**Scope:** `begb0037admin/agent-commons` issue #3 (cross-agent email/Teams style-learning pipeline), item 3/4 — bulk-ingest Kevin's own Sent items as the initial style corpus, via Graph API originally, redirected mid-task to reusing work-inbox's proven Outlook COM access instead.















**What was added:** `tools/sent_corpus_pull.py` — a new, separate script, NOT a change to `fetch_inbox.py` or the live 6x/day pipeline. Reuses the identical COM connection pattern already in production (`win32com.client.dynamic.Dispatch("Outlook.Application")` → `GetNamespace("MAPI")` → `GetDefaultFolder(5)` for Sent Mail), but pulls full body text over a month-chunked historical `[SentOn]` window (existing `fetch_inbox.py` Sent read is 7-day/100-char-preview only, feeding ephemeral AI-triage context — never persisted, and not touched by this addition).















**Redaction pass (automated, per Kevin's decision):** keyword/pattern-based, 4 categories (`health`, `bereavement`, `hr_case`, `absence`) — any match anywhere in subject+body excludes the whole message from the corpus. Redaction ledger records only `entry_id`/date/category/known-name-flag, never matched text. Tested against 13 synthetic cases (all 4 categories + 2 negative controls, including a "leave" false-positive check) — 13/13 passed. Chunking logic separately verified for gaps/overlaps across a year boundary.















**Not done yet:** no real pull has been run against live Outlook — this session's environment didn't have Outlook running, and starting it to pull real historical mail was treated as past the "build and report" checkpoint, needing Kevin's explicit go-ahead first. Proposed durable output location: `begb0037admin/agent-commons` `corpus/sent-items/` (scaffold README pushed there, no real corpus data yet).















**Update, same session — real dry run against live Outlook (Kevin started Outlook Classic mid-session):** ran in `--stats-only` mode (aggregate counts only, nothing written to disk) against the real last-90-days Sent folder. First run: `total_seen: 740` vs `clean_count(327) + redacted_count(76) = 403` — 337 items (45%) silently disappearing through a bare `except: continue`. Root cause: Sent Items also holds meeting requests/responses/cancellations (COM `Class` 53/54/55/56/57), which lack mail-style `Body`/`To` and threw an unhandled `AttributeError` indistinguishable from a real bug. Fixed by explicitly filtering to `Class == 43` (`olMail`) up front instead of relying on exception shape. Re-run fully reconciled: 403 = 327 clean + 76 redacted (health 60, hr_case 16, bereavement 3, absence 2), zero unexpected errors on real mail items. ~19% of real Sent Mail over 90 days matched a redaction category.















Full writeup and open questions (recipient-PII in the `to` field, redaction being pattern-based not NLP): `begb0037admin/drew` `memory/sent-items-corpus-investigation.md` and `begb0037admin/agent-commons` issue #3 comments. Cross-agent Outlook COM gotcha (Sent Items non-mail Classes) also logged to `begb0037admin/agent-commons` `memory/index.json`.















**Still not done, on purpose:** no content written to disk yet (all runs stats-only), nothing pushed to `agent-commons/corpus/sent-items/` beyond the design-doc README. Next: real (non-stats-only) pull to local staging, spot-check locally, then push only the reviewed redacted corpus.json.















---















## Session 2026-08-10 (continued) -- draft_final_diff_capture.py built, real baseline established (Drew)















**Scope:** `begb0037admin/agent-commons` issue #3, forward-going half of the corpus approach (item 3) -- capture principal's draft-to-final edits over time, not just the one-time Sent-items backfill.















**Feasibility investigated first (read-only structural probe, no content read/stored):** Outlook's `EntryID` is NOT a safe key to correlate a Drafts-folder item with its eventual Sent-folder counterpart -- sending mints a new MAPI entry. `ConversationID` is: present on 40/40 sampled items in both Drafts (103 total) and Sent Items (1585 total).















**Built:**







- `tools/style_corpus_common.py` -- redaction classifier (health/bereavement/hr_case/absence), `recipient_tier` mapping, and the `OL_MAIL_CLASS` non-mail-item filter, factored out of `sent_corpus_pull.py` now that a second script needs the identical logic.







- `tools/sent_corpus_pull.py` -- refactored to import the shared module instead of duplicating it. Re-ran the original 13-case synthetic redaction suite + chunking test against the refactor -- zero regression.







- `tools/draft_final_diff_capture.py` -- periodic snapshot-and-correlate (not an event-driven listener -- considered `Application.ItemSend` for perfect fidelity, rejected for v1 since it needs a persistently-running process, a different architecture from every other script here). Snapshots Drafts each run, diffs against the previous run's local-only ledger to find vanished drafts, correlates against Sent Items by `ConversationID` within a 72h window (earliest match wins, no fallback guessing), applies the same whole-pair redaction exclusion as Sent-items (either side sensitive excludes both), computes `recipient_tier`, classifies `edit_type`/`note` via claude-haiku-4-5 on the redacted pair (confirmed OK with Kevin, same model `fetch_inbox.py` already uses).















**Verification:** correlation logic -- 5 mocked-Outlook cases (window bounds, non-mail filtering, multiple-candidate tiebreak, no-match), 5/5 pass. Whole-pair redaction gate -- confirmed both directions (draft-only and final-only sensitivity both correctly exclude the pair) with the real classifier. `edit_type` classification -- 5 synthetic pairs against the real API, 5/5 valid enum, 4/5 exact intended match (1 legitimately ambiguous test case, not a classifier bug). **Real baseline run against live Outlook:** 96 drafts tracked into the local-only ledger (`C:/Users/admin/Documents/CorpusStaging/draft_watch/ledger.json`, confirmed outside any git working tree), 0 vanished/0 pairs -- expected and correct for a first run, not a bug (the mechanism is inherently forward-looking).















**Not done, on purpose:** no diff pairs exist yet (need a real send to happen between two runs), nothing pushed to `agent-commons/corpus/draft-final-diffs/`. Not yet wired into Task Scheduler -- holding for confirmation given it makes live Anthropic API calls per pair on an unattended schedule.















**Also this session:** Teams draft-staging design moved from proposal to concrete (surface confirmed as work-inbox by Kevin) -- new "Pending Teams Replies" panel, data cross-fetched from `agent-commons/pending-teams-drafts/drafts.json` (mirrors the existing CC-ticker cross-repo-fetch pattern; preserves the standing rule that Lauren never writes into work-inbox directly), reusing the existing `workInbox_ticks_v1` Cloudflare-Worker-synced tick mechanism for "mark as sent" rather than building new write-back infra. Design only, not built at the time this entry was written. **Superseded same day (10 Aug 2026):** Kevin explicitly decided against pursuing this at all -- "Teams access -- resolved: ad-hoc, no automation." No Teams read access, no automation panel; Teams replies stay manual/ad-hoc (paste to Lauren, paste the drafted reply back into Teams by hand), permanently, not "still deciding." This design is parked, not deleted, for reference if Teams-reply volume ever justifies revisiting -- but is not on the roadmap. Full detail: `begb0037admin/agent-commons` issue #3.















---















## Session 2026-08-10 (continued again) -- draft_final_diff_capture.py hardened and scheduled (Drew)















**Scope:** Kevin decided to schedule `draft_final_diff_capture.py` on Task Scheduler now rather than run it manually. Before scheduling anything that makes unattended, no-human-in-the-loop Anthropic API calls, hardened the script and verified the hardening, per Kevin's explicit ask to double-check error handling/rate-limit/cost safety first.















**Hardening added:**







- Decoupled correlation+redaction (cheap, local, must never be lost) from AI classification (has cost/rate considerations) into two phases. The ledger is now saved unconditionally before any AI-related code runs, so an Anthropic-side problem can never risk losing track of currently-open drafts.







- New local-only backlog (`pending_classification.json`) holds pairs that have passed redaction but haven't been classified yet -- either because a run found more pairs than its per-run cap, or a classification attempt failed. Nothing is dropped just because the AI step had a bad run.







- `--max-classifications-per-run` cap (default 25) bounds live Anthropic API calls per run -- protects against an unbounded cost/rate-limit spike if many drafts vanish at once (real burst, or a ledger bug). Overflow waits in the backlog for the next scheduled run.







- Anthropic client instantiation itself wrapped in try/except (bad key, package issue) -- degrades to `ai_unavailable_this_run: true` and preserves the backlog, rather than crashing the whole run.







- Each backlog item gets up to `MAX_CLASSIFICATION_RETRIES` (3) attempts across runs before being permanently logged to `draft_final_classification_failures.json` and dropped -- not retried forever.







- Proper exit codes: the `__main__` block now wraps the whole run in try/except and exits 1 on any unhandled failure, so Task Scheduler's own restart/failure detection actually sees a real failure rather than a silent no-op.















**Verified before trusting it (mocked Outlook + mocked Anthropic, no real data or real API calls):**







- Cap + carryover: 3 pairs found with cap=2 -> 2 classified immediately, 1 carried to the backlog and classified on the very next run, all 3 eventually published, none lost.







- Anthropic client init failure: run completes without crashing, `ai_unavailable_this_run: true`, the 1 pending pair correctly preserved in the backlog rather than lost.







- Persistent classification failure: pair retried across exactly 3 runs (retry_count 1, 2, 3), then permanently logged to classification_failures on the 3rd, backlog correctly empty afterward -- confirmed it doesn't retry forever.







- Re-ran against real live Outlook after hardening (`--stats-only`): 96 drafts tracked, 0 vanished/0 pairs -- consistent with the untouched baseline, no regression.















**Scheduled live, `.bat` launcher mirrors the existing Work Inbox Briefing pattern exactly** (fresh-pull-from-GitHub with an integrity check on each downloaded file, timestamped backup before overwrite, single last-run log via `Tee-Object`, exit code propagated) -- `D:/OneDrive - lelitte.com/Desktop/Run Draft Diff Capture.bat`, downloads both `draft_final_diff_capture.py` and `style_corpus_common.py` fresh each run. One deliberate improvement over the copied pattern: added explicit early-exit guards for `/update` and `/run` invocations so an unattended Task Scheduler run can never land on the original pattern's interactive `choice` menu prompt at the end.















**Task Scheduler task "Draft Diff Capture":** hourly, 7am-7pm, Mon-Fri (13 runs/weekday) -- picked to sit meaningfully more frequent than the existing 5x/day Work Inbox Briefing cadence, since this mechanism can only capture a draft if it's still open at the moment of a poll; hourly gives a real chance of catching messages that get genuine editing attention (which are also the most valuable ones for a style corpus) without over-polling for the many replies that are drafted and sent within minutes and were never going to be caught regardless of interval -- an inherent floor of the whole draft-snapshot approach, not something interval choice fully solves. Settings mirror Work Inbox Briefing (`StartWhenAvailable`, `RestartCount 2`, `RestartInterval 5min`, `ExecutionTimeLimit 15min`), plus `MultipleInstances IgnoreNew` so a slow run (e.g. many pairs to classify) can't stack overlapping runs.















**Verified live end-to-end** by running the `.bat` in `/update` mode manually (the exact invocation Task Scheduler uses) before registering the task: real GitHub download with integrity check passed for both files, real run against live Outlook completed (96 drafts tracked, 0 pairs -- consistent, expected), exit code 0, log written to `C:/Users/admin/Documents/Claude/Projects/work-inbox/tools/draft_diff_capture_last_run.log`. Task registered and confirmed `State: Ready`, next run 12:00 today.















**`sent_corpus_pull.py` stays manual** -- Kevin confirmed the one-time snapshot is sufficient, no scheduling needed there.















Full detail: `begb0037admin/agent-commons` issue #3.















---















## Session 2026-08-10 (final) -- Drew-to-Lauren wiring built and live: needs_reply flagging, needs_reply.json, Drafted Replies panel (Drew)















**Scope:** `begb0037admin/agent-commons` issue #3 step-3 brief, items 1/2/4 -- the actual drafting hand-off loop (Drew finds -> Lauren drafts -> Kevin reviews). Kevin gave final go-ahead after item 4 was decided as 4B (dashboard-only, no live-mailbox writes).















### Item 1 -- Phase 3.2 extended to flag needs_reply, real bug found and fixed twice















`fetch_inbox.py` Phase 3.2 (the existing per-email AI summary call over urgent+needs cards) now also returns `needs_reply: true/false`. Real production testing (not just unit tests) caught a genuine bug before it could reach the unattended scheduled run:















1. **First real run**: Phase 3.2 failed outright -- `Expecting ',' delimiter` JSON parse error. Root cause looked like a size problem, so `max_tokens` was raised 4096 -> 8000. Still failed on retest (`Unterminated string`).







2. **Root-caused properly** by reproducing the exact real 157-candidate payload with full diagnostics: `stop_reason: max_tokens`, `output_tokens: 8000` -- genuinely hitting the ceiling, but only 18KB of content, because the ~140-char raw Outlook EntryID used as the JSON map key for every entry was consuming most of the token budget before the model reached the actual summaries (hex strings tokenize far less efficiently than English text).







3. **Real fix**: switched to short sequential ids ("0","1","2"...) in the API exchange, mapping back to the real EntryID locally by array position. Confirmed on the identical real payload: `stop_reason: end_turn`, only 5947/8000 tokens used, all 157 entries parsed.















Verified against real live Outlook data across three full production `.bat` runs this session -- final state: 157/157 candidates get both `ai_summary` and `needs_reply`, 14-33 flagged true depending on the run (inbox contents change between runs).















### Item 2 -- work-inbox/data/needs_reply.json, published by tools/publish_needs_reply.py















New script, separate from `fetch_inbox.py` (keeps that script's single-file-pulled-fresh deployment model unchanged), fetches full body via Outlook COM for `needs_reply==true` entries only, applies the same redaction classifier already built for the corpora (`style_corpus_common.is_sensitive`), computes `sender_tier` (reusing `recipient_tier()` against the sender), writes `data/needs_reply.json`. Real runs this session: 20 flagged -> 16 published, 4 redacted; 21 flagged -> 16 published, 5 redacted -- redaction is doing real work, not a no-op. Self-consistency check (re-classifying every published entry) confirmed zero false negatives.















### Item 4 -- Drafted Replies panel, 4B (dashboard-only), plus a real architecture correction















Original design assumed the dashboard could cross-fetch `agent-commons/pending-email-drafts/drafts.json` the same way it already fetches `command-centre/data/tasks.json`. Tested empirically instead of assuming: `agent-commons` is a **private** repo (`gh api ... --jq .private` -> true), and the existing `github-proxy.lelitte.co.uk` Worker returned 404 for it (200 for the identical request against work-inbox) -- confirmed the shared proxy's own token can't read it. **Fix:** `tools/publish_drafted_replies.py`, a new script holding the real `GITHUB_PAT`, reads `agent-commons/pending-email-drafts/drafts.json` directly and mirrors only the already-redacted/tier-tagged content into `work-inbox/data/drafted_replies.json` -- the dashboard reads that as an ordinary same-repo file, agent-commons itself is never exposed to any client-side/anonymous reader.















Dashboard changes (`index.html`, `css/styles.css`, `js/app.js`): new "Drafted Replies" panel, distinct purple accent (not merged into the Today/Tomorrow/Week/Parked grid), per-card subject/`sender_tier` badge/timestamp/expandable draft text/Copy-to-clipboard/"Open original" (reusing the existing `openmail://` handler)/Mark sent/Discard. Mark sent/discard is bookkeeping only -- rides the exact same tick-sync mechanism (`getTicks`/`saveTicks`/`pushTicks`, existing `inbox-state` Worker route) already used for email cards, under a `draft_` key prefix so it doesn't collide with per-day briefing ticks. No new Worker route, nothing writes to a mailbox or sends anything.















Verified with a temporary synthetic seed pushed to `agent-commons/pending-email-drafts/drafts.json`: confirmed the mirror script correctly picks it up, and ran the actual `renderDraftedReplies()` function (not a reimplementation) in a Node DOM-stub harness against the real mirrored payload -- correct escaping (apostrophes/ampersands), correct tier badges, correct action wiring. No real browser was available in this environment to screenshot (Chrome extension not connected) -- flagged as a real limitation, not glossed over. Test seed reverted from agent-commons afterward; only Lauren's real content should live there.















### Full chain verified live, three times, via the actual production `.bat`















`Run Inbox Briefing.bat` now chains `fetch_inbox.py` -> `publish_needs_reply.py` -> `publish_drafted_replies.py` in one run (each downstream step non-fatal to the overall briefing if it fails). Final confirmed run: fetch_inbox.py succeeded, needs_reply.json published (16 entries, byte-identical verified), drafted_replies.json published (correctly empty, `source_missing: true`, since Lauren hasn't written anything yet), exit code 0 throughout.















One process-hygiene lesson from this session: nesting `run_in_background` (the Bash tool) around a command that ALSO backgrounds itself with a trailing `&` produces an orphaned, untracked process -- it happened here and briefly locked `inbox_briefing_last_run.log` for a real still-running `fetch_inbox.py` instance. Resolved by waiting for the orphaned PID to exit naturally rather than killing it (it was doing real, legitimate work, just detached from the tool's own tracking).















**Not done, on purpose:** nothing pushed to `agent-commons/corpus/draft-final-diffs/`-adjacent locations by this session; Lauren's own `pending-email-drafts/drafts.json` doesn't have real content yet, so the Drafted Replies panel is correctly empty in production right now -- that's expected, not a bug.















Full detail: `begb0037admin/agent-commons` issue #3.















---















## Session 2026-08-10 (final) -- Absences bug fixed and verified live (Drew)















**Scope:** Kevin reported the sidebar Absences list showing duplicates ("Simon" and "Simon Burford" as separate entries) and "date unknown" on 8 of 10 entries. Root-caused and proposed on agent-commons issue #3 before building.















**Root cause:** two unreconciled detection passes feeding one dict. The calendar pass keyed entries by whatever's left in the calendar item's subject after stripping leave keywords (often just a first name); the email-OOO-fallback pass keyed by the full Outlook sender display name and was **hardcoded** to always label "date unknown" -- it never attempted any date extraction at all. Different string keys for the same real person produced duplicates; the hardcoded fallback explained the date-unknown rate.















**Fix A -- name reconciliation via the Organizer field.** Calendar items already carry `item.Organizer` (was being pulled, just never read by absence-detection). Verified live before building: `Organizer` holds the exact same full display name Outlook uses as the email sender name (`'Simon Burford'`, `'Athena Artuso'`, confirmed against real calendar items in the detection window). Now used as the primary name source for calendar-derived entries, falling back to the subject-derived name only when Organizer is empty.















**Fix B -- best-effort OOO-text date extraction**, explicitly non-exhaustive: tries a handful of common phrasings ("until 18 August", "back Monday", "returning 18/08") before falling back to "date unknown". Genuinely unparseable text still correctly falls back rather than guessing wrong. Guessed dates are labeled "(best guess from email text)" so they're never confused with a calendar-verified date.















**Verified, real data, real production run:** live `briefing.json` absences count went from 10 (with duplicates) to 8 (deduplicated). "Athena"/"Athena Artuso" and "Simon"/"Simon Burford" each correctly merged into one real-dated entry. Two previously-"date unknown" entries (Crispin Muncaster, James Salas Guillen) now show best-effort guessed dates, clearly labeled. The remaining four (Christopher Sanders, Julie Hickman, Marie Cooksey, Sarah Rowles) genuinely have no extractable date and correctly stay honest about it rather than guessing.















**Also this session (same production run, already reported separately on issue #3):**







- `tools/publish_drafted_replies.py` schema bug fixed -- Lauren's real entries use `composed_at` not `drafted_at`, were being silently dropped; also now surfaces `confidence`/`inline_flags` on the dashboard, which the original design never accounted for. Verified with real content: 4/4 of Lauren's real drafts now publish and render correctly.







- `needs_reply` precision investigated (Lauren found ~20/24 flagged entries were false positives) -- root cause: no cc-vs-primary-recipient signal and no staleness signal reach Phase 3.2's classifier at all (neither is captured/passed). Proposed fix posted to issue #3, not yet built -- needs Kevin's sign-off on staleness-cutoff specifics first.















Full detail on all of the above: `begb0037admin/agent-commons` issue #3.















---















## Session 2026-08-10 (final, continued) -- Absences: calendar-only sourcing per Kevin's decision (Drew)















**Scope:** Kevin corrected the earlier absences fix -- he doesn't want OOO-email-guessed dates at all; his own Calendar plus the "People Department - HR Systems" calendar (confirmed real and enumerable earlier this session) are the absence source of truth. If someone's leave isn't logged in either, he does not want it surfaced.















**Built:** Phase 1 now also pulls the "People Department - HR Systems" calendar (an "Other Calendar" nested under Kevin's own primary mailbox, reached via the same COM session, wrapped in try/except so a folder-structure change degrades gracefully rather than failing Phase 1). Its items merge into the same `calendar` list Kevin's own primary calendar already populates, so the existing (Organizer-based) absence-detection logic picks them up with no separate code path. The OOO-auto-reply-email fallback -- and the best-effort date-guessing built for it earlier the same day -- were deliberately deleted, not just left unused: with calendar-only sourcing, every remaining absence entry has a real calendar-verified date by construction, so "date unknown" can no longer appear at all.















**A real, production-only edge case was caught and fixed.** A live run produced a bogus absence entry -- "People Department - Hr Systems - off today..." -- where a calendar item's `Organizer` field held the department's own name rather than a real person (likely how a particular admin-booked half-day/full-day entry was created). The existing organizer-placeholder pre-check should have caught this but didn't reproduce when replicated with identical logic moments later in the same session -- most likely a non-deterministic Outlook COM quirk specific to expanding a recurring series via `IncludeRecurrences`, not a pinned-down logic bug. Rather than keep chasing an intermittent trigger, added a defense-in-depth output-side guard in `_add_absence()`: reject any cleaned name that still contains obviously-non-person terms ("department", "systems", "team"), regardless of which mechanism produced it. Re-ran production after this fix: the bogus entry is gone, correctly replaced by "Kevin" (the real underlying person, via subject-derived fallback).















**Verified, real production data, three consecutive real runs today:**







- Run 1 (calendar-only sourcing, no fallback): 7 real entries, zero "date unknown", but included the bogus department-name entry.







- Run 2 (first placeholder-organizer fix attempt): bogus entry persisted -- confirmed the first fix attempt was insufficient on its own.







- Run 3 (defense-in-depth guard added): bogus entry gone, replaced by the real person ("Kevin"). Final live state: `Athena Artuso`, `David Johnson`, `Henry Acheampong`, `Julie Hickman`, `Kevin`, `Simon Burford`, `Susan Pratt` -- all real, calendar-verified dates, zero "date unknown", zero non-person entries.















One observation, not acted on unilaterally: Kevin's own leave now legitimately appears in his own Absences panel ("Kevin - off today..."), since he's tracked in the same calendar as everyone else. Not something he asked to exclude -- flagging it as a minor, possibly-odd-but-correct side effect rather than silently filtering it.















Full detail: `begb0037admin/agent-commons` issue #3.















---















## Fix list















3. **Drag reorder animation** — No visual feedback during drag. Cards need to visually shift in real time as Kevin drags — placeholder in the DOM during `dragover`.















4. **Phase 3.8 calendar-summary mismatch on days starting with an all-day event** — investigated by Drew 2026-08-04, root cause confirmed, NOT fixed (Phase 3.8 is closed — needs Kevin to explicitly reopen it before any code change). See "Session 2026-08-04" below for full detail and full writeup in `begb0037admin/drew` repo, `memory/calendar-summary-offset-bug.md`.















---















## Session 2026-08-04 — Calendar-summary offset bug investigated, NOT fixed (Drew)















**Scope:** Flagged during unrelated meeting-records work — live `data/briefing.json` `calTomorrow` items had AI-generated `summary` text describing a *different* meeting than the one it was attached to. Investigated whether this is a Python index bug in `fetch_inbox.py` Phase 3.8.















**Confirmed against live `data/briefing.json` (Wednesday 5 August briefing), pulled fresh with a cache-buster:**







- `calToday` (9 items, first item is a real 09:30 meeting, no preceding all-day item): zero mismatches, every summary correctly self-referential.







- `calTomorrow` (9 items, idx 0 is an all-day event — "Simon out of the office - funeral"): idx 1, 2, 3 each carry the summary content that rightfully belongs to the *next* item (idx 1 shows idx 2's title, idx 2 shows idx 3's title, idx 3 shows idx 4's actual topic while idx 4 itself is left with no summary). idx 6 and 7 are correctly self-referential — the mismatch does not persist for the whole day.















**Root cause (confirmed, not guessed):** the Python index bookkeeping in Phase 3.8 (`_cal_for_summary` construction and the `target[item["idx"]] = ...` write-back, ~line 1168 onward) is correct — re-read line by line, positions match. The mismatch correlates exactly with whether the day's `idx` sequence fed to the model starts at 0 or not. `calToday` starts at idx 0 (no shift). `calTomorrow`'s first non-all-day item carries idx 1 (idx 0 was filtered out as an all-day event) — and the model's own generation of the `"day_idx"` JSON response keys (`tomorrow_1`, `tomorrow_2`, ...) mismatches the content it writes for the first few real items before self-correcting by idx 6. This is a **prompt/model reliability issue** (claude-haiku-4-5, the model Phase 3.8 is locked to, appears to fall back to counting output position from 0 rather than reliably echoing the literal `idx` value whenever that value doesn't start at 0) — not a deterministic code bug.















**Proposed fix (not applied):** decouple the index shown to the model from the index used for write-back — renumber the `idx` sent to the model to always start at 0 within `_cal_for_summary` (sequential by array position), and keep the original `cal_today_items`/`cal_tomorrow_items` position in a separate field never exposed to the model, used only for the write-back. Small, contained diff, same file.















**Why this was not pushed:** Phase 3.8 is marked closed in this file and in `CLAUDE.md` — "do not modify without Kevin explicitly opening a new approved phase." No message in the task that surfaced this constituted that explicit reopening, so this was investigated and root-caused only, not fixed. Full writeup: `begb0037admin/drew` repo, `memory/calendar-summary-offset-bug.md`.















---















## Session 2026-08-02 — CC ticker done-task filtering fix (Drew's first task)















**Scope:** Fix bug where the "Command Centre Focus" sidebar ticker (`loadCcTicker()` in `js/app.js`) counted ALL Command Centre tasks per tier, including tasks marked `done: true`, so it disagreed with Command Centre's own "Daily Focus" tile, which correctly counts only open (`!t.done`) tasks per tier.















**Confirmed against live `command-centre/data/tasks.json` before the fix:** 39 tasks total, 13 marked done. All-tasks counts (old, wrong) vs. open-only counts (new, correct — matches CC's own tile):







| Tier | All (old) | Open only (new/CC) |







|---|---|---|







| Today | 10 | 5 |







| Tomorrow | 6 | 4 |







| Week | 13 | 8 |







| Parked | 10 | 9 |















**Root cause:** none of the four `tasks.filter(t=>t.tier===...)` calls in `loadCcTicker()` excluded `t.done`. Command Centre's own `js/app.js` `renderBoard()` does `tasks.filter(t=>t.tier==='today'&&!t.done)` — work-inbox's ticker didn't match that.















**Fix (commit `8582608`):** added `const openTasks=tasks.filter(t=>!t.done);` right after the tasks array is built, and switched all four tier-count filters (`cc-today-count`, `cc-tmrw-count`, `cc-week-count`, `cc-parked-count`) plus the age-based stats (`ages`, `stalled`, `oldest`, `avg`, `twoWeeks` — i.e. `cc-stalled`, `cc-oldest`, `cc-avg`, `cc-twoweeks`) to run over `openTasks` instead of the full `tasks` array. Judgement call, not explicitly requested by Kevin: extended the fix to the age stats too, for consistency — a completed-but-old task shouldn't be able to drag "Oldest task"/"Avg age" up or count toward "stalled"/"2+ weeks old", since those exist to flag *open* work going stale. Worth Kevin's explicit confirmation if this reads wrong once he's looking at real numbers.















**Verification:** pushed via GitHub Contents API PUT, base SHA `cbf52b72a7b84ceed9df287ceb9d5436d55ccc09` → new SHA `d07b6c97b65d5b9496466d011dbd4fa2071f1f55`. Re-fetched the pushed blob via `git/blobs/{sha}` (not `raw.githubusercontent.com`, which caches) and diffed byte-for-byte against the intended patched file — exact match. Re-fetched live `command-centre/data/tasks.json` after the fix and simulated the new `loadCcTicker()` logic: Today 5 / Tomorrow 4 / Week 8 / Parked 9 — matches Command Centre's own `renderBoard()` tier-count logic (`tier===x && !done`) exactly on the same live data. `node --check` confirmed the pushed file is syntactically valid JS.















**Diff scope confirmed minimal:** the only changes in the file are inside `loadCcTicker()` — one added line (`openTasks`) and eight filter calls switched from `tasks` to `openTasks`. Nothing else in `js/app.js` touched.















**Also flagged to Kevin, not actioned this session (his call to prioritize):**







- **Command Centre "Save failed — HTTP 502" + no way to exit inbox-suggestions view** — worker-side 502 from `cc-tasks-writer.kevinlelitte.workers.dev` on `persistTasks()`, cause not yet confirmed (possibly inbox auto-promotion write frequency, possibly cold start). Separately, `showView('inbox')` in command-centre's `js/app.js` has no matching `showView('board')` control in the markup — clicking "From your inbox" strands the user on the Inbox Suggestions view with no way back except F5. Kevin previously asked to hold other command-centre work until this is resolved — likely next in line.







- **Work Inbox "Command Centre Focus" ticker redesign (6-across / drop Parked, widen to show Urgent + Needs response)** — dropped entirely by Kevin mid-session, 2026-08-02: the two tiles (CC's own Daily Focus vs. work-inbox's cross-reference into CC) are supposed to show different things by design; the redesign ask was based on a misunderstanding, not a real gap. Nothing from that thread was ever pushed to GitHub — it only existed as unapproved scratchpad edits (`wi_app_fresh.js`, `wi_index_fresh.html`, `wi_styles_fresh.css`) — so there is nothing to revert on the live site. Do not resurrect this thread without Kevin raising it again.















---















## Session 2026-07-04 — Absence tomorrow-detection fix (commit `3aab85c`)















**Scope:** `fetch_inbox.py` absence detection extended to surface tomorrow's leave in the sidebar absences panel. Weekend-aware labelling added.















**What changed:**







- Absence detection block replaced with version that scans both today and next working day.







- Today's absences on weekends/Sundays show `"(next week)"` suffix — avoids "today" implying a working day when today is Saturday.







- Absences starting on `tomorrow` (= `next_workday(today)`) labelled `"(tomorrow)"` on Mon–Thu, `"(next week)"` on Fri/Sat/Sun.







- Shared `_extract_absence_name()` helper removes duplication from the name-stripping logic.







- No duplicate checking needed: date logic naturally prevents double-listing the same person.















**Kevin approval:** "Yep, approved."















---















## Session 2026-07-04 - Pipeline hardening review follow-ups















**Scope:** Apply quick review follow-ups after Granola rollout.















**What changed:**







- `fetch_inbox.py`: Added a shared GitHub API timeout for script GitHub reads/writes.







- `fetch_inbox.py`: Made Phase 3.6 task action append idempotent by skipping exact duplicate action text.







- `fetch_inbox.py`: Renamed the Granola comment to Phase 3.7b to reduce diagnostic ambiguity; behaviour unchanged.







- `js/app.js`: Added HTML escaping for calendar times, titles, organisers, and summaries before rendering.















**Remaining non-blocking improvement:** A first-class DRY_RUN mode would still make future diagnostics safer because Phase 3.6, Phase 4, and Phase 5 can write to GitHub.















---















## Session 2026-07-04 - Granola calendar context fix (CLOSED — do not reopen)















**Scope:** Fix Phase 3.7 Granola context and improve Phase 3.8 meeting prep summaries.















**What changed:**







- `fetch_inbox.py`: Granola note detail extraction now falls back from `summary` to `summary_text` / `summary_markdown`.







- `fetch_inbox.py`: Granola context passed into Phase 3.8 increased from 500 to 1500 characters.







- `fetch_inbox.py`: Phase 3.8 now asks for 2-3 concise prep sentences and has a 900 token response budget.















**Validation:** Local debug smoke test confirmed `FA Team Daily Catchup` matched `FA Team Catch-up - 03/07`; dashboard smoke test used `Company 90 - Status Update` and confirmed the calendar summary display works.















**Not included:** No title matching changes, no forced debug matches, no diagnostic logging spam, no phase skip flags, and no `fetch_inbox_debug.py` changes in production.















---















## Session 2026-07-03 — Calendar scroll (approved, pushed to main)















**Scope:** Replace expand/collapse toggle on Today and Tomorrow calendar columns with independent vertical scrolling. Keep fixed height (260px), same size and position.















**What changed:**







- **`css/styles.css`** (commit `dc3544b`): Removed expand/collapse styles (`.cal-col-body` with `overflow:hidden`, `.cal-expand-footer`, `.cal-expand-btn`). Added scroll styles — `.cal-col-body { max-height: 260px; overflow-y: auto; overflow-x: hidden }` with 4px webkit scrollbar (`#d1d9e6` thumb, hover `#94a3b8`).







- **`js/app.js`** (commit `6589384`): `renderBlock()` inside `renderCalPanel()` — return statement no longer includes `cal-expand-footer` div. `toggleCalExpand()` function removed entirely. Both Today (`calBodyToday`) and Tomorrow (`calBodyTom`) columns now scroll independently via the same `renderBlock` code path.















**Kevin approval:** "perfect, approved ensure that it's on both columns today and tomorrow."















---















## Session 2026-07-03 — Granola 0-matches investigation (superseded — see CLOSED phase above)















**Scope:** Diagnosing why Phase 3.7 Granola fetch returns 10 notes but matches 0 calendar items.















**Resolution:** Fixed 2026-07-04. Root cause was `summary_text`/`summary_markdown` fallback missing. See CLOSED phase entry above.















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







- fetch_inbox.py — all phases confirmed working







- **Granola calendar context (Phase 3.7b + 3.8)** — COMPLETE. Matching via keyword overlap; summary extracted from `summary_text`/`summary_markdown`. Do not modify.







- **Absence detection** — today's leave + tomorrow's leave (weekend-aware labelling). Commit `3aab85c`.







- Task Scheduler — `WorkInbox-0900` / `WorkInbox-1200` / `WorkInbox-1500` (Mon–Fri)







- Dashboard loads live briefing.json on load, falls back to localStorage archive







- Oxford navy sidebar — crest (external `images/oxford-crest.jpg`), branding, live clock, filter, CC ticker, absences, all 6 links populated







- 3-column calendar panel (Today `7fr` | Tomorrow `7fr` | July+August mini-cals in one card `4fr`)







- **Calendar columns scroll independently** — Today and Tomorrow each have `max-height: 260px; overflow-y: auto` with 4px scrollbar. Expand/collapse removed.







- Rotating context strip with "Briefing context" label, dot nav







- 2×2 priority grid with tier filter — flat `.card-ph` design, NEW/UPDATED badges on right







- CC ticker reads live from CC tasks.json every 60s







- drag-and-drop, tick sync, archive, show done, openmail:// all working







- Multi-machine setup complete (begb0037.AD-OAK)















### Known issues (fix next session)







- Drag reorder has no visual animation







- Phase 3.8 calendar-summary mismatch on days starting with an all-day event (see Fix list item 4 and Session 2026-08-04 above) — root cause confirmed, fix scoped, awaiting Kevin's explicit reopening of Phase 3.8















---















## localStorage Keys















| Key | Purpose |







|-----|--------|







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







- **Phase 3.7b and Phase 3.8 are closed** — do not modify without Kevin explicitly opening a new approved phase







