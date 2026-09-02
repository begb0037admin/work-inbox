# Checkpoint addendum — 2 Sept 2026 (afternoon), fold into HANDOVER.md top

**Interrupted write notice:** Drew (agent `ab5c74fc17985094a`) was mid-write on the proper
HANDOVER.md top-entry prepend when he hit his own session rate limit (resets 2:50pm
Europe/London). This file is a stopgap so the checkpoint isn't lost — fold this into
HANDOVER.md's top entry next session in the file's existing style, then delete this file.
Full detail also lives in local coordinator memory (`work-inbox-mail-pull-imap-migration.md`).

## Summary

**#31 (`drew/lane-b-cal-guard-snapshot-fix`) is very likely fixed. Formal re-confirmation
still outstanding. Write-safety hardening built on the same branch, unmerged. Calendar
stays on COM until the formal confirmation runs.**

1. Kevin confirmed: the ChatGPT connector "action control" UI setting was already checked
   extensively weeks ago — do not re-investigate, treat as closed/non-blocking.
2. Write-risk framing settled after extended discussion: prompt-injection-via-calendar-content
   is low-probability for Kevin's actual environment (internal Oxford colleagues, not
   public-facing); the realistic trigger is model over-reach during a degraded/retried
   connector call — exactly what happened in the original dry-diff failure. Standing rule now
   baked into Lane B's prompts: never decline/cancel/respond/send, zero recipients if a write
   is ever unavoidable.
3. Drew's scheduled-task-into-Kevin's-live-RDP-session proposal was correctly declined as a
   judgment call — cross-account command injection into a live domain session on an
   Oxford-managed, policy-monitored machine needs Kevin's own direct explicit yes with the
   mechanics understood, not an inference from "he's tired of pasting commands." Stays parked.
4. Drew's write-safety build this session, all on branch `#31`, nothing merged/live:
   - **approval_policy gate confirmed real** from actual codex-cli source
     (`mcp_tool_call.rs`, `exec/src/lib.rs`): a write-capable connector tool triggers an
     approval request; headless `codex exec` has no interactive channel so it auto-aborts
     rather than executing. Oxford's enterprise policy already forces `approval_policy` off
     `Never`.
   - **Toast notification, laptop + desktop, built** — extends the existing
     `Push-LaptopRunStatus.ps1` -> GitHub -> `Watch-BridgeBriefing.ps1` channel with a
     `lane_b_guard=halted` field, firing an independent toast on both machines the moment
     the guard trips (previously silent).
   - **Prompt hardened** with an explicit `SAFETY_RULE` on every Lane B prompt per point 2
     above.
   - **Retry-attempt audit clean by code inspection**: `guard_recontamination()` runs on
     every retry attempt, not just the final one.
5. A one-click diagnostic script (`docs/desktop-scripts/Lane-B-Dry-Diff-Isolate.ps1`, on
   branch `#31`) had two broken pushes — double-base64-encoding, then a
   `raw.githubusercontent.com` CDN stale-cache issue that persisted even after the git blob
   was correctly fixed (confirmed via the Contents API showing correct content while `curl`
   on the raw URL still served the old garbage). Worked around with an inline PowerShell
   here-string instead of a GitHub fetch. **Gotcha for future pulls:** don't trust
   `raw.githubusercontent.com` immediately after a fix-up push; verify via the Contents API
   or use a fresh path/inline content instead of waiting on cache invalidation.
6. **Manual re-test, run live in RDP:** two ad-hoc `--snapshot` pulls (`test1.json`,
   `test2.json`), both 51 events, **both** containing every annual-leave/non-working-day
   subject from the previous day's 8-diff failure list (Kevin/Chris/Michael/Julie/Marie K),
   just different JSON ordering. Formal `--diff`: `{"trips": [], "tripped": false}` —
   **CLEAN.** Strong evidence the original failure was a one-off connector flake during that
   run's retry storm (`list_events did not fire` x2 before succeeding), not a systematic bug
   in the connector or the guard's key-matching logic.
7. **Formal confirmation via the real `--dry-diff` entry point was attempted, not
   completed:** failed because the local `lane_b_cal_guard.py` on the laptop had reverted to
   the OLD main-branch version (missing `--dry-diff`/`--selftest` from its argparse usage),
   despite having been correctly pulled from the `#31` branch earlier the same session.
   **Cause not diagnosed — open question for next session:** what overwrote it? Check
   whether the `fetch_inbox.py` refresh, the shadow run, or anything else touching that
   directory this session could have reset it.

## Exact next action

```powershell
cd $env:USERPROFILE\work-inbox
iwr -UseBasicParsing "https://raw.githubusercontent.com/begb0037admin/work-inbox/drew/lane-b-cal-guard-snapshot-fix/lane_b_cal_guard.py?cb=$([guid]::NewGuid())" -OutFile .\lane_b_cal_guard.py
python -m py_compile .\lane_b_cal_guard.py
python .\lane_b_cal_guard.py --dry-diff
```

If `CLEAN`: `#31` is validated — merge it (Kevin's go-ahead), the write-safety commits land
with it, then only reconsider the connector cutover with Kevin's explicit go. If it fails
again: first check why the local guard file keeps reverting before re-diagnosing the
all-day-event question.

## Do not re-open

- The ChatGPT connector action-control UI check (Kevin: already tested weeks ago, closed).
- Prompt-injection-via-calendar-content as a primary risk (concluded low-probability here).
- The scheduled-task-into-live-session automation (parked, needs Kevin's own direct explicit
  request).
