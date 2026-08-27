# Mailbox delta-sweep kill-switch — `tools/codex_triage/mailbox_guard.py`

**Built + proof-fired 2026-08-27 (Drew).** Originally the hard prerequisite for
running the AI-triage phases on Kevin's live ChatGPT account via Codex (residual
write-risk accepted, detection-not-prevention). **The route has since pivoted to
headless Claude Code** (`docs/CLAUDE_CODE_HEADLESS_SCOPE.md`), which has **no
mailbox tool** — so this kill-switch is now a **lightweight optional regression
check**, not a gate. Kept because it is built, proven, and cheap to run as
before/after insurance around the first few live headless runs.

## What it does

Before/after Outlook COM snapshot around a run; on **any** mailbox mutation the
run did not intend (its only intended output is local files), it:

- **(a)** hard-disables the named scheduled task (`schtasks /Change /DISABLE`),
  with a hard refusal if the name is the live `Work Inbox Briefing` pipeline;
- **(b)** alerts Kevin — reuses the live pipeline's own
  `Show-TaskNotification.ps1` / BurntToast path + a loud timestamped console
  banner;
- **(c)** writes a timestamped incident record (`data/codex_runs/guard_incident_*.json`)
  and a `GUARD_TRIPPED.flag` sentinel that a run wrapper checks and refuses to
  proceed past until `clear-flag` is run.

Exit codes: `0` clean · `2` tripped · `3` refused (live-pipeline name) · `4` COM error.

## Sweep scope (Outlook COM, Kevin's primary mailbox)

- Top-level **Inbox** — newest 250 within 30 days, **plus** every message
  currently carrying a category or flag regardless of age.
- The five named **Inbox subfolder trees** the live Phase 1c sweep reads
  (Senior Management / Bi-monthly CDR/PD working group / H&S / Team / Projects).
- Per message: `Categories`, `FlagStatus`, `FlagRequest`, `UnRead`,
  `Importance`, parent folder path (a move = disappearance).
- **Sent Items** count + newest-5 identities; **Drafts** count + newest-5.
- **Default Calendar** + the **People Department - HR Systems** shared calendar,
  window `[today-1, today+8]`: per-event Start/End/Subject/Location hash + adds/
  removes.
- Subjects are **SHA1-hashed** in all snapshots and records — no raw email
  content is written or committed (`data/` is `.gitignore`d anyway).

### Teams gap (documented, accepted)

**Microsoft Teams is not covered** — it is not reachable through Outlook COM.
Under the Codex route this was a real residual gap (Teams `chat.ReadWrite` was
one of the connector's write tools). **Under the Claude Code route it is moot** —
there is no Teams tool for the model to call. No mitigation needed for the
current plan; if the route ever changes again, revisit.

### Stale-cache handling

A COM `Categories`/flag read within seconds of a Graph-side write returns a
false-unchanged value (reconfirmed 27 Aug). Before the AFTER snapshot the script
forces `Namespace.SyncObjects.Start()` on every sync object and sleeps
`--settle-seconds` (default 60). This only risks a false negative (missed write),
never a false trip.

### Sensitivity

`--sensitivity strict` (default) — any delta trips. `--sensitivity writes-only`
— ignores pure unread→read transitions (the one delta Kevin's own passive
email-reading on another device can cause mid-run that a tools-off model call
cannot); category/flag/move/send/draft/calendar changes always trip.

## CLI

```
python tools/codex_triage/mailbox_guard.py snapshot --out before.json --label before --no-settle
# ... run the thing ...
python tools/codex_triage/mailbox_guard.py snapshot --out after.json  --label after
python tools/codex_triage/mailbox_guard.py guard --before before.json --after after.json \
       --task "Work Inbox Codex Parallel" --sensitivity strict
python tools/codex_triage/mailbox_guard.py prove          # end-to-end proof-of-fire
python tools/codex_triage/mailbox_guard.py clear-flag     # after reviewing an incident
```

## Proof-of-fire — PASSED 2026-08-27 16:26–16:29 (admin machine)

`prove` mode: BEFORE snapshot → create a real dummy scheduled task → inject a
synthetic category via COM onto one disposable message → settled AFTER snapshot
→ diff → enforce → verify → remediate → clean up. Disposable message was
`mailbot@distrokid.com` (subject SHA1 `b03b9b731aa8`), category `''` before.

| # | Step | Result |
|---|---|---|
| 1 | select disposable message | OK — DistroKid, categories `''` |
| 2 | before snapshot | OK — 254 msgs, 74 cal events tracked |
| 3 | create real dummy scheduled task | OK — `Drew Guard Selftest Dummy` created |
| 4 | inject synthetic category via COM | OK — `'' -> Drew-guard-selftest-20260827-162646` |
| 5 | diff detects the synthetic write | OK — exactly 1 delta, `categories_changed` [critical], correct EntryID |
| 6 | enforce: disable task | OK — `schtasks /Change /DISABLE` returned SUCCESS |
| 7 | enforce: alert detail file written | OK |
| 8 | enforce: BurntToast alert | OK — `Show-TaskNotification.ps1` invoked rc=0 |
| 9 | enforce: incident record + sentinel | OK — `guard_incident_*.json` + `GUARD_TRIPPED.flag` written |
| 10 | verify dummy task state | OK — `schtasks /Query` shows `Status: Disabled` |
| 11 | remediate synthetic category (settled re-read) | OK — read back `''` after `SyncObjects` + 45s |
| 12 | `Restrict("[Categories] <> ''")` residue sweep | OK — 0 self-test residue |

Cleanup confirmed: dummy task deleted, `GUARD_TRIPPED.flag` cleared, config
baseline `~/.codex/config.toml` sha1 `b2a1a226…` untouched, mailbox clean.
Evidence JSON: `data/codex_runs/selftest_result_*.json` (local; `data/` is
gitignored).
