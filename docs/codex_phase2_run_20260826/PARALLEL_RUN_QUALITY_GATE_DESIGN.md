# Parallel-run quality-gate design — false-demotion + missing-importance (26 Aug 2026, Drew)

Design only, not built. Addresses the two quality findings from the 26 Aug dry run
(`DIFF_REPORT.txt`). During the eventual 7-day parallel run, Codex output goes only
to `codex_*` files and never feeds the dashboard — so for the parallel run itself
the "gate" is **measurement discipline**, not runtime blocking. The runtime-blocking
rules below are for any *future real use*, not the validation window.

## A. False-demotion (Codex marked the Simon Burford / Data Warehouse / REF thread `no_action_needed: true` where the real pipeline said `false`)

### During the parallel run — measure, don't act
Every run emits `data/codex_runs/<ts>_codex_disagreements.json`:
```
[{ "subject": "...", "matched": true,
   "real":  {"tier": "needs|urgent|fyi", "needs_reply": bool, "no_action_needed": bool},
   "codex": {"needs_reply": bool, "no_action_needed": bool, "demotion_reason": "..."},
   "agree_needs_reply": bool, "agree_no_action": bool,
   "direction": "match" | "codex_hides_work" | "codex_over_escalates" }]
```
- `codex_hides_work` = Codex said `no_action_needed: true` on an email the real pipeline kept in Needs/Urgent. **This is the disqualifying metric.** Track the running count across all 42 runs in a `data/codex_runs/_rollup.json`. Any single `codex_hides_work` case on a thread that (a) also appears by subject in that run's real context paragraph, or (b) contains an escalation marker (see list below), or (c) is from a VIP sender — is on its own sufficient to fail auto-cutover, not something to average away. The REF thread today would have tripped (a) and (b).
- `codex_over_escalates` is low-risk (extra noise, not hidden work) — track it as a nuisance metric only.
- Require Codex to return a one-line `demotion_reason` for every `no_action_needed: true` (it already did in the 25 Aug `codex_triage_ledger.json` format) so every disagreement is auditable.

### Runtime guardrail spec (for future real use, not the parallel run)
A Codex `no_action_needed: true` is **ignored** (item stays in its tier) if any of:
- subject or body contains an escalation marker: `at risk`, `no response`, `not had a response`, `chasing`, `overdue`, `deadline`, `by (close of business|COB|end of)`, `putting .* at risk`, `escalat`, or a named critical programme token (`REF`, `HESA`, `TRAC`, plus any Kevin adds);
- the same normalised subject appears in this run's context paragraph;
- sender is on the VIP list (reuse work-inbox's existing VIP-sweep sender list);
- `_ai_verdict_valid` equivalent is not cleanly true (mirror `fetch_inbox.py`'s existing rule that a non-schema-valid verdict never drives a demotion);
- the email is < 3 days old (staleness only ever *downgrades* urgency in the real pipeline — never let a fresh item be demoted to nothing on one model's say-so).
This mirrors the real pipeline's own conservatism (`STALENESS_CUTOFF_DAYS` only flips `needs_reply` true→false, `_ai_verdict_valid` gating, VIP sweep) rather than inventing a new posture.

## B. Missing importance field (Codex found 0 urgent vs the real pipeline's 3)

### Step 1 — try to actually get the field
The connector's `list_messages` did not return `importance` this run (`"Importance was not supplied by the Outlook connector"`). Before treating it as unavailable, the parallel-run build must test:
- explicitly `$select`-ing / requesting `importance` in the pull prompt;
- a per-id `fetch_message` full-detail call (which *does* return richer fields — it's how the write-gate verification read `categories`) for the pulled message set, rather than relying on the `list_messages` projection.
The Graph `message` resource has `importance` natively, so this is likely a projection/verbosity issue, not a hard gap.

### Step 2 — if genuinely unavailable
- The diff harness splits "urgent misses" by cause: keyword-driven (`URGENT_SUBJECTS` matched — Codex's ported `categorise()` would catch these too; a miss there is just the shallower 40-item pull, fixable by pulling deeper) vs importance-flag-driven (the real gap).
- Every parallel-run report states plainly: "Codex-graph cannot replicate importance-flag-driven Urgent classification" and **excludes the Urgent tier from the fidelity score** rather than scoring Codex down for a signal it structurally cannot see.
- Design note for any real cutover (not the parallel run): keep a **thin Outlook-COM shim** in the pipeline whose *only* job is to read `importance`/high-flag for the pulled message set, joined to Codex's pull on `normalised_subject + received_datetime_to_the_minute` (no COM EntryID dependency, so no cross-format ID problem). Codex does all content/judgement; COM supplies the one binary signal it can't. This keeps COM in the loop for exactly one field instead of the whole pull — a much smaller ongoing COM dependency than today's.

## C. What "parallel validation passed" would need to mean (so cutover isn't a vibe)
Over the 42 runs:
- `codex_hides_work` count on material threads (per the A guardrail triggers): **must be 0.** Any hit = no auto-cutover, human review of that case.
- `needs_reply` agreement on subject-matched overlap: track the rate; set the bar with Kevin (suggest >= 95% before it's even a conversation).
- New-task-suggestion precision: for each Codex `new_task` not matched by the real pipeline, a human (Lauren/Kevin) marks it genuine-or-noise at least weekly; Codex's task_updates volume (10 vs real 3 today) is checked for over-matching to existing tasks.
- Context-paragraph spot-check: Lauren/Kevin reads Codex's vs real's context paragraph a few times a week and records "would I have been equally well-briefed."
- Importance/Urgent: either resolved per B Step 1, or the shim per B Step 2 is built and verified, before Urgent-tier parity is claimed.
