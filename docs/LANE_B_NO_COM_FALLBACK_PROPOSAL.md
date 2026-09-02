# Lane B calendar — removing COM as the halt/uncertainty fallback (PROPOSAL, not implemented)

**Status:** 2 Sept 2026 (Drew). Design proposal only, per Kevin's explicit instruction to propose
before implementing. Nothing in this doc has been built. No code in `lane_b_cal_guard.py`,
`lane_b_call1.py`, `fetch_inbox.py`, or the wrapper has changed as a result of this doc.

**Standing decision this responds to (Kevin, verbatim):** *"If we find we can't trust the
connector read, then we find a way to trust it. I'm not going to rule Connector out of this. We
are going to use Connector one way or another... we are using Connector."* COM is no longer an
acceptable **permanent** fallback destination for a Lane B halt/uncertainty. This is architecture,
not up for re-litigation here.

---

## 1. The current behaviour (what's changing)

`lane_b_cal_guard.py --run` returns one of three codes; the wrapper (`Run Laptop Bridge
Briefing.ps1`) currently reacts to all three by falling back to `CAL_BACKEND=com` for that cycle:

| Exit | Meaning today | Wrapper reaction today |
|---|---|---|
| 0 | Clean — PRE and POST snapshots agree | proceed with `CAL_BACKEND=connector` |
| 1 | Persistent HALT — either a write/off-allowlist tool was actually observed (re-contamination), or the PRE/POST diff found a real change | disable the scheduled task, toast, **fall back to `CAL_BACKEND=com`** |
| 3 | Transient — connector didn't return `list_events` across all retries, can't verify | task stays enabled, **fall back to `CAL_BACKEND=com`** |

The problem Kevin is pointing at: exit 1 and exit 3 both silently swap in COM — a completely
different pull mechanism the guard has no visibility into and no ability to safety-check — as the
*permanent* answer to "the connector wasn't verified this cycle." That's the "silent safety net."

## 2. First split the two things "exit 1" currently conflates

They need different treatment, and treating them the same is itself part of the problem:

- **1a — re-contamination observed** (a write-verb or off-allowlist tool actually fired, caught by
  `lane_b_call1.py`'s `guard_recontamination()` on any attempt, PRE/Call-1/POST). This is a
  **safety** trip. More connector calls after this is the wrong direction — it increases exposure,
  it doesn't resolve uncertainty. This should stay a hard stop, unchanged in spirit, **only the
  destination for "what do we show this cycle" changes** (see §4) — never escalate/retry this one.
- **1b — snapshot diff found a change, no write tool ever observed**. This is an **uncertainty**
  trip: either (i) the connector gave a flaky/incomplete read on one side (the actual 1 Sept
  52-false-positive bug and, plausibly, yesterday's 8-diff all-day-event gap), or (ii) the calendar
  genuinely changed during the pull window (a real edit by Kevin or a colleague, nothing wrong).
  **This is exactly the case "retry/escalate until trustworthy" is suited to** — more connector
  reads can actually discriminate between (i) and (ii); more connector reads cannot help 1a.

Proposal: `cmd_run()` reports a distinct outcome for 1a vs 1b (e.g. keep exit 1 for 1a, use a new
exit 4 for 1b before escalation resolves it — see §3) so the wrapper (and any future caller) can
tell them apart instead of treating every HALT identically.

## 3. Escalation ladder for 1b (diff-with-no-write-tool-observed)

Today: PRE → Call-1 → POST → diff trips → immediate HALT.

Proposed: on a trip, **before** halting, take a third confirmation snapshot and use the
re-contamination guard's already-independent verdict as a second, orthogonal signal:

1. PRE → Call-1 → POST → diff trips, AND `guard_recontamination()` never flagged a write tool on
   any of PRE/Call-1/POST (already checked per-attempt today, confirmed by code read this
   session — see HANDOVER.md 2 Sept entry §6/audit).
2. Take **POST2** (one more `take_snapshot()` call, reusing existing retry-aware plumbing —
   cheap, ~1 extra codex call in the rare case this triggers at all).
3. Compare POST2 against PRE and POST:
   - **POST2 == PRE (not POST)** → the original POST was the flake. Treat PRE/POST2 as
     authoritative, proceed with `CAL_BACKEND=connector` this cycle, log the blip
     (`data/codex_runs/` — a visible, audited "transient diff self-resolved" record, not a silent
     pass), task stays enabled, no toast (nothing wrong happened).
   - **POST2 == POST (not PRE), reproducible across two independent reads** → the change is very
     likely genuine (a real exogenous edit), *not* a flaky read. Combined with re-contamination
     never having flagged a write tool across PRE/Call-1/POST/POST2, this is strong evidence
     (two independent signals, not one) that nothing self-inflicted happened. **Accept it**: proceed
     with `CAL_BACKEND=connector` using POST2's data, log clearly as "confirmed reproducible
     calendar change during pull window, no write tool observed at any point, accepted" — visible
     in the run log, not silent — task stays enabled, no toast (this is normal operation, a
     calendar changing during a 7-day read window is expected).
   - **Neither matches** (three reads disagree with each other) → genuine instability in what the
     connector is returning. This is the "can't get a clean read" case → §4.

**Known implementation gap, flagged not solved:** `take_snapshot()` today only returns a
*fingerprint* (subject/start/end/response_status/all_day), not the full raw event objects
`normalise_pull` needs to actually populate the briefing. The escalation path needs POST2's full
data usable downstream, not just its fingerprint for comparison — either extend `take_snapshot()`
to optionally persist the raw objects too, or have the escalation step re-run
`lane_b_call1.py --domain calendar` as the "confirmation fetch" (it already produces full,
correctly-shaped normalised data) instead of a bare snapshot. Worth resolving at implementation
time, not a blocker to reviewing the design.

## 4. When it's genuinely unverified — no COM, what instead

Applies to: 1a (re-contamination, always), 1b after escalation still disagrees, and exit-3
(transient/connector-unavailable, unchanged from today except for the destination).

- **Serve no calendar that cycle, honestly.** Reuse the pattern already proven and live for the
  laptop bridge (`WI_BRIDGE_ALLOW_EMPTY_CALENDAR`, used when there is genuinely no calendar
  source): empty calendar section + an explicit "Lane B calendar unavailable/unverified this
  cycle" warning in the briefing. This is materially different from a COM fallback — it's honest
  about the failure instead of silently substituting a different-trust-model data source the guard
  never checked.
- **Backoff / escalate the retry, not the fallback.** Two mechanisms worth considering (pick one,
  or start with the cheaper one and add the other if it's not enough):
  1. Increase `SNAP_RETRIES`/`CALL1_RETRIES`/timeout for the *next* attempt after a consecutive
     failure (cheap, no new scheduled infrastructure), or
  2. A short-interval out-of-band retry (e.g. 15–30 min later, calendar-pull only, not a full
     briefing re-run) rather than waiting for the next full 07/12/16 cadence slot.
- **Track and surface a consecutive-failure counter.** A single unverified cycle is unremarkable
  (already true today, exit 3 doesn't disable the task). A *streak* is a real signal something's
  wrong with the connector/account and deserves the same visibility the HALT toast now gets
  (built this session, see the ~09:2x/2 Sept entries) — e.g. after 2+ consecutive
  unverified/unresolved cycles, fire the same cross-machine toast channel with a distinct message
  ("Lane B calendar unverified for N consecutive cycles") even though this isn't a safety trip.
- **Never disable the scheduled task for pure unavailability** (exit 3 today, and 1b-still-unresolved
  after escalation) — only 1a (re-contamination) disables the task, unchanged from today.

## 5. What does NOT change

- The re-contamination guard's strictness (1a) — if anything this proposal argues for *keeping* it
  exactly as strict, since it's now the thing the whole 1b-escalation logic leans on as an
  independent signal.
- `SAFETY_RULE` prompt hardening, the toast build, the `approval_policy` structural finding — all
  already shipped on `#31` this session, unrelated to this proposal, still ready alongside it.
- Nothing here touches the live scheduled task or `CAL_BACKEND` on anything running. Calendar
  stays on COM as the *actual live pull mechanism* until Kevin gives an explicit cutover go-ahead
  — this doc only changes what happens *within the connector path itself* once that cutover has
  happened and Lane B is live.

## 6. Suggested build order (if/when Kevin says go)

1. Split exit 1 into 1a/1b in `lane_b_cal_guard.py` (small, low-risk — just better labeling of an
   outcome already computed).
2. Build the POST2 escalation ladder for 1b, resolving the raw-data gap in §3.
3. Replace the wrapper's `$CalBackend = 'com'` fallback lines (both the exit-1 and exit-3 cases)
   with the "serve empty + warn" path, reusing `WI_BRIDGE_ALLOW_EMPTY_CALENDAR`.
4. Add the consecutive-failure counter + its toast case (extends the `lane_b_guard` status field
   already wired this session with a new value, e.g. `unverified-streak-N`).
5. `--selftest`/`--dry-diff` coverage for the new escalation logic before any live exercise, same
   discipline as the existing guard fixes.

Each of these is independently reviewable and small — none of it needs to land in one PR, and none
of it should land without its own `--selftest` proof, same standard as the rest of `#31`.
