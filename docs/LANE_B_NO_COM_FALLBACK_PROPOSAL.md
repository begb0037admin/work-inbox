# Lane B calendar — removing COM as the halt/uncertainty fallback (PROPOSAL, not implemented)

**Status:** 2 Sept 2026 (Drew). Design proposal only. Nothing in this doc has been built —
distinct from `lane_b_cal_guard.py`/`lane_b_call1.py` themselves, which DID change today (commits
`d6a74de`, `a4582a0`, `6a681bd`) for a separate, related decision — see the update note below.
`fetch_inbox.py` and the wrapper are untouched by either.

**Standing decision this responds to (Kevin, verbatim):** *"If we find we can't trust the
connector read, then we find a way to trust it. I'm not going to rule Connector out of this. We
are going to use Connector one way or another... we are using Connector."* COM is no longer an
acceptable **permanent** fallback destination for a Lane B halt/uncertainty. This is architecture,
not up for re-litigation here.

---

## UPDATE (2 Sept 2026, later same day) — the snapshot-diff layer this doc's §2-3 was about is GONE

Kevin made a separate, further decision today: drop the PRE/POST snapshot-diff layer from the live
`--run` gate entirely, rather than keep chasing its reliability (it was the direct cause of the 1
Sept 52-false-positive bug, the all-day-event gap, and a "second connector call hangs" 360s-timeout
pattern that blocked every `--dry-diff` attempt this week — three separate incidents, none of them
ever a real write). `lane_b_call1.py`'s re-contamination guard — which inspects the ACTUAL tool
calls made during the fetch, not an inferred before/after diff — is now the SOLE live safety
mechanism. Assessed first whether that guard alone is sufficient (it is, with one real gap found
and closed the same day: partial/timed-out output is now scanned for tool calls too, via
`ReContaminationDetected`/`_scan_partial_output_for_writes`, so a write completing just before a
hang is no longer invisible). Full writeup: HANDOVER.md's 2 Sept "drop the snapshot-diff layer"
entry.

**Consequence for this doc:** §2-3 below (splitting exit-1 into "1a re-contamination" vs "1b diff
mismatch", and the POST2 confirmation-snapshot escalation ladder for 1b) is now **MOOT** — there is
no more 1b. Every HALT from `--run` is now unambiguously a re-contamination event: unlike a diff
mismatch, it has no ambiguity to resolve by reading again, so there is nothing left to escalate.
Left §2-3 in place below, struck through in spirit, for the record of what was considered and why
it's no longer needed — skip to §3-revised for the current design. §4 (no-COM destination) and §5-6
are updated to match.

---

## 1. The current behaviour (what's changing)

`lane_b_cal_guard.py --run` returns one of three codes; the wrapper (`Run Laptop Bridge
Briefing.ps1`) currently reacts to all three by falling back to `CAL_BACKEND=com` for that cycle —
**this part is UNCHANGED by today's snapshot-diff removal, still the thing this doc proposes to fix:**

| Exit | Meaning (as of the `6a681bd` redesign) | Wrapper reaction today |
|---|---|---|
| 0 | Clean — Call-1's re-contamination guard saw nothing unexpected | proceed with `CAL_BACKEND=connector` |
| 1 | Persistent HALT — the re-contamination guard actually observed a write/off-allowlist tool call | disable the scheduled task, toast, **fall back to `CAL_BACKEND=com`** |
| 3 | Transient — connector didn't return `list_events` across all retries, can't verify | task stays enabled, **fall back to `CAL_BACKEND=com`** |

The problem Kevin is pointing at, unchanged: exit 1 and exit 3 both silently swap in COM — a
completely different pull mechanism the guard has no visibility into and no ability to
safety-check — as the *permanent* answer to "the connector wasn't verified this cycle." That's the
"silent safety net." This doc is still about fixing exactly that.

## 2. ~~First split the two things "exit 1" currently conflates~~ — MOOT, see UPDATE above

*(Original §2 content preserved for the record; skip to §3-revised.)*

There is no longer a "1a vs 1b" distinction to make. Exit 1 from `--run` now means exactly one
thing: the re-contamination guard actually observed a write/off-allowlist tool call. That is
always a safety trip, always a hard stop, never something more connector reads would help resolve.

## 3-revised. HALT (exit 1) is unambiguous now — no escalation ladder needed

Today (post-`6a681bd`): Call-1 runs, its own re-contamination guard inspects every tool call made
across every retry attempt (including partial output from a killed/timed-out attempt). If it sees
anything outside the read-only allowlist for the two Lane B connector namespaces, Call-1 exits 1,
`cmd_run()` propagates that as exit 1, quarantines the normalised file, and returns.

**There is nothing to escalate or retry here** — unlike a diff mismatch (which could have been a
flaky read, and a second read could tell you), an observed write-tool call is a direct, first-hand
observation. Retrying after one increases exposure, it doesn't resolve ambiguity, because there
isn't any. This case goes straight to §4 below.

*(The original §3 "POST2 confirmation snapshot" escalation ladder, and its "known implementation
gap" about `take_snapshot()` only returning fingerprints not full event data, no longer apply —
there is no second read to confirm, and `take_snapshot()`/`diff_snapshots()` are diagnostic-only
now, not part of any live decision.)*

## 4. When it's genuinely unverified — no COM, what instead

Applies to: exit 1 (re-contamination — always, unconditionally now) and exit 3
(transient/connector-unavailable, unchanged from before today's redesign).

- **Serve no calendar that cycle, honestly.** Reuse the pattern already proven and live for the
  laptop bridge (`WI_BRIDGE_ALLOW_EMPTY_CALENDAR`, used when there is genuinely no calendar
  source): empty calendar section + an explicit "Lane B calendar unavailable/unverified this
  cycle" warning in the briefing. This is materially different from a COM fallback — it's honest
  about the failure instead of silently substituting a different-trust-model data source the guard
  never checked.
- **Backoff / escalate the retry, not the fallback** — this is about exit-3 (genuine connector
  unavailability), NOT exit-1 (never retry a re-contamination). Two mechanisms worth considering
  (pick one, or start with the cheaper one and add the other if it's not enough):
  1. Increase `CALL1_RETRIES`/timeout for the *next* attempt after a consecutive failure (cheap, no
     new scheduled infrastructure), or
  2. A short-interval out-of-band retry (e.g. 15–30 min later, calendar-pull only, not a full
     briefing re-run) rather than waiting for the next full 07/12/16 cadence slot.
- **Track and surface a consecutive-failure counter for exit-3 specifically.** A single unverified
  cycle is unremarkable (already true today, exit 3 doesn't disable the task). A *streak* is a real
  signal something's wrong with the connector/account and deserves the same visibility the HALT
  toast now gets (built earlier 2 Sept) — e.g. after 2+ consecutive exit-3 cycles, fire the same
  cross-machine toast channel with a distinct message ("Lane B calendar unverified for N
  consecutive cycles") even though this isn't a safety trip.
- **Never disable the scheduled task for pure unavailability** (exit 3) — only exit 1
  (re-contamination) disables the task, unchanged from before today.

## 5. What does NOT change

- The re-contamination guard's strictness — if anything today's redesign argues for keeping it
  exactly as strict (and it's now been hardened further — partial-output scanning), since it's the
  ONLY live safety mechanism now, not one of two.
- `SAFETY_RULE` prompt hardening, the toast build, the `approval_policy` structural finding — all
  already shipped on `#31`, unrelated to this proposal, still ready alongside it.
- Nothing here touches the live scheduled task or `CAL_BACKEND` on anything running. Calendar
  stays on COM as the *actual live pull mechanism* until Kevin gives an explicit cutover go-ahead
  — this doc only changes what happens *within the connector path itself* once that cutover has
  happened and Lane B is live.

## 6. Suggested build order (if/when Kevin says go) — simplified, no escalation-ladder work needed

1. Replace the wrapper's `$CalBackend = 'com'` fallback lines (both the exit-1 and exit-3 cases)
   with the "serve empty + warn" path, reusing `WI_BRIDGE_ALLOW_EMPTY_CALENDAR`.
2. Add the exit-3 consecutive-failure counter + its toast case (extends the `lane_b_guard` status
   field already wired earlier 2 Sept with a new value, e.g. `unverified-streak-N`).
3. Test coverage for the new wrapper behaviour before any live exercise, same discipline as the
   existing guard fixes — no `--selftest`/`--dry-diff` work needed here since the guard side of
   this is already done and tested (`lane_b_call1.py --selftest`, `lane_b_cal_guard.py --selftest`,
   both green as of `6a681bd`).

Smaller build than the original §6 — dropping the snapshot-diff layer also dropped the escalation
ladder this doc originally needed to design around it. What's left is just: stop falling back to
COM, serve empty+warn instead, and make a genuine unavailability streak visible.
