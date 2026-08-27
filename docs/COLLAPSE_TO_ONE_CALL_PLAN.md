# Collapse the 5 `claude_code` AI calls into ONE `claude -p` call — implementation spec

**Status:** Spec only, 2026-08-27 (Drew). **Not built.** Blocked on the auth prerequisite
below (needed to run the confirming end-to-end test that must gate this change). On branch
`claude/outlook-codecs-connector-upgrade-fe3dgf`.

This is the design for mitigation #2 in `docs/CLAUDE_CODE_BACKEND.md` ("Collapse the 5 calls
into 1"). Kevin approved building it *first*, before the direct cutover, because 6x/day
unmitigated tips the Pro weekly cap. Porting source: `tools/codex_triage/build_call2_brief.py`
on branch `drew/codex-phase2-ai-triage` (assembles all five phases into one prompt).

---

## HARD PREREQUISITE — not satisfied as of 27 Aug 2026

`fetch_inbox.py`'s `claude_code` helper (`_claude_code_once`) points `CLAUDE_CONFIG_DIR` at
`WI_CLAUDE_CONFIG_DIR` (primary) / `WI_CLAUDE_CONFIG_DIR_FALLBACK` (overflow). Those env vars
are set (`C:\WorkInboxAI\kevin`, `C:\WorkInboxAI\hope`) **but neither directory is
authenticated**:

- `C:\WorkInboxAI\kevin\` and `C:\WorkInboxAI\hope\` each contain only a 423-byte skeleton
  `.claude.json` (no `userID`-bearing login state that carries creds) and **no
  `.credentials.json`**.
- `claude -p` with `CLAUDE_CONFIG_DIR` set to either returns
  `{"is_error":true,"result":"Not logged in \u00b7 Please run /login"}`.
- `CLAUDE_CODE_OAUTH_TOKEN` is **not set** at User or Machine scope.

**Unblock (Kevin, interactive — Drew cannot):** for each account, in a shell:

```
set CLAUDE_CONFIG_DIR=C:\WorkInboxAI\kevin
claude            (then /login as kevin@lelitte.co.uk — completes the OAuth browser flow;
                   this writes C:\WorkInboxAI\kevin\.credentials.json)
```
then repeat with `CLAUDE_CONFIG_DIR=C:\WorkInboxAI\hope` and `/login` as hope@lelitte.co.uk.
Verify each with: `echo hi | claude -p --output-format json` under that `CLAUDE_CONFIG_DIR`
→ expect `"is_error":false`. (`claude setup-token` alone is not sufficient unless its minted
token is also exported as `CLAUDE_CODE_OAUTH_TOKEN` into the scheduled-task environment; the
`/login`-writes-`.credentials.json` path is what the current helper expects.)

Until both dirs return `is_error:false`, step 2 of the cutover (the confirming run, which
must force one hope@ failover) cannot be executed and **nothing about the collapse build
should be merged or cut over.**

---

## Why this needs a small reorder (and can't be a pure drop-in)

The `api` backend keeps 5 separate `client.messages.create()` calls; that path must stay
byte-identical. The 5 call sites are interleaved with deterministic processing and their
inputs become available at different points:

| Call | Site (branch line ~) | System prompt const | User payload built from | Earliest point all inputs exist |
|---|---|---|---|---|
| Phase 2 context | 963 | `SYSTEM` (929) | `inbox_for_api`, `sent`, `cal_today`, `cal_tomorrow` | ~928 |
| Phase 3.2 email summaries | 1213 | `EMAIL_SUMMARY_SYSTEM` (1168) | `emails_for_summary` ← `summary_candidates` (= `urgent`+`needs` cards) | after Phase 3 cards (~1113) |
| Phase 3.5 task triage | 2044 | `TRIAGE_SYSTEM` (2013) | `api_emails`/`email_candidates` (← `inbox`+`sent`+`categorise`), `task_summaries` (← `cc_content`) | after CC load (~1949) |
| Phase 3.7 task summaries | 2340 | `SUMMARY_SYSTEM` (2329) | `tasks_for_summary` ← `all_priorities` (= `priorities_*`) | after CC load (~1949) |
| Phase 3.8 calendar prep | 2539 | `CAL_SUM_SYSTEM` (2526) | `_cal_for_summary` ← `_all_day_candidates` (cal_*_items) + `_granola_context` | after Granola fetch (~2512) |

None of the five phases consumes another phase's **output** (confirmed against
`build_call2_brief.py`, which stages all five from one deterministic pass). They only need
deterministic inputs. So a single call is possible **if all five payloads are assembled at
one point** — which means hoisting the late deterministic prep up to just after Phase 3
card-building (~line 1113), for the `claude_code` path only.

---

## Design

### 1. New module-level state (near the existing `_ai_call_seq`, ~line 40)

```python
_CC_COMBINED = None          # dict with the 5 phase-slice keys, or None until assembled
_CC_COMBINED_USAGE = {}      # usage dict from the single combined claude -p call
_PHASE_KEY = {
    "context":       "context_phase",
    "email_summary": "email_summary_phase",
    "task_triage":   "task_triage_phase",
    "task_summary":  "task_summary_phase",
    "calendar_prep": "calendar_prep_phase",
}
```

### 2. Add a `_phase=None` kwarg to `_ai_create(...)`

`_ai_create` signature gains `_phase=None`. In the **api** branch it is ignored (behaviour
unchanged). In the **claude_code** branch:

```python
if AI_BACKEND == "claude_code":
    if _CC_COMBINED is None:
        raise RuntimeError("combined claude_code call not assembled before _ai_create; "
                           "call _cc_run_combined() first")
    slice_obj = _CC_COMBINED.get(_PHASE_KEY[_phase], {})
    return _AIText(json.dumps(slice_obj, ensure_ascii=True), _CC_COMBINED_USAGE)
```

Every downstream site already does `raw = resp.content[0].text.strip()` → strip ``` fences
→ `json.loads(raw)`. `json.dumps` of the slice round-trips through that untouched.

Add `_phase="context"` / `"email_summary"` / `"task_triage"` / `"task_summary"` /
`"calendar_prep"` to the 5 call sites (one kwarg each; no other change to those blocks).

### 3. `_cc_run_combined(...)` — the single call

New function. Assembles ONE prompt from the 5 `(system, payload)` pairs in the exact
structure of `build_call2_brief.py` (five `=== N. <phase> ===` sections, each with the
verbatim `fetch_inbox.py` system-prompt text and that phase's JSON payload, and a closing
"Return ONLY a single JSON object with exactly these five top-level keys:
context_phase, email_summary_phase, task_triage_phase, task_summary_phase,
calendar_prep_phase"). Then:

```python
obj = _claude_code_once("claude-haiku-4-5", COMBINED_SYSTEM, combined_user,
                        timeout_s=float(300.0), cfg_dir=CLAUDE_CFG_PRIMARY)
```

wrapped in the **same dual-account / timeout-failover loop** currently inside `_ai_create`
(factor that loop into a helper `_claude_code_call(system, user, timeout_s)` and reuse it
for both the combined call and any future single call). `max_tokens` is not passed to
`claude -p` (it self-manages); size the timeout for the whole batch (~300s; the 5-call
parallel run was ~451s wall but most of that was per-call cold-start overhead the single
call removes).

Parse: `text = obj["result"]` → strip ``` fences → `json.loads` → assert the 5 keys are
present (on any missing key, log and set that slice to `{}` so the corresponding phase
falls through to its existing `except`/skip path rather than crashing the run). Set
`_CC_COMBINED` and `_CC_COMBINED_USAGE`. Append ONE line to `ai_backend_usage.jsonl`
(`seq="combined"`, plus per-phase `output_chars` for eyeballing).

### 4. Hoist (claude_code path only) — insert right after Phase 3 `print("Phase 3 done ...")` (~line 1113)

```python
if AI_BACKEND == "claude_code":
    _cc_load_priorities()          # (a)
    _cc_build_cal_items_and_ids()  # (b)
    _cc_fetch_granola()            # (c)
    _cc_run_combined(              # (d)
        p2=(SYSTEM, USER),
        p32=(EMAIL_SUMMARY_SYSTEM, _cc_emails_for_summary_payload()),
        p35=(TRIAGE_SYSTEM, _cc_triage_payload()),
        p37=(SUMMARY_SYSTEM, _cc_tasks_for_summary_payload()),
        p38=(CAL_SUM_SYSTEM, _cc_cal_for_summary_payload()),
    )
```

To make (a)–(c) reusable without duplicating logic, extract each existing block into a
function and **call the function from both the early hoist point (guarded
`if AI_BACKEND == "claude_code"`) and the original location (guarded
`if AI_BACKEND != "claude_code"`, i.e. unchanged for the default api path)**:

- **(a) `_cc_load_priorities()`** — body = current lines ~1907–1949 (CC `tasks.json` GET +
  `priorities_today/tomorrow/week` + `cc_content`). Idempotent (guard on a
  `_cc_priorities_loaded` flag). Original site (~1900): `if AI_BACKEND != "claude_code": _cc_load_priorities()`. Everything downstream already just reads the module vars.
- **(b) `_cc_build_cal_items_and_ids()`** — body = current lines ~2362–2418
  (`build_cal_items` for the 4 columns + `_attach_cc_task_ids`). Needs `cc_content` (from a).
  `build_cal_items` def itself is already module-scoped earlier — verify; if not, hoist the
  `def` only. Original site guarded `!= "claude_code"`.
- **(c) `_cc_fetch_granola()`** — body = current lines ~2419–2512 (Phase 3.7b). Produces
  `_granola_context`. Network call; moving it earlier in the run is behaviourally
  transparent. Original site guarded `!= "claude_code"`.
- **System-prompt constants** `SYSTEM`, `EMAIL_SUMMARY_SYSTEM`, `TRIAGE_SYSTEM`,
  `SUMMARY_SYSTEM`, `CAL_SUM_SYSTEM` are currently assigned inside their phase blocks. Move
  each assignment to module scope (they reference nothing dynamic). Leaving the original
  in-block assignment is harmless (identical re-assign) but cleaner to remove.
- **Payload builders** `_cc_emails_for_summary_payload()` etc. — extract the small list
  comprehensions currently at lines ~1145–1167 (`emails_for_summary` + local `_age_days`),
  ~1976–2011 (`email_candidates`/`api_emails`/`task_summaries`), ~2320–2328
  (`tasks_for_summary`), ~2520–2523 (`_cal_for_summary`) into functions. Each phase block,
  in **both** backends, then calls the same function so the list the model is keyed against
  is provably identical to the list the downstream mapping code iterates. (Phase 3.2 keys
  responses `"0".."N-1"` by `enumerate(summary_candidates)`; Phase 3.5 keys by `email_n`
  index into `email_candidates`; Phase 3.7 by task `id`; Phase 3.8 by `"<day>_<idx>"`.
  Identical builder = identical keys.)

### 5. Phase 2 tail

Phase 2's `_ai_create` call at ~963 executes **before** Phase 3 cards, so in `claude_code`
mode it must not call out there. Guard the call: `if AI_BACKEND == "claude_code": pass`
(leave `context`/`subtitle` empty for now) `else: <existing try/except call+parse>`. Then
move the Phase 2 **tail** (the `same_briefing_date` preservation + `build_fallback_context`
/ `build_fallback_subtitle` block, ~lines 984–993) to run **after** `_cc_run_combined()` —
in api mode it stays where it is; in claude_code mode it runs post-combined, consuming
`_CC_COMBINED["context_phase"]` (`{"context": ..., "subtitle": ...}`) via the same parse.
Simplest: wrap that tail in a `def _finalise_context()` called from both places.

### 6. What stays exactly as-is

- The entire `api` backend path (all 5 `client.messages.create` calls, byte-identical).
- Every downstream parse/validate/`json.loads`/fence-strip block at the 5 sites.
- Phase 3.3/3.3b/3.3c demotion, Phase 3.5 suppression, Phase 3.6 CC sync, Phase 3.9
  scroll-out, Phase 4/5 push, `WI_AI_PARALLEL` guards.
- `_claude_code_once` (only refactor: pull the retry/failover loop out of `_ai_create` into
  a shared `_claude_code_call`).

---

## Test / acceptance (all require the auth prerequisite first)

1. `python -m py_compile fetch_inbox.py` clean.
2. `AI_BACKEND=api python fetch_inbox.py` (parallel or dry) — output identical to a pre-change
   api run (diff `data/briefing.json`). Proves the default path untouched.
3. `AI_BACKEND=claude_code WI_AI_PARALLEL=1 python fetch_inbox.py` — exactly **one**
   `claude -p` invocation in the log; `ai_backend_usage.jsonl` has one `seq="combined"`
   entry; all five phase outputs parsed; `data/claude_briefing.json` equivalent in shape to
   a recent api `data/briefing.json`; force one hope@ failover (temporarily point
   `WI_CLAUDE_CONFIG_DIR` at a broken dir) and confirm the loop switches accounts.
4. Record real per-run tokens (output + cache_read + cache_creation) from the combined
   entry; project to 6x/day x 5 weekdays vs the Pro weekly cap. Expected ~2.8–3.3M tok/week
   (down from ~4.3M at 5 calls). If still tight, report and recommend 3x/day — do not change
   cadence without Kevin.

## Rollback

`AI_BACKEND` stays `api` by default and the scheduled `Run Inbox Briefing.bat` does not set
it, so **merging this to `main` is inert** until the `.bat` is edited to
`set "AI_BACKEND=claude_code"` + `set "ANTHROPIC_API_KEY="`. One-line rollback after cutover:
remove those two `set` lines from the `.bat` (revert to metered api) — plus, if needed,
`git revert` the merge commit. Pre-change `Archive/` backup of `fetch_inbox.py` is a
mandatory first step of the build.
