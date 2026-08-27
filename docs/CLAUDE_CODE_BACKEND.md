# Headless Claude Code backend for `fetch_inbox.py`

**Status: CUT OVER 2026-08-27 (Drew).** Live on `main` (PR #29, merge `5423c83`).
The `\Work Inbox Briefing` scheduled task now runs the AI triage through **ONE
combined `claude -p` call** on Kevin's Claude subscription (kevin@ primary → hope@
overflow), not the metered Anthropic API. `AI_BACKEND` still **defaults to `api`**
in the code — the cutover switch is two `set` lines in `Run Inbox Briefing.bat`.
Supersedes the Codex route entirely.

**ONE-LINE ROLLBACK to the metered API:** delete the two `set` lines
(`AI_BACKEND=claude_code`, `ANTHROPIC_API_KEY=`) from `Run Inbox Briefing.bat`
(Kevin's Desktop). `fetch_inbox.py` then defaults to `AI_BACKEND=api` and uses the
still-present `ANTHROPIC_API_KEY` user env var. Pre-cutover `main` commit:
`c79d7c73956789087fad46f7bbaa132593bbb14c`. Pre-collapse `fetch_inbox.py` backup:
`Archive/fetch_inbox_backup_20260827_1746_pre_collapse_to_one_call.py`
(pre-backend backup: `..._20260827_1640_pre_claudecode_backend.py`).

## Cutover record (2026-08-27)

| Item | Detail |
|---|---|
| Auth | `C:\WorkInboxAI\kevin` + `C:\WorkInboxAI\hope` each `claude /login`-ed (real `.credentials.json`, `claude -p` → `is_error:false`, distinct accounts kevin@/hope@, both `pro`). User env vars `WI_CLAUDE_CONFIG_DIR` / `WI_CLAUDE_CONFIG_DIR_FALLBACK` point at them — **must stay logged in**. |
| Collapse 5→1 | `_cc_run_combined()` assembles all five phases (verbatim system prompts hoisted to `_SYS_*`, payloads built by the same deterministic logic each phase block rebuilds) into one prompt; `_ai_create(_phase=...)` returns that phase's slice of `_CC_COMBINED` so every downstream fence-strip + `json.loads` + validation runs unchanged. Fired once, right after Phase 3 cards (earliest point all payloads exist — needs CC-load + Granola + cal-candidates hoisted, done for the `claude_code` path only). |
| Confirming run | `WI_AI_PARALLEL=1`, primary → bogus dir to force failover. ONE call; hope@ failover proven; 5 slices parsed (`missing_keys=none`); `claude_briefing.json` schema-identical to api `briefing.json`; calendar summaries on the correct meetings. |
| Cutover run | Real scheduled-task trigger. kevin@ primary, call wall 255s. Briefing pushed (`a544d8a`); CC sync applied 7 updates (`command-centre` `099c6f11`); `triage_ledger.json` written (`40e5f121`); no mailbox effects. Task `ExecutionTimeLimit` raised `PT15M → PT20M`. |
| Usage | ~80k tokens/run (out ~23–28k incl. Haiku thinking; cache_creation ~55k; cache_read 0 cold) vs ~142k for the old 5-call path — **~44% less**. ~2.0M tok/week at 5×/day×weekdays (was ~3.55M). Shares Kevin's Pro pool. If still tight after a week: 3×/day (~1.2M/wk), not without Kevin. |

## What changed in `fetch_inbox.py` (original backend build — still accurate for the `api` path)

Original backend backup: `Archive/fetch_inbox_backup_20260827_1640_pre_claudecode_backend.py`.
The later 5→1 collapse (27 Aug) added `_cc_run_combined()` / `_claude_code_call()` /
`_cc_load_priorities()` / `_cc_build_cal_candidates_early()` / `_cc_fetch_granola()` /
`_p2_finalise()` and the `_SYS_*` hoisted prompts; `_ai_create()` gained `_phase=`.
See the cutover record above and `docs/COLLAPSE_TO_ONE_CALL_PLAN.md`.

| Change | Detail |
|---|---|
| Config block (after the `GITHUB_*` consts) | `AI_BACKEND` (`api` default / `claude_code`), `WI_AI_PARALLEL`, `PUSH_ENABLED`, `WI_CLAUDE_BIN`, `WI_CLAUDE_CONFIG_DIR`, `WI_CLAUDE_CONFIG_DIR_FALLBACK`, `_AI_CALL_LOG`. |
| `_ai_create(...)` helper | Drop-in for `client.messages.create()`. `api` → the real anthropic client, **byte-identical behaviour**. `claude_code` → `_claude_code_once()` + a 2-try / dual-account loop. Returns an `_AIText` shim exposing `.content[0].text` and `.usage`. |
| `_claude_code_once(...)` | `claude -p --model claude-haiku-4-5 --system-prompt <verbatim> --exclude-dynamic-system-prompt-sections --disallowedTools "Bash,Edit,Write,Read,Glob,Grep,WebFetch,WebSearch,NotebookEdit,Task,TodoWrite,SlashCommand" --strict-mcp-config --mcp-config '{"mcpServers":{}}' --permission-mode default --no-session-persistence --output-format json`, user prompt on stdin. Env: `ANTHROPIC_API_KEY` **stripped** (forces subscription billing) plus all `CLAUDE_CODE*` / `CLAUDECODE` / `CLAUDE_PID` / `CLAUDE_EFFORT` / `AI_AGENT` stripped (so the nested `claude -p` starts a clean independent session). `CLAUDE_CONFIG_DIR` set per account. |
| `client = ... if AI_BACKEND == "api" else None` | The anthropic client is only constructed for the `api` backend (it would raise with `ANTHROPIC_API_KEY` unset). |
| 5 call sites → `_ai_create(` | Phase 2 (ctx), 3.2 (email summaries), 3.5 (task triage), 3.7 (task summaries), 3.8 (calendar prep). Kwargs unchanged; the surrounding JSON-fence-strip + `json.loads` + validation logic **untouched**. |
| Parallel-validation guards | `WI_AI_PARALLEL=1` ⇒ `PUSH_ENABLED=False`, Phase 3.6 CC-sync skipped, Phase 3.9 forced dry-run, Phase 4 writes `data/claude_briefing.json` locally (no GitHub PUT), Phase 5 writes `data/claude_inbox_suggestions.json` locally. **No shared file, ledger, or Command Centre state is touched.** |
| `ai_backend_usage.jsonl` | Every `claude_code` call appends `{ts, seq, account, try, model, wall_s, usage, cost_usd, num_turns}` — the usage-measurement surface (gitignored). |

Default path (`AI_BACKEND` unset, `WI_AI_PARALLEL` unset): `PUSH_ENABLED == bool(GITHUB_PAT)`,
`_ai_create` calls `client.messages.create(**kw)` exactly as before. `python -m py_compile` clean.
`~/.codex/config.toml` sha1 `b2a1a226…` untouched (Codex not involved).

## Verified against the admin machine (27 Aug 2026)

- **Headless subscription auth: works.** `claude 2.1.247`. With `ANTHROPIC_API_KEY`
  unset, `claude -p` authed off `~/.claude/.credentials.json` (`claudeAiOauth`)
  and returned `is_error:false / subtype:success`. Active account:
  **`subscriptionType: pro`**, `rateLimitTier: default_claude_ai`. **Pro, not Max.**
- **Haiku 4.5 headless: works.** Response `canonicalModel: claude-haiku-4-5`
  (`claude-haiku-4-5-20251001`), `provider: firstParty`.
- **No write path: confirmed.** `permission_denials: []`, zero tools loaded, zero
  MCP servers (the machine's one global `github` MCP server not loaded under
  `--strict-mcp-config`). `claude -p` has no Outlook/Exchange/Graph tool at all;
  `fetch_inbox.py`'s COM stays its own Python. **The COM delta-sweep kill-switch
  (`mailbox_guard.py`) is NOT required for this route** — kept only as optional
  belt-and-braces (`OPTION1_KILLSWITCH.md`).
- **ToS:** headless/scripted use is a documented first-class feature
  (`-p` / `setup-token` / GH Action). Personal self-consumed automation at
  3–6×/day on one account is within terms. hope@ overflow is a **Kevin-confirmed
  permanent standing arrangement** (not cap-evasion), so the failover stays.

## Full parallel run — real numbers (1 run, 5 calls, 27 Aug ~16:52)

| Call | Phase | wall | output tok | cache_read | cache_creation | list-$ |
|---|---|---|---|---|---|---|
| 1 | 2 context | 117s | 11,590 | 23,625 | 0 | 0.074 |
| 2 | 3.2 summaries | 63s | 6,092 | 5,936 | 7,060 | 0.048 |
| 3 | 3.5 triage | 142s | 13,498 | 5,936 | 18,574 | 0.119 |
| 4 | 3.7 task summaries | 106s | 8,331 | 5,936 | 18,176 | 0.093 |
| 5 | 3.8 calendar prep | 23s | 1,774 | 5,936 | 9,675 | 0.034 |
| **run** | | **451s (~7.5 min)** | **41,285** | **47,369** | **53,485** | **$0.368** |

Notes:
- `list-$` is Claude Code's `total_cost_usd` = *equivalent metered-API cost*, **not**
  a real subscription charge (subscription is flat-fee). It is only useful as a
  rough size signal. The saving vs the current ~£36/mo API bill is still real.
- **Output tokens are thinking-inflated.** Haiku via `claude -p` appears to use
  extended thinking by default (the probe showed `thinking_tokens`); the current
  API pipeline does not. An 11.6k-token "context paragraph" call is that.
- **First run after a cold gap** stalled: two 150s timeouts on call #1 at 16:44
  (Pro plan rate-limit back-off under load, not a crash). The retry loop now
  treats a timeout as a usage-limit signal → switches to the fallback account;
  base per-call budget raised to `(timeout or 90) + 150`s. The 16:52 re-run then
  completed clean. **A warm-up call is NOT needed** (a single `claude -p` cold is
  ~2–20s); the stall was contention, which the account failover is the answer to.

## Is 6×/day viable on Pro?

**Not on the Pro plan alone, unmitigated.** Conservative load ≈ output +
cache_read + cache_creation ≈ **~142k tokens/run**:

| Cadence | ~tokens/week (weekdays) | Verdict on Pro (shared with all Kevin's agent work, already near-limit) |
|---|---|---|
| 6×/day | ~4.3M | Very likely to tip the weekly cap. |
| 3×/day | ~2.1M | Survivable but still competes with interactive/agent use. |

**Mitigations that make it fit (recommended combination in bold):**
1. **3×/day cadence** (e.g. 07:00 / 11:00 / 15:00) — briefing is a morning/interim
   artefact, not real-time. Halves the load.
2. **Collapse the 5 calls into 1** — the old Codex "Call 2" combined-brief design
   (`tools/codex_triage/build_call2_brief.py` on `drew/codex-phase2-ai-triage`
   already assembles all five phases into one prompt). Removes 4× the
   cache-creation + per-call harness overhead — the single biggest reduction
   (~60%+). Larger build; flagged, not done here.
3. **Suppress extended thinking** if `claude -p` exposes a thinking-budget flag —
   would cut the ~41k output ~5×. Not investigated (scope).
4. **hope@ failover** (permanent, confirmed) — absorbs spikes and any single-account
   weekly-cap hit mid-day.

**Cleanest route to a safe cutover: 3×/day + collapsed single call + hope@ failover.**
If Kevin wants to keep 6×/day, move the automation to a Max plan or a dedicated
account.

## Dual-account failover

Built and in the helper already. `WI_CLAUDE_CONFIG_DIR` = primary (kevin@),
`WI_CLAUDE_CONFIG_DIR_FALLBACK` = overflow (hope@). On a usage-limit error **or a
timeout stall** on the primary, the call retries once on the fallback account.
Each is just a `CLAUDE_CONFIG_DIR` pointing at a dir that has been `claude
setup-token`/`login`-ed to that one account. No usage introspection or state
tracking — cheap, and it degrades gracefully (if neither account answers, the
phase's existing `except` path takes over exactly as the API path does today).

## What Kevin has to do before this can replace the API pipeline

1. **Two interactive logins** (Drew cannot — they need his browser):
   - `set CLAUDE_CONFIG_DIR=C:\WorkInboxAI\kevin` then `claude setup-token` (or
     `claude` → `/login`) signed in as **kevin@lelitte.co.uk**.
   - `set CLAUDE_CONFIG_DIR=C:\WorkInboxAI\hope` then `claude setup-token` signed
     in as **hope@lelitte.co.uk**.
   Then set `WI_CLAUDE_CONFIG_DIR=C:\WorkInboxAI\kevin` and
   `WI_CLAUDE_CONFIG_DIR_FALLBACK=C:\WorkInboxAI\hope` as Windows user env vars.
2. **Cadence decision:** confirm 3×/day (recommended) or insist on 6×/day (then
   plan for Max / dedicated account per above).
3. **Cutover go-ahead:** after a short parallel-validation window where Kevin /
   Lauren eyeball `data/claude_briefing.json` vs the live `data/briefing.json`
   for a few days and are satisfied the triage quality matches (same model, same
   prompts — expected), Kevin gives an explicit go-ahead to either
   (a) point the existing `\Work Inbox Briefing` task's wrapper at
   `AI_BACKEND=claude_code` with `ANTHROPIC_API_KEY` unset, or
   (b) stand up a new task first. No `main` write, no scheduled-task change,
   without that go-ahead.
4. **Optional pre-cutover:** decide whether to also do mitigation #2 (collapse to
   one call) — recommended if staying on Pro.

## How to run a parallel-validation run by hand

```
cd <repo>
set AI_BACKEND=claude_code
set WI_AI_PARALLEL=1
set WI_CLAUDE_CONFIG_DIR=C:\WorkInboxAI\kevin           # once the logins exist
set WI_CLAUDE_CONFIG_DIR_FALLBACK=C:\WorkInboxAI\hope
python fetch_inbox.py
```
Produces `data/claude_briefing.json` + `data/claude_inbox_suggestions.json`
locally, appends `ai_backend_usage.jsonl`, pushes nothing, touches no ledger.
