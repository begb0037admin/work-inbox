# Headless Claude Code for the six AI-triage phases — SCOPE (no build)

**Status:** Scoping only, 2026-08-27 (Drew). Nothing here is built, configured,
scheduled, or deployed. Supersedes the Codex-connector route (see
`docs/CODEX_CONNECTOR_MIGRATION_RESEARCH.md` Section 9, 27 Aug "PIVOT" entry, and
`docs/OPTION3_BUILD_PLAN.md` — both now dormant).

## The pivot in one line

Kevin's decision (27 Aug 2026, via coordinator): stop pursuing Codex. Instead run
`fetch_inbox.py`'s AI-triage model calls through **headless Claude Code on
Kevin's Claude subscription** (flat monthly fee) instead of the **metered
Anthropic API** (`ANTHROPIC_API_KEY`, ~£36/mo). **Same model, same prompts** —
this is a billing-path swap, not a model swap, so it does **not** need the
7-day A/B validation the Codex plan required.

## Why this is a much better fit than Codex was

| Problem that killed the Codex route | Under headless Claude Code |
|---|---|
| `codex exec` auto-loads the ChatGPT account's Outlook/Calendar/Teams **connector write tools**; every account-side and machine-local control to disable them was tested 26–27 Aug and **failed** (Section 9). | Claude Code has **no mailbox tool of any kind**. Its toolset is Bash/Read/Write/Edit/Glob/Grep/WebFetch/WebSearch + configured MCP servers — none can touch Outlook. |
| `codex exec --help` (0.149.1) has **no tool-scoping flag** — you cannot tell it "no tools". | `claude` **does** support `--allowedTools` / `--disallowedTools` / `--tools` and `--strict-mcp-config`. You can run it with an empty toolset. |
| Different model (`gpt-5.6-terra`) vs the pipeline's `claude-haiku-4-5` → required a full quality A/B before cutover. | `--model claude-haiku-4-5` → **identical model to today**. No triage-quality revalidation needed. |
| Connector could not supply `importance` → lost the Urgent tier. | No change to the data path — `fetch_inbox.py` Phase 1 Outlook COM pull is untouched; `importance` still native. |

---

## 1. Feasibility — headless invocation from Task Scheduler on a *subscription*

**Confirmed available on this machine (admin, 27 Aug 2026):**

- **Claude Code `2.1.247`**, `claude` on `PATH`.
- Headless mode: **`claude -p "<prompt>"`** (`--print`), with `--output-format
  json` (returns `{ "result": "<assistant text>", "total_cost_usd", "usage":
  {...}, "num_turns", "session_id", ... }`) or `stream-json`. Prompt can also be
  piped on stdin (`echo "<prompt>" | claude -p`), and `--input-format` supports
  structured input.
- Non-interactive hardening flags all present: `--allowedTools`,
  `--disallowedTools`, `--tools`, `--mcp-config`, `--strict-mcp-config`,
  `--permission-mode`, `--system-prompt` / `--system-prompt-file` /
  `--append-system-prompt`, `--exclude-dynamic-system-prompt-sections`,
  `--model`, `--fallback-model`, `--no-session-persistence`, `--settings`.
- **`claude setup-token`** subcommand exists — "Set up a long-lived
  authentication token" — the sanctioned path for headless/CI auth against a
  subscription.

**Auth model — this is the one real gotcha.** Two credential paths coexist:

1. **Subscription (what we want):** `~/.claude/.credentials.json` →
   `claudeAiOauth` block with `accessToken` / `refreshToken` / `expiresAt` /
   `scopes` / **`subscriptionType`** / **`rateLimitTier`**. This is the Claude.ai
   subscription OAuth token. Present and valid on this machine right now (this
   very session runs on it — `CLAUDECODE=1`, `CLAUDE_CODE_ENTRYPOINT` set, and
   "usage limit" is a subscription concept).
2. **Metered API:** **`ANTHROPIC_API_KEY` is currently set as a Windows user
   environment variable** and is what `fetch_inbox.py` uses today. **Claude Code
   prefers `ANTHROPIC_API_KEY` when it is present in the environment** and will
   then bill the API — which would defeat the entire purpose of the swap.

**Therefore the scheduled invocation MUST run in an environment where
`ANTHROPIC_API_KEY` is unset**, with subscription auth supplied by either the
existing `~/.claude/.credentials.json` (auto-refreshing) or, better for
unattended use, a **`CLAUDE_CODE_OAUTH_TOKEN`** minted by `claude setup-token`
and stored as a Windows user env var (never in a file, per project rules).

**Task Scheduler shape (proposed, not built):**

```
Program:   powershell.exe
Arguments: -NoProfile -WindowStyle Hidden -File "<repo>\Run_Inbox_Briefing_CC.ps1"
```

and `Run_Inbox_Briefing_CC.ps1` does, in order:

```powershell
$env:ANTHROPIC_API_KEY = $null          # force subscription billing
# $env:CLAUDE_CODE_OAUTH_TOKEN already set as a user env var (setup-token)
git fetch origin; git checkout origin/main -- fetch_inbox.py
python fetch_inbox.py                     # AI_BACKEND=claude_code (see section 6)
```

No blocker. `claude -p` from a hidden PowerShell under Task Scheduler is a
standard, documented pattern.

---

## 2. Model — identical to today, so no A/B revalidation

- `fetch_inbox.py` makes **exactly 5** `client.messages.create()` calls, **all
  `model="claude-haiku-4-5"`** (lines 805, 1055, 1886, 2180, 2379 — Phase 2
  context, Phase 3.2 email summaries, Phase 3.5 task triage, Phase 3.7
  priority-task summaries, Phase 3.8 calendar prep). Phases 3.3/3.3b are
  deterministic Python post-processing on Phase 3.2's output, not model calls —
  hence "six phases, five calls".
- `claude --model claude-haiku-4-5` runs **the same model**. Same model + the
  **verbatim same system/user prompt text** = the same output distribution as
  the current API calls. **This is a billing-path change only. No triage-quality
  A/B, no 7-day parallel validation, is required** — state this plainly to Kevin.
- **One parity nuance to check once (not a 7-day exercise):** `claude -p` wraps
  the model in Claude Code's own system prompt / harness. To get as close as
  possible to a bare Messages API call, pass the pipeline's exact system prompt
  via **`--system-prompt "<verbatim>"`** together with
  **`--exclude-dynamic-system-prompt-sections`** (strips Claude Code's dynamic
  preamble). Then do a **one-off** side-by-side on a single day's Phase 3.2 batch
  (API output vs `claude -p` output) to confirm the summaries/verdicts match in
  practice. If they do — done. This is a sanity check, not a gate.
- `--fallback-model` is available if a Haiku capacity blip needs a graceful
  degrade (e.g. to `claude-haiku-4-5` → a Sonnet fallback), but note the
  pipeline deliberately locked to Haiku because "Sonnet timed out on this inbox
  size" (`CLAUDE.md`) — so any fallback needs the same care.

---

## 3. Which subscription, and does 6×/day fit

- Kevin has three Claude logins — `kevin@` / `hope@` / `adam@lelitte.co.uk`
  (Kevin / Hope / Adam personas; see `~/.claude/CLAUDE.md` memory
  "Claude login accounts"). work-inbox is **Kevin's** project → the automation
  should run on **`kevin@lelitte.co.uk`'s** subscription.
- **`~/.claude/.credentials.json` carries `subscriptionType` and `rateLimitTier`
  fields** — Kevin should read those (or check the Claude web UI → Settings) to
  confirm the exact plan before a build. Not dumped here (his token file).
- **Load added by the automation:** 5 model calls per run × 6 runs/day (the live
  `\Work Inbox Briefing` task fires 7/9/11/1/3/5 Mon–Fri) = **30 Haiku calls/day
  ≈ 150/week**. Each is a short, single-turn, tools-off call over a bounded
  inbox — small token counts, no agentic loop. In Claude Code subscription
  terms (rolling 5-hour windows + a weekly cap, measured in tokens/messages),
  30 small Haiku calls spread across a working day is a **modest** but **not
  free** addition.
- **The real consideration Kevin must weigh:** this automation **shares the same
  subscription usage pool as Kevin's interactive Claude Code / coordinator /
  agent work** (Drew, Lauren, Matthew, etc. all run on it). Kevin is **near a
  usage limit right now**. 30 Haiku calls/day is unlikely to be the thing that
  tips him over on its own, but it is real consumption he cannot see itemised.
  Mitigations, in order of preference:
  1. **Cadence cut to 3×/day** (e.g. 7/11/3). The briefing is a morning/interim
     artefact, not real-time — halving the call volume for little practical loss.
  2. **A dedicated or secondary Claude account** for the automation only (mirrors
     the Codex "Option A" logic — isolation from the interactive pool). Cost of a
     second plan vs the £36/mo API saving needs to net out.
  3. **The Edu / university-associated account**, *if* one exists for Claude and
     its terms permit automated use (unconfirmed — Kevin to check; the Codex Edu
     account was ChatGPT, not Claude).
- **Recommendation:** start on `kevin@`'s existing plan at **3×/day**, watch the
  weekly usage for one week, raise to 6×/day only if there is clear headroom.

---

## 4. Terms of service

- Headless / scripted Claude Code is a **documented, first-class feature**:
  `claude -p`, the Claude Agent SDK, `claude setup-token` "for CI", and the
  official GitHub Actions integration all exist specifically for
  non-interactive/automated use.
- A **personal morning-briefing automation that Kevin himself consumes** is
  individual use. The ToS boundaries that matter here — and that this stays well
  inside:
  - **no reselling / no serving the output to third parties** (it is Kevin's own
    inbox briefing, for Kevin);
  - **no running multiple accounts to evade rate limits** (so: pick one account
    and stay on it — do **not** shard the 6 runs across `kevin@`/`hope@`/`adam@`
    to dodge a cap; that would breach terms);
  - **no automated bulk/abusive volume** — 3–6 short calls a few times a day is
    nowhere near that.
- **Auth method:** use `claude setup-token` (the sanctioned long-lived headless
  token) rather than lifting the interactive `accessToken` out of
  `.credentials.json` — same effect, but it is the supported path and it
  refreshes cleanly for unattended runs.
- **Net:** scheduled automated use at this cadence, on one account, for Kevin's
  own workflow, is within terms. If Kevin wants belt-and-braces certainty he can
  confirm with Anthropic support, but nothing here is a grey area.

---

## 5. Write-risk — the Codex blocker does **not** exist here

**Assessment: headless Claude Code has NO path to write to Kevin's live Oxford
mailbox.**

- **No mailbox tool.** Claude Code ships Bash/Read/Write/Edit/Glob/Grep/WebFetch/
  WebSearch + whatever MCP servers are configured. There is **no Outlook /
  Exchange / Graph / calendar tool** in that set. The thing that made Codex
  dangerous — an auto-provisioned `microsoft_outlook_email.set_message_categories`
  et al. from a connected ChatGPT app — has **no analogue** in Claude Code.
- **The COM calls are the pipeline's own Python.** `fetch_inbox.py` talks to
  Outlook via `win32com` in its own process. The model is handed **plain
  text/JSON** and returns **plain text/JSON**. It never sees, and cannot invoke,
  the COM layer. Swapping the *billing path* of the 5 model calls changes
  nothing about that boundary.
- **Tools can be positively locked off.** The scheduled invocation runs:
  ```
  claude -p --model claude-haiku-4-5 \
    --system-prompt-file <phase_system_prompt> --exclude-dynamic-system-prompt-sections \
    --allowedTools "" --strict-mcp-config --mcp-config "{}" \
    --permission-mode default --output-format json --no-session-persistence
  ```
  Empty allowed-tools + `--strict-mcp-config` with an empty MCP config ⇒ the
  model has **zero tools and zero MCP servers** (in particular the one global
  MCP server on this machine, `github`, is **not** loaded). It is a pure
  text-completion call.
- **Even a hypothetical misbehaving model cannot reach the mailbox** — there is
  no tool to call, and `fetch_inbox.py` never passes model output to any COM
  write (it only ever *reads* via COM; the only writes it performs are to
  GitHub, local JSON, and — Phase 3.6 — `command-centre/data/tasks.json` via the
  Cloudflare Worker, none of which is the mailbox).

**Consequence for the kill-switch (`tools/codex_triage/mailbox_guard.py`,
already built + proof-fired this session):** it drops from **hard prerequisite**
(which it was for the Codex route) to **lightweight optional regression check**.
Recommended use: run it as a before/after wrapper around the **first ~5 days**
of live headless runs to positively confirm "no mailbox delta, ever" under the
new backend, then retire it or keep it as near-zero-cost insurance. It is **not**
a gate on the Claude Code build. Full details:
`docs/OPTION1_KILLSWITCH.md`.

**Residual risks worth a line each (all minor):**
- A future change that adds an MCP server or re-enables tools for this
  invocation would need the same "no tools" review — bake `--allowedTools ""
  --strict-mcp-config` into the wrapper and document why.
- The `ANTHROPIC_API_KEY`-unset requirement is a footgun: if a future edit lets
  the key leak into the scheduled env, billing silently reverts to metered API
  (a cost regression, not a safety one). The wrapper should assert the key is
  unset and log it (timestamped, per standing requirement).

---

## 6. The swap — smallest change that works

**One helper, five call sites, one feature flag.**

`fetch_inbox.py` today: one module-level `client = anthropic.Anthropic(timeout=60.0)`
and `anthropic_available` flag (line ~800), then 5 `client.messages.create(
model="claude-haiku-4-5", system=..., messages=[{"role":"user","content":...}],
max_tokens=..., temperature=...)` calls, each already wrapped in its own
try/except + JSON-extract/repair.

**Change:**

1. Add `AI_BACKEND = os.environ.get("AI_BACKEND", "api")` near the top.
2. Introduce `def _ai_text(system, user, max_tokens, temperature=0.0) -> str`
   that returns the assistant's text and nothing else:
   - `AI_BACKEND == "api"` → the existing `client.messages.create(...)` path,
     returns `resp.content[0].text`. **Unchanged default — zero behavioural
     risk until the flag is flipped.**
   - `AI_BACKEND == "claude_code"` → `subprocess.run(["claude", "-p",
     "--model", "claude-haiku-4-5", "--system-prompt", system,
     "--exclude-dynamic-system-prompt-sections", "--allowedTools", "",
     "--strict-mcp-config", "--mcp-config", "{}", "--output-format", "json",
     "--no-session-persistence"], input=user, capture_output=True, text=True,
     timeout=120, env={**os.environ, "ANTHROPIC_API_KEY": ""})`, then
     `json.loads(stdout)["result"]`. Log `usage` / `total_cost_usd` from that
     JSON every call (feeds the usage-projection Kevin asked for, and the
     standing timestamped-output rule).
3. Replace each of the 5 `client.messages.create(...)` + `.content[0].text`
   blocks with a `_ai_text(system, user, max_tokens, temperature)` call. The
   surrounding JSON-repair/validation logic is **untouched** — it operates on
   the returned string exactly as now.
4. `anthropic_available` becomes `backend_available`: for `api`, the existing
   probe; for `claude_code`, a one-off `claude -p "reply OK"` pre-flight (also
   doubles as the cold-start warm-up — Claude Code does not have Codex's
   infra-hang problem, but a pre-flight is cheap and confirms auth/token
   validity before the real calls).
5. **Rollout:** land the helper with `AI_BACKEND` defaulting to `api` (no-op).
   Flip to `claude_code` first for **one manual run**, diff that run's
   `data/briefing.json` against an `api` run of the same inbox (the section-2
   parity check). If clean, point the new scheduled task (or the existing one's
   wrapper) at `AI_BACKEND=claude_code`. Keep `api` as the instant rollback —
   one env var.

**Files touched:** `fetch_inbox.py` (one new helper + 5 mechanical call-site
swaps), one new wrapper `.ps1`, one Task Scheduler change (or a new parallel
task first). No change to Phase 1 COM, Phase 3 card logic, Phase 3.6 CC sync,
the dashboard, or any data schema.

---

## What remains before a build go-ahead

1. **Kevin confirms the plan tier** on `kevin@lelitte.co.uk` (`subscriptionType`
   / `rateLimitTier`) and the **cadence** (recommend 3×/day to start).
2. **Kevin runs `claude setup-token`** on the admin machine and stores
   `CLAUDE_CODE_OAUTH_TOKEN` as a Windows user env var (Drew cannot mint this —
   it needs his interactive login).
3. **Decide the account question:** `kevin@` shared pool (simplest, recommended
   with the cadence cut) vs a dedicated Claude plan for the automation.
4. Then Drew builds section 6's helper behind `AI_BACKEND`, does the one-off
   Phase 3.2 parity diff, reports, and waits for the flip go-ahead.
5. Optional: wrap the first ~5 scheduled `claude_code` runs in
   `mailbox_guard.py` (before/after COM sweep) to positively confirm zero
   mailbox effect, then retire it.

**Cost outcome to confirm post-cutover:** Anthropic console usage for the
work-inbox `ANTHROPIC_API_KEY` should fall from ~£36/mo toward £0; the offset is
whatever headroom the `kevin@` subscription gives up. Net saving is the full
~£36/mo **if** it rides the existing plan without forcing an upgrade — which the
3×/day cadence is chosen to protect.
