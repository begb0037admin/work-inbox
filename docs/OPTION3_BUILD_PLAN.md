# Option 3 Build Plan — Connector-free `CODEX_HOME` + Outlook COM data pull

**Status:** Written 2026-08-27 (Drew). Build plan only — nothing in this file
has been built, logged in, configured, deployed, or scheduled. Same shape as
`docs/PHASE2_BRIEF.md`. Supersedes HANDOVER option 3 as the concrete plan for it.

**Kevin's decision this authorises planning for:** Option 3 from the HANDOVER
"~14:55" entry — *"Disconnect the Outlook/Calendar/Teams connectors from the
automation's ChatGPT account entirely, and give Phase 2 its read data via
Outlook COM instead. The six AI-triage phases still move to Codex — which is
what actually zeros the ~£36/mo — but the data pull stays on COM."* Codex
assessed this feasible (research doc Section 9, "Codex commissioned to attempt
a local fix", angle C, commit `c4ccbd1`).

---

## Exact next action

**Build Step 1 only, and stop for review:** create the dedicated connector-free
ChatGPT identity Kevin picks in "Decision 1" below, `codex login` it into a
separate `CODEX_HOME` (`C:\CodexAutomation\.codex`), and run one **read-only
verification** `codex exec` under it that lists its available tools — confirming
**zero `microsoft_outlook_*` / `microsoft_teams_*` / `microsoft_outlook_calendar_*`
tools are present**. Report that tool list back before building anything else.
Everything downstream (COM adapter, wrapper, quality gate, parallel schedule)
is gated on that one check passing. No `fetch_inbox.py` edit, no Task Scheduler
entry, no `main` write until Kevin gives a fresh explicit go-ahead for the
7-day run (which also requires the quality gate built first).

---

## The mission this plan is anchored on

**Zero the ~£36/mo Anthropic spend by moving the six AI-triage phases off the
Anthropic API onto Codex (ChatGPT-subscription auth, not per-token metered).**
Everything else is secondary and tradeable. Specifically traded away, with
Kevin's steer, to keep the path simple:

| Given up | Replacement | Cost of the trade |
|---|---|---|
| Graph `web_link` opener for codex-sourced tasks | Existing COM `openmail://<entryId>` path (already live, command-centre `5054906` / `986584e`) | None — COM EntryID is now the *only* id format in play, so codex tasks use the default unchanged opener. The `sourceType:'codex-graph'` / `openEmailWeb()` machinery already shipped goes dormant, not removed. |
| Outlook/Calendar/Teams **connector reads** (breadth: shared calendars, Teams messages, cross-mailbox) | `fetch_inbox.py`'s existing Outlook COM Phase 1 pull (inbox 50 newest + today/tomorrow calendar + subfolder/VIP sweeps) | None for the pipeline as it stands today — Phase 1 COM is already the production read path; nothing downstream consumed connector-only breadth. |
| Connector-read parity as a validation metric | Dropped from the quality gate | None — was only ever relevant to a connector-based variant. |
| **Missing-importance quality gap** (dry run: Codex saw 0 urgent vs real pipeline's 3) | **Fixed for free** — COM supplies `importance`/high-flag natively; `categorise()`'s `imp == 2 -> "urgent"` path works again | Net *positive*: Option 3 removes a known quality problem the connector variant had. |

If any later choice pits a nice-to-have against a simpler/faster path to the
cost saving, take the simpler path and note the trade for Kevin. Do not
gold-plate. The one risk that still gets real engineering attention is
**false demotion** (Codex marking real work `no_action_needed: true`) — scoped
proportionately in the quality gate below, not expanded.

---

## Decision 1 — the connector-free ChatGPT identity (Kevin decides; recommendation given)

**Why this is a decision and not an assumption:** the connector-free property
comes from the **ChatGPT account**, not from `CODEX_HOME`. The plugin-disable
test (27 Aug) proved the `microsoft_outlook_*` tools re-provision from the
account on every `codex exec` session regardless of local state. So a separate
`CODEX_HOME` logged into an account that *has* the Microsoft apps would still
expose the write tools. The automation needs an identity with **no Microsoft
apps connected, ever.**

**What local records show (all from admin-machine file inspection, 27 Aug):**

- **`eb7a812e-1b9d-4586-b1a4-02a4ed7ca116`** — the account `~/.codex/auth.json`
  is currently signed into (`auth_mode=chatgpt`, no `OPENAI_API_KEY`). This is
  the account carrying **all three Microsoft connectors** (Outlook Email
  `connector_4aaab285…`, Calendar `connector_e6a73946…`, Teams
  `connector_246af094…`), linked to `kevin.lelitte@admin.ox.ac.uk` in
  `.codex-global-state.json`'s sidebar catalog. Per prior sessions this is
  Kevin's **personal ChatGPT Plus**. **NOT connector-free.**
- **`cc80356f-959e-449f-9721-add87a9ba0a5`** — the account in both `auth.json`
  backups (31 Jul) and `cloud-config-bundle-cache.json` (21 Aug),
  `chatgpt_user_id=user-p4k3NWfqPlngDwubaxSCuWVF`. This one is
  **enterprise/workspace-managed** (`requirements_toml.enterprise_managed`:
  group approval-policy limits, `chronicle=false`, workspace model policy
  `gpt-5.6-luna`). This is the **Edu / managed-workspace** account. **Its
  Microsoft-connector state is not visible in any local file** — the sidebar
  catalog on disk was captured while signed into `eb7a812e`, not this account.
  Cannot be confirmed connector-free without Kevin checking the ChatGPT web UI
  while signed into it.

**Neither existing account is confirmed connector-free.** `eb7a812e` definitely
is not; `cc80356f` is unknown and, being workspace-managed, its connector state
could be changed by a workspace admin who may not be Kevin.

### Options for Kevin

| # | Identity | Cost impact on the £36/mo saving | Robustness |
|---|---|---|---|
| **A (recommended)** | **A dedicated new personal ChatGPT Plus account**, used only by the automation's `CODEX_HOME`, with Microsoft apps **never** connected. | ~£16/mo Plus fee → **net saving ~£20/mo** (still eliminates the Anthropic spend; Plus fee is the offset). | **Highest.** Fully Kevin-controlled, connector-free by construction, isolated from his interactive Plus/Edu quota, no admin can add apps to it. |
| B | **The existing Edu/managed account `cc80356f`**, *if* Kevin confirms in the ChatGPT web UI that it has **no** Microsoft apps connected **and** confirms no workspace admin can add them. | **Full £36/mo saving**, zero added cost. | Medium. Relies on a state Kevin must verify and may not fully control (workspace-managed). Also: its `enterprise_managed` approval-policy / model-policy requirements would apply to the automation runs — acceptable for a read-only text task, but note it. |
| C | Strip all three Microsoft connectors from `eb7a812e` (Plus) entirely and point the automation at it via the separate `CODEX_HOME`. | Full £36/mo saving, zero added cost. | Low. Kevin loses those connectors for his own interactive Codex use; re-adding them for interactive use later silently re-arms the write path if he ever runs the automation flow from his main `~/.codex`. Fragile. Not recommended. |

**Recommendation: Option A.** A dedicated identity is the only one that is
connector-free *by construction* rather than by a revocable/adminable setting,
and it keeps automation usage off Kevin's interactive quota (which matters for
the usage-headroom question in "Cost validation" below). The ~£16/mo Plus fee
is a real offset but the primary objective — no per-token Anthropic metering,
no live-mailbox write path — is fully met. If Kevin would rather take the full
saving and can satisfy himself on `cc80356f`'s connector state and control,
Option B is acceptable; Option C is not.

---

## Architecture — Option 3 is the existing Phase 2 dry-run machinery minus "Call 1"

The 26 Aug Phase 2 dry run (branch `drew/codex-phase2-ai-triage`) already built
and validated the hard parts. Option 3 **reuses them almost unchanged** and
deletes the connector dependency:

| Component | Phase 2 dry run (connector variant) | Option 3 |
|---|---|---|
| Data pull | "Call 1" — three `codex exec` connector pulls (inbox/sent/calendar), split to dodge the connector output-size limit | **Deleted.** `fetch_inbox.py`'s existing Outlook COM Phase 1 pull, reshaped by a thin adapter (below). No `codex exec` involved in the read at all. |
| Deterministic split | `tools/codex_triage/categorise_and_stage.py` — verbatim port of `fetch_inbox.py`'s `categorise()` / `badge_for()` / `make_card()` | **Unchanged**, except it reads the COM-shaped adapter output instead of `call_inbox_result.json`. |
| Granola calendar-prep context | `tools/codex_triage/build_granola_context.py` — verbatim port of Phase 3.7b, no `codex exec` needed | **Unchanged.** |
| The AI call | "Call 2" — one `codex exec` session, `build_call2_brief.py` assembles a brief that already states *"This call does NOT need any connector or tool access at all… Do not use any connector, do not attempt any write action of any kind"* | **Unchanged brief-builder.** This is the single `codex exec` invocation in the whole pipeline, and it is **already connector-free by design.** Under Option 3 it runs under the connector-free `CODEX_HOME`, so "no connector available" is enforced structurally, not just instructed. |
| System prompts | `EMAIL_SUMMARY_SYSTEM` / `TRIAGE_SYSTEM` / `SUMMARY_SYSTEM` / `CAL_SUM_SYSTEM` / Phase-2 context `SYSTEM` copied **verbatim** from `fetch_inbox.py`, only "Anthropic API" → "you" reframed | **Unchanged.** All six phases (the sixth — Phase 3.7 priority-task summaries — was added in the dry run, in `call2_judgement_output.json`'s `task_summary_phase`). |
| Deterministic post-processing | `fetch_inbox.py`'s Phase 3.3 / 3.3b / 3.3c demotion / staleness / thread-collapse and Phase 3.9 scroll-out — **ported as Python, not delegated to Codex** | **Unchanged.** Codex only does the five/six language-judgement phases; every categorise/badge/demotion/staleness/scroll-out rule stays deterministic Python so the tuned business logic is preserved exactly. |
| Output | `docs/codex_briefing.json` / `docs/codex_suggestions.json` / `docs/codex_triage_ledger.json` (produced once, dry run) | **Same three files**, formalised as the parallel-run outputs. Never touches `data/briefing.json` / `data/tasks.json` / `data/triage_ledger.json`. |

**So the genuinely new build for Option 3 is small:** (1) the COM→Codex input
adapter, (2) the connector-free `CODEX_HOME` + identity, (3) the warm-up/retry
wrapper, (4) formalised output writers + dedup ledger, (5) the quality-gate
harness, (6) the parallel Task Scheduler job (last, and separately gated).

---

## Build steps

### Step 1 — connector-free `CODEX_HOME` + identity  *(do first, then stop for review)*

1. Kevin creates / nominates the identity per Decision 1.
2. Create `C:\CodexAutomation\.codex\` (the automation `CODEX_HOME`). Nothing
   else on the machine points here; Kevin's interactive `C:\Users\admin\.codex\`
   is untouched.
3. `set CODEX_HOME=C:\CodexAutomation\.codex` then `codex login` — sign in as
   the connector-free identity. This writes `auth.json` into the automation
   `CODEX_HOME` only.
4. **Verification (read-only, no state change):** `set CODEX_HOME=C:\CodexAutomation\.codex`
   then a single `codex exec -s read-only --skip-git-repo-check` whose entire
   instruction is *"List every tool available to you in this session, grouped
   by namespace. Do not call any of them."* Confirm the output contains **no**
   `microsoft_outlook_email.*`, `microsoft_outlook_calendar.*`, or
   `microsoft_teams.*` entries. Also run `codex --version` / `codex doctor`
   under that `CODEX_HOME` and record it.
5. Confirm the interactive baseline is untouched: `~/.codex/config.toml` still
   sha1 `b2a1a22661b3596b92384e081b6625f786346f0e`, `codex doctor` under the
   default `CODEX_HOME` still clean.
6. **Report the tool list back. Do not proceed to Step 2 without Kevin's
   acknowledgement that the connector-free check passed.**

Isolation guarantees to state in the report:
- The automation `CODEX_HOME` has its own `auth.json`, `config.toml` (can start
  empty / minimal), `history.jsonl`, sessions, rules — zero overlap with
  `~/.codex`.
- The scheduled task sets `CODEX_HOME` in its own environment block only; it
  does not export it user-wide, so an interactive `codex` Kevin runs still uses
  `~/.codex`.
- If Kevin ever runs the automation flow manually he must `set CODEX_HOME` first
  — document this in the run `.bat`.

### Step 2 — COM → Codex input adapter

New script `tools/codex_triage/com_pull_adapter.py` (or fold into the wrapper).
It does **not** change `fetch_inbox.py`; it consumes `fetch_inbox.py`'s existing
Phase 1 output.

- **Source:** run `fetch_inbox.py` in a mode that emits its Phase 1 COM pull
  (inbox 50 newest, sent items, today/tomorrow calendar, subfolder + VIP
  sweeps) as JSON. If `fetch_inbox.py` has no such switch today, the smallest
  change is to add a `--dump-phase1 <path>` flag that writes the already-built
  in-memory Phase 1 structures and exits before any Anthropic call — this is an
  additive, non-behavioural change to the live script and needs its own small
  review + backup, but it is far less invasive than re-implementing the COM
  pull. Flag it to Kevin as the one `fetch_inbox.py` touch Option 3 needs.
- **Transform** to the shapes the reused scripts expect:
  - `categorise_and_stage.py` expects an object with an `inbox` array whose
    items have `subject`, `from_email`, `from_name`, `is_read`, `importance`,
    `received_utc`, `body_preview`, `id`, `kevin_is_primary_recipient`.
  - **`importance`** — populate from the COM high-flag path (`item.Importance ==
    2` → `"high"`). This is the field the connector could not supply; COM has
    it. Restores `categorise()`'s `imp == 2 -> "urgent"` classification and
    closes the dry run's "0 urgent vs 3" gap. No COM-shim join needed — the
    importance value rides in the same pull as everything else.
  - **`id`** — set to the COM **EntryID**. Under Option 3 there is no Graph id
    anywhere, so this is the canonical id end to end: the existing
    `triage_ledger.json` semantics, the `openmail://<entryId>` opener, and
    Phase 3.9's live `GetItemFromID` re-lookups all keep working unchanged for
    any eventual cutover.
  - Same transform for `sent` and `calendar` into the shapes
    `build_call2_brief.py` / `build_granola_context.py` read
    (`stage_triage_api_emails.json`, `real_briefing.json`, etc.).
- **Deterministic pre-processing before Codex:** run the ported
  `categorise()` / `badge_for()` / `make_card()` and the Phase 3.3/3.3b/3.3c/3.9
  logic here, in Python, exactly as the dry run does. Codex never sees these
  decisions as open questions.

### Step 3 — the single Codex call + output isolation

1. `build_call2_brief.py` assembles `brief_call2_judgement.txt` from the staged
   files (unchanged from the dry run — the six phases, verbatim system prompts).
2. One invocation:
   `CODEX_HOME=C:\CodexAutomation\.codex codex exec -s read-only --skip-git-repo-check -o call2_judgement_output.json < brief_call2_judgement.txt`
   (exact flags per the dry run). Read-only sandbox + connector-free `CODEX_HOME`
   ⇒ no filesystem write outside the workspace, no connector tools to call.
3. **Output writers** produce, from `call2_judgement_output.json` + the
   deterministic Python results:
   - `docs/codex_briefing.json` — parallel to `data/briefing.json`, never written to `data/`.
   - `docs/codex_suggestions.json` — parallel to `data/inbox_suggestions.json`.
   - `docs/codex_triage_ledger.json` — **separate dedup namespace**, keyed on
     the COM EntryID, its own `applied` / `promoted` / `tracked_needs_urgent`
     keys. It must never read from or write to `data/triage_ledger.json`. (The
     dry run already produced all three files once — this formalises them.)
   - Each codex-sourced suggestion carries `sourceType: "codex-graph"` for the
     eventual-cutover discriminator (already resolved + live, command-centre
     `986584e`) plus the COM `entryId` — so at cutover its opener is the
     **default unchanged `openmail://` path**; `web_link` / `openEmailWeb()` is
     not populated and not needed.
4. During the parallel run these files are **never** read by the dashboard.
   They exist only to be diffed against the live pipeline's real output.

### Step 4 — warm-up / retry wrapper

Cold `codex exec` after a gap, or after daemon churn, **reliably hangs on infra
startup** — hit repeatedly across 27 Aug (once ~90s on a bypass-trust warm-up).
The scheduled wrapper (`tools/codex_triage/run_codex_triage.ps1` or `.bat`) must:

1. `set CODEX_HOME=C:\CodexAutomation\.codex` (never rely on inherited env).
2. **Pre-flight warm-up:** one throwaway `codex exec -s read-only
   --skip-git-repo-check` with a trivial instruction (*"reply OK"*), timeout
   ~120s. Discard output.
3. **Retry loop on the real call:** up to 3 attempts, each with a hard timeout
   (~300s — the dry run's Call 2 ran ~42s, so 300s is generous headroom for a
   slow start). Between attempts: kill any lingering `codex` procs under the
   automation `CODEX_HOME`, wait 30s, warm up again.
4. If all 3 attempts fail: write a `codex_triage_FAILED_<ts>.json` marker, do
   **not** write partial `codex_*.json`, print a timestamped failure line
   (standing timestamp requirement), exit non-zero. A failed run is a no-op for
   the parallel comparison, not a corrupt output.
5. Every run prints a timestamped start/end line to the console/log.

### Step 5 — `PARALLEL_RUN_QUALITY_GATE_DESIGN.md` (hard prerequisite for the 7-day run)

A design already exists at
`docs/codex_phase2_run_20260826/PARALLEL_RUN_QUALITY_GATE_DESIGN.md` (branch
`drew/codex-phase2-ai-triage`). It is **design only, not built**, and it
pre-dates Option 3. Bring it onto this branch and finish it, scoped
proportionately — the real risk is false demotion; do not expand beyond that.

**What carries over unchanged:**
- **False-demotion measurement.** Every run emits
  `data/codex_runs/<ts>_codex_disagreements.json` + rolling `_rollup.json`.
  Disqualifying metric: `codex_hides_work` = Codex `no_action_needed: true` on
  an email the live pipeline kept in Needs/Urgent. A **single** such case on a
  *material* thread — appears by subject in that run's real context paragraph,
  OR contains an escalation marker (`at risk`, `no response`, `chasing`,
  `overdue`, `deadline`, `escalat`, named critical programme token `REF` /
  `HESA` / `TRAC`, …), OR VIP sender — fails auto-cutover on its own, not
  averaged away. The 26 Aug dry run's Simon Burford / Data Warehouse / REF
  thread would have tripped this.
- **`needs_reply` agreement rate** on subject-matched overlap — track it,
  suggested ≥ 95% bar, set with Kevin.
- **New-task-suggestion precision** — weekly human (Lauren/Kevin) genuine-or-noise
  marking of Codex `new_task`s the live pipeline didn't raise; watch
  `task_updates` volume for over-matching (dry run: 10 vs real 3).
- **Context-paragraph spot-check** — Lauren/Kevin read Codex's vs real's a few
  times a week: *"would I have been equally well-briefed."*
- **Repeated runs, not one.** The judgement is over the whole parallel window
  (≈ 6/day × 7 days ≈ 42 runs), not a single run — the dry run explicitly could
  not be trusted on n=1.

**What Option 3 changes / simplifies in the gate:**
- **Drop the missing-importance section (B) entirely.** COM supplies
  `importance`; Urgent-tier parity is expected, not excluded. The gate simply
  checks Codex's urgent count tracks the live pipeline's (both now driven by
  the same COM high-flag).
- **Drop connector-read-parity checks.** Not applicable — same COM source feeds
  both sides now; the only variable under test is **Codex's model vs
  Haiku 4.5 on identical data and identical rules**. This is the cleanest
  possible A/B and the gate should say so.
- **Volume parity is now a real check, not noise.** In the dry run, pull-depth
  differences (connector 40 vs COM 50 + Phase 3.9 carry-forward) muddied every
  comparison. Option 3 feeds *the same pulled set* to both — so a divergence in
  candidate counts is a genuine signal, not an artefact.

**Decision rule for trusting it (so cutover isn't a vibe):** over the parallel
window — `codex_hides_work` on material threads **must be 0**; `needs_reply`
agreement ≥ the bar set with Kevin; urgent-count parity within a small tolerance;
context-paragraph spot-checks pass; new-task precision acceptable to Lauren/Kevin.
Any material `codex_hides_work` hit = no auto-cutover, human review of that case.

### Step 6 — parallel Task Scheduler job  *(NOT this session; needs a fresh, separate go-ahead)*

Only after Steps 1–5 are built and verified, and only on Kevin's fresh explicit
authorisation for the 7-day run:

- A **new** scheduled task, separate from the live 7am/9am/11am/1pm/3pm/5pm
  `fetch_inbox.py` job. Suggest it runs a few minutes *after* each live run so
  it can diff against that run's fresh `data/briefing.json`.
- Runs the Step 4 wrapper under `CODEX_HOME=C:\CodexAutomation\.codex`.
- Writes only `docs/codex_*.json` + `data/codex_runs/*` + the disagreement
  files. Never writes `data/briefing.json` / `data/tasks.json` /
  `data/triage_ledger.json`. No cutover. Parallel only.
- Timestamped logging on every run.

---

## Opener impact (item 5) — confirmed clean

- Codex-sourced tasks under Option 3 carry a **COM EntryID** (no Graph id
  exists anywhere in this variant). Their Open-email button therefore uses the
  **existing, unchanged `openmail://<entryId>` → `open_email.py` →
  `GetItemFromID`** path — the same path every current task uses. Nothing new
  to build or test for the opener.
- The `source` / `sourceType` discriminator (command-centre `986584e`, live) is
  unaffected: `sourceType: "codex-graph"` can still tag codex-sourced tasks for
  provenance, but the opener's `sourceType === "codex-graph"` → `openEmailWeb()`
  branch simply won't fire because no `web_link` is populated — it falls through
  to the default `openEmail(e, entryId)` path. Confirmed by reading the shipped
  logic: `openEmailWeb()` requires a valid allowlisted `web_link` /
  `display_url`; absent that it shows the de-emphasised button + explanatory
  alert, never a broken open. With a valid `entryId` present the default path is
  taken.
- **Lost breadth, noted for Kevin:** the Graph connector's `web_link` (OWA deep
  link, openable from any machine/browser) is gone; codex tasks are only
  openable on a machine with classic Outlook + the `openmail://` handler
  registered — exactly the same constraint as every existing task. Also lost:
  connector reads of shared calendars and Teams messages. **Nothing in the
  current pipeline consumed either** — Phase 1 COM has always been the read
  source and the dashboard only ever shows Kevin's own inbox/calendar. No
  code path depended on connector breadth; this is breadth the migration
  never actually used.

---

## Cost validation (item 6)

The parallel run is also the cost measurement. Over the ≈ 7-day window:

1. **Anthropic side — does it actually go to ~zero?** The live `fetch_inbox.py`
   keeps running on Anthropic during the parallel window (that's the
   comparison baseline), so spend does **not** drop yet. What the parallel run
   proves is that Codex's output is *trustworthy enough to cut over*. The
   saving is realised only at cutover, when `fetch_inbox.py`'s six Anthropic
   calls (Phase 2 context, 3.2 summaries, 3.3/3.3b demotion, 3.5 triage, 3.7
   summaries, 3.8 calendar prep) are switched off. Measure it by: (a) the
   Anthropic console usage for the `work-inbox` key in the full calendar month
   *after* cutover — expect it to fall from ~$36.57/mo toward $0 (only any
   residual non-triage call remains, if any); (b) a code check that no
   `anthropic` client call remains on the six-phase path.
2. **ChatGPT side — usage headroom under 6×/day headless.** This is the
   unresolved caveat from research doc Section 4 (OpenAI retired per-message
   pricing April 2026 for token credits; heavy headless use may hit plan
   limits). During the parallel run, the wrapper logs per run: wall-clock
   duration, attempt count, and any rate-limit / quota error string from
   `codex exec`. After 7 days (≈ 42 real calls + 42 warm-ups) check: zero
   quota/limit errors, and — if Option A was chosen — that the dedicated
   account's usage page shows comfortable headroom at 42 calls/week
   extrapolated to a month (~180). If Option B (shared Edu account), also
   confirm the automation load didn't degrade Kevin's interactive use. If
   headroom is thin, the mitigation is to drop cadence (e.g. 3×/day) — the
   briefing is a morning artefact, not real-time.
3. **Net figure to report to Kevin:** £36/mo Anthropic eliminated, minus the
   identity cost (£0 for Option B/C, ~£16/mo for a dedicated Plus under Option
   A). State the net plainly once the identity is chosen.

---

## Deploy sequence (item 9) — manual discipline, no CI

Each of these is a discrete, separately-authorised step. HR-style, same as
command-centre's norms.

1. **Step 1 (identity + connector-free `CODEX_HOME`)** — build, run the
   read-only tool-list verification, report the tool list, **stop.** Restore
   point: nothing changed on the interactive side; `~/.codex/config.toml` sha1
   `b2a1a226…` unchanged; the automation `CODEX_HOME` is a brand-new directory
   that can simply be deleted to undo.
2. **Steps 2–4 (COM adapter, reused scripts on-branch, wrapper)** — build on
   branch `claude/outlook-codecs-connector-upgrade-fe3dgf` (or a fresh
   `drew/option3-*` branch). The one live-script touch is the
   `fetch_inbox.py --dump-phase1` flag: back up `fetch_inbox.py` first
   (timestamped, committed), make the additive flag change, verify it changes
   no existing behaviour (run the live path once, diff `data/briefing.json`
   byte-for-byte against a run without the flag), get Kevin's go-ahead for that
   specific one-line change before it lands on `main`.
3. **Step 5 (quality gate)** — doc + harness on branch. No production effect.
4. **Dry run** — run Steps 2–4 end to end once, by hand, under the automation
   `CODEX_HOME`. Produce `codex_*.json` + one `_codex_disagreements.json`
   against that day's real `data/briefing.json`. Review with Kevin.
5. **Step 6 (7-day parallel schedule)** — only on Kevin's **fresh explicit
   separate go-ahead**, given after the quality gate exists and the dry run
   looked right. Stated restore point: delete the new scheduled task; the live
   `fetch_inbox.py` job is untouched throughout. Verify the first scheduled run
   fired, produced the three files, wrote no `data/` production file, logged a
   timestamp.
6. **Cutover** — **not in scope of this plan at all.** A separate decision after
   the parallel window, judged against Step 5's decision rule, with its own
   restore point and go-ahead.

---

## Hard gates (unchanged, restated)

- No build beyond Step 1's verification without Kevin's acknowledgement.
- No `codex login`, no `CODEX_HOME` creation, no `config.toml` change, no
  `fetch_inbox.py` edit, no deploy, no Task Scheduler entry **this session**.
- The 7-day run needs Kevin's fresh, separate, explicit go-ahead **after** the
  quality gate is built.
- No Phase 2 Codex task-writer, no `source:'codex-graph'` **write** to
  `data/tasks.json`, no PAT rotation, no Oxford IT escalation, no `main` writes
  without the per-change go-ahead above.
- Leave the machine at the `b2a1a226…` `~/.codex/config.toml` baseline.
- Every run's console/log output prints a timestamp.
