# Email automation security mitigations

**Status:** decision brief, 28 August 2026. Written by the coordinator with Kevin. Sits alongside `IMAP_OAUTH2_SPIKE_20260828.md`, `MAIL_BACKEND_MIGRATION_PLAN.md`, `CODEX_CONNECTOR_MIGRATION_RESEARCH.md`.

## Context

work-inbox needs to pull mail (and ideally calendar) for AI triage 5x/weekday. Three routes to the Microsoft 365 cloud mailbox exist:

| Route | Token holder | Runs on whose funding? | Status |
|---|---|---|---|
| IMAP (OAuth2, Thunderbird public client) | our own read-only Python | n/a (own code, no AI cost) — but rides a Microsoft first-party preauthorization Oxford didn't grant | **Proven, parked** (`IMAP_OAUTH2_SPIKE_20260828.md`) |
| Graph-direct (OAuth2) | our own read-only Python | n/a — untested, likely admin-gated | Not attempted |
| **ChatGPT M365 connector** | **Codex (`codex exec`)** | **Oxford** — ChatGPT Edu (Oxford-funded) covers Codex CLI use; also the only route that reaches Teams + calendar | **CHOSEN 28 Aug.** Write-gate handled by the layered mitigation model below, not by gating connector tools. |

Kevin's position (28 Aug): this is a **funding** choice, not a compliance one. Using Claude is allowed at Oxford; the issue is that `claude -p` runs on Kevin's **personal** Claude subscription, whereas Oxford pays for **ChatGPT Edu**, and ChatGPT Edu's entitlement **does** cover Codex CLI / programmatic use (confirmed). Moving the pipeline to Codex shifts the running cost from Kevin's personal spend onto Oxford's paid entitlement. Kevin also has a personal ChatGPT Plus account. Separately, the ChatGPT M365 connector is the only route that also reaches **Teams** (currently invisible to work-inbox) and calendar in one place.

## The core risk

The connector exposes mailbox **write** tools to an autonomous agent:
- Confirmed exposed (proof-fired Aug): `set_message_categories`, flags, read-state, folder move.
- Unconfirmed but plausible: send / reply / forward / draft-create. Codex could not enumerate this from reasoning alone; it needs a live tool-manifest read (never run). A connector/account update could add them later even if absent now.

In headless `codex exec` these tools **cannot be reliably denied**. Every local and account-side control failed 26–27 Aug: `config.toml [apps]`, per-connector "allow read actions only", top-level "always ask", plugin-disable (config + physical cache), PreToolUse hooks (built, tested — write executed, hook never fired), execpolicy `.rules`, `codex mcp` filter. No `--allowed-tools` / `--deny-tool` flag in codex-cli 0.149.1.

**Specific failure mode:** an unintended **send** triggered by prompt injection from hostile email body content during an unattended run — not the agent choosing to send. Codex second-opinion rating: "low, but not negligible — a genuine tail risk, not theoretical." Dominant vector is prompt injection; hallucinated/confused tool calls and retry loops are lower.

**Codex verdict:** the connector route is "VIABLE ONLY WITH tenant-enforced read-only Graph consent" (`Mail.Read`, no `Mail.ReadWrite`, no `Mail.Send`). Kevin has confirmed Oxford IT will not provide that.

## Mitigation layers (defence in depth — none sufficient alone)

| # | Layer | Stops | Does not stop |
|---|---|---|---|
| 1 | **Connector = dumb fetch only.** Rigid instruction: "list messages in date range, return these fields." No open-ended reasoning over email content in the connector call. All AI triage runs afterward over sanitised data. | Most injection — the connector call is not interpreting hostile text | A steered send call if the model is diverted mid-fetch |
| 2 | **Sanitise / truncate bodies** before any AI step. Triage classifies mostly on subject / sender / metadata; full bodies rarely needed. | Injection payloads in body text reaching a reasoning step | Injection via subject or sender display name |
| 3 | **Connector / account set read-only** in ChatGPT settings. | Casual writes, *if* the setting is enforced for `codex exec` | Tried 27 Aug, did not hold for `codex exec` — must be re-verified against current CLI |
| 4 | **Draft-only rule** — system instruction: never send, only draft. | The agent *choosing* to send | Injection *tricking* a send — bypasses the instruction rather than obeying it |
| 5 | **Human review of every draft** the automation produces (fits the existing Lauren drafting-loop: drafts land in `drafts.json` / Outlook Drafts, Kevin sends manually, nothing auto-sent). | A bad draft leaving the mailbox | Nothing — this is the human backstop |
| 6 | **`mailbox_guard.py` post-run sweep** — COM delta-check of Sent / Drafts / categories / flags / read-state vs a pre-run baseline; hard-disables the scheduled task + BurntToast alert on any delta. Proof-fired on a synthetic category change. | Silent recurrence — you find out and the task stops | The send itself: message already delivered, irreversible; detection lag ≈ remaining runtime + sweep, i.e. minutes |

## Bottom line

- Layers 1–2 do the real work: if the connector never reasons over untrusted content, the injection surface nearly closes.
- Layers 3–6 are backstops of decreasing strength. Layer 5 (human-in-the-loop on every draft) is the only hard guarantee.
- Stacked, residual risk of an unintended external send is **small but non-zero**. There is no perfect enforcement of the connector's write tools; the layered model is what makes the residual acceptable.
- This is a **risk-acceptance decision for Kevin** (made 28 Aug: proceed with the connector): accept a small, mitigated residual to move running cost onto Oxford-funded ChatGPT Edu and gain Teams + calendar in one path — versus stay on COM + `claude -p` on Kevin's personal spend (zero external-send risk, no Teams, and a hard cliff for calendar when New Outlook is forced).

## Open questions to resolve before committing

1. **RESOLVED (28 Aug, Kevin): not a policy question — a funding one.** Claude is allowed; the point is that `claude -p` is Kevin's personal spend and Oxford-funded ChatGPT Edu covers Codex CLI use. Both the mail pull *and* the triage engine move to Codex so the running cost sits on Oxford's entitlement. `claude -p` stays live as the fallback until the Codex path is proven at parity.
2. **Does the connector actually expose a `send` / `reply` / `forward` / `draft-create` tool?** One read-only `codex exec` tool-manifest enumeration answers it definitively. — Drew.
3. **Does the connector "read-only" / "allow read actions only" setting hold for `codex exec` in the current codex-cli?** Re-test; it did not in 0.149.1 as of 27 Aug. — Drew.

## Q2 / Q3 findings -- read-only investigation, 2026-08-28 ~22:05 UTC (Drew)

**Baseline discipline held.** `~/.codex/config.toml` sha1 `35f8910382373d525598194b2649159cfeed3f6a` **before and after, unchanged**. No `codex login`, no `[apps]` edits, no writes, no Outlook/mail/calendar/Teams action. One `codex exec -s read-only --skip-git-repo-check --json` enumeration run (10s, exit 0; JSONL shows a single `agent_message`, **zero tool calls**), plus `codex mcp list` and `codex features list` (read-only status). codex-cli **0.149.1** (no 0.150.x on the machine; not updated). Artifacts in the session scratchpad, not committed.

### Q2 -- does the connector expose send / reply / forward / draft-create?

**Two-level answer:**

1. **In the connector's published manifest (ChatGPT account catalog, `~/.codex/.codex-global-state.json` -> `mcp-extension-sidebar-catalog`): YES.** The "Microsoft Outlook Email" connector (`connector_4aaab2856305417b993eca9a216aaf6e`) is still catalogued today; its tool set (per the 28 Aug enumeration in HANDOVER) includes `send_email`, `send_email_on_behalf`, `reply_to_email`, `forward_email`, `schedule_email`, `draft_email`, `create_reply_draft`, `create_forward_draft`, `create_shared_reply_draft`, `move_email`, `mark_email_read_state`, `set_message_categories` (46 tools, ~24 state-changing). "Microsoft Outlook Calendar" (`connector_e6a7394...`) and "Microsoft Teams" (`connector_246af09...`) are also catalogued. **So the connector definition does contain send/reply/forward/draft-create tools.**

2. **In an actual headless `codex exec -s read-only` session right now: NONE of them load.** The live tool manifest returned was **only**: `functions.exec`, `functions.wait`, `collaboration.spawn_agent`, `collaboration.followup_task`, `collaboration.interrupt_agent`, `collaboration.list_agents`, `collaboration.send_message`, `collaboration.wait_agent`. **Zero `microsoft_*` / `outlook` / `email` / `teams` / `calendar` / `github` tools.** This is a **change from 26-27 Aug**, when `microsoft_outlook_email.set_message_categories` was proven callable from `codex exec`.
   - `codex features list`: `apps` = stable/**true** (Apps enabled), `enable_mcp_apps` = under development/false.
   - `codex mcp list`: only `meeting-context`, `node_repl`, `openaiDeveloperDocs`, and `cua_repl` (**disabled** -- the `ChatGPT.exe` bridge). No connector MCP entries (connectors are "Apps", bridged via the ChatGPT app-server, not `mcp_servers`).
   - **Why the connectors aren't loading into `codex exec` is undetermined.** Candidates: connector-side auth expired ("not currently logged in for connectors"); the `cua_repl` / ChatGPT app-server bridge being disabled/not running; residual state from the 27 Aug plugin-disable tests. Resolving it needs `codex login` / launching the ChatGPT app -- which the brief says not to do -- so per instruction, **STOPPED here**.

**Net:** the send tools are real and defined, but as of this check they are **not reachable from `codex exec`**. Do NOT treat this as a durable safety property -- it is an unexplained current state, not an enforced control; a `codex login` / connector re-auth / ChatGPT-app launch could restore them.

### Q3 -- does the "read-only" / "allow read actions only" setting hold for `codex exec` now?

**Currently untestable, and the underlying mechanism is unchanged from 27 Aug.**
- codex-cli is still **0.149.1**. `exec_permission_approvals` (the feature that would gate connector tools) = still "under development / disabled". No `--allowed-tools` / `--deny-tool` in `codex exec --help`.
- The 27 Aug write-gate re-test could **not** be repeated: with zero connector tools in the `codex exec` session there is nothing to apply a read-only restriction to.
- **Assume the 27 Aug finding still stands** (the ChatGPT read-only setting did NOT remove the write tools from a headless `codex exec` tool list) until it can be re-verified against a session that actually loads the connectors -- nothing in the CLI has changed to fix it.

### Q1 -- RESOLVED (28 Aug, Kevin): FUNDING, not policy. Claude is allowed; `claude -p` is Kevin's personal spend while Oxford-funded ChatGPT Edu covers Codex CLI programmatic use. Both mail pull and triage move to Codex to put running cost on Oxford's entitlement. `claude -p` stays the live fallback until the Codex path is proven at parity.

## Not changing anything yet

COM + `Classic Outlook Keepalive` watchdog remains the live mail path. IMAP is parked as proven. This brief is analysis only — no route change without Kevin's explicit go-ahead.
