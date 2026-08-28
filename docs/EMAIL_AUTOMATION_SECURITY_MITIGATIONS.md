# Email automation security mitigations

**Status:** decision brief, 28 August 2026. Written by the coordinator with Kevin. Sits alongside `IMAP_OAUTH2_SPIKE_20260828.md`, `MAIL_BACKEND_MIGRATION_PLAN.md`, `CODEX_CONNECTOR_MIGRATION_RESEARCH.md`.

## Context

work-inbox needs to pull mail (and ideally calendar) for AI triage 5x/weekday. Three routes to the Microsoft 365 cloud mailbox exist:

| Route | Token holder | Sanctioned by Oxford? | Status |
|---|---|---|---|
| IMAP (OAuth2, Thunderbird public client) | our own read-only Python | No — rides a Microsoft first-party preauthorization | **Proven, parked** (`IMAP_OAUTH2_SPIKE_20260828.md`) |
| Graph-direct (OAuth2) | our own read-only Python | No — untested, likely admin-gated | Not attempted |
| **ChatGPT M365 connector** | **Codex (`codex exec`)** | **Yes — Oxford has consented the ChatGPT enterprise app** | Write-gate unresolved |

Kevin's position (28 Aug): a tenant that is locking down to "ChatGPT only" is exactly the kind that will later block unsanctioned OAuth clients — which would kill IMAP-direct and Graph-direct both. The ChatGPT connector may become the only governance-durable route to mail **and** calendar.

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
- Stacked, residual risk of an unintended external send is **small but non-zero**. There is no perfect enforcement on the Oxford-sanctioned path.
- This is a **risk-acceptance decision for Kevin**: accept a small, mitigated residual to stay on the sanctioned, calendar-capable, New-Outlook-proof path — versus stay on COM + the `Classic Outlook Keepalive` watchdog (zero external-send risk, but a hard cliff for calendar when New Outlook is forced).

## Open questions to resolve before committing

1. **Is Claude Code in-policy at Oxford, or is ChatGPT the only sanctioned AI tool?** The triage currently runs on headless Claude Code (`claude -p`, cut over 27 Aug). If only ChatGPT is sanctioned, the triage engine also has to move, not just the mail pull. — Kevin / governance.
2. **Does the connector actually expose a `send` / `reply` / `forward` / `draft-create` tool?** One read-only `codex exec` tool-manifest enumeration answers it definitively. — Drew.
3. **Does the connector "read-only" / "allow read actions only" setting hold for `codex exec` in the current codex-cli?** Re-test; it did not in 0.149.1 as of 27 Aug. — Drew.

## Not changing anything yet

COM + `Classic Outlook Keepalive` watchdog remains the live mail path. IMAP is parked as proven. This brief is analysis only — no route change without Kevin's explicit go-ahead.
