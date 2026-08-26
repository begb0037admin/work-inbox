# Connector write-path investigation — can the Codex Outlook/Calendar/Teams connectors be made read-only? (26 Aug 2026, Drew)

Follow-up to the failed write-gate test (`WRITEGATE_TEST_INCIDENT.md`). Investigation only — **no consent, scope, or config change was made.** All findings are from local file inspection on the admin machine plus `codex` CLI help/feature output plus two read-only `codex exec` probes.

## 1. Where the scope grant actually lives — NOT local, NOT user-adjustable

The three connectors are OpenAI-managed "curated remote" apps (marketplace `openai-curated-remote`), each with a fixed connector id declared in its `.app.json`:

| Connector | id | Declared capabilities (`plugin.json` `interface.capabilities`) |
|---|---|---|
| Outlook Email | `connector_4aaab2856305417b993eca9a216aaf6e` | `["Interactive", "Write"]` |
| Outlook Calendar | `connector_e6a7394682e24467ac68c60696f275a4` | `["Interactive", "Write"]` |
| Microsoft Teams | `connector_246af0940da3457da0e751171dc1ce60` | `["Interactive", "Write"]` |

`C:\Users\admin\.codex\auth.json` holds **only a ChatGPT OAuth session** (`id_token` / `access_token` / `refresh_token` / `account_id`). There is **no Microsoft/Graph token anywhere on this machine.** The connector's OAuth grant to Microsoft Graph is held entirely on OpenAI's connector backend, keyed to the ChatGPT `account_id`. 

**Consequence:** the Graph delegated-permission set these connectors hold (`Mail.ReadWrite` / `Mail.Send` / `Calendars.ReadWrite` / `Chat.ReadWrite` / Planner write, etc. — matching the write surface the Phase 1 audit observed: draft/send/reply/forward/move/categorize, calendar create/update/cancel/RSVP/attach, Teams send channel+chat, Planner update) is fixed by **OpenAI's connector app registration**. It cannot be narrowed from this machine, from `config.toml`, or from Kevin's normal ChatGPT settings. Azure AD delegated-permission consent is all-or-nothing per app — a user consenting to the connector consents to the whole requested set; there is no "consent to the read subset" path for an end user.

## 2. What can actually close the write path, and who must do it

Ranked cleanest-first:

### Option A — Oxford tenant admin restricts the OpenAI enterprise app (true Graph-level fix)
Whoever administers Oxford's Azure AD tenant can, for the OpenAI enterprise application:
- In **Enterprise applications → (OpenAI app) → Permissions**, revoke the admin-consented **write-scoped** Graph delegated permissions (`Mail.ReadWrite`, `Mail.Send`, `Calendars.ReadWrite`, `Chat.ReadWrite`, `Tasks.ReadWrite`) and leave only the `.Read` equivalents, then re-consent to that reduced set. A write call then fails at Graph with `403`/`Authorization_RequestDenied` regardless of any Codex-side approval behaviour.
- **Whether reads still function after that is empirical** — must be tested, because the connector may hard-require its full declared scope set to initialise at all.
- **Prerequisite fact to establish first:** is this connector working against `kevin.lelitte@admin.ox.ac.uk` via *Kevin's own user consent* or via *Oxford tenant admin consent*? Given work-inbox/CLAUDE.md already documents that Oxford blocks the standard Graph consent flow for this account, admin consent is plausible — in which case admin already has the handle to narrow or pull it. **Action: Kevin raises an Oxford IT / IdM request to check the OpenAI app's consent type and current granted scopes for his account, and ask whether write scopes can be revoked while keeping read.** Kevin cannot self-service this.

### Option B — ChatGPT connector settings (check first, 5 minutes, Kevin)
Kevin should open **ChatGPT → Settings → Connectors →** each of the Outlook Email / Outlook Calendar / Teams connectors and look for any control beyond the "Always ask" toggle: a "read-only", "permissions", or per-capability enable/disable option. My local inspection cannot see the ChatGPT web UI, so this must be a manual check. If such a toggle exists it is by far the easiest path. (Note: the "Always ask" toggle is already known **not** to gate headless `codex exec` — that is what the write-gate test proved — so a plain "Always ask" is not sufficient.)

### Option C — disconnect the write-capable connectors from the automation's ChatGPT account
Fully remove the Outlook/Calendar/Teams connectors from the ChatGPT account that the scheduled automation uses, and obtain the read data another way. This defeats the point of Phase 2 unless OpenAI offers a read-only connector variant (none visible in the local marketplace snapshot).

## 3. Local fallback mitigations if A/B/C aren't achievable

Confirmed facts about the local surface:
- `codex exec` has **no CLI flag** that governs MCP/connector-tool approval. `-s`/`--sandbox` governs only model-generated local shell/filesystem. `--approve-for-me` and `--dangerously-bypass-approvals-and-sandbox` both loosen, not tighten. (`codex exec --help`, full text checked.)
- `config.toml` **does** support per-app, per-tool overrides under an `[apps.<connector_id>...]` table — this is the exact structure that was found set to `approval_mode = "approve"` for GitHub write tools in the cc93c7b incident (Section 8). No `[apps.*]` table is present in the current config (confirmed clean).
- Feature flags of note (`codex features list`): `apps` (stable, on), `guardian_approval` (stable, on), `network_proxy` (experimental, off), `request_permissions_tool` / `exec_permission_approvals` / `guardianv2` (all under development, off).

### Fallback 1 (most promising, NOT yet verified) — `approval_mode` deny overrides in `config.toml`
Add an `[apps.<connector_id>...]` block for each of the three connectors setting every **state-changing** tool to a deny/reject value, mirroring (in reverse) the auto-approve override from the incident. 
- **Unverified:** whether Codex's config schema accepts a `"deny"` / `"reject"` / `"never"` value for `approval_mode` (only `"approve"` has been observed in the wild). If it only accepts `"approve"` / `"ask"`, then `"ask"` in a headless `codex exec` session *might* hard-fail the call (no TTY to answer) — but that is exactly the behaviour that did NOT happen in the write-gate test with the default, so it needs a controlled test.
- **Tool names to deny** (partial list from a read-only enumeration probe; the email write tools — draft/send/reply/forward/move/categorize — did not come back in the probe and would need a fuller enumeration):
  - `microsoft_outlook_calendar_create_event`, `..._cancel_or_delete_event`, `..._cancel_or_delete_shared_calendar_event`, `..._add_event_attachment`, `..._add_shared_calendar_event_attachment`, `..._create_contact`, `..._create_contact_folder`
  - `microsoft_teams_send_channel_message`, `microsoft_teams_send_chat_message`, `microsoft_teams_update_planner_task`
  - (email: `microsoft_outlook_email.*` — the draft/send/reply/forward/move/categorize/create-folder tools — names to be fully enumerated before writing the block)
- **Recommended test protocol (discrete, backed-up, reversible):** (1) `cp config.toml config.toml.bak-<ts>-drew-writegate-mitigation`; (2) add the deny block; (3) re-run the exact write-gate test (categorize one throwaway email); (4) confirm via Outlook COM the category did NOT apply; (5) re-run a read pull to confirm reads still work; (6) keep the block if it holds, restore the backup if it breaks reads or doesn't block. This is a config tightening, not a consent/scope change — but given this file's incident history it should be a named, explicitly-authorised step, not folded silently into other work.

### Fallback 2 (needs verification) — disable / remove the write-capable connector plugins
`codex plugin remove outlook-email` (etc.) or `[plugins."outlook-email@..."] enabled = false`. **Unverified whether this removes the underlying `codex_apps/microsoft_outlook_*` tools** — the plugin is a UX/skill wrapper; the tools come from the account-level connected app, so removing the plugin may leave the tools reachable. Needs a test run.

### Fallback 3 — dedicated automation identity
Run the scheduled job under a separate OS user (or separate `CODEX_HOME`) logged into a ChatGPT account that has the connectors connected — but this hits the same scope problem (a fresh connection still grants read-write). Only helps stacked on Fallback 1, or if a read-only connector variant exists.

### Fallback 4 (detection, not prevention — weakest, only as a stacked backstop)
After every scheduled run, a read-only sweep (Outlook COM is available on this machine) compares message categories / flags / read-state / folder membership and the Sent Items + Drafts count against a stored baseline; any delta hard-alerts and disables the scheduled task. This cannot un-send a real send, so it is a backstop on top of a real preventive control, never a substitute.

## 4. Bottom line for the go/no-go

- A genuine read-only re-scope **is not something Kevin can do himself** and is **not adjustable locally**. It requires either an OpenAI-side control that may not exist (Option B — check first) or an **Oxford tenant-admin action on the OpenAI enterprise app's Graph permissions** (Option A — the real structural fix, needs an IT request).
- The best **local** mitigation (Fallback 1, `approval_mode` deny overrides) is **plausible but unproven** and needs a discrete backed-up test before it can be relied on.
- Until one of those is in place and verified, the 6x/day-for-7-days unattended automation stays blocked — nothing currently guarantees a scheduled run cannot write to Kevin's live mailbox/calendar/Teams.
