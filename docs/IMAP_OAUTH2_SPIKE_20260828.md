# IMAP + OAuth2 feasibility spike — RESULT: **PASS**

**Date:** 28 August 2026, ~21:46–21:49 UTC
**Run by:** coordinator (Claude Code), while Drew's session was rate-limited (resets 22:50 Europe/London)
**Context:** After the 28 Aug 13:30 reboot broke both Outlook-COM pipelines (see HANDOVER top entries), Kevin chose "Option B" — get the mail pull off Outlook COM via IMAP+OAuth2 **without involving Oxford IT** — over "Option A" (ask Oxford IT for read-only Graph consent so the ChatGPT connector route becomes safe). Codex second opinion had concluded the connector route is "VIABLE ONLY WITH tenant-enforced read-only Graph consent; otherwise NOT VIABLE" for a mailbox that cannot risk a send.

## What was proven

| # | Question | Result |
|---|---|---|
| 1 | Can a token be obtained with **no Oxford IT / no admin consent / no app registration**? | **YES** — device-code flow, user-consented, using **Thunderbird's public client id**. |
| 2 | Is IMAP enabled for `begb0037@ox.ac.uk` specifically? | **YES** — `SELECT INBOX` returned **558 messages**; a message header was fetched successfully. |
| 3 | Does XOAUTH2 IMAP auth work against Exchange Online? | **YES** — `AUTHENTICATE completed.` |
| 4 | Does the cached token refresh **silently** (reboot / unattended survival)? | **YES** — second run used `acquire_token_silent` with **no prompt**, full read succeeded. |
| 5 | Enough IMAP capability for the pull? | **YES** — `CAPABILITY`: `IMAP4rev1 AUTH=XOAUTH2 SASL-IR UIDPLUS MOVE ID IDLE NAMESPACE LITERAL+`. |

### Evidence (spike log, both runs)
```
21:46:35  token OK. scopes granted: EWS.AccessAsUser.All IMAP.AccessAsUser.All POP.AccessAsUser.All SMTP.Send User.Read
21:46:35  mailbox identity (preferred_username): begb0037@ox.ac.uk
21:46:36  IMAP AUTHENTICATE OK: OK [b'AUTHENTICATE completed.']
21:46:36  SELECT INBOX (readonly) -> OK [b'558']
21:46:36  INBOX message count: 558
21:46:36  latest message headers:
    From: Marie Cooksey <marie.cooksey@admin.ox.ac.uk>
    Subject: POC in my Absence
    Date: Fri, 28 Aug 2026 16:35:29 +0100
21:46:36  === PASS ===
--- second run, fully non-interactive (cached token) ---
21:49:46  cached account found: begb0037@ox.ac.uk -- trying silent token
21:49:46  SILENT token acquisition OK (reboot/refresh survival path works)
21:49:46  SELECT INBOX (readonly) -> OK [b'558']
21:49:47  === PASS ===
```

## Config that works

- **Client id:** `9e5f94bc-e8a4-4e73-b8be-63364c29d753` (Mozilla Thunderbird — preauthorized for the O365 Exchange Online resource, user-consentable).
- **Authority:** `https://login.microsoftonline.com/organizations`
- **Scope requested:** `https://outlook.office365.com/IMAP.AccessAsUser.All` (+ `offline_access` implicit via MSAL)
- **IMAP host:** `outlook.office365.com:993`, SASL `XOAUTH2`, string `user=<upn>\x01auth=Bearer <token>\x01\x01`
- **Token cache:** MSAL `SerializableTokenCache` to a file; `acquire_token_silent` on subsequent runs; survives process/host restart.
- **Library:** `msal` 1.37.0 (already installed), stdlib `imaplib`. Python 3.14.

### What did NOT work
- Microsoft Office first-party client id `d3590ed6-52b3-4102-aeff-aad2292ab01c` → **`AADSTS65002`** ("must be configured via preauthorization") — Exchange Online does not accept IMAP tokens from that client. This is why Thunderbird's id is used.
- The shortlink `https://login.microsoft.com/device` misbehaved once (`AADSTS900561` GET-on-POST-endpoint). Use `https://microsoft.com/devicelogin`.

## Caveats / open decisions for the migration

1. **The granted token bundle includes `SMTP.Send`.** Thunderbird's client requests the whole mail bundle (IMAP+POP+SMTP+EWS); Exchange returns all of it. So the token *technically* can send mail.
   **Why this is still acceptable (unlike the ChatGPT connector route):** there is no autonomous agent with tools on this path. `fetch_inbox.py` would call `imaplib` for reads only and never construct an SMTP session. Nothing in the pipeline can *accidentally* invoke send the way an LLM holding a send-tool can. The connector route's problem was an ungateable tool surface exposed to a model; this route has no such surface.
   **If zero send-capability in the token is nonetheless required:** register a dedicated app (single-tenant, delegated `IMAP.AccessAsUser.All` only). Open question whether the Oxford tenant permits user (non-admin) app registration — check before assuming. Not required to proceed.
2. **Periodic re-auth still happens** (no Primary Refresh Token on this device — see the no-PRT confirmed-fact memory). But with IMAP it surfaces as a clean catchable `invalid_grant` on `acquire_token_silent` → emit a specific toast → Kevin does one device-code sign-in. It is **not** a wedged Outlook GUI. Refresh-token rolling lifetime ~90 days absent a CA sign-in-frequency trigger.
3. **What IMAP does not give us — stays on COM:**
   - **Outlook Categories** — not an IMAP concept. Audit Phase 1 for how much it relies on categories (the pull, the dedup, the card build).
   - **`importance` / high-flag** — derivable from MIME headers (`Importance:`, `X-Priority:`), moderate effort.
   - **`EntryID` + `openmail://` opener** — gone. Switch to an OWA deep-link keyed on `Message-ID` (precedent: command-centre `sourceType=codex-graph`).
   - **Calendar (Phases 3.7 / 3.8)** — no calendar over IMAP at all. A small **calendar-only COM surface remains**. Far less exposed than today's full mail pull, but it means classic Outlook still needs to be runnable for the calendar phases — the WS1 watchdog (`Classic Outlook Keepalive` task) stays relevant.

## Recommendation

Proceed to a **phased migration of the mail-pull half of `fetch_inbox.py` Phase 1 to IMAP+OAuth2.** This removes the OST-mount / classic-Outlook-must-be-running / interactive-sign-in-hang failure mode for the mail pull — the exact thing that broke today. Calendar phases stay on COM for now (smaller, separable follow-up).

## Exact next action (Drew, after session reset)

1. Read this doc + the WS1/WS2/WS3 HANDOVER entries. Fold this into the HANDOVER top entry.
2. Decide client-id strategy: ship on Thunderbird's id (pragmatic, widely used for exactly this) vs. attempt a dedicated single-tenant app registration (check if Oxford allows non-admin registration first).
3. Audit Phase 1's dependence on Outlook **Categories** and **importance**; design the IMAP equivalents.
4. Build behind a flag (`MAIL_BACKEND=com|imap`, `com` default) mirroring the `AI_BACKEND` pattern — cautious-change-pace: parallel-run and diff `data/briefing.json` against a COM run before any cutover.
5. Design the re-auth toast + device-code helper (reuse the WS1 toast rate-limiter).

## Spike artefacts (local, throwaway — not committed)

- `…/scratchpad/imap_spike.py` — the spike script (MSAL device flow + silent path + read-only IMAP test).
- `…/scratchpad/imap_spike_token_cache.bin` — a **live MSAL token cache for `begb0037@ox.ac.uk`** with the mail bundle incl. `SMTP.Send`. **Delete this** once the migration has its own credential handling, or sooner. It is in the session scratchpad, not the repo, not synced.
