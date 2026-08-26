# Oxford IT / IdM request — ChatGPT (OpenAI) Outlook/Calendar/Teams connector: confirm consent type and restrict to read-only scopes

**Prepared for Kevin Lelitte to send / forward to Oxford IT Services (IAM / Microsoft 365 team).**
**Prepared 26 August 2026.**

---

**To:** Oxford University IT Services — Identity & Access Management / Microsoft 365 administration
**From:** Kevin Lelitte, HR Systems (kevin.lelitte@admin.ox.ac.uk)
**Subject:** Request to confirm consent type and restrict a third-party app's Graph permissions to read-only — OpenAI "ChatGPT connectors"

Hello,

I use OpenAI's ChatGPT with its Microsoft 365 "connectors" feature against my Oxford account (`kevin.lelitte@admin.ox.ac.uk`) for read-only inbox and calendar triage. I want to make sure this integration **cannot write** to my mailbox, calendar or Teams — only read. I have two asks:

## 1. Please confirm how these connectors are authorised for my account

The OpenAI ChatGPT connectors in question are three registered applications (OpenAI-operated). In OpenAI's own connector metadata they carry these connector identifiers:

- Outlook Email — `connector_4aaab2856305417b993eca9a216aaf6e`
- Outlook Calendar — `connector_e6a7394682e24467ac68c60696f275a4`
- Microsoft Teams — `connector_246af0940da3457da0e751171dc1ce60`

In Azure AD / Entra ID these will appear as one or more **Enterprise Applications** published by **OpenAI, L.L.C.** (publisher domain `openai.com`). Please tell me:

- Are these apps consented for my account via **my own individual user consent**, or via **tenant-wide admin consent**?
- What **delegated Microsoft Graph permissions** are currently granted to them for my account (e.g. `Mail.Read`, `Mail.ReadWrite`, `Mail.Send`, `Calendars.Read`, `Calendars.ReadWrite`, `Chat.Read`, `Chat.ReadWrite`, `ChannelMessage.Send`, `Tasks.ReadWrite`, `offline_access`, etc.)?
- Is user consent for third-party apps in our tenant generally allowed, restricted to an admin-approved list, or blocked? (This is relevant because the standard Microsoft Graph app-registration consent flow is already blocked for my account under current MDM/app-consent policy — I want to understand whether these connectors are working via an exception, admin consent, or a different path.)

## 2. Please restrict these apps to read-only Graph scopes for my account

I would like the **write-capable** delegated permissions removed and the integration left with read-only access. Specifically, if the following (or equivalents) are granted, please **revoke** them and, where a read equivalent is needed for the connector to function, **re-consent to the read-only scope only**:

| Revoke (write) | Keep / re-consent (read only) |
|---|---|
| `Mail.ReadWrite` | `Mail.Read` |
| `Mail.Send` | (no send at all) |
| `MailboxSettings.ReadWrite` | `MailboxSettings.Read` (if needed) |
| `Calendars.ReadWrite` | `Calendars.Read` |
| `Chat.ReadWrite`, `ChatMessage.Send`, `ChannelMessage.Send` | `Chat.Read` (if needed) |
| `Tasks.ReadWrite` | `Tasks.Read` (if needed) |
| `Contacts.ReadWrite` | `Contacts.Read` (if needed) |

If the app registration is structured so that its delegated permissions cannot be partially granted (i.e. it must be consented all-or-nothing to the full requested set), please tell me — in that case the options are (a) an app-specific Conditional Access / permission restriction limiting it to read scopes, (b) blocking the write scopes at the tenant level for this app, or (c) confirming it simply cannot be made read-only, in which case I will stop using the connector for this purpose.

**Why:** this integration runs partly through automated/headless tooling where the interactive "ask before writing" safety prompt does not reliably fire, so I want the write path closed at the identity layer rather than relying on the application's own approval UI. I am not asking to remove the integration — read access is genuinely useful — only to make it structurally incapable of writing.

Please let me know if you need anything else from me (e.g. a screenshot of the connector settings in ChatGPT, or the exact time window of recent activity) to locate the app grants.

Thanks very much,
Kevin Lelitte
HR Systems

---

## Notes for Kevin (not part of the message)

- Before sending, you may want to do the 5-minute ChatGPT-side check first: **ChatGPT → Settings → Connectors →** each of the three connectors, and look for any "read-only" / per-permission toggle beyond "Always ask". If one exists, that's faster than an IT ticket. The "Always ask" toggle alone is **not** sufficient — testing on 26 Aug 2026 showed a headless `codex exec` write went through with no prompt despite that setting.
- If IT asks "which OpenAI enterprise app object" — there may be a single "OpenAI ChatGPT" enterprise app rather than three separate objects; the three `connector_...` ids are OpenAI's internal connector references, which IT may not see directly. The Azure AD side is more likely to show one OpenAI app with a bundle of Graph delegated permissions. Either way the ask is the same: strip the `.ReadWrite` / `.Send` scopes.
- Full technical background is in `begb0037admin/work-inbox` `docs/codex_phase2_run_20260826/CONNECTOR_WRITE_PATH_INVESTIGATION.md` and `docs/CODEX_CONNECTOR_MIGRATION_RESEARCH.md` Section 9.
