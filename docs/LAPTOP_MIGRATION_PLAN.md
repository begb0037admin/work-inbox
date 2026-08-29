# LAPTOP_MIGRATION_PLAN.md — work-inbox off Outlook COM, two-lane, whole pipeline on Kevin's Oxford laptop

**Date:** 2026-08-29 (Drew). **Rev 2** — folds in Kevin's four decisions (calendar via connector not EWS/Graph; Lane B identity = Oxford ChatGPT Edu stripped to Calendar+Teams; consent resolved; whole pipeline on the laptop).
**Status:** PLAN + Kevin's action list + Phase-1 command sequence + Lane B B1 findings (calendar + Teams). **No build. No cutover. No `.bat` / scheduled-task / `main` default-behaviour change.** Awaiting Phase 1 execution by Kevin, then STOP before Phase 2.
**Companion docs (all on `main`):** `MAIL_BACKEND_MIGRATION_PLAN.md` (the `MAIL_BACKEND=com|imap` flag, already built), `PHASE1_IMAP_MIGRATION_AUDIT.md` (what IMAP breaks), `IMAP_OAUTH2_SPIKE_20260828.md` (mail feasibility PASS), `CONNECTOR_SAFEGUARDS.md` (connector tool enumeration + the Codex "NOT SOUND for an unattended connector fetch that holds write tools" verdict — still applies, see §5f), `CODEX_CONNECTOR_PIPELINE_PLAN.md` + `EMAIL_AUTOMATION_SECURITY_MITIGATIONS.md` (connector history — now scoped to Calendar+Teams, Lane B). **New:** `LANE_B_TEAMS_CAL_DESIGN.md` (the dumb-fetch design, produced alongside this rev — design only).

---

## 0. Shape

| | Lane A | Lane B | Triage | Downstream |
|---|---|---|---|---|
| **What** | **Mail only** — inbox / VIP sweep / 5 subfolder trees / Sent | **Calendar + Teams** — primary + "People Department - HR Systems" shared calendar; Teams chat / channel / in-meeting chat / meeting recordings+transcripts | unchanged | unchanged |
| **How** | our own read-only Python: `imap_mail.pull()` over IMAP+XOAUTH2, `EXAMINE` (read-only). **No send tool exists in the code path.** | ChatGPT M365 connector via `codex exec`, dumb-fetch only. Connector holds write tools while Call-1 runs (Codex "NOT SOUND" caveat) — contained by identity isolation + two kill-switches + re-contamination guard (§5). | `claude -p` headless (Kevin's personal Claude subscription) | `fetch_inbox.py` builds `briefing.json`, pushes to GitHub, syncs command-centre |
| **Auth** | MSAL **broker** (`enable_broker_on_windows=True`, `pymsalruntime`), `acquire_token_silent` off the **laptop PRT** → target zero prompts | `codex login` to the **Oxford ChatGPT Edu** account, stripped to Calendar + Teams connectors only | `claude login` to Kevin's personal account (+ optional kevin/hope failover dirs, §7 note) | `GITHUB_PAT` env var on the laptop |
| **Host** | **Kevin's Oxford laptop, profile `begb0037.AD-OAK`** — the entire pipeline. Desktop retires from work-inbox at cutover. |

Why the laptop: verified this session — `AzureAdJoined YES / DomainJoined YES / **AzureAdPrt YES**` (tenant `cc95de1b-97f5-4f93-b4ba-fe68b852cf91`, PRT auto-renews ~14d on the Oxford network). The admin desktop is only Workplace-Joined with **no PRT** — that refresh churn caused the 28 Aug outage. Laptop stays docked and on.

**Kevin's decisions, 29 Aug (this rev):**
1. Calendar comes through the **connector (Lane B)**, not EWS and not Graph — Oxford does not allow Graph as an auth method, so a Graph self-calendar-read test would hit the same wall. **EWS is removed from the plan entirely.** Lane A is mail-only. Lane B is calendar + Teams. Calendar's blast radius is higher than Teams (a misfire declines / cancels / RSVPs real meetings, fires invites at real attendees) → **its kill-switch HALTS the pipeline on ANY detected calendar change, not just logs.**
2. Lane B identity = **the Oxford ChatGPT Edu account**, becoming the automation's **dedicated** identity **from 1 September**. Kevin moves his interactive AI work off Edu (to personal Plus / Claude) at that point and strips the Edu connectors to **Calendar + Teams only** (all connectors are user-removable). No new/paid account, no failover. The Edu usage limit Kevin hit on 29 Aug was his prior *interactive* use, not the automation — from 1 Sept, Lane B's ~50 light `codex` calls/week don't contend for quota, and because Kevin won't be using Edu interactively the re-contamination risk largely disappears in practice. **The full-manifest auto-disable guard stays mandatory anyway as cheap insurance** (§5e).
3. **Consent — resolved.** User consent already works at `ox.ac.uk` for the Outlook Calendar and Teams connectors (Kevin has run "list my calendar" / "show Teams messages" in ChatGPT Edu — works, no admin prompt). No blocker. Q2 (identity) is fully settled — no paid account, no ongoing quota concern.
4. **Whole pipeline on the laptop** — contingent on Phase 1/2 proving it works. IMAP pull + `codex exec` Lane B + `claude -p` triage + briefing push + command-centre sync, all on `begb0037.AD-OAK`. Desktop retires from work-inbox at cutover.

**Sequencing (coordinator, 29 Aug):**
- **Now (before 1 Sept):** Lane A only — Phase 1 laptop toolchain, then Phase 2(i) the MSAL broker IMAP silent-auth proof. Plus this doc + `LANE_B_TEAMS_CAL_DESIGN.md` as **design**. None of this touches the Edu account or its quota. **`codex login` to Edu is NOT part of Phase 1.**
- **From 1 Sept:** Edu quota resets and Edu becomes dedicated. Kevin strips the Edu connectors to {Calendar, Teams} and runs `codex login` on the laptop. Then Phase 2(ii) (`codex exec` connector-load proof) and the Lane B build increments (`LANE_B_TEAMS_CAL_DESIGN.md` §9) begin.

The desktop pipeline (COM + `claude -p` + `Classic Outlook Keepalive` watchdog) **stays live throughout** as the fallback. Nothing retires until both lanes are proven and Kevin gives a fresh explicit cutover go-ahead (Phase 6).

---

## 1. Host — Kevin's Oxford laptop

| Fact | Value | Verified |
|---|---|---|
| Machine / profile | Oxford laptop, Windows user `ad-oak\begb0037` (the pipeline account — **standard user**, Oxford AD domain account) | Phase 1, Kevin |
| Join state | `AzureAdJoined: YES`, `DomainJoined: YES`, **`AzureAdPrt: YES`** — tenant `cc95de1b-97f5-4f93-b4ba-fe68b852cf91`; PRT auto-renews ~14 days on the Oxford network | Phase 1, Kevin |
| **Admin account split** | Kevin can only elevate as a **separate local admin account `begb0037-a`** (no PRT, not domain-joined). The pipeline account `ad-oak\begb0037` is a standard user **and holds the PRT**. All Phase 1 per-user installs succeeded as `ad-oak\begb0037` with no elevation. **Consequence:** the scheduled task runs as `ad-oak\begb0037`; anything that genuinely needs elevation is a separate manual step done as `begb0037-a` (which would NOT have the PRT — so no auth work runs there). | Phase 1, Kevin |
| Power | stays **docked and on** | Kevin's commitment |
| `%LOCALAPPDATA%` | `C:\Users\begb0037.AD-OAK\AppData\Local` — token caches + toast stamps live here (`WorkInboxAI\`) | Phase 1 |
| PowerShell | 5.1.26100.8875; ExecutionPolicy `RemoteSigned` (CurrentUser) | Phase 1 |

**Execution reality:** Drew runs on the desktop and cannot execute on the laptop. Every laptop step is copy-paste-ready PowerShell 5.1 / a single Python script for Kevin to run and paste back. Kevin is the hands for installs, `claude login`, `codex login`, and the first interactive auth consent.

### Phase 1 — COMPLETE (verified, 29 Aug)

| Component | Installed |
|---|---|
| Python | 3.12.10 (per-user, `%LOCALAPPDATA%\Programs\Python\Python312`) |
| Node / npm | 24.19.0 (Program Files) / 11.17.0 |
| Git | 2.55.0.3 (Program Files) |
| Claude Code | 2.1.251 (`%USERPROFILE%\.local\bin`) — `claude login` done (Kevin's personal account); `claude -p` returns "ready" |
| Codex CLI | **0.151.0** (npm global, `%APPDATA%\npm`) — **no `codex login`** (deferred to 1 Sept). Note: newer than the desktop's 0.149.1; the Lane B design must re-verify tool-gating on 0.151.x (still no `--allowed-tools` as of 0.150.1 per `CONNECTOR_SAFEGUARDS.md` §B6 — re-check). |
| Python pkgs (pip `--user`) | `msal` 1.38.0 + `pymsalruntime` import OK; `pywin32` + `anthropic` import OK |
| Env | `GITHUB_PAT` set (User scope); `ANTHROPIC_API_KEY` unset (User + Machine both blank) → `claude -p` bills the subscription |
| Scripts in `%USERPROFILE%\work-inbox\` | `fetch_inbox.py`, `imap_mail.py`, `reauth_imap.py`, `diff_mail_pull.py` — all `py_compile` exit 0 |

### Phase 2(i) — PASSED (29 Aug)

`broker_imap_proof.py` v2, two runs on the laptop:
- **Broker interactive is NOT available for the Thunderbird client id** — `enable_broker_on_windows=True` + `acquire_token_interactive` returned `broker_error / Status_ApiContractViolation` instantly, twice (with `CONSOLE_WINDOW_HANDLE` and with a real HWND). The Thunderbird public client `9e5f94bc-…` has no WAM/broker redirect registered. **We keep this client** — it is the only one that gets `IMAP.AccessAsUser.All` at Oxford.
- **PATH B works:** plain `acquire_token_interactive` (system browser, no broker) → **one SSO account click, no password, no MFA** (~22s) on the PRT-joined laptop. Token carried `IMAP.AccessAsUser.All`; cache persisted (9298 bytes); IMAP `EXAMINE INBOX` → 558 messages.
- **Run 2 (cold) = FULL WIN:** `silent[broker]: SILENT token OK` — the broker CAN serve a silent token from the browser-seeded file cache. `auth_path: silent(broker-app)`, IMAP OK, `=== PASS ===`.

**Operational reality:** day-to-day scheduled runs are **fully silent** (`acquire_token_silent`, broker-app path). The **first-time seed and the periodic re-auth** (CA sign-in-frequency / ~90-day refresh-token roll, weeks apart) are **one system-browser SSO click on the laptop — no password, no MFA**, via `reauth_imap.py`. Not in the scheduled path.

### Phase 3 — Lane A code (commits `0900b3f` + `d5447b9`, on `main` behind the unset flag — NOT cut over)

- **`imap_mail.py`** — `acquire_token_silent()` now tries broker-app silent then plain-app silent off the shared cache (`%LOCALAPPDATA%\WorkInboxAI\msal_imap_token_cache.bin`); `ImapReauthRequired` with the combined error otherwise. `_broker_app()` returns `None` when `msal[broker]`/`pymsalruntime` is absent → the admin desktop is plain-app-only, unchanged.
- **`reauth_imap.py`** — default is now plain system-browser `acquire_token_interactive` (the proven PATH B); `--device-code` kept as an any-other-device fallback.
- **`fetch_inbox.py`** — `win32com`/`pywintypes`/`anthropic` imports guarded (byte-identical where installed = both current machines); `_COM_ERROR` alias keeps the `except` sites valid COM-free. **`CAL_BACKEND=com|connector`** flag added, default `com`. **`com` / default behaviour unchanged.**
- **`fetch_inbox.py` (`d5447b9`) — `MAIL_BACKEND=imap` must never open classic Outlook.** The first laptop parallel run (29 Aug) auto-launched `OUTLOOK.EXE` because the COM connect ran unconditionally at Phase 1 start, before the `WI_MAIL_PARALLEL` skip-exit. Fixed: under `imap`, COM is attempted **only** when a COM calendar source is genuinely in scope (`CAL_BACKEND=com` AND not `WI_MAIL_PARALLEL` AND `CAL_BACKEND` not requested as `connector`), and even then with **`connect_to_outlook(allow_launch=False)`** — a missing/not-connected classic Outlook degrades to *empty calendar + warning*, never auto-starts the app. The "open Outlook" failure toast is gated on `allow_launch` too. `CAL_BACKEND=connector` is recorded as `CAL_CONNECTOR_NYI` (calendar empty + warning, no COM) rather than coerced to `com`. `com`/default path byte-identical (`allow_launch` defaults `True`; every new branch gated on `not allow_launch` / `CAL_CONNECTOR_NYI`).
- **First laptop parallel capture (29 Aug, `WI_MAIL_PARALLEL=1`):** `IMAP - silent OAuth2 token OK ... via broker-app` → INBOX 38 → +VIP 0 → +subfolders 10 → Sent 10 → `IMAP mail pull: inbox 48 (unread 19) sent 10` → `imap_inbox_raw.json (48)` / `imap_sent_raw.json (10)` → `Exiting 0`. `INBOX/Bi-monthly CDR/PD working group` skipped (`/` in name — §4b, unresolved/accepted). **Re-run needed after `d5447b9` to confirm no Outlook launch.**
- **Still pending in Phase 3:** the dashboard JS `mail_backend==="imap"` → OWA-opener branch — needs a screenshot for Kevin (command-centre-style UI gate) before it ships.
- **Phase 3 re-run (29 Aug, after `d5447b9`) — CLEAN:** `Phase 1 - MAIL_BACKEND=imap: NOT connecting Outlook COM (WI_MAIL_PARALLEL mail-only capture); classic Outlook will not be opened` → `silent OAuth2 token OK ... via broker-app` → `inbox 48 (unread 19) sent 10` → `Exiting 0`. `Get-Process OUTLOOK` → nothing. **No Outlook launched.**

### Phase 5 — mail parity (`parity_vs_briefing.py`, commit `4a7ce21`)

- **Strict same-window field parity was already PROVEN on the admin desktop** (29 Aug, `diff_mail_pull.py`): INBOX common 48/52, SENT 10==10, **REAL parity issues 0** (+31 benign X.500→SMTP, +5 read-cap churn). The Phase 3 code changes (broker auth, guarded imports, no-launch) do not touch `imap_mail.pull()`'s message logic, so that result stands. A fresh strict re-confirm would need a desktop COM capture (`Run Mail Parity Test.bat`) — optional, not blocking.
- **`parity_vs_briefing.py`** is the **self-contained laptop** check for ongoing confidence across inbox states: pulls the live desktop `data/briefing.json` (+ optionally last N commits) from GitHub, runs a fresh `MAIL_BACKEND=imap WI_MAIL_PARALLEL=1` capture, and checks the IMAP pull **surfaces the same messages, attributed the same way**. `briefing.json` is a triaged artifact (no `message_id`, no per-card `is_read`/`importance`/`has_attachments`, sender is a display name) and a snapshot from an earlier run, so this is a **coverage + attribution sanity check**, not a byte diff — drift (new mail since, items filed/read) is reported as *expected*, not failure.
  - **Real flags:** `only_in_briefing` in `needs`/`urgent` (COM surfaced it, IMAP missed it) · `only_in_imap` unread `needs`/`urgent` the briefing lacked · `kevin_is_primary_recipient` mismatch on a matched pair.
  - **Soft / expected (reported, not counted):** `only_in_briefing` `fyi`/`low` predating the snapshot · `only_in_imap` arrived-after-snapshot · grouped-thread siblings of a matched `fyi` card · read-cap boundary churn · derived-tier differences (the script uses `diff_mail_pull._tier()`; `briefing.json` uses the full `categorise()`).
  - Also folds in a **folder diagnostic** (`NAMESPACE` + `LIST` rows near the CDR folder) — see §4b.
- **Run:** `python parity_vs_briefing.py` once now, then ~once a day for 3–4 days; Kevin/Lauren eyeball. Command in §7 "Phase 5 parity".

---

## 2. Target architecture

```
              KEVIN'S OXFORD LAPTOP  (begb0037.AD-OAK, PRT present) — the WHOLE pipeline
  ┌──────────────────────────────────────────────────────────────────────────────────────┐
  │                                                                                      │
  │  LANE A — MAIL ONLY  (SAFE: our own read-only Python, no send tool in the path)       │
  │  ┌────────────────────────────────────────────────────────────────────────────────┐  │
  │  │ MSAL broker (enable_broker_on_windows=True, pymsalruntime)                      │  │
  │  │   acquire_token_silent → off the laptop PRT, ZERO prompts                       │  │
  │  │        │ scope: https://outlook.office365.com/IMAP.AccessAsUser.All             │  │
  │  │        ▼                                                                        │  │
  │  │   imap_mail.pull()  →  outlook.office365.com:993, SASL XOAUTH2, EXAMINE (RO)    │  │
  │  │   INBOX + VIP sweep + 5 subfolder trees + Sent  →  exact Phase-1 dict shape     │  │
  │  └───────────────┬────────────────────────────────────────────────────────────────┘  │
  │                  │ inbox / sent lists (sanitised at source: preview truncation)      │
  │                  ▼                                                                    │
  │  LANE B — CALENDAR + TEAMS  (eyes-open residual, Kevin accepted; connector via codex) │
  │  ┌────────────────────────────────────────────────────────────────────────────────┐  │
  │  │ codex login → Oxford ChatGPT Edu account, connectors = {Calendar, Teams} ONLY  │  │
  │  │ dedicated CODEX_HOME.  Call 1 (connector attached): rigid list/read only —      │  │
  │  │   calendar: list_calendars → list_events (window) → fetch_events_batch         │  │
  │  │   teams: list_chats/list_channels → list_*_messages / transcripts             │  │
  │  │   NO reasoning over content. Fixed field list. Hard output cap.                │  │
  │  │        │                                                                       │  │
  │  │        ▼ normalise_pull.py (deterministic sanitiser — HTML strip, truncate,    │  │
  │  │          neutralise instruction-like text, strip zero-width/bidi, record hits) │  │
  │  │        │                                                                       │  │
  │  │   ┌────┴───────────────────────┐   ┌──────────────────────────────────────┐    │  │
  │  │   │ CALENDAR kill-switch       │   │ TEAMS kill-switch                     │    │  │
  │  │   │ pre/post snapshot of the   │   │ pre/post "messages from = me" in the │    │  │
  │  │   │ calendar window. ANY diff  │   │ window. New one → disable NEXT run + │    │  │
  │  │   │ (new/missing/modified/RSVP │   │ toast (blast radius = 1 Teams msg).  │    │  │
  │  │   │ /cancel) → HALT: disable   │   └──────────────────────────────────────┘    │  │
  │  │   │ task + exit(1) + toast +   │   ┌──────────────────────────────────────┐    │  │
  │  │   │ GUARD_TRIPPED. Manual      │   │ RE-CONTAMINATION guard (mandatory):  │    │  │
  │  │   │ re-enable only.            │   │ full tool manifest logged every run. │    │  │
  │  │   └───────────────────────────┘   │ ANY tool outside {calendar-read,     │    │  │
  │  │                                   │ teams-read} available → HALT + toast. │    │  │
  │  │                                   └──────────────────────────────────────┘    │  │
  │  └───────────────┬────────────────────────────────────────────────────────────────┘  │
  │                  │ calendar events + teams messages/transcripts (sanitised)          │
  │                  ▼                                                                    │
  │        ┌──────────────────────────────────────────────────────────────┐              │
  │        │  claude -p  (headless Claude Code, Kevin's PERSONAL sub)      │  ← triage    │
  │        │  UNCHANGED. Consumes Lane A + Lane B after sanitisation.      │              │
  │        │  ANTHROPIC_API_KEY must be UNSET (subscription billing).      │              │
  │        └───────────────────────────┬──────────────────────────────────┘              │
  │                                    ▼                                                  │
  │           fetch_inbox.py → data/briefing.json → GitHub (GITHUB_PAT) → dashboard       │
  │                            + command-centre sync (Phases 3.5/3.6)                     │
  └──────────────────────────────────────────────────────────────────────────────────────┘
```

Everything downstream of `briefing.json` (dashboard, command-centre sync, `data/ticks.json`, Lauren's drafting loop) is **unchanged** — it consumes the JSON regardless of what produced it.

---

## 3. What already exists on `main` (do not rebuild)

Behind the **unset** `MAIL_BACKEND` flag (`com` default = byte-identical to today):

| File | State | Gap for the laptop |
|---|---|---|
| `imap_mail.py` | `pull()`, `_pull_sent()`, `message_still_in_inbox()`, `acquire_token_silent()`, OWA-search `web_link`, mUTF-7 subfolder resolution, `Importance:`/`X-Priority:` → 0/1/2, `_kevin_is_primary_recipient()`. Parity harness fixes (**Message-ID join, Sent meeting-response filter, subfolder LIST — DONE**, per the 29 Aug ~01:15 HANDOVER: last diff INBOX common 48/52, **SENT 10==10, real parity issues 0**). | Auth is **device-code only** (`authority=.../organizations`, Thunderbird client `9e5f94bc-…`, scope `IMAP.AccessAsUser.All`, cache `%LOCALAPPDATA%\WorkInboxAI\msal_imap_token_cache.bin`). Needs a **broker path** for the PRT machine (Phase 2). |
| `reauth_imap.py` + `docs/desktop-scripts/Re-auth Work Inbox IMAP.bat` | device-code priming + read-only `EXAMINE INBOX` verify + timestamps | becomes the **fallback** only; broker makes the happy path promptless |
| `diff_mail_pull.py` | field-by-field COM↔IMAP parity diff → `data/parallel/parity_<ts>.json` | reused as-is for Phase 5 (mail parity) |
| `fetch_inbox.py` | `MAIL_BACKEND=com\|imap` + `WI_MAIL_PARALLEL=1`; 4 COM mail loops guarded `for X in ([] if MAIL_BACKEND=="imap" else <orig>)`; `connect_to_outlook()` non-fatal under `imap`; `mapi is None` calendar guard; `_imap_reauth_toast_due()` 1/hr stamp. `com` path proven byte-identical. **Restore point:** `main` `9a52b07`, blob `bd02b41089850678b8268318a0afab5e6d457e8a`, snapshot `Archive/fetch_inbox_backup_20260828_*_pre_mail_backend_flag.py`. | top-of-file `import win32com.client` / `pywintypes` / `anthropic` are **unconditional** — Phase 3 makes them lazy; Phase 1 installs them anyway. **The calendar phases (3.7/3.8) still call COM** — Phase 3 re-points them at a Lane B calendar-source file behind a `CAL_BACKEND=com\|connector` flag. |
| `tools/codex_triage/` (branch `drew/codex-phase2-ai-triage`) | `normalise_pull.py` skeleton, `categorise_and_stage.py`, `build_granola_context.py`, `build_call2_brief.py`, `mailbox_guard.py` (COM sweep — **not reused**), the 26-Aug dry-run machinery | the sanitiser + Call-1 runner get reused/adapted for Lane B calendar+Teams; the COM `mailbox_guard.py` is replaced by the two connector kill-switches (§5d) |

**Genuinely new work:** broker auth in `imap_mail.py`; `CAL_BACKEND` flag + a Lane B calendar-source adapter feeding Phases 3.7/3.8; lazy COM imports; Lane B in full (Call-1 runner for calendar+Teams, sanitiser, the two kill-switches, re-contamination guard, the parallel scheduled task); laptop toolchain + scheduled task + `powercfg`; `claude -p` failover replication.

---

## 4. Lane A design detail — MAIL ONLY

### 4a. Auth — MSAL broker, silent off the PRT

- `pip install "msal[broker]"` (pulls `pymsalruntime`).
- `PublicClientApplication(CLIENT_ID, authority="https://login.microsoftonline.com/organizations", enable_broker_on_windows=True, token_cache=<serializable cache at %LOCALAPPDATA%\WorkInboxAI\msal_imap_token_cache.bin>)`.
- **First run:** `acquire_token_interactive(SCOPES, parent_window_handle=PublicClientApplication.CONSOLE_WINDOW_HANDLE)` → on a PRT machine this is a **WAM broker** flow: no password, no device code, at most one consent click, often zero.
- **Every subsequent run:** `acquire_token_silent(SCOPES, account=accounts[0])` → served from the broker/PRT, no prompt, survives reboot; with the laptop's ~14-day PRT auto-renew it should not surface interactive re-auth the way the no-PRT desktop did.
- Scope: `["https://outlook.office365.com/IMAP.AccessAsUser.All"]` only. Client id stays Thunderbird's public `9e5f94bc-e8a4-4e73-b8be-63364c29d753` (pre-authorised for Exchange Online; the 28 Aug spike proved silent IMAP against it). **No calendar scope on this token — calendar is Lane B now.**
- Failure path unchanged: silent failure → `ImapReauthRequired` → one 1/hour toast → `SystemExit(1)`. Never an interactive hang in the scheduled context. Fallback is `reauth_imap.py` (device-code).

### 4b. Lane A residuals (from `PHASE1_IMAP_MIGRATION_AUDIT.md`, still true)

- **`openmail://` opener dies** under IMAP → replaced by an OWA search deep-link on `Message-ID` (`imap_mail._owa_search_link()`, already written). **Dashboard JS branch not written** — until it is, IMAP cards have a dead opener. Phase 3 deliverable; screenshot for Kevin (command-centre-style UI gate).
- **Phase 3.9** (inbox-resolution tracking) keys on Outlook `EntryID`; under IMAP it degrades to fail-open-carry unless re-wired to `message_id` + `imap_mail.message_still_in_inbox()`. Open decision: re-wire before cutover, or accept fail-open-carry for week 1.
- **`SMTP.Send` in the token bundle** — Thunderbird's client returns the whole mail bundle. Mitigation is architectural and unchanged: `imap_mail.py` imports `imaplib` only, never `smtplib`; no agent-with-tools on this path. (This is the exact property that makes Lane A "safe" and Lane B "eyes-open".)
- **Outlook Categories** — zero dependence confirmed by full-repo grep; nothing to replace.
- **Subfolder `INBOX/Bi-monthly CDR/PD working group` — CUTOVER BLOCKER, fix scoped (diagnostic-gated).** The `/` in the Outlook folder name is also the Exchange Online IMAP hierarchy separator. `imap_mail.pull()` builds `target = "INBOX/" + tree` and requires `list_name == target` (or a child); the current skip means either (a) the server nests it as a 3-segment path with no intermediate `Bi-monthly CDR` folder so `SELECT` fails, or (b) the server substitutes the `/` in the returned `LIST` name. **Cannot fix blind.** `parity_vs_briefing.py`'s folder diagnostic prints `NAMESPACE` + every `LIST` row containing `cdr`/`working group`/`bi-monthly` + all `INBOX/` children. Once we see the real server name: change `imap_mail.pull()`'s subfolder matching so a `tree` containing `/` matches any `LIST` entry whose name — with `/` and `&-` stripped and lowercased — contains the tree's normalised form, then `SELECT` that entry's **exact server string verbatim** (handles both (a) and (b) without guessing). Must be closed and re-verified before cutover — not carried.
- **Classic Outlook on the laptop** — IS installed + configured (Oxford standard image) and works, BUT the pipeline must never depend on it or launch it. `MAIL_BACKEND=imap` now never opens it (fix in commit `d5447b9`). It is uninstalled at cutover per §6 "Cutover"; until then it is simply left untouched.

---

## 5. Lane B design detail — CALENDAR + TEAMS (design only; full design in `LANE_B_TEAMS_CAL_DESIGN.md`)

### 5a. Identity — the Oxford ChatGPT Edu account, dedicated from 1 September

- From **1 Sept**, `codex login` on the laptop targets **Kevin's Oxford ChatGPT Edu** account, which becomes the automation's **dedicated** identity — Kevin moves his interactive AI work to personal Plus / Claude.
- **Kevin removes GitHub + Outlook Email connectors himself at the 1-Sept kickoff**, before `codex login` (action list §6). Target end state: connectors = **{Outlook Calendar, Microsoft Teams}** and nothing else.
- Single account, **no failover**. Lane B volume ≈ 50 light fetch calls/week; from 1 Sept it does not contend with Kevin's interactive quota (he's off Edu). A failed run degrades calendar/Teams that cycle like any outage — acceptable.
- Kevin won't use Edu interactively → the re-contamination risk is small in practice, but the **manifest guard (§5e) stays mandatory as cheap insurance** (a connector update, or a future account change, could still widen the surface).

### 5b. B1 findings — connector tool manifests (live, read-only)

Source: `~/.codex/.codex-global-state.json` → `electron-persisted-atom-state` → `mcp-extension-sidebar-catalog` → `codex_apps` server (292 tools merged), read 2026-08-29 on the desktop. `~/.codex/config.toml` sha1 `35f8910382373d525598194b2649159cfeed3f6a` recorded before and after — unchanged; plain file read, no `codex exec`, no `codex login`.

**`microsoft_outlook_calendar` — 34 tools, 16 write / 18 read.**

| | Tools |
|---|---|
| **Read — Lane B Call-1 allowlist (calendar):** | `list_calendars` (incl. shared calendars → covers "People Department - HR Systems"), `list_events` (date range — the Phase 3.7 pull), `fetch_event`, `fetch_events_batch`, `list_event_instances` (recurring occurrences/exceptions), `list_recurring_series`, `get_mailbox_settings` (preferred timezone). Optionally `search_events`. |
| Read, not needed | `get_schedule`, `find_available_slots`, `get_profile`, all `*_contact*` / `*_contact_folder*` reads, `search_people`, `search_directory_users`, `search_mailbox_contacts` |
| **Write — NEVER in the allowlist; ANY appearance → HALT (§5d):** | **Tier-1 external-comms:** `create_event`, `create_shared_calendar_event`, `update_event`, `update_shared_calendar_event`, `cancel_or_delete_event`, `cancel_or_delete_shared_calendar_event`, `respond_to_event`, `respond_to_shared_calendar_event` (these fire invites / updates / cancellations / RSVPs at real attendees — the higher blast radius Kevin flagged). Plus `add_event_attachment`, `add_shared_calendar_event_attachment`, and 6 `*_contact*` mutations. |

**`microsoft_teams` — 33 tools, 9 write / 24 read.**

| | Tools |
|---|---|
| **Read — Lane B Call-1 allowlist (teams):** | `list_teams`, `list_channels`, `list_chats`, `resolve_team`, `resolve_channel`, `resolve_chat`, `list_chat_messages`, `list_channel_messages`, `fetch` (one message/conversation by path), `search`, `resolve_scheduled_online_meeting` (from an event ID / join URL — keyed off the Lane B calendar pull), `list_online_meeting_transcripts`, `get_online_meeting_transcript_content` (size-bounded WebVTT). Optionally `resolve_user`, `get_chat_members`, `list_online_meeting_recordings`. |
| Read, not needed | `get_profile`, all `*_planner_*` reads, `get_online_meeting_recording`, `get_online_meeting_transcript`, `validate_write_target` |
| **Write — NEVER in the allowlist; ANY appearance → disable NEXT run + toast (§5d):** | `send_chat_message`, `send_channel_message`, `reply_to_message`, `reply_to_channel_message`, `create_chat`, `create_channel`, `create_planner_task`, `update_planner_task`, `delete_planner_task`. |

Constraint: recordings/transcripts are **scheduled meetings only** and must be resolved from an event ID / join URL — so Lane B's transcript pull chains off the Lane B calendar pull (event id → `resolve_scheduled_online_meeting` → transcript).

### 5c. B1 consent — RESOLVED (Kevin, 29 Aug)

User consent already works at `ox.ac.uk` for the Outlook Calendar and Teams connectors — Kevin has exercised both interactively in ChatGPT Edu with no admin prompt. **No tenant-admin-consent blocker.** (This is the ChatGPT *app*; whether headless `codex exec` on the laptop loads the same connector tools is the separate Phase-2 make-or-break — §5f, §6.)

### 5d. The two kill-switches — deliberately asymmetric

| | **Calendar kill-switch — HALT on ANY change** | **Teams kill-switch — disable next run** |
|---|---|---|
| Blast radius it guards | Tier-1 calendar writes: a stray `respond_to_event` / `cancel_or_delete_event` / `update_event` sends decline/cancel/reschedule notices to real attendees | a stray `send_chat_message` / `reply_to_*`: one Teams message |
| Pre-run baseline | `list_events` over the working window (today + 6 days) on **primary + the HR Systems shared calendar**, snapshot per event `{id, subject, start, end, lastModifiedDateTime, isCancelled, responseStatus, seriesMasterId}` → `data/codex_runs/cal_baseline_<ts>.json`. Taken **seconds before** Call-1. | `list_chats` + touched-channel `list_channel_messages` filtered `from = me` in a trailing window → `{message id set}` → `data/codex_runs/teams_baseline_<ts>.json` |
| Post-run check | repeat the exact same read **seconds after** Call-1 + sanitiser finish. **Trip = ANY event where post ≠ pre**: new id, missing id, changed `lastModifiedDateTime`, changed `responseStatus`, `isCancelled` flip. | repeat; trip = any `from = me` message id present post- not pre- |
| On trip | **HALT:** `Disable-ScheduledTask` on the Lane B task **AND** the whole pipeline run `exit(1)` (no `briefing.json` this cycle) **AND** BurntToast **AND** `data/codex_runs/GUARD_TRIPPED_cal_<ts>.json` with the diff. **Manual investigation + explicit re-enable by Kevin only.** | `Disable-ScheduledTask` + BurntToast + `GUARD_TRIPPED_teams_<ts>.json`. Next run is blocked; the current briefing still completes. |
| Known residual | The pre/post window is ~2–5 min. A **genuine third-party edit** to an event in that window (a colleague reschedules a meeting) also trips it → a false HALT; Kevin eyeballs and re-enables. Accepted: calendar blast radius justifies favouring safety over uptime, per Kevin's instruction. | net-zero (delete+resend), delayed send, on-behalf/shared Sent — documented blind spots; containment not prevention |
| Identity | the kill-switch reads run in their **own** short `codex exec` with the same {Calendar, Teams} connector, read tools only, manifest-asserted (§5e) | same |

### 5e. Re-contamination guard — MANDATORY (Kevin, 29 Aug)

Because Lane B rides Kevin's interactive Edu account, a connector he re-adds (or a connector update) must not silently widen the surface.

- **Every Lane B run** (Call-1 and each kill-switch session) begins with a tool-manifest enumeration turn ("list your tools, call nothing"), parsed from the JSONL → `data/codex_runs/<ts>_manifest.json`.
- **Assert the available tool set is a subset of** `{ the calendar-read allowlist } ∪ { the teams-read allowlist }`. i.e. zero `github.*`, zero `microsoft_outlook_email.*`, zero `canva.*` / `sites.*` / `granola.*`, **and zero write tools** from calendar or teams.
- **Any violation → HALT** (disable the Lane B task + `exit(1)` + BurntToast "work-inbox Lane B: unexpected connector tool `<name>` — pipeline disabled" + `GUARD_TRIPPED_manifest_<ts>.json`). Fail-closed.
- Weekly rollup: manifest histogram, any drift.
- **Also asserts the connector read tools we need ARE present** — if `list_events` / `list_chats` are absent, the connector didn't load into `codex exec` (the 28 Aug Q2 state) → "connector not available", no run this cycle, toast. This is the §5f risk in operational form.

### 5f. Codex "NOT SOUND" caveat still applies — what Lane B relies on instead

`CONNECTOR_SAFEGUARDS.md` §D: Call-1 holds the connector's write tools the whole time it runs; a rigid prompt is not a capability boundary; a single `codex exec` invocation can make a write tool call after hostile content is in context and before it returns its JSON; kill-switches detect after, not prevent. **That objection is not resolved here.** Lane B proceeds anyway because:
1. **Narrower blast radius than mail** — worst case is one Teams message or one calendar-notice to attendees, not an external HR email. Kevin has explicitly accepted this.
2. **Dedicated 2-connector identity** — no GitHub (89 tools, 41 write), no Outlook Email send tools on the account.
3. **Calendar kill-switch HALTS** (not just logs) on any change — the strictest containment available without an enforced boundary.
4. **Re-contamination guard** — the surface can't silently widen.
5. **No other route to calendar or Teams** exists at Oxford (Graph disallowed, EWS retiring, Teams has never had a non-connector route).
6. **`claude -p` triage (Call-2 equivalent) has no connector** — the reasoning step that ingests hostile content cannot act.

### 5g. `codex exec` connector-loading — the OTHER make-or-break

28 Aug Q2 finding: a plain headless `codex exec` on the **desktop** loaded **zero** connector tools (only `functions.*` + `collaboration.*`; cause undetermined — connector auth / ChatGPT app-server bridge / residual state). Kevin's "tested working" is the ChatGPT **app**, interactive — not `codex exec`. So Phase 2 must also prove: **headless `codex exec` on the laptop, logged into the stripped Edu account, loads the Calendar + Teams read tools and nothing else.** If it does not, Lane B has no route and calendar/Teams stay out of scope until it does. Co-equal with the IMAP broker proof.

---

## 6. Phases — increment and report after each; do NOT run to completion silently

### Kevin's action list — NOW

- Run the Phase 1 command sequence (§7) on the laptop and paste all output back. (No Edu-account interaction in Phase 1.)

### Kevin's action list — FROM 1 SEPT (Lane B kickoff)

- **(a)** Move interactive AI work off the Edu account (to personal Plus / Claude).
- **(b)** In ChatGPT Edu → Settings → Connectors: **remove the GitHub connector and the Outlook Email connector.** Leave Outlook Calendar + Microsoft Teams. Confirm back what remains.
- **(c)** On the laptop: `codex login` targeting the Edu account (§7 block 1.9, deferred).

### Lane A — NOW

| # | Phase | Output | Gate |
|---|---|---|---|
| **1** | Laptop toolchain | **DONE — §1 "Phase 1 COMPLETE"** | — |
| **2(i)** | MAKE-OR-BREAK #1 — silent IMAP token off the PRT | **DONE / PASS — §1 "Phase 2(i) PASSED"** (`broker_imap_proof.py` v2; cold run 2 = silent FULL WIN; first-time/periodic seed = one browser click) | — |
| **3** | Wire Lane A in | **Code DONE — `0900b3f` + `d5447b9`.** Re-run 29 Aug CLEAN (no Outlook launched, inbox 48 / sent 10, `Exiting 0`). **Pending: dashboard JS `mail_backend==="imap"` → OWA opener branch — screenshot for Kevin before it ships.** | screenshot approval for the JS branch |
| **4** | Scheduled task(s) on the laptop as **`ad-oak\begb0037`** (the PRT-holding standard user — **not** `begb0037-a`), 5×/weekday matching the current cadence (confirm against the live desktop `\Work Inbox Briefing` task); `powercfg` never-sleep; every run timestamped. **Parallel — writes `docs/*` / `data/parallel/*` / `data/codex_runs/*` only, never `data/briefing.json`.** Phase 2(i) settled the WAM question: day-to-day silent runs are fine unattended; the periodic re-auth needs a logged-in session for the one browser click (laptop stays logged in — docked + on). | task XML + `powercfg` | Kevin go-ahead |
| **5** | **Mail parity.** Strict field parity already PROVEN on the desktop (`diff_mail_pull.py`, 0 real issues, 29 Aug). `parity_vs_briefing.py` (`4a7ce21`) = self-contained laptop coverage check vs live `data/briefing.json` — Kevin runs it now + ~daily for 3–4 days, Kevin/Lauren eyeball. **Cutover blocker surfaced: the `Bi-monthly CDR/PD working group` `/`-in-name subfolder gap — fix scoped (§4b), diagnostic-gated.** | 3–4 clean daily runs (0 real flags) + the CDR subfolder fix landed & re-verified |

### Lane B — FROM 1 SEPT

| # | Phase | Output | Gate |
|---|---|---|---|
| **B1** | Connector tool enumeration — calendar + Teams | **DONE — §5b** | — |
| **B2** | Identity + consent | **DONE — §5a / §5c** (Edu dedicated from 1 Sept, stripped to Calendar+Teams; user consent already works; no quota contention) | — |
| **B3** | Dumb-fetch design — Call-1 rigid read + allowlists, `normalise_pull.py` sanitiser, the two kill-switches (calendar = HALT-on-any-change), re-contamination guard — **design doc, not build** | **DONE — `docs/LANE_B_TEAMS_CAL_DESIGN.md`** | Kevin approves the Lane B design before build |
| **2(ii)** | **MAKE-OR-BREAK #2** (from 1 Sept, after `codex login`). Headless `codex exec` on the laptop (stripped Edu account) loads exactly the {Calendar, Teams} read tools — `list_events` returns real events, `list_chats` returns real chats, manifest has **no** other connector / no write tools. | proof log | **STOP + report.** If it fails, Lane B has no route — calendar/Teams stay out of scope. |
| **B4** | Build Lane B per `LANE_B_TEAMS_CAL_DESIGN.md` §9 behind the parallel task (writes `docs/codex_*.json` / `data/codex_runs/*` only) | — | Kevin's fresh explicit separate go-ahead |

### Cutover — Phase 6, separately gated

Both lanes proven → **Kevin's fresh explicit go-ahead for the cutover step specifically** → flip flags → retire, in order: desktop `\Work Inbox Briefing` task; classic Outlook (uninstall); `Classic Outlook Keepalive` task + `Ensure-ClassicOutlook.ps1` preflight; the COM calendar pull; `openmail://` opener → OWA deep-link; **the desktop leaves work-inbox entirely.** `claude -p` + the desktop pipeline stay live until this step.

---

## 7. Phase 1 — exact command sequence for Kevin (PowerShell 5.1, on the laptop)

> Run each numbered block, **paste the full output back to Drew** before moving on. `;` not `&&`. `&` call-operator for quoted paths. Where a block says "NEW SHELL", close the window and open a fresh PowerShell 5.1 so PATH updates apply. Nothing here touches the live desktop pipeline — it is all additive on a fresh machine.

```powershell
# ============================================================================
#  work-inbox laptop migration - PHASE 1 (toolchain).  Laptop, profile
#  begb0037.AD-OAK.  Start in an ELEVATED Windows PowerShell 5.1 for 1.0-1.2.
# ============================================================================

# --- 1.0  Confirm we are on the right machine ---
$env:COMPUTERNAME
whoami
$PSVersionTable.PSVersion
dsregcmd /status | Select-String 'AzureAdJoined|DomainJoined|AzureAdPrt|TenantId'

# --- 1.1  winget present? (if this errors, install "App Installer" from the
#          Microsoft Store first, then re-run) ---
winget --version

# --- 1.2  Install Python 3.12, Node LTS, Git  (accept the UAC prompts) ---
winget install --id Python.Python.3.12 -e --source winget --accept-package-agreements --accept-source-agreements
winget install --id OpenJS.NodeJS.LTS   -e --source winget --accept-package-agreements --accept-source-agreements
winget install --id Git.Git             -e --source winget --accept-package-agreements --accept-source-agreements

# =====>  NEW SHELL: close this window, open a fresh NON-elevated PowerShell 5.1  <=====

# --- 1.3  Verify the toolchain ---
python --version
python -m pip --version
node --version
npm --version
git --version

# --- 1.4  Python packages for Lane A (broker auth + pipeline imports) ---
python -m pip install --upgrade pip
python -m pip install "msal[broker]" pywin32 anthropic
python -c "import msal, pymsalruntime; print('msal', msal.__version__, 'pymsalruntime OK')"
python -c "import win32com.client, anthropic; print('pywin32 + anthropic import OK')"

# --- 1.5  Install Claude Code (native installer - same method as the desktop) ---
irm https://claude.ai/install.ps1 | iex
#   (fallback if blocked:  npm install -g @anthropic-ai/claude-code )

# --- 1.6  Install Codex CLI (install ONLY - do NOT 'codex login' yet; the Edu
#          account is not the dedicated automation identity until 1 Sept) ---
npm install -g @openai/codex
codex --version

# =====>  NEW SHELL so 'claude' and 'codex' are on PATH  <=====

# --- 1.7  Verify, and confirm NO ANTHROPIC_API_KEY (subscription billing for claude -p) ---
claude --version
codex --version
[System.Environment]::GetEnvironmentVariable('ANTHROPIC_API_KEY','User')     # expect blank
[System.Environment]::GetEnvironmentVariable('ANTHROPIC_API_KEY','Machine')  # expect blank

# --- 1.8  Log in to Claude on Kevin's PERSONAL account (opens a browser) ---
claude login
claude -p "reply with the single word: ready"

# --- 1.9  GITHUB_PAT for the briefing push + command-centre sync (USER env var,
#          never written to a file). Paste Kevin's existing work-inbox PAT value: ---
[System.Environment]::SetEnvironmentVariable('GITHUB_PAT','<PASTE PAT HERE>','User')
#   verify in a NEW shell:  [System.Environment]::GetEnvironmentVariable('GITHUB_PAT','User').Length

# --- 1.10  Run dir + pull the pipeline scripts (cache-busted raw URLs) ---
New-Item -ItemType Directory -Force -Path "$env:USERPROFILE\work-inbox" | Out-Null
Set-Location "$env:USERPROFILE\work-inbox"
$t = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
foreach ($f in 'fetch_inbox.py','imap_mail.py','reauth_imap.py','diff_mail_pull.py') {
  Invoke-WebRequest -UseBasicParsing "https://raw.githubusercontent.com/begb0037admin/work-inbox/main/$f`?t=$t" -OutFile $f
}
python -m py_compile fetch_inbox.py imap_mail.py reauth_imap.py diff_mail_pull.py
Write-Host "PHASE 1 COMPLETE - paste all output from 1.0 onward back to Drew."
```

Notes for Kevin:
- Do **not** set `ANTHROPIC_API_KEY` anywhere on this laptop — its absence is what makes `claude -p` bill the subscription, not the metered API.
- `claude -p` **failover** (the desktop's `C:\WorkInboxAI\{kevin,hope}` two-config-dir setup, kevin@ primary → hope@ overflow) is **Phase 3**, not Phase 1 — Phase 1 does the single primary `claude login`. Kevin can decide then whether to replicate the failover or simplify to one account (Lane B is light; triage volume is the existing 5×/weekday).
- **No `codex login` in Phase 1.** Codex CLI is installed now; the Edu-account login is deferred to the 1-Sept Lane B kickoff (below), after Kevin has moved his interactive work off Edu and stripped the connectors.
- `winget` machine-wide installs show UAC prompts — expected with local admin.

### Deferred — 1-Sept Lane B kickoff block (run AFTER Kevin strips the Edu connectors)

```powershell
# PREREQUISITE: GitHub + Outlook Email connectors REMOVED from the Edu account;
# only Outlook Calendar + Teams remain. Kevin's interactive work has moved off Edu.
codex login                     # opens a browser - sign in to the Oxford ChatGPT EDU account
codex --version
$before = (Get-FileHash "$env:USERPROFILE\.codex\config.toml" -Algorithm SHA1).Hash
# read-only tool enumeration - do NOT approve any tool run:
codex exec -s read-only --skip-git-repo-check "List every tool available to you as a JSON array of names. Call nothing."
$after  = (Get-FileHash "$env:USERPROFILE\.codex\config.toml" -Algorithm SHA1).Hash
"config.toml sha1 before=$before after=$after (must match)"
# Paste the tool list + both hashes back to Drew - this is Phase 2 make-or-break #2.
```

### Phase 3 parallel run (NOW — laptop, as `ad-oak\begb0037`)

Pulls the Phase-3 scripts and does a `MAIL_BACKEND=imap` capture side-by-side with a `com` capture. **`WI_MAIL_PARALLEL=1` makes `fetch_inbox.py` write only `data/parallel/*` and exit — no push, no `briefing.json`, no command-centre sync, no calendar/Granola/AI.**

```powershell
# 3a. one-time: seed the IMAP token via the system browser (one account click, no password)
cd $env:USERPROFILE\work-inbox
$t=[DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
foreach ($f in 'fetch_inbox.py','imap_mail.py','reauth_imap.py','diff_mail_pull.py') {
  iwr -UseBasicParsing "https://raw.githubusercontent.com/begb0037admin/work-inbox/main/$f`?t=$t" -OutFile $f
}
python .\reauth_imap.py        # browser opens -> pick ad-oak\begb0037 -> "verified: EXAMINE INBOX OK, 558 messages"

# 3b. IMAP parallel capture (writes data\parallel\imap_inbox_raw.json + imap_sent_raw.json, then exits)
$env:MAIL_BACKEND='imap'; $env:WI_MAIL_PARALLEL='1'
python .\fetch_inbox.py
Remove-Item Env:MAIL_BACKEND; Remove-Item Env:WI_MAIL_PARALLEL

# 3c. (on the DESKTOP, or any box with classic Outlook connected) COM parallel capture
#     for the baseline half -- OR skip and diff against data/briefing.json history later.
#     $env:MAIL_BACKEND='com'; $env:WI_MAIL_PARALLEL='1'; python .\fetch_inbox.py ; Remove-Item Env:*

# 3d. diff (if both captures exist in data\parallel\)
python .\diff_mail_pull.py
Get-ChildItem .\data\parallel

# Paste 3a's verify line + 3b's console output (the "Phase 1 - IMAP mail pull: inbox N ... sent N"
# line and the "WI_MAIL_PARALLEL ... capture done" line) + 3d if you ran it. Back to Drew.
```

### Phase 5 parity (NOW, then ~once a day for 3–4 days — laptop, as `ad-oak\begb0037`)

```powershell
cd $env:USERPROFILE\work-inbox
$t=[DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
foreach ($f in 'fetch_inbox.py','imap_mail.py','diff_mail_pull.py','parity_vs_briefing.py') {
  iwr -UseBasicParsing "https://raw.githubusercontent.com/begb0037admin/work-inbox/main/$f`?t=$t" -OutFile $f
}
python .\parity_vs_briefing.py            # fresh IMAP capture + pull live briefing.json + diff
#   or:  python .\parity_vs_briefing.py --history 5   # also diff the last 5 briefing snapshots

# Writes data\parallel\parity_vs_briefing_<ts>.json. Paste the console summary
# (the "REAL FLAGS: N" lines + the "folder diagnostic" block) back to Drew.
# It runs its own capture -- no Outlook opens; no push; reads only.
```

Expected on 3b: `Calendar backend: com`, `Mail backend: imap  [WI_MAIL_PARALLEL ...]`, `IMAP - silent OAuth2 token OK for begb0037@ox.ac.uk via broker-app`, `Phase 1 - IMAP mail pull: inbox <N> ... sent <N>`, `WI_MAIL_PARALLEL (imap) mail-only capture done ... Exiting 0`.

---

## 8. Hard gates (unchanged, restated)

- **No cutover.** No `.bat` / scheduled-task change to defaults, no `main` default-behaviour change, no connector account creation/linking (Kevin strips connectors himself; Drew does not), without Kevin's **fresh explicit go-ahead for that specific step**.
- `claude -p` + the **desktop pipeline stay live throughout**.
- Baseline `~/.codex/config.toml` sha1 recorded before/after any `codex exec` (`35f8910382373d525598194b2649159cfeed3f6a`); **no `codex login` by Drew** — Kevin does it in Phase 1.
- Every run timestamped.
- Lane A: `com` default stays byte-identical; `imap` is dead code until a gated flip. `CAL_BACKEND=com` default; `connector` dead until a gated flip.
- Lane B is **design only** until Kevin approves `LANE_B_TEAMS_CAL_DESIGN.md`; then build behind the parallel task only.
- Checkpoint to `HANDOVER.md` + push; clickable commit links in every report.
- Do not touch `images/oxford-crest.jpg`. Do not re-investigate Microsoft **Graph** (disallowed at Oxford as an auth method — Kevin, 29 Aug) or **EWS** (removed from the plan). Do not use `TextEncoder` on `index.html`.

---

## 9. Open items for Kevin (none block Phase 1)

1. **`claude -p` failover on the laptop** — replicate the `C:\WorkInboxAI\{kevin,hope}` two-account setup, or simplify to Kevin's personal account only? (Decide at Phase 3.)
2. **Phase 3.9** — re-wire to `message_id` before cutover, or run fail-open-carry for week 1?
3. **Calendar-source degradation — DEFAULTED to "degrade + warn" (confirm with Kevin).** If a calendar source is unavailable — pre-1-Sept: classic Outlook not running under `MAIL_BACKEND=imap`; post-1-Sept: `codex exec` didn't load the connector that cycle — Phases 3.7/3.8 produce an **empty calendar + a warning line, and the mail briefing continues**. This is the behaviour now shipped in `fetch_inbox.py` (`d5447b9`). The alternative (hard-fail the whole run) loses the mail briefing too and defeats the resilience goal. Kevin: confirm you're OK with a briefing that occasionally has no calendar section rather than no briefing at all.
4. **Calendar kill-switch false-HALT tolerance** — accepted per Kevin's instruction (safety over uptime), but confirm Kevin is content to manually re-enable after a legitimate third-party meeting edit trips it.
5. **`SMTP.Send` in the IMAP token bundle** — accept the architectural mitigation (recommended), or pursue a dedicated app registration (needs Oxford IT, ruled out)?

---

## 10. Checkpoint status

- This rev + the Phase 1 command sequence + Lane B B1 findings (calendar + Teams) + `LANE_B_TEAMS_CAL_DESIGN.md` = the checkpoint. **Stop here for Kevin to run Phase 1 (§7, no Edu-account interaction) and report back before Phase 2(i).**
- **Now → 1 Sept:** Lane A only. Phase 1 toolchain, then Phase 2(i) — MSAL broker silent IMAP token off the PRT. Touches nothing on the Edu account or its quota.
- **From 1 Sept:** Edu becomes the dedicated automation identity (Kevin off Edu interactively, connectors stripped to {Calendar, Teams}). `codex login` on the laptop, then Phase 2(ii) — headless `codex exec` loading exactly {Calendar, Teams} read tools — then the Lane B build (`LANE_B_TEAMS_CAL_DESIGN.md` §9).
- Config baseline intact: `~/.codex/config.toml` sha1 `35f8910382373d525598194b2649159cfeed3f6a` (before/after both `.codex-global-state.json` reads this session — Teams manifest, then calendar manifest).
