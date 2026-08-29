# LAPTOP_MIGRATION_PLAN.md — work-inbox pull off Outlook COM, two-lane, on Kevin's Oxford laptop

**Date:** 2026-08-29 (Drew)
**Status:** PLAN + Phase-1 command sequence + Lane B B1/B2 findings. **No build. No cutover. No `.bat` / scheduled-task / `main` default-behaviour change.** Awaiting Kevin's review of this doc before Phase 2 starts.
**Companion docs (all on `main`):** `IMAP_OAUTH2_SPIKE_20260828.md` (mail feasibility PASS), `MAIL_BACKEND_MIGRATION_PLAN.md` (the `MAIL_BACKEND=com|imap` flag, already built), `PHASE1_IMAP_MIGRATION_AUDIT.md` (what IMAP breaks), `CONNECTOR_SAFEGUARDS.md` (§A Teams tool enumeration; §D Codex "NOT SOUND" for an unattended connector fetch that holds write tools), `CODEX_CONNECTOR_PIPELINE_PLAN.md` + `EMAIL_AUTOMATION_SECURITY_MITIGATIONS.md` (connector history — now scoped down to Teams-only, Lane B).

---

## 0. Why this supersedes the prior direction

The 29 Aug ~03:30 HANDOVER entry left the connector route at "NOT SOUND, back to Kevin". Kevin's decision **this session** resolves it:

- The **mail + calendar** pull moves to our own **read-only Python** (Lane A) — structurally no write tool exists in the code path. This is the enforced read-only boundary Codex said was missing.
- The **ChatGPT M365 connector is retained for Teams only** (Lane B) — chat, in-meeting chat, meeting recordings/transcripts, which have **no other route**. Kevin has explicitly accepted the residual risk here, contained to "worst case a stray Teams message" via a dedicated single-connector identity + dumb-fetch + kill-switch.
- **Triage stays on `claude -p`** (headless Claude Code, Kevin's subscription), unchanged. Both lanes feed it after sanitisation.
- **Host moves from the admin desktop (`DESKTOP-MJDJM64`) to Kevin's Oxford laptop (`begb0037.AD-OAK`).** The desktop is Azure-AD *registered* with **no PRT** — every token refresh is a periodic interactive hang (root cause of the 28 Aug outage). The laptop is Azure-AD **joined + domain-joined with `AzureAdPrt: YES`**, so MSAL can broker tokens **silently off the PRT**. This is the single biggest reliability change.

The desktop pipeline (`claude -p` + COM + `Classic Outlook Keepalive` watchdog) **stays live throughout** as the fallback. Nothing is retired until both lanes are proven and Kevin gives a fresh explicit cutover go-ahead (Phase 6).

---

## 1. Host — Kevin's Oxford laptop

| Fact | Value | Verified |
|---|---|---|
| Machine / profile | Oxford laptop, Windows user `begb0037.AD-OAK` (Oxford AD domain account) | this session (Kevin) |
| Join state | `AzureAdJoined: YES`, `DomainJoined: YES` | this session |
| **PRT** | **`AzureAdPrt: YES`** — tenant `cc95de1b-97f5-4f93-b4ba-fe68b852cf91`; auto-renews ~14 days on the Oxford network | this session |
| Python | **not installed** | this session |
| Node / Claude Code | **not installed** | this session |
| Admin rights | Kevin has **local admin only** (not domain admin) | this session |
| Power | laptop stays **docked and on** | Kevin's commitment |
| `%LOCALAPPDATA%` | `C:\Users\begb0037.AD-OAK\AppData\Local` — token caches + toast stamps live here (mirrors the desktop's `%LOCALAPPDATA%\WorkInboxAI\`) | inferred from profile name |

**Execution reality:** Drew runs on the desktop and cannot execute on the laptop. Every laptop step is delivered as copy-paste-ready PowerShell 5.1 for Kevin to run and paste back. Kevin is the hands for installs, `claude login`, and the first interactive auth consent.

---

## 2. Target architecture

```
                         KEVIN'S OXFORD LAPTOP  (begb0037.AD-OAK, PRT present)
  ┌───────────────────────────────────────────────────────────────────────────────┐
  │                                                                               │
  │  LANE A  —  mail + calendar  (SAFE: our own read-only Python, no send tool)   │
  │  ┌─────────────────────────────────────────────────────────────────────────┐  │
  │  │ MSAL broker  (enable_broker_on_windows=True, pymsalruntime)             │  │
  │  │   acquire_token_silent  →  off the laptop PRT, ZERO prompts             │  │
  │  │        │                          │                                    │  │
  │  │        ▼ IMAP.AccessAsUser.All     ▼ EWS.AccessAsUser.All  (or Graph    │  │
  │  │   imap_mail.pull()                 Calendars.Read — bake-off, §4b)      │  │
  │  │   outlook.office365.com:993        cal reader: FindItem on Calendar     │  │
  │  │   EXAMINE (read-only)              (read-only)                          │  │
  │  └──────────┬──────────────────────────────────┬───────────────────────────┘  │
  │             │ inbox / sent lists               │ calendar events              │
  │             ▼                                  ▼                              │
  │        fetch_inbox.py  Phase 1 (MAIL_BACKEND=imap) + Phases 3.7/3.8           │
  │        (CAL_BACKEND=ews|graph)  —  COM code paths guarded out                 │
  │                                                                               │
  │  LANE B  —  Teams only  (eyes-open residual, Kevin accepted)                  │
  │  ┌─────────────────────────────────────────────────────────────────────────┐  │
  │  │ DEDICATED ChatGPT identity — ONLY the Microsoft Teams connector linked  │  │
  │  │ `codex exec` dumb-fetch: rigid "list these chats / channels / meeting   │  │
  │  │   transcripts, return these fields" — NO reasoning over content         │  │
  │  │   → sanitiser → Teams-specific post-run kill-switch                     │  │
  │  │   → tool-manifest logged every run; auto-disable on any unexpected tool │  │
  │  └──────────┬──────────────────────────────────────────────────────────────┘  │
  │             │ teams messages / transcripts (sanitised)                        │
  │             ▼                                                                 │
  │        ┌──────────────────────────────────────────────────────────┐          │
  │        │  claude -p   (headless Claude Code, Kevin's subscription) │  ← triage │
  │        │  UNCHANGED. Consumes Lane A + Lane B after sanitisation.  │          │
  │        └───────────────────────────┬──────────────────────────────┘          │
  │                                    ▼                                          │
  │                     data/briefing.json → GitHub → dashboard                   │
  └───────────────────────────────────────────────────────────────────────────────┘
```

Everything downstream of `briefing.json` (dashboard, command-centre sync, `data/ticks.json`, Lauren's drafting loop) is **unchanged** — it consumes the JSON regardless of what produced it.

---

## 3. What already exists on `main` (do not rebuild)

From the 28–29 Aug IMAP work, behind the **unset** `MAIL_BACKEND` flag (`com` default = byte-identical to today):

| File | State | Gap for the laptop |
|---|---|---|
| `imap_mail.py` | `pull()`, `_pull_sent()`, `message_still_in_inbox()`, `acquire_token_silent()`, OWA-search `web_link`, mUTF-7 subfolder resolution, `Importance:`/`X-Priority:` → 0/1/2, `_kevin_is_primary_recipient()`. Parity harness fixes (**Message-ID join, Sent meeting-response filter, subfolder LIST — all DONE**, per the 29 Aug ~01:15 HANDOVER entry; last diff run: INBOX common 48/52, **SENT 10==10, real parity issues 0**). | **Auth is device-code only** (`AUTHORITY=.../organizations`, Thunderbird client `9e5f94bc-…`, scope `IMAP.AccessAsUser.All`, cache `%LOCALAPPDATA%\WorkInboxAI\msal_imap_token_cache.bin`). Needs a **broker path** added for the PRT machine (Phase 2). |
| `reauth_imap.py` + `docs/desktop-scripts/Re-auth Work Inbox IMAP.bat` | device-code priming, read-only `EXAMINE INBOX` verify, timestamps | becomes the **fallback** only; broker makes the happy path promptless |
| `diff_mail_pull.py` | field-by-field COM↔IMAP parity diff, `data/parallel/parity_<ts>.json` | reused as-is for Phase 5 |
| `fetch_inbox.py` | `MAIL_BACKEND=com\|imap` + `WI_MAIL_PARALLEL=1`; 4 COM mail loops guarded `for X in ([] if MAIL_BACKEND=="imap" else <orig>)`; `connect_to_outlook()` non-fatal under `imap`; `mapi is None` calendar guard; `_imap_reauth_toast_due()` 1/hr stamp. `com` path proven byte-identical. **Restore point:** `main` `9a52b07`, blob `bd02b41089850678b8268318a0afab5e6d457e8a`, snapshot `Archive/fetch_inbox_backup_20260828_*_pre_mail_backend_flag.py`. | top-of-file `import win32com.client` / `import pywintypes` / `import anthropic` are **unconditional** — a COM-free laptop still needs those importable. Phase 3 makes them lazy; Phase 1 installs `pywin32` + `anthropic` anyway so nothing blocks. **No EWS/Graph calendar reader exists** — Phase 3 builds it + a `CAL_BACKEND` flag. |

**Genuinely new work:** broker auth in `imap_mail.py`; a calendar reader module (EWS or Graph — §4b) + `CAL_BACKEND` flag; lazy COM imports in `fetch_inbox.py`; the laptop scheduled task; Lane B in full (design-only until Kevin approves).

---

## 4. Lane A design detail

### 4a. Auth — MSAL broker, silent off the PRT

- Library: `msal` with `enable_broker_on_windows=True`; requires `pymsalruntime` (`pip install "msal[broker]"`).
- `PublicClientApplication(CLIENT_ID, authority=..., enable_broker_on_windows=True, token_cache=...)`.
- First run: `acquire_token_interactive(SCOPES, parent_window_handle=msal.PublicClientApplication.CONSOLE_WINDOW_HANDLE)` → on a PRT machine this is a **WAM broker** flow: no password, no device code, at most **one consent click** for a first-party client, often zero. Then the account is in the cache.
- Every subsequent run: `acquire_token_silent(SCOPES, account=accounts[0])` → served from the broker/PRT, **no prompt**, survives reboot and (per the laptop's ~14-day PRT auto-renew) survives far longer than the desktop's no-PRT refresh churn.
- Failure path unchanged: silent failure → `ImapReauthRequired` → one 1/hour toast → `SystemExit(1)`. Never an interactive hang in the scheduled context.
- **Client-id question for Phase 2:** the Thunderbird public client (`9e5f94bc-…`) is pre-authorised for the **Exchange Online** resource (IMAP + EWS + POP + SMTP — the 28 Aug spike token bundle showed `EWS.AccessAsUser.All IMAP.AccessAsUser.All POP.AccessAsUser.All SMTP.Send User.Read`). It is **not** authorised for Microsoft Graph. If calendar goes via **Graph** (§4b) we need a different pre-consented client — the **Microsoft Office** first-party id `d3590ed6-52b3-4102-aeff-aad2292ab01c` (broadly pre-consented, includes Graph delegated; note it *fails* for IMAP with `AADSTS65002`, so it cannot replace Thunderbird for the mail half). Phase 2 proves which client(s) the broker can get tokens for silently.

### 4b. Calendar — EWS vs Graph bake-off (Phase 2 decides)

Phases 3.7 (raw pull) / 3.8 (AI summaries) currently read the primary calendar **and** the "People Department - HR Systems" shared calendar via COM. Replacement options:

| Path | Scope | Client | Pros | Cons / risk |
|---|---|---|---|---|
| **EWS `FindItem`** on `Calendar` (+ shared cal by SMTP) | `EWS.AccessAsUser.All` | Thunderbird `9e5f94bc-…` (already in the bundle) | one client for mail + calendar; `exchangelib` or a hand-rolled SOAP `FindItem` is small; read-only | **Microsoft has announced EWS for Exchange Online will be blocked from 1 October 2026** (verify against current MS guidance — this is ~5 weeks out). Building the long-term calendar path on EWS is a dead man walking. |
| **Graph `GET /me/calendarView`** (+ `/users/{shared}/calendarView`) | `Calendars.Read` (delegated — **your own** calendar read generally needs **no admin consent**; the Graph dead-end was about *app-registration* consent on the non-joined desktop, a different thing) | Office `d3590ed6-…` | survives past Oct 2026; trivial REST, no extra lib; JSON maps cleanly to the Phase 3.7 dict | needs its own client + a broker token proof (Phase 2); shared-calendar read may need the mailbox owner to have shared it / `Calendars.Read.Shared` |
| **stay on COM for calendar only** | — | — | zero new work | reintroduces "classic Outlook must run" on the laptop — the exact thing we're removing. **Only** as a stopgap if both above fail Phase 2. |

**Phase 2 proves both** an EWS `FindItem` and a Graph `calendarView` call returning real events with a broker-acquired token, then we pick the survivor. Given the EWS retirement clock, **Graph is the presumptive choice unless its broker token or shared-calendar read fails.** `CAL_BACKEND=ews|graph|com` flag mirrors `MAIL_BACKEND`; `com` stays default until cutover.

### 4c. Known Lane A residuals (from `PHASE1_IMAP_MIGRATION_AUDIT.md`, still true)

- **`openmail://` opener dies** under IMAP — replaced by an OWA search deep-link on `Message-ID` (`imap_mail._owa_search_link()`, already written). The **dashboard JS branch is not written** — until it is, IMAP cards have a dead opener. Phase 3 deliverable; needs Kevin screenshot approval per command-centre-style UI gate.
- **Phase 3.9** (inbox-resolution tracking) keys on Outlook `EntryID` today; under IMAP it degrades to fail-open-carry unless re-wired to `message_id` + `imap_mail.message_still_in_inbox()`. Open decision: re-wire before cutover, or accept fail-open-carry for week 1.
- **`SMTP.Send` in the token bundle** — Thunderbird's client returns the whole mail bundle. Mitigation is architectural and unchanged: `imap_mail.py` imports `imaplib` only, never `smtplib`; no agent-with-tools on this path. Accept, or (only if a zero-send token is mandated) attempt a dedicated single-tenant app registration — needs Oxford IT, currently ruled out.
- **Outlook Categories** — zero dependence confirmed by full-repo grep; nothing to replace.

---

## 5. Lane B design detail (DESIGN ONLY — no build until Kevin approves the Lane B design)

### 5a. Scope — Teams, and only Teams

Mail and calendar are Lane A. Lane B exists solely for what the connector uniquely reaches:
1. 1:1 / group chat messages
2. Channel messages
3. In-meeting chat (surfaces as a Teams chat tied to the meeting)
4. Meeting **recordings** metadata + **transcripts** (WebVTT text) for **scheduled** meetings

### 5b. B1 findings — Teams connector tool manifest (live, read-only)

Source: `~/.codex/.codex-global-state.json` → `electron-persisted-atom-state` → `mcp-extension-sidebar-catalog` → `codex_apps` server, read 2026-08-29 on the desktop. **`microsoft_teams` = 33 tools, 9 write / 24 read.** `~/.codex/config.toml` sha1 `35f8910382373d525598194b2649159cfeed3f6a` recorded before and after this read — unchanged; no `codex exec` run, plain file read.

**Write (9) — never in Lane B's allowlist; any appearance in a run manifest = auto-disable:**
`send_chat_message`, `send_channel_message`, `reply_to_message`, `reply_to_channel_message`, `create_chat`, `create_channel`, `create_planner_task`, `update_planner_task`, `delete_planner_task`.
Tier-1 (external-comms, irreversible): the four send/reply tools + arguably `create_chat`/`create_channel`. **This is the whole of Lane B's accepted blast radius — "worst case a stray Teams message."**

**Read (24) — the fetch surface Lane B actually uses:**

| Purpose | Tools |
|---|---|
| enumerate + resolve | `list_teams`, `list_channels`, `list_chats`, `resolve_team`, `resolve_channel`, `resolve_chat`, `resolve_user`, `get_profile`, `get_chat_members` |
| read messages | `list_chat_messages`, `list_channel_messages`, `fetch` (one message/conversation by path), `search` (chats + channel messages, `sender_name`/`recipient_name` filters) |
| meetings — recordings/transcripts | `resolve_scheduled_online_meeting` (from an event ID or join URL), `list_online_meeting_recordings`, `get_online_meeting_recording`, `list_online_meeting_transcripts`, `get_online_meeting_transcript`, `get_online_meeting_transcript_content` (size-bounded WebVTT) |
| Planner reads (not needed; noted) | `list_planner_plans`, `list_planner_buckets`, `list_planner_tasks`, `fetch_planner_task` |
| preflight helper | `validate_write_target` (read-only, but its existence signals the connector expects write flows) |

**Lane B Call-1 allowlist (proposed):** `list_teams, list_channels, list_chats, resolve_team, resolve_channel, resolve_chat, list_chat_messages, list_channel_messages, fetch, search, resolve_scheduled_online_meeting, list_online_meeting_transcripts, get_online_meeting_transcript_content`. Anything else called → discard output + disable task + alert.

Constraints noted from the tool descriptions: recordings/transcripts are **scheduled meetings only** and must be resolved from an event ID or join URL first — so Lane B's meeting-transcript pull is keyed off the Lane A calendar pull (event ID → `resolve_scheduled_online_meeting` → transcript). Ad-hoc call transcripts are not reachable this way.

### 5c. B2 findings — which ChatGPT identity carries the Teams connector

**How ChatGPT connectors work:** each connector ("app") is an independent per-account OAuth link. On any one ChatGPT account you choose which connectors to link — "Teams only, nothing else" is achievable simply by not linking anything else on that account. The Microsoft-side identity (who you sign in as when linking) is separate from the ChatGPT account identity.

| Option | Assessment |
|---|---|
| **Kevin's personal ChatGPT Plus** (current `codex` account `eb7a812e-…`) | **Disqualified by the brief.** Live manifest read this session shows it already carries **GitHub (89 tools), Outlook Email (46), Outlook Calendar (34), Teams (33), Canva, Granola** — 292 Apps tools total, incl. 41 GitHub write tools and the mail send tools. Using it for Lane B puts all of that one connector-load away from an unattended run. Not isolated. |
| **Kevin's Oxford ChatGPT Edu account** | **Not recommended as the dedicated identity.** (1) It's Kevin's *interactive* account — automation shares his quota and a workspace admin can change connector policy under the run. (2) It is almost certainly where the Oxford-consented **Outlook** connector already lives (the funding-rationale connector) — so it is not "Teams only". (3) Whether the Edu **workspace admin has enabled the Teams app** for members is **unknown** — needs Kevin or an Oxford ChatGPT workspace admin to check Settings → Connectors/Apps. |
| **A dedicated, separate ChatGPT account** (new personal Plus ~£16/mo, or a second identity if Oxford can provision one — unlikely; Edu seats are per-person) with **ONLY** the Microsoft Teams connector linked, authenticated to Kevin's Oxford Entra identity (`begb0037@ox.ac.uk`) | **Recommended.** This is the only configuration that is genuinely single-connector and isolated from Kevin's interactive quota and from the GitHub/Outlook write surface. Cost: ~£16/mo if Plus (small vs the containment benefit). |

**Single blocking unknown before any Lane B build — must de-risk first:** can the Microsoft **Teams** connector actually be *consented* for `begb0037@ox.ac.uk` — i.e. is the OpenAI/ChatGPT enterprise app permitted for the Teams/Graph `Chat.Read`/`ChannelMessage.Read`/`OnlineMeetingTranscript.Read` scopes under Oxford's Entra **app-consent policy**, or does it need tenant-admin consent the way Graph-direct did? **This is the same failure class that killed Graph-direct.** The Outlook connector got consent somehow — Phase B2 follow-up is to establish whether that was *admin* (tenant-wide) or *user* consent, because Teams needs the same. If it needs admin consent and Oxford won't give it, Lane B is dead and Teams stays invisible to work-inbox.

**Not done, per the brief:** no account created, no connector linked, no `codex login`. Scope only.

### 5d. B3 — dumb-fetch design (to be written as a design doc after Kevin approves the Lane B direction)

Mirrors the (now retired for mail) `CODEX_CONNECTOR_PIPELINE_PLAN.md` two-call split, scoped to Teams:
- **Call 1** (connector attached, dedicated `CODEX_HOME` on the dedicated account): rigid list/read instruction over the B1 allowlist, no reasoning over message content, fixed field list, hard output cap.
- **sanitiser** (`normalise_pull.py`-style, deterministic): HTML strip, truncate, neutralise instruction-like text / role markers, strip zero-width + bidi, record hits.
- **Teams-specific post-run kill-switch:** pre/post read of "messages `from = me`" in the run window across the chats/channels touched; any new one → `Disable-ScheduledTask` + BurntToast + `GUARD_TRIPPED_<ts>.json`. Documented blind spots (net-zero, delayed) — it is containment, not prevention.
- **tool-manifest logging every run** + **auto-disable** if any tool outside the B1 allowlist (esp. any of the 9 writes) appears in the available or called set.
- **`~/.codex/config.toml` sha1 before/after every `codex exec`; no `codex login`.**
- **Codex "NOT SOUND" caveat (`CONNECTOR_SAFEGUARDS.md` §D) still applies** — Call 1 holds the write tools while it runs; the prompt is not a capability boundary. Lane B's justification is *narrower blast radius* (Teams message vs external HR email), *dedicated single-connector identity*, *Kevin's explicit risk acceptance*, and *no other route to Teams* — not that the objection was resolved.

---

## 6. Phases — increment and report after each; do NOT run to completion silently

### Lane A

| # | Phase | Output | Gate |
|---|---|---|---|
| **1** | **Laptop toolchain** — Python 3.12, Node LTS, Git, Claude Code + `claude login` (Kevin's personal account), Python packages, first pull of the pipeline scripts, `py_compile` clean | pasted command output | this doc's review |
| **2** | **MAKE-OR-BREAK: broker silent-auth proof.** A small script (`scratchpad`, throwaway): MSAL `enable_broker_on_windows=True` acquires an **IMAP** token and a **calendar** token (EWS *and* Graph attempt) silently off the PRT — target zero prompts after ≤1 first-run consent click. Prove: IMAP `SELECT INBOX`; and a calendar read (`FindItem` on Calendar for EWS, `GET /me/calendarView` for Graph) returning real events. | proof log + which calendar path won | **STOP + report.** If broker can't get these silently, the "no prompts" premise fails — reassess. |
| **3** | Wire it in behind flags: add the broker path to `imap_mail.py`; new calendar reader module (`ews_calendar.py` or `graph_calendar.py`) + `CAL_BACKEND=com\|ews\|graph` flag (mirror `MAIL_BACKEND`; `com` default byte-identical); make `fetch_inbox.py`'s `win32com`/`pywintypes`/`anthropic` imports lazy; dashboard JS `mail_backend==="imap"` → OWA opener branch (screenshot for Kevin). Phases 3.7/3.8 consume the calendar source under the flag. | diffs + restore point | Kevin go-ahead to push code |
| **4** | Scheduled task on the laptop under `begb0037.AD-OAK`, "run whether logged on or not", 5×/weekday matching the current cadence (07:00/09:00/11:00/13:00/15:00/17:00 Mon–Fri — confirm against the live desktop task); `powercfg` never-sleep; every run timestamped. **Parallel task — writes `docs/*` / `data/parallel/*` only, never `data/briefing.json`.** | task XML + `powercfg` settings | Kevin go-ahead |
| **5** | Parallel-run: laptop IMAP+calendar pull vs the live desktop `data/briefing.json` GitHub history — field-level diff (`diff_mail_pull.py` + a calendar diff), several days, Kevin/Lauren eyeball | parity reports | — |

### Lane B (parallel workstream)

| # | Phase | Output | Gate |
|---|---|---|---|
| **B1** | Teams connector tool enumeration | **DONE — §5b above** | — |
| **B2** | Dedicated identity analysis | **DONE — §5c above.** Follow-up: Kevin/Oxford-admin to confirm (a) Teams app enabled on the Edu workspace, (b) whether the ChatGPT/OpenAI enterprise app can get **user** consent for Teams Graph scopes at Oxford or needs **admin** consent | Kevin/Oxford check |
| **B3** | Dumb-fetch `codex exec` Teams pull + sanitiser + Teams kill-switch + manifest-logging/auto-disable — **design doc, not build** | `docs/LANE_B_TEAMS_DESIGN.md` | Kevin approves the Lane B design before it is written up in full / built |

### Cutover — Phase 6, separately gated

Both lanes proven → **Kevin's fresh explicit go-ahead for the cutover step specifically** → flip flags → retire, in order: desktop `\Work Inbox Briefing` task; classic Outlook (uninstall); `Classic Outlook Keepalive` task + `Ensure-ClassicOutlook.ps1` preflight; COM calendar pull; `openmail://` opener → OWA deep-link. `claude -p` + the desktop pipeline stay live until this step.

---

## 7. Phase 1 — exact command sequence for Kevin (PowerShell 5.1, on the laptop)

> Run each numbered block, **paste the full output back to Drew** before moving on. `;` not `&&`. `&` call-operator for quoted paths. Where a block says "NEW SHELL", close the window and open a fresh PowerShell 5.1 so PATH updates apply.

```powershell
# ============================================================================
#  work-inbox laptop migration — PHASE 1 (toolchain).  Laptop, profile
#  begb0037.AD-OAK.  Start in an ELEVATED Windows PowerShell 5.1 for 1.0–1.2.
# ============================================================================

# --- 1.0  Confirm we are on the right machine ---
$env:COMPUTERNAME
whoami
$PSVersionTable.PSVersion
dsregcmd /status | Select-String 'AzureAdJoined|DomainJoined|AzureAdPrt|TenantId|PRT '

# --- 1.1  winget present? (Windows 11 ships App Installer; if this errors, install
#          "App Installer" from the Microsoft Store first, then re-run) ---
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

# --- 1.4  Python packages for Lane A (broker auth + calendar + pipeline imports) ---
python -m pip install --upgrade pip
python -m pip install "msal[broker]" pywin32 anthropic exchangelib
python -c "import msal, pymsalruntime; print('msal', msal.__version__, 'pymsalruntime OK')"
python -c "import exchangelib; print('exchangelib', exchangelib.__version__)"
python -c "import win32com.client, anthropic; print('pywin32 + anthropic import OK')"

# --- 1.5  Install Claude Code (native installer — same method as the desktop) ---
irm https://claude.ai/install.ps1 | iex

# =====>  NEW SHELL again so 'claude' is on PATH  <=====

# --- 1.6  Verify Claude Code, and confirm NO ANTHROPIC_API_KEY (subscription billing) ---
claude --version
[System.Environment]::GetEnvironmentVariable('ANTHROPIC_API_KEY','User')     # expect blank
[System.Environment]::GetEnvironmentVariable('ANTHROPIC_API_KEY','Machine')  # expect blank

# --- 1.7  Log in to Claude on Kevin's PERSONAL account (opens a browser) ---
claude login
#   then prove it runs on the subscription:
claude -p "reply with the single word: ready"

# --- 1.8  Create the run dir and pull the pipeline scripts (cache-busted raw URLs) ---
New-Item -ItemType Directory -Force -Path "$env:USERPROFILE\work-inbox" | Out-Null
Set-Location "$env:USERPROFILE\work-inbox"
$t = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
foreach ($f in 'fetch_inbox.py','imap_mail.py','reauth_imap.py','diff_mail_pull.py') {
  Invoke-WebRequest -UseBasicParsing "https://raw.githubusercontent.com/begb0037admin/work-inbox/main/$f`?t=$t" -OutFile $f
}
python -m py_compile fetch_inbox.py imap_mail.py reauth_imap.py diff_mail_pull.py
Write-Host "PHASE 1 COMPLETE — paste all output from 1.0 onward back to Drew."
```

Notes for Kevin:
- If `irm https://claude.ai/install.ps1 | iex` is blocked or fails, fallback: `npm install -g @anthropic-ai/claude-code` then a new shell.
- Do **not** set `ANTHROPIC_API_KEY` anywhere on this laptop — its absence is what makes headless `claude -p` bill the subscription instead of the metered API.
- `winget` machine-wide installs will show UAC prompts — that's expected with local admin.
- Nothing here touches the live desktop pipeline. This is all additive on a new machine.

---

## 8. Hard gates (unchanged, restated)

- **No cutover.** No `.bat` / scheduled-task change to defaults, no `main` default-behaviour change, no connector account creation/linking, without Kevin's **fresh explicit go-ahead for that specific step**.
- `claude -p` + the **desktop pipeline stay live throughout**.
- Baseline `~/.codex/config.toml` sha1 recorded before/after any `codex exec` (`35f8910382373d525598194b2649159cfeed3f6a`); **no `codex login`**.
- Every run timestamped.
- Lane A code paths: `com` / `ews`(interim) stay default; `imap` / `graph` are dead code until a gated flip.
- Checkpoint to `HANDOVER.md` + push; clickable commit links in every report.
- Do not touch `images/oxford-crest.jpg`. Do not re-investigate Microsoft **Graph app-registration** consent (delegated self-calendar read is a different mechanism — §4b). Do not use `TextEncoder` on `index.html`.

---

## 9. Open decisions / risks for Kevin

1. **EWS retires ~1 Oct 2026 for Exchange Online** (verify against current MS guidance). If Lane A calendar lands on EWS it has weeks of life. **Recommendation: Graph delegated `Calendars.Read` for calendar** (self-calendar read does not need admin consent), with EWS only as a throwaway Phase-2 comparison. Phase 2 proves the Graph broker token.
2. **Lane B identity spend** — a dedicated ChatGPT Plus account (~£16/mo) is the only genuinely isolated option. Approve, or accept the Edu-account compromises (shared quota, admin-policy exposure, not single-connector)?
3. **Lane B consent risk** — if the OpenAI/ChatGPT Teams connector needs *tenant-admin* consent at Oxford (same class as Graph-direct), Lane B is blocked and Teams stays out of scope. Kevin / an Oxford ChatGPT workspace admin to check before any Lane B build.
4. **`SMTP.Send` in the IMAP token bundle** — accept the architectural mitigation (recommended), or hold out for a dedicated app registration (needs Oxford IT, currently ruled out)?
5. **Phase 3.9** — re-wire to `message_id` before cutover, or run fail-open-carry for week 1?
6. **Calendar under `imap`/COM-free** — accept graceful degradation (dead calendar source → empty + warning, briefing continues) vs hard-fail the run?
7. **`codex exec` connector loading on the laptop** — the 28 Aug Q2 finding was that a headless `codex exec` on the *desktop* loaded **zero** connector tools (cause undetermined; needs `codex login` / the ChatGPT app-server bridge). Lane B on the laptop will hit the same question — a Lane B B3 prerequisite, flagged now.

---

## 10. Checkpoint status

- This doc + the Phase 1 command sequence + Lane B B1/B2 findings = the first checkpoint. **Stop here for Kevin's review before Phase 2.**
- Phase 2 (broker silent-auth proof) is the make-or-break and must not start until Kevin has read this.
- Config baseline intact: `~/.codex/config.toml` sha1 `35f8910382373d525598194b2649159cfeed3f6a` (before/after the manifest read this session).
