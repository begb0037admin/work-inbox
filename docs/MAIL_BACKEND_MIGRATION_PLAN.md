# MAIL_BACKEND migration plan — Outlook COM → IMAP+OAuth2 for the mail pull

**Date:** 28 August 2026 (Drew)
**Companion docs:** `PHASE1_IMAP_MIGRATION_AUDIT.md` (what breaks), `IMAP_OAUTH2_SPIKE_20260828.md` (feasibility PASS)
**Status: BUILT behind a flag, default OFF. NOT cut over. No `.bat` / scheduled-task change. Awaiting parity verification + Kevin's fresh explicit go-ahead.**

---

## 1. Goal and non-goals

**Goal:** move the daily briefing's **mail** half off Outlook COM so a closed / wedged / sign-in-prompting classic Outlook can no longer break it — the exact failure of the 28 Aug 13:30 reboot.

**Non-goals (unchanged, COM-only, forever on this plan):**
- Calendar — Phases 3.7 (raw pull) and 3.8 (AI summaries). IMAP has no calendar.
- Therefore classic Outlook must still be *runnable*; the WS1 `Classic Outlook Keepalive` task stays relevant. IMAP **shrinks** the Outlook dependency, it does not remove it.

---

## 2. The flag (mirrors `AI_BACKEND`)

| Env var | Values | Default | Effect |
|---|---|---|---|
| `MAIL_BACKEND` | `com` \| `imap` | `com` | `com` = today's COM mail pull, **byte-identical**. `imap` = mail pull via `imap_mail.py`. |
| `WI_MAIL_PARALLEL` | `1` | unset | Do the mail pull, dump raw lists to `data/parallel/<com\|imap>_{inbox,sent}_raw.json`, **push nothing, mutate nothing** (folds into the existing `WI_AI_PARALLEL` no-write posture). For `diff_mail_pull.py`. |

Implemented in `fetch_inbox.py` as:
- `MAIL_BACKEND` / `MAIL_PARALLEL` defined next to `AI_BACKEND`; `MAIL_PARALLEL` ORs into `AI_PARALLEL` so every existing no-write guard (push, triage ledger, Command-Centre sync, `claude_` output prefix) also holds for a mail-parallel run.
- The four COM mail loops guarded with `for X in ([] if MAIL_BACKEND == "imap" else <original expr>)` — under `com` the ternary returns the original expression unchanged.
- Under `imap`, after the (empty) COM loops, `imap_mail.pull(...)` populates `inbox` / `sent`.
- `connect_to_outlook()` is wrapped under `imap` so a COM failure is non-fatal (calendar degrades to empty + warning; mail briefing continues). Under `com` it is unchanged (hard-fails the run, as today).
- On `imap_mail.ImapReauthRequired`: log, fire ONE rate-limited toast (`_imap_reauth_toast_due()`, 1/hour stamp file — mirrors the WS1 keepalive stamp mechanism), `raise SystemExit(1)`. **Never an interactive prompt, never a hang.**

**Verification that `com` is byte-identical:** `python -m py_compile` clean; every guard is `… else <verbatim original>`; `mapi is None` is only reachable under `imap`; no downstream COM call path changes when `MAIL_BACKEND` is unset or `com`.

---

## 3. Credential handling — DECISION: ship on Thunderbird's public client id

### Thunderbird client id vs a dedicated single-tenant app registration

| | Thunderbird public client `9e5f94bc-…` | Dedicated single-tenant app registration |
|---|---|---|
| Proven at Oxford | **Yes** — spike: 558 msgs, silent refresh confirmed | Untested; **Oxford has a confirmed history of blocking app-registration consent flows** (that is exactly why MS Graph is a dead end here — see `CLAUDE.md` "What Was Tried and Abandoned" and the no-PRT / Graph confirmed-fact memory) |
| Needs Oxford IT | No | **Almost certainly yes** (admin consent / app-consent-policy) — and Kevin has **ruled out Oxford IT** (27 Aug, verbatim "NOT going to Oxford org IT") |
| Client secret to store | **None** — public client, device-code / PKCE only | None if kept public, but the registration itself is the blocker |
| Token scope bundle | Exchange returns the whole mail bundle incl. `SMTP.Send` (Thunderbird's client is pre-authorised for all mail scopes) | Could be narrowed to delegated `IMAP.AccessAsUser.All` only |

**Decision: Thunderbird's public client id.** Rationale: it is the only option proven to work at Oxford without IT; it is the standard, widely-used approach for exactly this (IMAP+OAuth2 tools); there is no secret to store or rotate.

### The `SMTP.Send` scope in the returned bundle — why it is acceptable here

The spike's granted token carried `EWS.AccessAsUser.All IMAP.AccessAsUser.All POP.AccessAsUser.All SMTP.Send User.Read`. Mitigation is **architectural, not token-level**:
- `imap_mail.py` imports `imaplib` only. `smtplib` is never imported anywhere in the module. There is no code path — not in retry, not in error recovery — that can construct an SMTP session.
- There is **no autonomous agent with tools** on this path (this is the entire difference from the rejected ChatGPT-connector route). It is our own deterministic Python doing read-only `EXAMINE` + `FETCH` + `SEARCH`.
- We request scope `https://outlook.office365.com/IMAP.AccessAsUser.All` only; MSAL still receives the broader bundle back because of how Thunderbird's client is registered, but nothing consumes it.

**If a zero-send-capability token is later mandated:** that — and only that — is the trigger to attempt a dedicated single-tenant registration with delegated `IMAP.AccessAsUser.All` only, and it is the one scenario where asking Oxford IT would be worth it. **Not required to proceed.** Documented as a known residual, accepted.

### Token cache

- **Location:** `%LOCALAPPDATA%\WorkInboxAI\msal_imap_token_cache.bin` (same dir as the WS1 toast stamps). Outside the repo tree.
- **Never committed:** it lives in `%LOCALAPPDATA%`, and `.gitignore` additionally now lists `msal_imap_token_cache.bin` and `*.bin` defensively.
- **Refresh:** every run `PublicClientApplication(...).acquire_token_silent(SCOPES, account=accounts[0])`; cache re-persisted only if `has_state_changed`. Survives reboot (confirmed in the spike).
- **Re-auth failure path:** `acquire_token_silent` returning `None` / an `error` dict (`invalid_grant`, `AADSTS50173` token-revoked, `AADSTS700082` refresh-token expired past its ~90-day rolling window) → `imap_mail` raises `ImapReauthRequired` → `fetch_inbox.py` logs it, raises ONE 1/hour toast, exits `1`. The periodic re-auth still happens (this device has no Primary Refresh Token — see the no-PRT confirmed-fact memory) but as a **loud, non-blocking notify** instead of a silently wedged Outlook GUI.

---

## 4. Re-auth helper (`reauth_imap.py` + `Re-auth Work Inbox IMAP.bat`)

- Device-code flow: prints a short code + `https://microsoft.com/devicelogin` (the canonical URL — the shortlink `login.microsoft.com/device` misbehaved in the spike), waits (`acquire_token_by_device_flow` blocks up to ~15 min), writes the refreshed token to the same cache, then **verifies** with a read-only `EXAMINE INBOX` and prints the message count.
- PowerShell 5.1-callable: `python .\reauth_imap.py`, or double-click `Re-auth Work Inbox IMAP.bat` (which also `git checkout origin/main -- reauth_imap.py imap_mail.py` first, then `pause`s so Kevin sees the result).
- Prints timestamps on start, token acquisition, verification, done.
- Kevin runs it (a) once before `MAIL_BACKEND=imap` is ever used, (b) whenever the "mail sign-in expired" toast fires.

---

## 5. Verification design — parallel run + field-by-field diff (cautious-change-pace)

Per the 17 Aug 2026 regression+revert: **no cutover until parity is demonstrated over several real cycles.**

### Capture (same scheduled window, both backends)
```
set WI_MAIL_PARALLEL=1
set MAIL_BACKEND=com   & python fetch_inbox.py     # writes data/parallel/com_*.json, pushes nothing
set MAIL_BACKEND=imap  & python fetch_inbox.py     # writes data/parallel/imap_*.json, pushes nothing
```
(For a genuine same-window comparison these should run back-to-back, ideally wired as one extra step in a *copy* of the run wrapper used only for testing — not the live scheduled task.)

### Diff
```
python diff_mail_pull.py
```
Compares, per message (matched on `Message-ID`, falling back to subject+sender+time):
- `subject`, `from_email`, `is_read`, `has_attachments`, `importance`, `kevin_is_primary_recipient` — exact
- `received` — within 120 s
- **derived tier** (`categorise()` logic mirrored in the differ) — exact
- set differences: messages present in only one backend

Writes `data/parallel/parity_<ts>.json` + a console report. Exit 0 only on total parity.

### Deployment (done 28 Aug 2026 — the parity test is self-contained)
The scheduled `\Work Inbox Briefing` task pulls **only** `fetch_inbox.py` fresh from `raw.githubusercontent.com/.../main/` into the run dir `C:\Users\admin\Documents\Claude\Projects\work-inbox` — it does not fetch sibling modules. So:
- `imap_mail.py`, `reauth_imap.py`, `diff_mail_pull.py` were pulled from `main` into the run dir directly, and `fetch_inbox.py` there was refreshed to `main` (byte-identical `com` behaviour; a `.backup-*-pre-mailbackend-siblings` copy was kept). All four compile in place.
- **`D:\OneDrive - lelitte.com\Desktop\Re-auth Work Inbox IMAP.bat`** — primes the token. Pulls `imap_mail.py` + `reauth_imap.py` fresh into the run dir, then runs `reauth_imap.py`. Reference copy: `docs/desktop-scripts/Re-auth Work Inbox IMAP.bat`.
- **`D:\OneDrive - lelitte.com\Desktop\Run Mail Parity Test.bat`** — one-click parity run. Pulls `fetch_inbox.py` + `imap_mail.py` + `diff_mail_pull.py` fresh into the run dir, runs the `com` capture, the `imap` capture, then `diff_mail_pull.py`. Sets `MAIL_BACKEND` per-process only; unsets it before the diff. Reference copy: `docs/desktop-scripts/Run Mail Parity Test.bat`.
- The run-dir local git clone is heavily drifted (HEAD `c8ab371`, large uncommitted delete diff) — the `.bat`s deliberately use `curl` raw pulls, never `git pull`, on that dir.

### Exact commands for Kevin (PowerShell 5.1)
```powershell
# (a) prime the IMAP token once — approve the device code in a browser
& "D:\OneDrive - lelitte.com\Desktop\Re-auth Work Inbox IMAP.bat"

# (b) run the COM + IMAP parallel capture and diff (classic Outlook must be
#     running and "Connected to: Microsoft Exchange" for the com half)
& "D:\OneDrive - lelitte.com\Desktop\Run Mail Parity Test.bat"

# (c) parity output:
#     C:\Users\admin\Documents\Claude\Projects\work-inbox\data\parallel\
#       com_inbox_raw.json / imap_inbox_raw.json / com_sent_raw.json / imap_sent_raw.json
#       parity_<timestamp>.json  (+ the console report the .bat echoes)
Get-ChildItem "C:\Users\admin\Documents\Claude\Projects\work-inbox\data\parallel"
```
Repeat (b) across 3–4 windows over 2–3 days. The `.bat` pushes nothing and mutates nothing.

### Acceptance gate before cutover
1. `diff_mail_pull.py` shows **no `only_in_*` and no field mismatches** (bar the known-benign `from_email` X.500→SMTP improvement and sub-second `received` jitter) across **at least 3–4 scheduled cycles over 2–3 days**.
2. Kevin and/or Lauren eyeball a couple of those `imap` briefings on the dashboard for "does this look right".
3. **Dashboard JS opener branch shipped** (`mail_backend === "imap"` → OWA-web opener) and screenshot-approved by Kevin — otherwise IMAP cards have a dead `openmail://`.
4. Phase 3.9 re-wired to key on `message_id` + use `imap_mail.message_still_in_inbox()` (follow-up #1), OR an explicit accepted decision to run with Phase 3.9 in fail-open-carry mode initially.
5. **Kevin gives a fresh explicit go-ahead for the cutover step specifically.**

### Cutover (only after all of the above)
- Update the run wrapper (`Run Inbox Briefing.bat`, live Desktop, not repo-tracked) to `set MAIL_BACKEND=imap` **and** to add a `curl` raw-pull of `imap_mail.py` into the run dir alongside the existing `fetch_inbox.py` pull (otherwise `imap_mail.py` is never freshened before a scheduled run — same reason the parity `.bat` pulls it).
- Timestamped `.bat` backup first, same convention as the 28 Aug preflight changes.
- Watch the next 2–3 live runs. Rollback = flip `MAIL_BACKEND` back to `com` (or unset) in the `.bat` — one line, instant, no code revert needed.

---

## 6. Files in this build

| File | New/changed | Risk |
|---|---|---|
| `imap_mail.py` | **new** | none — dead code unless `MAIL_BACKEND=imap` |
| `reauth_imap.py` | **new** | none — only run by hand |
| `diff_mail_pull.py` | **new** | none — read-only analysis tool |
| `docs/desktop-scripts/Re-auth Work Inbox IMAP.bat` | **new** | none — reference copy; live copy on the Desktop |
| `docs/desktop-scripts/Run Mail Parity Test.bat` | **new** | none — reference copy; live copy on the Desktop |
| `fetch_inbox.py` | flag + 4 loop guards + imap injection + non-fatal COM under `imap` + `mapi is None` calendar guard + `_imap_reauth_toast_due()` | **`com` path byte-identical**; `imap` path is new and unverified |
| `.gitignore` | added `msal_imap_token_cache.bin`, `*.bin` | none |
| `docs/PHASE1_IMAP_MIGRATION_AUDIT.md`, this file | **new** | none |

Deployed to the machine 28 Aug: the three `.py` siblings into the run dir; both `.bat`s onto `D:\OneDrive - lelitte.com\Desktop\`. The stale `git fetch`-based repo-root `Re-auth Work Inbox IMAP.bat` from the first commit was removed — the `docs/desktop-scripts/` copies are canonical.

**Restore point for the `fetch_inbox.py` change:** `main` `9a52b07`, `fetch_inbox.py` blob `bd02b41089850678b8268318a0afab5e6d457e8a`, snapshot `Archive/fetch_inbox_backup_20260828_*_pre_mail_backend_flag.py`. Rollback = `git revert <this commit>` or restore the blob.

---

## 7. Open decisions for Kevin (none block the parity run)

1. **Graceful calendar degradation under `imap`.** This build makes a dead classic Outlook degrade the calendar to *empty + warning* under `imap` (instead of failing the whole run). That is what delivers the resilience benefit, but it means a broken calendar is less loud. Acceptable? (Alternative: keep calendar failure hard, losing half the benefit.)
2. **Phase 3.9 initial mode.** Ship the `message_id` re-wire (follow-up #1) *before* cutover, or accept fail-open-carry for the first week and re-wire after?
3. **`SMTP.Send` in the token bundle** — accept the architectural mitigation (recommended), or hold out for a dedicated app registration (needs Oxford IT, currently ruled out)?
