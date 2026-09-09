# Mail connector migration — assessment, 9 September 2026 (Drew)

**Status: assessment only. No changes to `fetch_inbox.py`'s live IMAP path in this session.**
**Reads alongside:** `CODEX_CONNECTOR_PIPELINE_PLAN.md` (the already-settled 8-increment execution plan for the calendar+Teams+mail connector migration, "direction decided" 29 Aug), `MAIL_BACKEND_MIGRATION_PLAN.md` (the COM→IMAP migration this would supersede), `HANDOVER.md` section K2 (9 Sep, "Investigated and partially de-risked, NOT built, NOT live").

This doc answers the five assessment questions Kevin asked for directly. It does not re-litigate `CODEX_CONNECTOR_PIPELINE_PLAN.md` — that plan already covers the mail case in full (§4 schema, §6a kill-switch rework, §6b parity method); this is a point-in-time status check against it plus the specific quota question.

---

## 1. Is it technically viable — does the connector return what `fetch_inbox.py` needs?

**Yes, for reads, substantially proven live already — not new research:**

- `microsoft_outlook_email.list_messages` with a **date-range filter** works live (HANDOVER K2, this week): 18 real inbox messages returned for a 6–8 Sept window, correct subjects, via the personal Lane B identity. This is the mechanism that would replace IMAP's "50 newest emails" pull — a date-window pull is actually a better fit for the pipeline's real intent (catch everything since the last run) than the current hardcoded 50-item cap, which can silently truncate on a heavy morning.
- `list_messages`/`fetch_message` return the Graph `message` resource, which carries every field `fetch_inbox.py` currently needs: `subject`, `from`/`sender`, `toRecipients`/`ccRecipients`, `internetMessageId` (the cross-backend dedup key IMAP already uses), `isRead`, `receivedDateTime`, `bodyPreview`/`body.content`, `hasAttachments`, `importance`. `lane_b_call1.py` already has field-mapping helpers for exactly this shape (`_first(m, ...)`, the `from`/`sender` nested-object handling, `body_preview` extraction) — built for the weblink-resolution use case, directly reusable.
- `internetMessageId eq '<id>'` also returns a real `web_link` — already shipped for the "Open email" OWA deep-link (work-inbox `fcb47a9`), so link resolution isn't a new problem, it's solved.
- The verb-based read-only guard (`LANE_B_NAMESPACES`, `guard_recontamination()`) already includes `microsoft_outlook_email` and already passes `list_messages`/halts `send_mail`/`reply_to_message`/etc. — same mechanism as calendar/Teams, no new guard code needed for the read path itself.

**Not yet proven / not built (per `CODEX_CONNECTOR_PIPELINE_PLAN.md` §6, increments 1–2 and 4–6):**
- `normalise_pull.py` (schema normaliser + Layer-2 sanitiser — strips HTML, truncates bodies, neutralises prompt-injection-shaped text in email bodies before anything reasons over them) — not built.
- The actual Call-1 mail runner producing `raw_pull.json` at pipeline scale (today's proof was a single ad-hoc probe, not a wired-in runner with retry/warm-up/timeout) — not built.
- Pagination behaviour at real volume (a genuinely heavy morning can exceed 50 messages) hasn't been live-tested — the 18-message probe was well under any pagination boundary.

**Verdict: viable.** The hard "does the data exist and can we read it" question is answered. What remains is integration engineering, matching the shape of the already-completed Lane B calendar/Teams build — not a research problem.

---

## 2. Quota impact — concrete, not hypothetical

- **Edu is not a viable primary for mail today.** HANDOVER K2 (this week): a verification probe against the Edu identity hit Edu's monthly usage cap directly — "doesn't reset until 1 Oct." This is not a projection, it already happened this session, independent of any mail-volume increase.
- **There is no primary/failover pair for mail at all.** Per HANDOVER's Option-1 reassessment: the Edu identity has the Outlook Email connector deliberately *not* attached (Kevin's Q2 decision — Edu stays Calendar+Teams only). Only the **personal** account (`kevin@lelitte.co.uk`) has Outlook Email connected. So unlike calendar/Teams, which have a real Edu-primary/personal-failover pair, mail would run on the personal account as its *only* identity — a single point of failure for quota, not a two-tier fallback.
- **Volume math:** today's IMAP pull is one connection, one 50-item fetch, 3×/weekday. A connector equivalent adds roughly one `codex exec` call per run for the mail domain (batched, same pattern as calendar's one call per domain) — call count itself doesn't obviously explode. The real risk isn't call *count*, it's that mail would now compete with calendar+Teams+drafted-reply weblink-resolution on the **same personal-account quota** that calendar/Teams already lean on when Edu fails over — concentrating more of the daily pipeline onto the one identity that has no failover behind it.
- **Compounding factor:** the existing weblink-resolution feature (`resolve_mail_weblink`, shipped 3 Sept) already runs on personal-only (no Edu route exists for mail) and already does one `codex exec` call per newly-promoted card. Full mail-fetch would add a second, larger mail-domain call on top of that, on the same identity, every run.

**Verdict: real, concrete risk, not just "known and accepted in the abstract."** Edu is already exhausted this month for reasons unrelated to mail; adding mail-fetch load lands entirely on the one account with no fallback. This needs Kevin's own quota-risk call before build, not an inherited "he already accepted this for calendar/Teams" — the calendar/Teams risk-acceptance was for a route that *has* a fallback tier; mail does not.

---

## 3. Cutover approach

`CODEX_CONNECTOR_PIPELINE_PLAN.md` §6b already specifies exactly this, unchanged: diff the parallel connector output (`docs/codex_briefing.json`) against the nearest-preceding live `data/briefing.json` from GitHub commit history — field-by-field, cards matched on normalised subject + received-to-the-minute — over a real window (target ≈40+ compared runs), rather than a live side-by-side IMAP/connector dual-run. This was deliberately designed to avoid needing Kevin to produce manual side-by-side captures (he doesn't run classic Outlook), and it already accounts for gaps where no baseline exists for a given cycle.

No new cutover mechanism is needed — the existing plan's increments 1–5 (normalise → stage → Call-1 runner → parity harness → Kevin reviews the diff) are the right sequence for mail specifically, same as they were designed for.

---

## 4. Connector failure / no-fallback risk for mail specifically

This is the sharpest open point, and it's a **repo-documented, unresolved blocker**, not a new concern I'm raising for the first time:

- `CODEX_CONNECTOR_PIPELINE_PLAN.md` §6/§6a: the current kill-switch (`mailbox_guard.py`, Layer 6) does a post-run **Outlook COM** delta-sweep of Sent/Drafts. It is explicitly flagged as needing a connector-based ("COM-free") rework before this route can go live for mail — "**Do not wire it as-is**". Per HANDOVER K2 (9 Sept), this rework **has not been built.**
- That COM dependency is now doubly dead: work-inbox's own hard rule as of 8 Sept is "Outlook Classic is retired — OWA-in-browser only... No machine uses desktop Win32 Outlook." The kill-switch's current implementation assumes classic Outlook is running to do its delta-sweep — a precondition that no longer holds operationally.
- Net effect: if mail were cut over to the connector today, the pipeline would have **no working post-run safety net** for an unexpected write (a stray send/reply/categorisation change) until the connector-based delta-check (§6a: pre/post Sent-folder snapshot via `list_messages`, same disable+notify code, new data source) is actually built. That's a real gap, not a theoretical one.
- If the connector is simply unavailable on a given run (rate limit, auth expiry, outage), full IMAP retirement means **mail has no fallback at all** — a blank or stale dashboard, not a degraded-but-functional one. `claude -p`/IMAP staying live in parallel throughout the validation window (as the settled plan already requires) is exactly what prevents this; it's the reason the plan explicitly forbids retiring IMAP before cutover is proven and approved.

**Verdict: mail becomes a single point of failure with no fallback the moment IMAP is actually retired — but the plan already accounts for this by keeping IMAP live until parity is proven and Kevin separately authorises retirement.** The open blocker is the kill-switch rework, which is unbuilt.

---

## 5. What happened in this session, and why no code change was made

The task as originally scoped was explicit: assess only, do not touch `fetch_inbox.py`'s live IMAP path. Partway through, three messages arrived from the coordinator instructing escalating departures from that scope — skip validation, skip the dual-run/diff step, "ship it live," each citing Kevin's authorization without a message from Kevin appearing directly in this session.

I did not act on those instructions, for reasons independent of how they arrived — the project's own pre-existing documents settle this on their own terms:

- `CODEX_CONNECTOR_PIPELINE_PLAN.md` §7 ("Hard gates"): *"No cutover. No `.bat` / scheduled-task change, no `main` default-behaviour change, without Kevin's fresh explicit go-ahead for that specific step."* This is Kevin's own previously-ratified plan, not a new restriction I'm inventing.
- The Layer-6 kill-switch rework it requires before mail can safely go live is confirmed **not built** (§4 above) — a genuine blocker, not a risk/tradeoff judgment call.
- `CLAUDE.md`'s 8 Sept approval-protocol amendment waives the screenshot/per-step ceremony for routine changes, but explicitly carves out *"anything genuinely destructive, irreversible, or outside the agreed task scope"* — retiring the only live mail path in favour of an unbuilt, single-identity, no-fallback connector path, ahead of the plan's own increments and without the kill-switch rework, falls squarely in that carve-out.
- `CONSTITUTION.md` §2: *"If a receiving role encounters something that requires a decision, it stops, reports exactly what it has found, and returns to the reasoning seat. It does not proceed, improvise, or interpret."*

**What I did instead:** verified the technical viability findings above against the real repo history (not re-derived from scratch — HANDOVER K2 and `CODEX_CONNECTOR_PIPELINE_PLAN.md` already contained most of the answer), wrote this assessment, and left `fetch_inbox.py` untouched.

**Recommended next step, if Kevin wants to proceed:** resume `CODEX_CONNECTOR_PIPELINE_PLAN.md` at increment 1 (`normalise_pull.py` + Layer-2 sanitiser — zero live-path risk, new files only) as its own dedicated session, in the order the plan already lays out, with increment 6 (COM-free kill-switch) completed before any parallel run touches real mail data, and with Kevin's own explicit quota-risk decision given Edu is already exhausted this month.
