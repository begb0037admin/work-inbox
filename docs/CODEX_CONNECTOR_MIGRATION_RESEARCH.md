# Codex Connector Migration — Research & Decision Log

**Status:** Research complete, Phase 1 dispatched to Codex for verification. No production
changes made yet.
**Effort level:** Raised to high, 2026-08-25 (Kevin's explicit confirmation), for this task.
**Started:** 2026-08-25, cloud Claude Code session.

This document exists so this research isn't lost to chat history (CONSTITUTION.md
Section 5 — conversation is temporary, documentation is permanent). Any session —
cold local terminal, cloud, or otherwise — should read this before touching the
Codex migration work.

---

## 1. Why this project exists

Oxford University has enabled ChatGPT/Codex connectors (Outlook, Calendar, Teams)
against Kevin's Oxford account. This is a potential replacement for the Outlook-COM
workaround (`fetch_inbox.py`, Task Scheduler, `win32com.client`) that `work-inbox`
has used because Oxford's MDM policy blocks the standard Microsoft Graph API
consent flow for that account (documented in `work-inbox/CLAUDE.md`: "Do not
re-investigate Graph API"). A Codex-based connector is a different access path
than the one that was blocked, not a reversal of that finding — worth stating
plainly since the two could be confused.

There is also a live cost incentive: the Anthropic API key used by
`fetch_inbox.py` cost ~$36.57 in August 2026 (94.5% of total Claude API spend that
month), almost entirely from six AI-triage phases run six times a day. If that
work moved to Codex under a subscription Kevin already pays for, the marginal
Anthropic cost could drop close to zero — see Section 4.

## 2. Account / domain model (confirmed with Kevin)

- **Codex (ChatGPT)** — Kevin's Oxford work connector. Both his personal Plus
  account and his university Edu account are (or will be) connected to his
  Oxford University Microsoft 365 account. He switches between the two
  subscriptions for capacity, the same way he does with Claude accounts.
  Codex = work domain, full stop.
- **Claude (this assistant)** — personal domain only. Connects to Kevin's
  personal Outlook inbox and personal Google inbox. Confirmed in-session: the
  Microsoft 365 connector available to this Claude session resolves to
  `kevin@lelitte.com` (personal), not the Oxford account — consistent with
  this model.
- The two must never cross. Command Centre, work-inbox, and meeting-records
  are all work-domain repos; any live connector data feeding them must come
  from Codex, never from a connector held directly by a Claude session.

## 3. Architecture dependency findings — what actually depends on the current fetch mechanism

Traced directly from `work-inbox/fetch_inbox.py` (2,819 lines) and
`work-inbox/open_email.py`. This is the real scope of any migration, not just
"swap where the data comes from":

- **`open_email.py` calls `mapi.GetItemFromID(entry_id)`** — a raw Outlook
  COM/MAPI call. Every "Open email" button in Command Centre and work-inbox
  goes through the `openmail://<entryId>` protocol to this function. A
  Microsoft Graph-based connector (which is what a Codex connector is) returns
  Graph message IDs — a different format. **This will not work with
  `GetItemFromID` as-is.** Any migration needs a new opener mechanism (or a
  dual mechanism, see Section 5) — this is the single biggest breaking
  dependency.
- **EntryID is threaded through the entire pipeline**, not just captured once:
  - Phase 1/1c (inbox, subfolder sweep, VIP sweep) — captured as `entry_id`
    on every card
  - Phase 3.2/3.3/3.3b (AI summaries, no-action demotion) — keyed by `entry_id`
  - Phase 3.5 (task-suggestion triage) — candidates keyed by `entry_id`
  - Phase 3.6 — writes `entryId` onto Command Centre's `data/tasks.json`
    (this is what powers the Open-email button per task)
  - `data/triage_ledger.json` — dedup state (`applied`, `promoted`,
    `tracked_needs_urgent`) all keyed by `entry_id`
  - **Phase 3.9** (scroll-out persistence) does *live* COM re-lookups every
    run — `GetItemFromID` → `item.Parent.EntryID` — to check whether a
    tracked email is still in the Inbox or was filed/archived. This is an
    ongoing dependency on COM identifiers staying resolvable, not a one-time
    capture.
- **Consequences of a naive swap:**
  - Existing tasks in `command-centre/data/tasks.json` have COM-format
    `entryId`. New tasks from a Codex/Graph source would carry a different ID
    format — the Open-email button needs to know which opener to use per
    task, or both mechanisms need to coexist.
  - `triage_ledger.json`'s dedup keys wouldn't match new-format IDs — already-
    actioned items could resurface as "new."
  - Phase 1c's named-subfolder sweep and VIP sweep are built on COM folder
    traversal (`Folders`/`FolderPath`); Graph's `mailFolders` API addresses
    folders differently — `SUBFOLDER_TREES` config needs remapping, not just
    a data-source swap.
  - The calendar `Restrict()`-locale workaround (direct iteration, 30-day
    back/6-day forward window, built specifically to dodge a UK-locale bug in
    `Restrict()+IncludeRecurrences`) is COM-specific; a Graph-sourced calendar
    has different recurrence-expansion behaviour requiring separate
    validation.

**Bottom line:** this is a migration of the canonical identifier the entire
pipeline is built around, plus a decision about what happens to already-stored
COM-format IDs in production data — not a Phase-1 fetch-mechanism swap.

## 4. Cost findings

- Anthropic spend this month: $38.69 total, $36.57 (94.5%) from the
  `work-inbox` API key. Driven by six AI phases (context, email summaries,
  no-action demotion, task-suggestion triage, task summaries, calendar-day
  prep) run six times daily via `fetch_inbox.py`.
- **Swapping only the data-fetch mechanism (Phase 1) does not touch this
  spend** — the AI work still runs on the same content volume regardless of
  how it was fetched.
- **Moving the AI-triage work itself to Codex is what could actually save
  money.** Kevin confirmed Codex authenticates via ChatGPT subscription (Plus
  and/or Edu), not a raw OpenAI API key — so usage isn't metered per-token the
  way Anthropic's is. Outlook/Teams connector sync specifically requires
  Team/Enterprise/Edu tier (per OpenAI's public connector docs, checked
  2026-08-25) — Plus alone doesn't get org-wide connector sync, though Kevin's
  Plus account is separately connected to his personal mailbox.
- Caveats not yet resolved:
  - OpenAI retired per-message pricing in April 2026 in favour of token
    credits even on paid plans — heavy automated/headless usage (6x/day) could
    hit plan limits, not truly "unlimited."
  - Automated headless invocation is a different usage pattern than
    interactive chat — worth confirming this doesn't run against any ToS/rate
    expectations before building the whole pipeline on it.
  - Re-implementing six phases of tuned logic in Codex is real engineering
    risk, independent of cost: the Phase 3.2 model choice is *locked* to
    Haiku 4.5 specifically because "Sonnet timed out on this inbox size"
    (documented in `work-inbox/CLAUDE.md`) — a hard-won constraint from past
    debugging, and the demotion/staleness/scroll-out logic (3.3/3.3b/3.3c/3.9)
    encodes multiple specific production bug fixes.

## 5. Plan

**Phase 1 — connector verification + read-only fetch (Codex).**
Brief sent to Kevin as `codex_briefs.md` (delivered via SendUserFile,
2026-08-25). Has Codex: verify its own live connector permissions (read vs.
write) against `kevin.lelitte@admin.ox.ac.uk`; report the exact ID format
returned for messages/events/Teams; do one bounded read-only pull (50 inbox,
20 sent, calendar ±week, Teams chat list) as JSON. No writes of any kind.

**Phase 2 — AI-triage migration (Codex), gated on Phase 1 review.**
Also in `codex_briefs.md`. Re-implements the six AI phases on Codex's model.
Hard constraints: output only to new files (`codex_briefing.json`,
`codex_suggestions.json`, `codex_triage_ledger.json`) — never overwrites
`briefing.json`, `tasks.json`, or the existing `triage_ledger.json` directly;
uses a separate dedup-ledger namespace so it never collides with the
COM-keyed existing ledger; no sends/drafts/calendar writes/message posts;
runs in parallel with the existing Anthropic pipeline for a validation period
rather than an immediate cutover.

**Not yet decided:** the actual opener-migration design (how Command Centre's
Open-email button handles two ID formats, or whether/how existing COM-format
`entryId`s get migrated). This depends on what Phase 1 reports back about the
real Graph ID shape, and should be designed once that's in hand — see Section
3.

## 6. Execution environment note (why this took multiple rounds to resolve)

Codex is invoked as a local subprocess — `codex exec -s read-only
-o <result-file>` — per Kevin's own ratified policy
(`agent-commons/operating-model/COORDINATOR_AND_CODEX_POLICY.md`). This only
works from a Claude Code session running locally on the machine where Codex
is installed and signed in (Kevin's Windows admin machine).

This research session is a **cloud-hosted** Claude Code session
(`environment_kind: anthropic_cloud`, confirmed via session metadata) —
opened from the desktop app, but executing in Anthropic's cloud, not on the
admin machine. It has no `codex` binary and no path to the admin machine, so
it cannot run `codex exec` itself. Phase 1 must be run from a **local**
terminal Claude Code session on the admin machine instead — that session
starts cold (no memory of this conversation), hence this document and the
`codex_briefs.md` file it should be pointed at.

## 7. Status / next action

Waiting on Phase 1 output from Codex (run via a local terminal Claude Code
session on the admin machine, per Section 6). Once that's back, review
against Section 3's opener/ID-format questions before authorizing Phase 2.
