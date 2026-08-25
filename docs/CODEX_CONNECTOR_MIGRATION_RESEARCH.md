# Codex Connector Migration — Research & Decision Log

**Status:** Phase 1 verified and reviewed against real committed output. Opener-migration
design drafted below. Phase 2 not yet briefed — one follow-up check needed first (Section 8).
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
  dependency. **Confirmed live in Section 7 below — no longer a hypothesis.**
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
    validation. **Confirmed in Section 7: Graph expands recurring events into
    `occurrence`/`exception` items with their own IDs — a different
    complexity to design around, not the same locale bug, but real.**

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

**Phase 1 — connector verification + read-only fetch (Codex). COMPLETE.**
Run 2026-08-25 via `codex exec -s read-only`, output committed as
`docs/phase1_result.json` / `docs/phase1_brief.txt` (commit
`c28c19166d40ee98072b804257592be607811ed6`). No writes made. See Section 7
for findings and Section 8 for the one follow-up before Phase 2.

**Phase 2 — AI-triage migration (Codex), gated on Phase 1 review.**
Re-implements the six AI phases on Codex's model. Hard constraints: output
only to new files (`codex_briefing.json`, `codex_suggestions.json`,
`codex_triage_ledger.json`) — never overwrites `briefing.json`, `tasks.json`,
or the existing `triage_ledger.json` directly; uses a separate dedup-ledger
namespace so it never collides with the COM-keyed existing ledger; no
sends/drafts/calendar writes/message posts; runs in parallel with the
existing Anthropic pipeline for a validation period rather than an
immediate cutover. **Not yet briefed — see Section 8.**

**Opener-migration design (resolved, see Section 7 for why):** Command
Centre's Open-email button needs to branch on a `source` field per task,
not assume one ID format:
- Existing tasks (no `source` field, or `source: "outlook-com"`) — keep
  using the current `openmail://<entryId>` → `open_email.py` →
  `GetItemFromID` path. Nothing about these changes.
- New Codex-sourced tasks (`source: "codex-graph"`) — do **not** attempt to
  route these through `GetItemFromID` at all; it will not accept a Graph
  id. Instead, store whatever the connector's `webLink` field provides (see
  Section 8) and open it as a plain hyperlink in the browser (Outlook Web
  Access), no custom protocol handler needed. This is simpler than the
  original worry — it avoids building any Windows-registered equivalent for
  Graph IDs.
- `command-centre/data/tasks.json` gets one new optional field, `source`,
  defaulting to absent/`"outlook-com"` for every existing task (no
  migration needed for old data) and set to `"codex-graph"` only on tasks
  Phase 3.6-equivalent Codex logic creates.
- `triage_ledger.json` stays untouched by Codex; Codex's dedup ledger is a
  separate file (`codex_triage_ledger.json`) keyed on Graph `id`, so the two
  dedup systems coexist without collision, per Section 5's Phase 2
  constraints.

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
it cannot run `codex exec` itself. Phase 1 was run from a **local** terminal
Claude Code session on the admin machine instead (dispatched as "Drew," the
accountable lead for work-inbox per `AGENT_DIRECTORY.md`).

**Going forward: results should flow back via a GitHub push to this
branch, not via pasted chat content.** The cloud session watches this PR
and reads pushed files directly — see Section 7, read straight from the
real committed `phase1_result.json`, not a relayed summary.

## 7. Phase 1 findings (verified against the real committed file, not a relayed summary)

Read directly from `docs/phase1_result.json` at commit
`c28c19166d40ee98072b804257592be607811ed6`:

**a. Live scope** — read confirmed across Inbox + 63 other folders, Sent
Items, calendar (list/search/recurrence, 1,825-day connector search limit),
Teams (chats/channels/messages/meeting transcripts/recordings). Write-capable
actions exposed but not invoked: email draft/send/reply/forward/move/
categorize/create-folder; calendar create/update/cancel/RSVP/add-attachment;
Teams create chat or channel/send/reply; Planner create/update/delete. No
email-delete action was exposed at all. **Important: this write-capable
surface exists at the connector level — it was kept unused by the
`-s read-only` sandbox flag, not by an inherent account restriction.** Phase
2's tooling must enforce the same discipline explicitly; it cannot assume
the connector itself is read-only.

**b. ID formats — the answer Section 3 hinged on:**
- Email: field `id`, Graph-style opaque value. An `internetMessageId` /
  `internet_message_id` field was explicitly requested and did **not** come
  back in either casing — there is no bridging ID between this connector and
  anything COM-based.
- Calendar: fields `id` (Graph-style) and `iCalUId` (present in schema, but
  `null` on every live event returned).
- Teams: chat `id` like `19:meeting_...@thread.v2` or `19:...@unq.gbl.spaces`;
  canonical message routes `/chats/{chat_id}/messages/{message_id}` and
  `/teams/{team_id}/channels/{channel_id}/messages/{root_id}` (with
  `/replies/{reply_id}` for channel replies).
- **Confirms the Section 3 risk exactly: this is not an Outlook-COM EntryID,
  `mapi.GetItemFromID` will not accept it, and the obvious dedup workaround
  (match on `internetMessageId`) is not available either.**

**c. Bounded pull — partially incomplete, honestly flagged by the run
itself:** server-side the pull succeeded (50 inbox / 20 sent / 144 calendar
events / 100 Teams chats), but the full literal per-record JSON (314 records)
was truncated by the connector's response transport at ~50,062 tokens before
it could be written out. What's actually in `phase1_result.json` is schema
shape + counts + error stats, not record-by-record JSON. A chunked re-run
would be needed to get the full literal data, but is **not required** to
proceed with the opener-migration design above — the ID-format answer is
what mattered, and that came through clearly.

**d. Failures:** 1 Teams membership lookup denied (`403 Forbidden /
InsufficientPrivileges` — account not in that chat's roster), 6 rate-limited
(`429 TooManyRequests`) — `member_count` missing for 7/100 chats. No email or
calendar failures. Calendar recurrence expands into `occurrence`/`exception`
events with their own IDs and typically `null` `recurrence` on the expanded
item — status must be read from `type`/`seriesMasterId` instead.

**No writes occurred.** Only `docs/phase1_brief.txt` and
`docs/phase1_result.json` were added to this repo; `data/briefing.json`,
`data/tasks.json`, and `data/triage_ledger.json` are untouched, confirmed via
diff.

## 8. One follow-up before Phase 2 is briefed

Phase 1 didn't check whether the connector's email/calendar responses
include a **`webLink`** field (the standard Microsoft Graph property that
opens an item directly in Outlook Web Access). The opener design in Section
5 assumes this exists — it's a standard Graph field, but per this project's
own rule ("verify live, don't assume"), it needs one direct confirmation
before being relied on for the Open-email button.

**Next single action:** have Codex answer one question — "does the message/
event JSON already pulled in Phase 1 include a `webLink` field, and what
does a sample value look like?" — via `codex exec -s read-only` against the
same account, no new pull needed (it can inspect what Phase 1 already
returned in its own session, or do one bounded re-pull of 5 items if that
context isn't available). Once that's confirmed, Phase 2 can be briefed
using the Section 5 opener design as-is.

## 9. Status / next action

Confirm `webLink` (Section 8) → brief Phase 2 with the Section 5 design
folded in → run Phase 2 via the local terminal session/Codex → review its
output the same way Phase 1 was reviewed here (read the real committed
file directly, not a relayed summary).
