# work-inbox — Roadmap

**Last updated:** 2026-09-01
**Status:** Active — pipeline fully working, roadmap clear

---

## Backend migration (two-lane) — status as of 2026-09-01

Kevin's locked decision (via coordinator, 1 Sept 2026): **email = IMAP (Lane A), calendar = ChatGPT M365 connector (Lane B), Teams = ChatGPT M365 connector (Lane B).** Build order: calendar, then Teams. Connector write-risk (`docs/CONNECTOR_SAFEGUARDS.md` §D) explicitly accepted by Kevin and closed. The 1 Sept (c)+(e) "ship Lane A only, park Lane B" recommendation is **rejected**.

**Headless connector = PROVEN WORKING on the Edu account (1 Sept, coordinator, direct on the admin desktop).** The three "MAKE-OR-BREAK" failures were false negatives: a bare tool-enumeration does not surface the lazily-loaded `codex_apps` connector tools, and the connectors had been toggled off in the Edu account (now re-enabled). codex-cli version is not a factor. `codex exec -s read-only --json` on `begb0037@ox.ac.uk` returns real `mcp_tool_call` results from `microsoft_teams.list_chats` and `microsoft_outlook_calendar.list_events`.

| Lane | Scope | Status |
|---|---|---|
| **A** | Mail pull via IMAP+OAuth2 (`MAIL_BACKEND=imap`) | Proven (Phase 2(i)); shipped behind default-unset flag; **live now via the laptop bridge briefing**; formal Phase 6 cutover still pending |
| **B — calendar** | `CAL_BACKEND=connector`, Phases 3.7/3.8 fed by a `codex exec --json` Call-1 against `codex_apps` (`microsoft_outlook_calendar.list_events` etc.), `normalise_pull.py` sanitiser, **calendar kill-switch HALTs on any change**, re-contamination guard asserts on the actual `mcp_tool_call` server/tool events | **BUILDING.** `normalise_pull.py` shipped (`3d59dc3`). Next: `lane_b_call1.py` (runner + JSONL parser + guard), `lane_b_cal_guard.py`, then the `fetch_inbox.py` `CAL_BACKEND=connector` wiring. **Edu (`begb0037@ox.ac.uk`) is the PRIMARY Lane B identity; personal ChatGPT Plus is failover only.** Config-hash baseline for the guard: `4FD8EF763BF0A8DDAD9A138B6679A84FE8536F73`. |
| **B — Teams** | new Teams briefing section, same `codex_apps` surface (`microsoft_teams.*`) + safeguards | Queued behind calendar (build order: calendar first) |

See `HANDOVER.md` top entry (1 Sept, "CORRECTION: MAKE-OR-BREAK #2 AND #3 WERE FALSE NEGATIVES") for the verified evidence and the real tool surface.

---

## Current State — Pipeline Complete ✅

| Phase | Description | Status |
|---|---|---|
| 1 | Outlook COM pull — inbox, sent, calendar | ✅ Complete |
| 2 | Anthropic API triage → briefing.json | ✅ Complete |
| 3 | Task Scheduler — 7am/9am/11am/1pm/3pm/5pm Mon-Fri | ✅ Complete |
| 3.5 | Inbox suggestions → command-centre (draggable cards) | ✅ Built — confirmed unused in practice; retirement planned (see below) |
| 3.6 | Auto-apply dated action entries to command-centre tasks.json | ✅ Complete |
| 4 | Multi-machine — replicate on work machine (begb0037.AD-OAK) | ✅ Complete |

---

## Planned — AI Chat Panel ⏳ (post command-centre migration)

**Prerequisite:** command-centre file split & Cloudflare Pages migration complete and stable. work-inbox follows the same migration pattern once command-centre is proven (per migration plan dated 2026-06-25).

**Summary:** Add an embedded AI chat interface to the work-inbox dashboard. Kevin can type freeform notes about inbox items, emails, or calendar events directly in the dashboard; Claude responds and logs relevant actions — no separate Claude session required.

### What gets built

| Component | Detail |
|---|---|
| "Ask Claude" nav item | New entry in the sidebar nav |
| Chat panel (main area) | Multi-turn conversational UI alongside the existing briefing view |
| `js/chat.js` | Chat UI logic and thread management (clean new file in modular codebase post-migration) |
| Worker `/chat` route | Shared route on `cc-tasks-writer` — same route as command-centre chat; briefing.json context passed alongside tasks context |
| `data/chat_history.json` | Persistent rolling conversation history (~20 exchanges). Separate file from command-centre's history. GitHub-backed — works from any browser/machine. |
| `ANTHROPIC_API_KEY` | Shared Worker secret — added once for both dashboards |

### Behaviour — Phase 1

- **Freeform input** — type anything about inbox items, emails, or calendar; Claude asks clarifying questions if needed
- **Context-aware** — Claude receives current `briefing.json` alongside the conversation history, so it knows what's in today's inbox without Kevin summarising it
- **Actions-only writes (Phase 1)** — can append dated action entries to command-centre tasks.json where an inbox item maps to an existing task. No writes to briefing.json in Phase 1.
- **Persistent memory** — last ~20 exchanges in `data/chat_history.json`; loaded on next visit

### Also in this phase

- Retire Phase 3.5 of `fetch_inbox.py` (inbox suggestions to command-centre — unused)
- Archive `data/inbox_suggestions.json`

### Governance gates

| Gate | Requirement |
|---|---|
| Before build | command-centre migration confirmed stable; work-inbox migration confirmed stable |
| UI change | Screenshot approved by Kevin before push to main |
| Worker change | Kevin approves shared `/chat` route and `ANTHROPIC_API_KEY` secret |

### Combined with command-centre

This is a joint feature. Both dashboards share the same Worker `/chat` route and `ANTHROPIC_API_KEY`. Each has its own `data/chat_history.json`. The intent is a unified AI assistant accessible from whichever dashboard Kevin has open. See `command-centre/ROADMAP.md` — Module 1.5.

---

## Future — Phase 2 Chat (extended authority)

Once Phase 1 chat is stable in both dashboards:

- Chat can move task tiers, update summaries, and add new tasks (not just append actions)
- Requires a separate planning session and governance gate before implementation
- Chat in work-inbox may gain authority to mark briefing items as actioned

---

## Deferred — Drag-and-drop mechanics improvements (Tier 2 / Tier 3)

**Recorded 12 Aug 2026.** Card drag-and-drop review (`memory/wi-dragdrop-review-12aug.md`) found the mechanism fragile — two real bugs (Show/Hide Done reset, title-collision card vanish) traced to the same root cause: every drag gesture triggers a full destroy-and-rebuild of all six board sections. Tier 1 (cheap fixes — rAF-batched drag movement, consistent cross-browser drag ghost via `setDragImage()`, reorder-boundary hysteresis) was approved and shipped 12 Aug 2026 (commit `9ef7f176`). Tiers 2 and 3 were reviewed but explicitly deferred — Kevin wants them recorded for a later decision, not built now.

| Tier | Scope | Why deferred |
|---|---|---|
| 2 — Moderate | Replace `priDragEnd`'s unconditional full rebuild with a targeted DOM patch — only the affected zones, only on an actual drop — matching the pattern `toggleTick` already uses successfully elsewhere in `js/app.js`. Meaningfully closes the "any drag wipes unrelated UI state" risk class without a full rewrite. | Real engineering effort, not a quick patch; worth a dedicated session |
| 3 — Most invasive | Replace the hand-rolled native HTML5 drag-and-drop with a proper library (e.g. SortableJS) or a state-driven diffed render. Removes the bug class structurally and adds free touch/mobile support. | Largest change of the three; needs a staged, Codex-reviewed build, not a same-session patch |

**Status:** Both await Kevin's go-ahead. Revisit if drag-and-drop bugs recur after Tier 1, or when there's appetite for a dedicated dashboard-engineering session.

---
