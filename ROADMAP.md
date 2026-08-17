# work-inbox — Roadmap

**Last updated:** 2026-06-25
**Status:** Active — pipeline fully working, roadmap clear

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
