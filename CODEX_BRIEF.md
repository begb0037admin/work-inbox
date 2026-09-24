# Codex brief — work-inbox: single section toggle (Kevin, 24 Sep)

Branch `drew/wi-cc-label-and-columns` (HEAD `be713fb`). Work only in C:/Users/admin/github/work-inbox. Only touch `js/app.js`, `css/styles.css`, `HANDOVER.md`. Don't read other repos, `data/`, `Archive/`, `js/vendor/`. Keep UTF-8 intact. Commit locally; no push; short final message.

Existing code: `getCollapsedSecs()`/`toggleSecCollapse()`/`applySecCollapse()` (`workInbox_collapsedSecs_v1`, header onclick + arrow), and `priExpandSection()` + the `.expand-all` header buttons (bulk drawer expand — remove).

## Kevin's rule (24 Sep, final): ONE section toggle per section — no bulk card expand
Kevin: "I would expect it to expand the section and click it again to collapse the section... not
expand the already opened tiles and show me the information. I can do that with whichever tile I
want to read."
- Replace the per-section "Expand all"/"Collapse all" (which bulk-opens card drawers) with a
  SECTION toggle button in the section header. Label: **"Collapse"** when the section is open,
  **"Expand"** when it is folded. Clicking folds/unfolds the whole section (all its cards hidden /
  shown). Merge it with the existing section-fold mechanism (reuse the existing remembered
  collapse state and storage key so Kevin's current folded/open choices carry over) so there is
  ONE clear control per section: remove the separate small fold chevron/arrow glyph (or put the
  glyph inside the same button, e.g. "Collapse ▾" / "Expand ▸"). Clicking the header row itself
  may keep toggling the same state, but the button is the visible control and its label must
  always match the state (after reload, drag, re-render, and header click).
- When folded, keep the header showing the section name and its card count, so it's obvious the
  section has hidden cards.
- Remove the bulk card-drawer expand function and its button entirely. Each card's own › still
  opens/closes that card's details, remembered per card as now.
- Button: `type="button"`, `aria-expanded` true/false, `aria-controls` the section's list,
  keyboard Enter/Space, `event.stopPropagation()` so it doesn't double-toggle with the header.

Checks: `node --check js/app.js`; UTF-8 non-ASCII unchanged. Top-of-HANDOVER entry.
