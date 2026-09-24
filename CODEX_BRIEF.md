# Codex brief — work-inbox: section counts must equal rendered cards (Kevin bug) + toggle style

Branch `drew/wi-section-counts` off main. Work only in C:/Users/admin/github/work-inbox. Only
touch `js/app.js`, `css/styles.css`, `tests/` (new), `HANDOVER.md`. Don't read other repos (except
the read-only style reference `C:/Users/admin/github/kevin-task-tracker/public/style.css`), `data/`,
`Archive/`, `js/vendor/`. Keep UTF-8 intact. Commit locally; no push; short final message.

## Bug (Kevin, screenshot): "Urgent – action required today" shows 3 but only 2 cards
Cause: `_secHeadHtml(sec,…,priSecs.X.length)` counts the section list BEFORE the renderer drops
cards: `_priRenderOneCard` returns '' for a true `del_<key>` tick, and handled (ticked) cards get
`card-hidden` when "Show done" is off. So deleted/handled cards (including a Command-Centre-backed
card deleted "here only") still count.
Fix: one pure predicate `_priCardVisible(p, ticks, showingDone)` (not deleted; and not handled
unless showing done) used by BOTH the renderer and the count — the count passed to
`_secHeadHtml` must be `priSecs.X.filter(visible).length`. (For FYI keep the existing
"threads (messages)" raw-count label logic but base `count` on the visible list.) Counts are
recomputed on every render, so they follow drag, archive, delete, undo, restore, section toggle.

## Regression test — `tests/section_count_test.js` (Node, no deps)
Extract `_priCardVisible` (and any tiny helpers it needs) from `js/app.js` by source (like
command-centre's `tests/tier_order_test.js` does) and assert: deleted → invisible; handled →
invisible unless showingDone; plain → visible; a section of [plain, handled, deleted] counts 1
(3 with showingDone? no — deleted never counts: 2).

## Toggle style (Kevin/Jacob: identical to the tracker)
The section toggle label becomes the tracker's style: small muted text, "Collapse ▾" when open,
"Expand ▸" when folded (same font size/colour/weight as the tracker's `.section-toggle-label`,
no border/background). Update wherever the label is set (render + `applySecCollapse`).

Checks: `node --check js/app.js`; `node tests/section_count_test.js`. Top-of-HANDOVER entry.
