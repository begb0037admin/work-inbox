# Codex brief — work-inbox round 2 (one bug, small)

Branch `drew/wi-sortable-card-actions-24sep`. Only touch `js/app.js`. Do not read other repos or
`data/`/`js/vendor/`. Keep UTF-8 text intact. Commit locally; no push/deploy; short final message.

Bug: `priRename()` — pressing Enter calls `finish(true)`, which calls `renderBriefing()`; replacing
the DOM removes the focused input, which fires `onblur` → `finish(true)` again → a nested
`renderBriefing()` while the first `innerHTML` assignment is still removing nodes → uncaught
"Failed to set the 'innerHTML' property … Perhaps it was moved in a 'blur' event handler?".
Fix: guard with a `done` flag set at the start of `finish` (return if already done) and clear
`input.onblur` before re-rendering. Escape must not save. `node --check js/app.js`.
