# Codex brief — work-inbox dashboard: SortableJS drag, card actions, expand/collapse (Kevin, 24 Sep 2026)

Repo: this working tree, branch `drew/wi-sortable-card-actions-24sep` (checked out). Kevin approved
build + deploy. You implement; Drew reviews, browser-tests, deploys. **Dashboard only**: change
`index.html`, `css/styles.css` (additive rules), `js/app.js`; `js/vendor/Sortable.min.js` is already
vendored (1.15.6). **Do not touch** `fetch_inbox.py` or any Python/PowerShell/pipeline file,
`data/`, `Archive/`, `images/` (never touch `images/oxford-crest.jpg`), and never add any Microsoft
Graph / COM / IMAP / connector call — these actions act ONLY on dashboard data. Do not search or
print `js/vendor/`, `data/`, `Archive/`. Keep UTF-8 text intact (em-dashes etc.; never re-encode
the file). Commit locally; do not push or deploy. Short final message.

Kevin's instruction: **no layout changes** — keep the six sections (Urgent, Tomorrow, This week,
Today, Needs response, FYI/Parked), columns, card look and styling exactly as they are. Port ONLY
these behaviours from `C:\Users\admin\github\kevin-task-tracker\public\app.js` and
`C:\Users\admin\github\command-centre\js\app.js` (read-only references; command-centre just shipped
the same port — mirror its approach and icon SVGs).

## 0. Stable card identity (prerequisite — real bug)
Connector-era email cards have `entry_id: ""` but a `message_id`, so `_priGetKey()` falls back to
a title slug and `_tickStorageKey()` scopes that to the day (`currentKey+'_'+slug`) — a ticked card
reappears the next day. Add `if(p.message_id) return 'mid_'+p.message_id;` to `_priGetKey()`
(after `entry_id`/`entryId`, before `id`), and make `_tickStorageKey()` treat `mid_` as stable like
`eid_`/`id_`. Back-compat: section overrides already fall back to the legacy title key; add the same
legacy fallback in the order map (`om[k] ?? om[legacyTitleKey]`) and in `isTicked` (if the new key
is unset but the old day-scoped key `currentKey+'_'+legacyTitleKey` is true, treat it as ticked and
migrate it to the new key once).

## 1. Drag → SortableJS
Replace the hand-rolled HTML5 priority-card drag (`priDragStart/End`, `priCardDragOver/Leave/Drop`,
`priZoneDragOver/Leave/Drop` for card drags, the rAF reorder machinery, `draggable="true"` +
`ondrag*` attributes on `.card-ph`) with one Sortable per `.pri-drop-zone`: shared group,
`animation:150`, `forceFallback:true`, `fallbackOnBody:true`, ghost = dashed faded placeholder gap,
fallback clone solid white with blue 2px border + shadow, drag from anywhere on the card (no `handle`, like the tracker), `filter` for buttons/links/chevron with `preventOnFilter:false`, touch
`delay:150, delayOnTouchOnly:true`; destroy old instances on every re-render; defer re-renders
while a drag is active (tracker pattern). Within-zone reorder and cross-zone moves both allowed;
the card lands exactly where the placeholder shows (supersedes the old "cross-zone lands at top"
rule — Kevin wants the tracker's accurate drop). Persist exactly as today via `_priSetOrder` (DOM
order of every zone) + `_priSetOverride` for a cross-zone move. Toast "Moved to <section label>" /
"Order saved" with **Undo** (restores the previous order/override and re-renders). Keep the
email-card → board drop path (`emailCardDragStart/End`, `_emailDragData`, zone `ondrop` for that
case) working if that source is present.

## 2. Card actions (same icons as command-centre) — Kevin's 2x2 layout
- Remove the `.card-done-btn` circle. Keep the existing badges and `CC→` button where they are.
- **Layout (Kevin's mock-up, identical to command-centre's new layout — mirror its CSS/markup in
  `C:/Users/admin/github/command-centre/css/styles.css` and `js/app.js`, look for the 2x2 action
  grid):** the expand chevron `›` sits on its own at the left of the action group; then a **2x2
  grid**: top row **Archive, Delete**; bottom row **Email ✉ (open email), Edit ✎**. Keep the current
  26px icon-button size and current gap — do NOT make them bigger. With no email link, keep the
  grid shape with an invisible same-size placeholder in the Email slot so every card aligns. On a
  handled card, Restore takes Archive's slot. Inline SVG, `aria-label`s.
- Archive = "handled" = the existing tick (`toggleTick` → ticks.json via the existing
  cc-tasks-writer `inbox-state` sync, which already does CC done-sync for `id_` keys — unchanged).
  This is also the pipeline's existing dismissal signal (Phase 3.9 in `fetch_inbox.py` drops a
  carried-forward card whose `eid_<entry_id>` tick is true), so archived cards are not re-added.
  Toast "Marked handled" + Undo. "Show done" remains the archived view; handled cards there show
  Restore + Delete.
- Delete = in-page confirm ("Remove this card from the dashboard permanently?" Cancel / Remove —
  never `window.confirm()`), then set tick key `del_<storageKey>` = true AND, for email cards only
  (keys starting `eid_` or `mid_`, never `id_` Command-Centre cards), also set the normal tick key
  true so Phase 3.9's existing dismissal signal drops it from the triage ledger — all through the
  same `saveTicks()`/`pushTicks()` path. Cards with a true `del_` key are never rendered, in any
  view, on any day — so the next pipeline run re-emitting the same email does not bring it back.
  Toast "Removed" + Undo (sets the keys back to `false` — don't delete keys, the Worker merge would
  resurrect them). A Command-Centre-mirrored card's Delete hides it here only; it never marks or
  deletes the Command Centre task. Never touches the mailbox.
- Edit = inline rename of the card's display title (like command-centre's rename): Enter saves,
  Escape cancels; stored as tick key `title_<storageKey>` = new title (string) via the same sync;
  the render uses it in place of the pipeline title; an empty value clears it.
- Replace any other native `confirm()` in the file with the same in-page confirm.
- Undo/toast: reuse `wiNotify` styling or add a small toast with an action button, ~10 s.

## 3. Expand / collapse (match the tracker/command-centre)
Every card gets a chevron `<button>` in a fixed spot (`aria-expanded`, `aria-controls`,
Enter/Space, focus kept). Clicking it or the card body toggles an expanded panel under the card row
(no layout change to the collapsed card): full text with paragraph breaks preserved
(`white-space:pre-wrap`, escaped, `<strong>` tags from `sub` may stay via `sanitizeSub`-style
whitelisting but keep newlines): for email cards `ai_summary`, the full `sub`, from + received; for
Command-Centre-mirrored cards `description` then the `actions` log (newest first). Must not start a
drag, open an email, or fire within ~250 ms after a drag. Remember expanded card keys in
`localStorage` `workInbox_expanded_v1` (try/catch). "Expand all / Collapse all" small button in each
section header (stopPropagation so it doesn't collapse the section).

## 4. Checks
`node --check js/app.js`; confirm `index.html` and `js/app.js` still contain their original
non-ASCII characters (e.g. `–`, `—`) unchanged. Add a short section at the top of `HANDOVER.md`
(what changed, checks, exact next action = Drew's browser verification).
