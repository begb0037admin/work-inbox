# Codex brief — URGENT work-inbox card layout + Expand all hotfix (Kevin, 24 Sep)

Branch `drew/wi-card-layout-hotfix` off main. Work only in C:/Users/admin/github/work-inbox. Only touch `js/app.js`, `css/styles.css` (additive
rules / edits to the rules added in PR #37), `HANDOVER.md`. Do not read other repos except the
read-only reference `C:/Users/admin/github/command-centre/css/styles.css`. Don't read `data/`,
`Archive/`, `js/vendor/`. Keep UTF-8 intact. Commit locally; no push; short final message.

## Bugs (reproduced by Drew in Chromium at 1440px)
1. `.card-ph` is `display:flex` with `flex-wrap:nowrap`, so the expanded `.pri-detail`
   (flex-basis 100%) sits BESIDE the header: `.card-ph-body` collapses to 0 px wide and the title
   and meta wrap one word per line; the action buttons overlap the title.
2. "Expand all" in each section header never changes its label to "Collapse all" (it still
   toggles, but the label is static), so Kevin reads it as broken.

## Required layout (Kevin's annotated screenshot — exact)
- Card header row = full card width: drag grip, then title + meta taking ALL remaining width
  (`flex:1; min-width:0`, normal word wrapping), then a fixed action area on the right.
- Action area = ONE grid, **3 columns x 2 rows**, 26px cells, same gap as now (Kevin's final
  annotated mock-up — supersedes any 2x2/3-row version):
  - Row 1: **› (expand chevron) | Archive | Delete** (Restore takes Archive's slot on handled cards)
  - Row 2: **CC→ | Email ✉ | Edit ✎**
  i.e. DOM order: chevron, archive, delete, cc, email, edit; `grid-template-columns:repeat(3,26px)`.
  Missing CC→ or Email → an invisible same-size placeholder so every card's grid has the same
  shape. The CC→ button moves out of the title/meta flow into row 2 left; make it a 26px square
  `.card-icon` showing "CC→" in small text (keep its existing click behaviour, `title` +
  `aria-label` "Open in Command Centre"). The › keeps its aria attributes and toggle. Nothing
  from the action area may sit inside or on top of the title/meta.
- Expanded `.pri-detail` goes BELOW the header row, spanning the card's full width
  (`flex-basis:100%` with `flex-wrap:wrap` on the card, or make the header its own row element),
  with long text wrapping (`overflow-wrap:anywhere`), no horizontal overflow.
- Below 600 px: the grid moves under the title row, right-aligned (same as command-centre).

## Expand all
Label per section = "Collapse all" when every visible card in that section is expanded, else
"Expand all"; recomputed on every render (so it stays right after a drag, a reload, or a single
card toggle). Clicking expands/collapses every visible card in that section, persisted in
`workInbox_expanded_v1`, and must not collapse the section itself.

## Checks
`node --check js/app.js`; UTF-8 non-ASCII characters unchanged. Top-of-HANDOVER entry.
