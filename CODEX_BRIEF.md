# Codex brief — work-inbox: CC→ label fit + right column clipped below ~1400px (small)

Branch `drew/wi-cc-label-and-columns` off main. Work only in C:/Users/admin/github/work-inbox.
Only touch `css/styles.css`, `js/app.js` (only if needed), `HANDOVER.md`. Don't read other repos,
`data/`, `Archive/`, `js/vendor/`. Keep UTF-8 intact. Commit locally; no push; short final message.

1. **CC→ overflows its 26px slot** (text pokes out to the left of the button edge). Make it fit
   inside the 26px square on every card, identical across cards: e.g. a small inline SVG icon
   (a box with an outward arrow) or a label like "CC" at ~9px, `white-space:nowrap`, no padding,
   `overflow:hidden`. Keep `title`/`aria-label` "Open in Command Centre", same size/gap as the
   other five buttons.
2. **Right column clipped below ~1400px.** At 1100–1400px the Priorities tab's two-column layout
   (left: Urgent/Tomorrow/This week; right: Today/Needs response/FYI) is wider than the main area,
   which has `overflow:hidden`, so the right column is cut off (measured: right column's right edge
   1252px at a 1100px viewport; cards also narrower than their content). Fix: the two columns must
   always fit the available width — use `minmax(0,1fr)` columns / `min-width:0` on the column
   children, and stack into one column when the main area is narrower than ~900px. No horizontal
   scrolling, nothing clipped, card text wraps. Don't change the look at ≥1440px.
`node --check js/app.js` if touched. Top-of-HANDOVER line.
