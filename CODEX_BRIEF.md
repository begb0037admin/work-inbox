# Codex brief — work-inbox round 2 on this branch: remove every Outlook Classic path (Kevin bug)

Branch `drew/wi-section-counts` (HEAD `50b436a`, keep that work). Work only in
C:/Users/admin/github/work-inbox. Only touch `js/app.js`, `tools/publish_drafted_replies.py`,
`HANDOVER.md` (and a new small test if useful). Don't read other repos, `data/`, `Archive/`,
`js/vendor/`. Keep UTF-8 intact. Commit locally; no push; short final message.

Hard rule (8 Sep 2026): Outlook Classic is retired; every open-email path opens OWA in the browser;
never `openmail://`, never COM. Kevin: Draft Replies → "Open original" opens Outlook Classic.

## Dashboard (`js/app.js`)
- Delete `openEmail()` (the `openmail://` opener, ~line 334) and every call to it:
  - Draft Replies (~1755-1766): `open_mode==='com'` renders `openEmail(source_entry_id)`. Instead:
    always use `draftWebUrl(e)` (validated OWA host); if it returns a URL → open in a new tab (same
    helper as Priorities); otherwise render the existing honest muted "No linked original" button
    (disabled look, tooltip "The original email link isn't available yet"). Ignore `open_mode`
    for the decision.
  - Priority card link fallback (~431) and icon fallback (~977): when no OWA link resolves, no
    `openEmail` — render the Email slot as the muted/disabled state (keep the grid shape).
- `grep -n "openmail" js/app.js` must return only comments (or nothing) afterwards; reword comments
  so they state the rule, not a live fallback.

## Source (`tools/publish_drafted_replies.py`)
- Never emit `open_mode:"com"`. `open_mode` = "web" when a validated OWA link exists, else "none".
- For drafts with no `web_link`/`display_url`, try to resolve one with the existing connector helper
  `lane_b_call1.resolve_mail_weblink_by_subject(subject, received)` (read-only), capped at
  `WI_DRAFT_WEBLINK_MAX_RESOLVES` (default 3) per run, results cached in
  `data/drafted_replies_weblinks.json` keyed by `draft_id` (only real https outlook.office(365).com
  links; failures cached with a timestamp and retried after 24h) so the connector isn't hit every
  run. Any exception → leave the link empty, never break the publish. Keep `source_entry_id` in the
  output as metadata only.
- `python -m py_compile tools/publish_drafted_replies.py`.
Top-of-HANDOVER entry (what changed; resolution can't be proven until the connector quota resets
27 Sep 10:25).
