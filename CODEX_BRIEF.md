# Codex brief — raise the Lane B inbox read cap 30 → 100 (small)

Branch `drew/wi-mail-read-cap-100`. Only touch `lane_b_call1.py` and `HANDOVER.md`. Do not read
other repos, `data/`, `Archive/` or `js/`. Commit locally; no push; short final message.

Why: `mail_truncation_risk` has been true on every briefing since 15 Sep — the read pass returns
exactly `WI_LANE_B_MAIL_MAX_READ` (default 30) every run, so older in-window read mail is dropped
and the dashboard shows "Mail fetch may be incomplete". The connector pages at 200 (see
`draft_diff_connector.py`), so 100 stays within one page.

1. `lane_b_call1.py`: change the default of `WI_LANE_B_MAIL_MAX_READ` from `"30"` to `"100"`
   (env override unchanged). Update the nearby comments that state the old number (e.g. the
   "(30)" in the 14 Sep rewrite comment) so they don't mislead, noting the 24 Sep change and why.
2. `python -m py_compile lane_b_call1.py`.
3. Short entry at the top of `HANDOVER.md` (what/why, how to revert: set the default back to 30 or
   set `WI_LANE_B_MAIL_MAX_READ=30` in the wrapper env).
