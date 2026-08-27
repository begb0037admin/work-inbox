# Phase 2 Brief — AI-Triage Migration to Codex

**Status:** Briefed 2026-08-25, by Kevin's explicit go-ahead, after Phase 1
verification and the cc93c7b incident were both closed out. Supersedes the
Phase 2 sketch in `docs/CODEX_CONNECTOR_MIGRATION_RESEARCH.md` Section 5 —
that sketch is a design reference; this file is the actual brief to execute.

**Precondition update, 2026-08-25 (later same day) — risk explicitly
accepted, not resolved:** Drew correctly verified live that neither part of
the original precondition below was met — the PAT was confirmed still
active (`GET /user` → 200) and the connector-level GitHub App access was
confirmed never reviewed or removed at the OpenAI/ChatGPT account level
(only the local `config.toml` auto-approval override had been stripped).
Drew correctly held and did not run Phase 2 on finding this. Kevin was told
the tradeoff plainly and asked to choose: close the gap first, or accept
the risk explicitly. **Kevin explicitly accepted the risk**, stated directly
in the cloud coordinating session (not through a file, given the day's
trust issues) — "I accept the rest. Please continue." This means: the PAT
remains active and the connector remains unreviewed at Kevin's informed,
explicit choice; Phase 2 may now proceed on that basis.

**This does not relax the structural constraint below.** Codex still gets no
GitHub-write tool for this phase, regardless of the PAT/connector decision —
that fix stands on its own merits (least-privilege for this specific task)
independent of whether the old exposed path is closed elsewhere.

**Original precondition (for the record, now superseded by the above):** the
GitHub PAT found exposed in `~/.codex/config.toml` during the cc93c7b
incident is rotated on GitHub's side, and the `apps.connector_76869538...`
connector has been reviewed at the Codex account level (not just had its
auto-approval stripped locally).

---

## What this phase does

Re-implements the six AI-triage phases currently run by `fetch_inbox.py`
against the Anthropic API, using Codex's own model instead, reading from
the same Oxford connector Phase 1 already verified:

1. **Context paragraph** (current Phase 2 equivalent) — short summary of
   today's inbox/calendar state.
2. **Per-email summaries** for urgent/needs items (current Phase 3.2), with
   a `needs_reply` boolean and one-line reason.
3. **No-action demotion** (current Phase 3.3/3.3b) — flag items that need no
   action despite keyword-matching as urgent/needs, with a reason. Be
   conservative — a false demotion hides real work from Kevin.
4. **Task-suggestion triage** (current Phase 3.5) — propose new Command
   Centre tasks or updates to existing ones, each carrying the source
   message's `id` (Graph format, per Phase 1), a proposed title, tier, and
   the evidence snippet it's based on.
5. **Calendar day-prep summaries** (current Phase 3.8) — same shape as
   today's four day-view columns.

## Hard constraints

**Credential/write-path constraint — new, from the cc93c7b incident, and
non-negotiable, unaffected by the risk-acceptance above:**
- Codex must run this phase with **no GitHub-write tool available to it at
  all** — not `-s read-only` sandboxing alone, which only constrains local
  shell/filesystem and did not stop the cc93c7b write. Confirm before
  running that no `mcp_servers.github`-style entry with write scope, and no
  `apps.*` connector with an auto-approved `github.create_file` /
  `github.create_pull_request` tool, is active in Codex's config for this
  session. If Codex needs to read anything from GitHub for this task, that's
  fine — the constraint is on writes.
- Codex writes its output to **local files only**. It does not push
  anything to GitHub itself, under any circumstance, for this phase. The
  orchestrating session (Drew) reads those local files, reviews them, and
  pushes via the GitHub API — the same pattern already used successfully for
  the Phase 1 output and the incident remediation. This isn't extra
  ceremony; it's the actual fix for what went wrong last time, not a policy
  restated in hope it holds.

**Data/scope constraints (from the original Section 5 design, unchanged):**
- Output only to new files: `codex_briefing.json`, `codex_suggestions.json`,
  `codex_triage_ledger.json`. Never overwrite `data/briefing.json`,
  `data/tasks.json`, or the existing `data/triage_ledger.json` directly.
- `codex_triage_ledger.json` is a separate dedup namespace, keyed on the
  Graph `id` — it must never read from or write to the existing
  `triage_ledger.json`'s `applied`/`promoted`/`tracked_needs_urgent` keys,
  which are keyed on Outlook-COM EntryIDs from the old pipeline.
- No sends, drafts, calendar writes, or Teams message posts under any
  circumstance — this is a read-and-suggest pipeline only.
- New task suggestions use the opener design confirmed in Section 8 of the
  research doc: store `web_link` (falling back to `display_url` if absent)
  and a `source: "codex-graph"` field. Do not attempt to construct an
  `openmail://` URI or an `entryId` for these — the ID format doesn't
  support it (Section 3/7 of the research doc).
- Run in parallel with the existing Anthropic-based pipeline for a
  validation period — do not cut over immediately. Kevin decides when
  enough parallel runs have been compared to trust a cutover; this brief
  does not authorize one.

## What to report back

- Confirmation that the credential precondition above was actually checked
  (not assumed) before running, and what was found.
- The three output files' content (or a representative sample if large),
  same as Phase 1's honest partial-completion reporting — do not claim
  completeness that isn't there.
- Any case where Codex's local model produced a materially different
  triage decision than what the equivalent Anthropic-based phase would have
  — this is the actual signal for whether parallel validation is going
  well, not just "did it run."
- Anything that required a write action to be exposed to Codex even though
  none was used — that's still worth knowing, the same way Phase 1's report
  was.
