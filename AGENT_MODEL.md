# AGENT_MODEL.md
# Runtime Operating Model

Version : 2.0
Status  : Ratified
Updated : 2026-06-10 (v2.0 — single-agent model)
Author  : Kevin Lelitte, HR Systems, University of Oxford

Governed by: CONSTITUTION.md
Scope      : All work repositories (begb0037admin)

---

## Preamble

This document defines the current implementation of the four-role
model established in CONSTITUTION.md Section 1. It assigns tools
and software to roles, defines execution mechanics, and records
platform context.

v2.0 replaces the four-tool dispatch model (Claude Chat / Kevin /
Cowork / Chrome) with a single-agent model built on Claude Code.
The constitutional roles are unchanged. What has changed is that
one agent now holds the reasoning, execution, and verification
seats, because Claude Code can perform all three natively. The
role boundaries survive as approval gates rather than as briefs
passed between tools.

This document changes more frequently than the constitution. When
the tooling changes, this document is updated. The constitutional
principles do not change with it.

If this document conflicts with CONSTITUTION.md, the constitution
wins. See CONSTITUTION.md Section 6.

---

## Section 1 — Platform Context

Two machines are in scope for work operations.

**Admin machine (Kevin)**
Operator : Kevin Lelitte
OS       : Windows
Runs     : Claude Code (primary agent), Outlook COM scripts via
           Task Scheduler (fetch_inbox.py)

**Personal machine (Hope)**
Operator : Hope (personal domain)
OS       : macOS
Scope    : AIMM and personal projects only — out of scope for
           all work repositories

The two machines do not share a local filesystem.

**GitHub is the sole authoritative source of truth and the only
working surface.** As of 2026-06-10, local clones are retired as
working copies — all reads and writes go through the GitHub API.
The single exception is scripts that must run locally (Outlook COM),
and those pull their own latest version from GitHub before every
run.

Never hardcode machine-specific paths in repository files. Use
GitHub URLs as the stable reference wherever possible.

---

## Section 2 — Role Assignments

The four constitutional roles are currently assigned as follows.

**Seat A — Reasoning Seat → Claude Code**
Thinks, plans, architects, and routes. All sessions begin here.

**Seat B — Human Seat → Kevin (work) / Hope (personal)**
Holds approval authority and oversight. May intervene at any
point; when invoked, supersedes all in-flight decisions. Is no
longer required to run scripts or paste output — Claude Code
executes directly — but retains the right to do either.

**Seat C — Execution Seat → Claude Code**
Implements approved changes directly: GitHub Contents API writes,
git operations, and script execution on the local machine where a
task requires it.

**Seat D — Verification Seat → Claude Code**
Verifies live behaviour by fetching deployed pages and data
(GitHub Pages, raw content, API state) and by browser inspection
where needed. Kevin remains the final visual check on dashboards.

One agent holding Seats A, C, and D does not collapse the
boundaries — it removes the brief-passing overhead between them.
The boundaries survive as the approval gates below.

**Approval gates — Kevin's explicit confirmation is required
before:**

1. Creating new tasks or changing task tiers in command-centre
2. Any destructive or hard-to-reverse operation — deletions,
   history rewrites, repository settings changes, credential
   changes
3. Publishing anything beyond the begb0037admin repositories
4. Any amendment to CONSTITUTION.md
5. Any action in the personal domain (Hope) from a work session,
   or vice versa

Everything else — reads, reversible writes under the backup rules,
verification — Claude Code executes without asking.

---

## Section 3 — Execution Protocol

The v1.x dispatch notation (🔵 RUN SCRIPT / 🟡 COWORK BRIEF /
🔴 CHROME BRIEF) is retired. Claude Code reasons and executes in
one loop.

The constitutional sequencing rules survive unchanged:

1. **One change at a time.** Verify the result of a write before
   making the next. No parallel writes to the same file.
2. **Restore point before change** (CONSTITUTION.md Section 4).
   Every API write is a commit — the prior commit SHA is the
   restore point and the rollback procedures in each repo's
   CLAUDE.md apply. Governed data files additionally require the
   datestamped Archive/ backup before any write.
3. **Stop and report** (CONSTITUTION.md Section 2). If a task
   requires a decision outside the approved scope or hits an
   approval gate, Claude Code stops and asks Kevin. It does not
   improvise past a gate.

---

## Section 4 — Write Standards

All GitHub writes follow these rules:

- **Contents API only** — GET fresh SHA immediately before PUT.
  On 409 conflict, re-fetch and retry once; on second failure,
  stop and report.
- **Archive/ backup first** for governed data files (command-centre:
  data/tasks.json and index.html; work-inbox: index.html). One
  datestamped backup per file per day; skip if today's exists.
- **Byte-level edits for non-ASCII content.** Files containing
  multi-byte characters are edited as bytes (base64 in, targeted
  byte replacement, base64 out). Never decode/re-encode whole
  files through a text layer — that is how the 2026-06 mojibake
  was created.
- **Cache-bust all raw reads.** Every raw.githubusercontent.com
  fetch carries `?t=<timestamp>`. API reads do not need it.
- **No secrets in any committed file.** PATs and API keys live in
  Windows user environment variables (scripts) and the gh CLI
  keyring (Claude Code).
- **Large outputs** are written to files in the repo, not pasted
  into chat.

---

## Section 5 — Session Discipline

1. Claude Code holds persistent memory across sessions and the
   repositories hold the documentation. Together these replace
   handover briefs for same-operator continuation. Handover briefs
   remain mandatory for cross-operator transfer (Section 6).
2. No session closes without documentation updated to reflect
   current state (CONSTITUTION.md Section 5). HANDOVER.md is
   updated at the end of every working session.
3. Decisions live in documentation, not in chat history. Anything
   worth keeping gets committed.

---

## Section 6 — Cross-Domain Model

Two operators share the same GitHub account and tooling.

**Kevin** — work domain. All Oxford HR Systems repositories.
**Hope**  — personal domain. AIMM and personal projects only.

Domain boundaries are strict. Work context is never carried into
personal sessions and vice versa. Mixed-domain work requires a
Cross-Domain Code Brief.

When an operator hands work to the other, the sending operator
generates a handover brief; the receiving operator works from the
brief alone and issues a return brief on completion. If asked to
pick up work without a handover brief, request one before
proceeding.

**Failover chain (work):** Kevin → Hope
**Failover chain (personal):** Hope → Kevin

---

## Section 7 — GitHub Access

All repositories are hosted under the begb0037admin GitHub account.

Claude Code authenticates via the gh CLI (keyring; repo, workflow,
gist, read:org scopes). Scripts (fetch_inbox.py) authenticate with
a PAT held in Windows user environment variables.

Repositories are currently public — required for GitHub Pages
hosting on the current plan. The historical access reason (Seats
A/C could not reach private repos) no longer applies under Claude
Code; visibility is now a data-exposure decision, tracked per repo
in CLAUDE.md.

Authentication secrets are never committed. When credentials are
rotated, both operator preferences must be updated on the same day.

---

## Section 8 — Repository Scope

The following repositories are currently governed by this model.
This table reflects current governance scope and may change
without constitutional amendment.

| Repository           | Status         | Notes                    |
|----------------------|----------------|--------------------------|
| clockify             | Active         | Gold standard / template |
| command-centre       | Active         | Task dashboard (Module 1)|
| work-inbox           | Active         | Inbox briefing pipeline  |
| hris-dashboard       | Active         | Complex — handle last    |
| hris-launcher        | Active         |                          |
| hr-fa-knowledge-base | Active         |                          |
| meeting-records      | Active         |                          |
| hr-projects          | Active         |                          |
| desktop-tutorial     | Decommissioned | Deletion pending         |
| aimm                 | Out of scope   | Personal domain — Hope   |
| personal-finance     | Out of scope   | Personal domain — Hope   |

---

## Version History

| Version | Date       | Change                              |
|---------|------------|-------------------------------------|
| 1.0     | 2026-06-06 | Initial ratification.               |
| 1.1     | 2026-06-09 | Updated work-inbox local path.      |
| 1.2     | 2026-06-09 | Cache-bust rule added to Section 4. |
| 2.0     | 2026-06-10 | Single-agent model: Claude Code     |
|         |            | holds Seats A, C, D; Kevin holds    |
|         |            | Seat B as approval authority.       |
|         |            | Dispatch notation retired; approval |
|         |            | gates defined. GitHub-only working  |
|         |            | surface ratified. Repo table        |
|         |            | updated (command-centre,            |
|         |            | hris-launcher added).               |
