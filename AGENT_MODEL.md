# AGENT_MODEL.md
# Runtime Operating Model

Version : 1.0
Status  : Ratified
Updated : 2026-06-09
Author  : Kevin Lelitte, HR Systems, University of Oxford

Governed by: CONSTITUTION.md
Scope      : All work repositories (begb0037admin)

---

## Preamble

This document defines the current implementation of the four-role
model established in CONSTITUTION.md Section 1. It assigns tools
and software to roles, defines dispatch mechanics, and records
platform context.

This document changes more frequently than the constitution. When
the tooling changes, this document is updated. The constitutional
principles do not change with it.

If this document conflicts with CONSTITUTION.md, the constitution
wins. See CONSTITUTION.md Section 6.

---

## Section 1 — Platform Context

Two machines are in scope for work operations.

**Work machine (Kevin)**
Operator : Kevin Lelitte
OS       : Windows
Username : begb0037 | Domain: AD-OAK
Path root: C:\Users\begb0037.AD-OAK\
work-inbox: C:\Users\admin\Documents\Claude\Projects\work-inbox\

**Personal machine (Hope)**
Operator : Hope (personal domain)
OS       : macOS
Scope    : AIMM and personal projects only — out of scope for
           all work repositories

The paths listed in this section are descriptive runtime context
only and are not authoritative configuration values.

The two machines do not share a local filesystem. Files stored on
one machine are not directly accessible from the other.

GitHub is the authoritative source of truth for all governed
repositories and acts as the shared storage layer between machines.
Repository content is authoritative; local copies are working
copies only.

A Cowork brief that relies on machine-specific paths, configuration,
or local files may not execute correctly on the other machine. Any
brief that touches the local filesystem must make the target machine
explicit.

Never hardcode machine-specific paths in repository files. Use
GitHub URLs as the stable reference wherever possible. Scripts that
require local paths must derive or parameterise them at runtime.

---

## Section 2 — Role Assignments

The four constitutional roles are currently assigned as follows.

**Seat A — Reasoning Seat → Claude Chat**
Thinks, plans, architects, and routes. All sessions begin here.
Produces all dispatch briefs. Makes all architectural decisions.
Does not write files. Does not execute commands. Does not operate
the browser.

**Seat B — Human Seat → Kevin (work) / Hope (personal)**
Executes read-only terminal commands on instruction from Seat A.
Pastes output back verbatim without interpretation or modification.
Human authority is available for oversight, approval, and
intervention at any point. When invoked, it supersedes all
in-flight decisions. See CONSTITUTION.md Section 1.

**Seat C — Execution Seat → Cowork**
The sole seat authorised to implement approved changes. Writes
files to disk, executes bash commands, makes git commits, and
controls the browser when browser automation is required. Acts
only on complete, explicit briefs from Seat A. Has no authority
to make decisions beyond the brief.

**Seat D — Verification Seat → Chrome (browser)**
Confirms live behaviour in a running environment. Read-only.
Reports what it observes. Does not interpret or decide.

Verification is requested only when the required answer cannot be
obtained from reasoning or implementation outputs.

Seat D is never the first seat reached in a workflow.

---

## Section 3 — Dispatch Protocol

Dispatches are issued by Seat A only. Each dispatch targets exactly
one seat. Dispatches are strictly sequential. Parallel dispatches
are not permitted. See CONSTITUTION.md Section 2.

**Dispatch notation:**

🔵 RUN SCRIPT   → Seat B
   Read-only terminal command. Kevin runs exactly as given and
   pastes output back verbatim.

🟡 COWORK BRIEF → Seat C
   Implementation instruction. Commands only — no prose. Must be
   fully self-contained with complete context. Cowork has zero
   assumed knowledge of the current session.

🔴 CHROME BRIEF → Seat D
   Numbered checklist of browser actions with specific expected
   outputs at each step. Used only when Seats A and C cannot
   resolve the verification need.

A brief is complete when the receiving seat requires no
architectural decisions to carry it out. An incomplete brief is
a Seat A failure. See CONSTITUTION.md Section 2.

---

## Section 4 — Cowork Brief Standards

Cowork briefs must be self-contained. The following are required
in every brief:

- **Restore point** — the SHA or file state to return to if the
  change fails. Must be stated explicitly. See CONSTITUTION.md
  Section 4.
- **Target machine** — Windows (work) or macOS (personal).
  Never assumed.
- **Complete file paths** — derived from GitHub URLs, never
  hardcoded local paths.
- **Exit condition** — a clear statement of what done looks like.

Assumptions are prohibited. Any information required to complete
the task must be present in the brief.

Large audit or recon output must be written to a file by Cowork,
not pasted to chat. Seat A reads it surgically on demand via
targeted fetch.

---

## Section 5 — Session Discipline

These rules apply every session.

1. Large output → Cowork writes to file. Never pasted to chat.
   Seat A requests specific sections only.
2. Trigger "prep handover" before any large execution task —
   not after. Large execution task is determined by operator
   judgement and includes any task expected to generate substantial
   output, prolonged execution, or significant context accumulation.
3. Open a fresh session for any task that starts with a large
   data load.
4. After a sustained or complex session, Seat A flags proactively:
   "This session has been running for a while — should I generate
   a handover brief now as a precaution?"
5. No session closes without documentation updated to reflect
   current state. See CONSTITUTION.md Section 5.

---

## Section 6 — Cross-Domain Model

Two operators share the same GitHub account and tooling.

**Kevin** — work domain. All Oxford HR Systems repositories.
**Hope**  — personal domain. AIMM and personal projects only.

Domain boundaries are strict. Work context is never carried into
personal sessions and vice versa. Mixed-domain briefs are not
valid.

Shared tooling does not create shared authority. Domain separation
remains in force regardless of platform, repository ownership, or
account configuration.

When an operator hits a session limit, they generate a handover
brief and pass it to the other operator. The receiving operator
works from the brief alone — zero assumed knowledge from the
sending session. On completion, the receiving operator issues a
return brief.

The sending operator always generates the handover brief. The
receiving operator never generates it on their behalf.

If a receiving operator is asked to pick up work without a
handover brief, the correct response is to request one before
proceeding.

**Failover chain (work):** Kevin → Hope
**Failover chain (personal):** Hope → Kevin

---

## Section 7 — GitHub Access

All repositories are hosted under the begb0037admin GitHub
account. Private repositories are accessed via the GitHub
Contents API.

URL pattern : https://api.github.com/repos/begb0037admin/
              {repo}/contents/{path}?ref=main
Auth header : Authorization: token {PAT}

Authentication secrets are held outside repository files and are
never committed. When credentials are rotated, both operator
preferences must be updated on the same day.

---

## Section 8 — Repository Scope

The following repositories are currently governed by this model.
This table reflects current governance scope and may change
without constitutional amendment.

| Repository           | Status         | Notes                    |
|----------------------|----------------|--------------------------|
| clockify             | Active         | Gold standard / template |
| hris-dashboard       | Active         | Complex — handle last    |
| hr-fa-knowledge-base | Active         |                          |
| work-inbox           | Active         |                          |
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
