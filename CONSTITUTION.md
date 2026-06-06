# CONSTITUTION.md
# The Operating Constitution

Version : 1.0
Status  : Published — under review against ratification gates (see Section 9)
Ratified: 2026-06-06
Author  : Kevin Lelitte, HR Systems, University of Oxford

---

## Preamble

This document defines the enduring principles under which all work
is conducted across all repositories and projects. It is not a
description of current tools, workflows, or conventions. Those live
in AGENT_MODEL.md and CLAUDE.md respectively.

When this document conflicts with any other document, Section 6
governs.

This constitution changes rarely. When it must change, the amendment
process defined in Section 7 applies. Edits outside that process
are not valid.

---

## Section 1 — The Separation of Concerns

The system operates across four distinct roles. Each role has a
defined authority boundary. No role operates outside its boundary
without explicit instruction from the reasoning seat.

**The four roles are:**

1. **The Reasoning Seat** — thinks, plans, designs, and routes.
   All sessions begin here. No other role acts without a dispatch
   from this seat.

2. **The Human Seat** — executes read-only operations on instruction.
   Returns output verbatim. Does not interpret, modify, or decide.
   Human authority is available for oversight, approval, and
   intervention at any point. When invoked, it supersedes all
   in-flight decisions.

3. **The Execution Seat** — the sole authority for implementing
   approved changes. Acts only on complete, explicit briefs from
   the reasoning seat. Has no authority to make decisions beyond
   the brief.

4. **The Verification Seat** — confirms live behaviour in a running
   environment. Read-only. Reports what it observes. Does not
   interpret or decide.

The assignment of tools and software to these roles is an
implementation detail defined in AGENT_MODEL.md, not here.

---

## Section 2 — Dispatch Quality

The reasoning seat owns solution design. A dispatch must be complete
enough that the receiving role requires no architectural decisions
to carry it out.

An incomplete dispatch is a reasoning seat failure, not an executor
failure.

If a receiving role encounters something that requires a decision,
it stops, reports exactly what it has found, and returns to the
reasoning seat. It does not proceed, improvise, or interpret.

One dispatch at a time. No role acts until the previous dispatch
has returned a result.

---

## Section 3 — Implementation Authority

One role, and only one role, is responsible for implementing
approved changes. This is an absolute authority boundary. It does
not change based on urgency, convenience, or context.

If any other role believes a change needs to be made, the output
is a handoff — not an implementation attempt.

---

## Section 4 — Rollback Before Change

Before any change, a restore point must be recorded. The restore
point must be explicitly stated in the dispatch brief. If the change
introduces a failure, the immediate response is to restore to that
point — not to attempt a fix.

This principle applies regardless of how small the change appears.

---

## Section 5 — Documentation Permanence

Conversation is temporary. Documentation is permanent.

The authoritative record of any project is its documentation, not
session history. Any decision, change, or discovery that is not
committed to documentation does not exist for the purposes of
future sessions.

End-of-session discipline is not optional. Before any session
closes, the current state must be recorded in the project's
documentation. A session that closes without updated documentation
has created debt.

---

## Section 6 — The Source of Truth Hierarchy

When documents conflict, the following hierarchy applies:

1. The operator's current AI preferences — the living source of
   truth for the reasoning seat
2. This constitution — enduring principles
3. AGENT_MODEL.md — current runtime operating model
4. CLAUDE.md — project-specific rules and context
5. STATUS.md / HANDOVER.md — current project state

A conflict between levels is a signal that a lower document has
drifted. The correct response is to update the lower document to
align with the higher — not to compromise the higher document
downward.

---

## Section 7 — Amendment Process

This constitution may only be amended by explicit decision.

An amendment requires:
1. A documented reason stating why the principle is insufficient
   or incorrect
2. A session note recording the decision
3. A version bump — minor for clarifications, major for principle
   changes
4. Propagation to all repositories that carry this file

Edits made without this process are not valid amendments. If a
principle feels wrong in practice, the correct response is to raise
it through the amendment process — not to quietly adjust the file.

Operational details that have leaked into this document are not
amendments — they are corrections and may be removed at any time
without a version bump. The operator determines whether a change
is a correction or an amendment. That determination must be stated
in the session note.

---

## Section 8 — Universal Applicability

This constitution applies to every repository, every project, and
every session regardless of technology stack, workflow, or tooling.

If a principle in this document does not apply to a specific
project, that is a signal that either the principle needs amendment
or the project is operating outside the agreed model. It is not a
reason to ignore the principle.

---

## Section 9 — Review Period

This constitution is published and authoritative from the date of
ratification. It is simultaneously under active review against the
following gate criteria:

1. AGENT_MODEL.md drafted and validated against these principles.
2. At least one CLAUDE.md aligned to the full source of truth
   hierarchy.
3. The governance stack exercised in at least one repository under
   real working conditions.

When all three gates are met, v1.0 is confirmed stable. Until then,
the constitution is authoritative but the operator retains the right
to make structural revisions without the full amendment process,
provided each revision is recorded in the version history.

---

## Version History

| Version | Date       | Change                              |
|---------|------------|-------------------------------------|
| 1.0     | 2026-06-06 | Initial ratification                |
|         |            | Corrections applied from governance |
|         |            | review: Preamble, Sections 1, 3, 4, |
|         |            | 7. Review period declared.          |
