#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
codex_model_policy.py -- shared model/effort selection for scripted `codex exec`
calls across the estate's Codex-connector consumers.

Source of truth: begb0037admin/constitution/MODEL_POLICY.md, specifically its
"Implementation and enforcement" 5-point spec. This module is the machine-
readable implementation of that policy's workload-class table -- it does not
replace that document. If the two ever disagree, MODEL_POLICY.md wins; update
this file to match, never the other way around.

Built by Drew, 10 Sep 2026 (coordinator handover Priority 4). Consumers:
  - work-inbox/lane_b_call1.py's run_codex_json() (this repo, local import)
  - hris-dashboard/fetch_osm_report_connector.py (cross-repo import, same
    pattern it already uses for run_codex_json/extract_tool_calls/etc.)

CONFIRMED LIVE, 10 Sep 2026, before this change existed (see work-inbox
HANDOVER.md for the full trail): neither Lane B's primary CODEX_HOME
(C:\\Users\\begb0037.AD-OAK\\.codex on the real production host, 101l-de013193 --
config.toml sha1 ba0184e864ffd081069820cc7a6f8f19acf5c845, matching the
baseline lane_b_call1.py already checks) nor its failover CODEX_HOME
(C:\\WorkInboxAI\\codex-laneb, which has no config.toml at all) had any
model_reasoning_effort or model override. Production was running on
`codex exec`'s own undocumented default -- NOT the desktop's personal
interactive ~/.codex/config.toml's `xhigh` setting, which is a different file
Lane B never reads. This module is what makes the estate's actual policy
choice (Luna, workload-appropriate effort, hard ceiling at "high") the thing
that is actually running, instead of leaving it to an unexamined default.
"""

from __future__ import annotations

# Luna is the estate-preferred model for scripted Codex connector calls
# (MODEL_POLICY.md "Preferred model"). Confirmed present in the failover
# CODEX_HOME's cached model list (models_cache.json, fetched 2026-09-10,
# client_version 0.151.0): slug "gpt-5.6-luna".
PREFERRED_MODEL = "gpt-5.6-luna"

# Fallback when Luna is unavailable (MODEL_POLICY.md Implementation point 5 --
# previously undecided; decided here, 10 Sep 2026, Drew). gpt-5.5 is a
# confirmed-available, distinctly prior-generation model on the SAME
# account's model list, not another untested gpt-5.6-family sibling
# (gpt-5.6-sol / gpt-5.6-terra) picked with no evidence either suits this
# workload. If Luna's unavailability ever turns out to be an account/
# entitlement problem rather than a transient outage, revisit this choice
# with Kevin rather than assuming it's still correct.
FALLBACK_MODEL = "gpt-5.5"

# Hard ceiling (MODEL_POLICY.md "Hard ceiling: high"). No scripted call may
# ever select these, regardless of workload class, override, or caller bug.
FORBIDDEN_EFFORTS = frozenset({"xhigh", "max"})
ALLOWED_EFFORTS = frozenset({"low", "medium", "high"})

# Workload-class -> effort tier, per MODEL_POLICY.md's table. The model
# choice (Luna vs fallback) is independent of workload class -- only effort
# varies by class.
_WORKLOAD_EFFORT = {
    "low": "low",
    "medium": "medium",
    "high": "high",
}
_TIER_RANK = {"low": 0, "medium": 1, "high": 2}


class ModelPolicyViolation(RuntimeError):
    """Raised the instant a caller (directly, via a bad workload_class, or via
    an override) would cause an out-of-policy model/effort selection -- in
    particular anything in FORBIDDEN_EFFORTS. This is the runtime enforcement
    MODEL_POLICY.md's own independent review flagged as the concrete gap:
    'should source' is not a control until it is implemented as one. Callers
    must let this propagate -- it must abort the call, not be swallowed."""


def resolve_model_effort(
    workload_class: str,
    *,
    luna_available: bool = True,
    effort_override: str | None = None,
) -> tuple[str, str]:
    """Return (model, effort) for a declared workload_class, per
    MODEL_POLICY.md's table and hard ceiling.

    An unrecognised workload_class fails CLOSED to "high" (MODEL_POLICY.md's
    own "a workload that does not clearly fit any row ... fails closed to
    High -- never defaults to Low or Medium on uncertainty") rather than
    raising -- a bad class string is a caller bug, not grounds to crash a
    scheduled fetch, but it must never silently run cheaper/lower-scrutiny
    than intended.

    effort_override: TEST-ONLY escape hatch (no current production caller
    passes this). Still validated against FORBIDDEN_EFFORTS / ALLOWED_EFFORTS
    below -- there is no path in this function that can ever select xhigh or
    max, override or not. It is ALSO rejected if it would be a DOWNGRADE
    below the workload_class's own mapped tier (touchpoint-1 Codex review
    finding, 10 Sep 2026: an unvalidated override could otherwise defeat
    fail-closed-to-High for an unrecognised class, e.g.
    resolve_model_effort("typo-class", effort_override="low"), or silently
    understate a genuinely High-classified call). An override that is EQUAL
    to or HIGHER than the mapped tier is allowed (never less scrutiny than
    the class implies, more is fine for a deliberate test)."""
    mapped_tier = (_WORKLOAD_EFFORT.get(workload_class) or "high")
    effort = (effort_override or mapped_tier).strip().lower()
    if effort_override and effort not in FORBIDDEN_EFFORTS and effort in _TIER_RANK:
        if _TIER_RANK[effort] < _TIER_RANK[mapped_tier]:
            raise ModelPolicyViolation(
                f"refusing effort_override={effort!r} -- it is a DOWNGRADE below "
                f"workload_class {workload_class!r}'s own mapped tier {mapped_tier!r}. "
                f"effort_override may only raise scrutiny, never lower it; this exists "
                f"to stop an override from silently defeating the fail-closed-to-High "
                f"rule for an unrecognised class, or understating a genuinely High call."
            )
    if effort in FORBIDDEN_EFFORTS:
        raise ModelPolicyViolation(
            f"refusing to build a codex exec call with reasoning effort {effort!r} -- "
            f"MODEL_POLICY.md's hard ceiling is 'high'; xhigh/max are never permitted for "
            f"a scripted call. If a workload genuinely needs more than high, that is a "
            f"signal to stop and get Kevin's explicit sign-off for that specific call, "
            f"the same pattern as Constitution Section 10's human-seat effort gate -- not "
            f"to raise this ceiling in code."
        )
    if effort not in ALLOWED_EFFORTS:
        raise ModelPolicyViolation(
            f"refusing to build a codex exec call with unrecognised reasoning effort "
            f"{effort!r} -- must be one of {sorted(ALLOWED_EFFORTS)}."
        )
    model = PREFERRED_MODEL if luna_available else FALLBACK_MODEL
    return model, effort


def build_codex_effort_args(
    workload_class: str,
    *,
    luna_available: bool = True,
    effort_override: str | None = None,
) -> list[str]:
    """The actual `codex exec` argv fragment: -m <model> -c
    model_reasoning_effort=<effort>. `codex exec` has no dedicated
    reasoning-effort flag (confirmed live against installed codex-cli
    0.151.0/0.152.0 --help) -- effort is set via the generic config-override
    mechanism, a dotted-path TOML override of <CODEX_HOME>/config.toml's own
    model_reasoning_effort key."""
    model, effort = resolve_model_effort(
        workload_class, luna_available=luna_available, effort_override=effort_override
    )
    return ["-m", model, "-c", f"model_reasoning_effort={effort}"]


# --------------------------------------------------------------------------- #
#  Self-test -- proves the rejection in ModelPolicyViolation actually fires.
#  MODEL_POLICY.md Implementation point 4: "A test or check (even a simple
#  one) proving a caller cannot silently pass an unapproved model/effort
#  through without the rejection ... firing." Plain print-based pass/fail,
#  same style as lane_b_call1.py's own cmd_selftest() and
#  fetch_osm_report_connector.py's cmd_selftest_guard() -- no new test
#  dependency, matches existing repo convention.
# --------------------------------------------------------------------------- #
def _selftest() -> int:
    fails: list[str] = []

    def check(name: str, cond: bool) -> None:
        print(("  ok   " if cond else "  FAIL ") + name)
        if not cond:
            fails.append(name)

    # Each workload class resolves to Luna + the matching effort tier.
    for cls in ("low", "medium", "high"):
        model, effort = resolve_model_effort(cls)
        check(f"resolve_model_effort({cls!r}) -> Luna + {cls} effort",
              model == PREFERRED_MODEL and effort == cls)

    # An unrecognised class fails CLOSED to high, never to low/medium.
    model, effort = resolve_model_effort("nonsense-class-nobody-declared")
    check("unrecognised workload_class fails closed to 'high' (not low/medium)",
          effort == "high")

    # luna_available=False selects the fallback model, effort unaffected.
    model, effort = resolve_model_effort("high", luna_available=False)
    check("luna_available=False -> FALLBACK_MODEL, effort still honoured",
          model == FALLBACK_MODEL and effort == "high")

    # The hard ceiling: xhigh and max must be rejected, not silently clamped.
    for banned in ("xhigh", "max", "XHIGH", " Max "):
        raised = False
        try:
            resolve_model_effort("high", effort_override=banned)
        except ModelPolicyViolation:
            raised = True
        check(f"effort_override={banned!r} raises ModelPolicyViolation (case/whitespace-insensitive)",
              raised)

    # A garbage effort string (not xhigh/max, just invalid) is also rejected,
    # not silently passed through to `codex exec` as a literal string.
    raised = False
    try:
        resolve_model_effort("high", effort_override="ultra-mega")
    except ModelPolicyViolation:
        raised = True
    check("effort_override='ultra-mega' (not a real tier) also raises, not passed through", raised)

    # Downgrade rejection (touchpoint-1 Codex review finding, 10 Sep 2026):
    # an override may never select LESS scrutiny than the workload_class's
    # own mapped tier -- this is what stops effort_override from defeating
    # either fail-closed-to-High or a genuinely High classification.
    raised = False
    try:
        resolve_model_effort("nonsense-class-nobody-declared", effort_override="low")
    except ModelPolicyViolation:
        raised = True
    check("effort_override='low' on an unrecognised (fail-closed-to-High) class is "
          "rejected as a downgrade, not allowed to slip below High", raised)

    raised = False
    try:
        resolve_model_effort("high", effort_override="low")
    except ModelPolicyViolation:
        raised = True
    check("effort_override='low' on an explicitly High-classified call is rejected as a downgrade", raised)

    # An override that only RAISES scrutiny (equal or higher tier) is fine.
    model, effort = resolve_model_effort("low", effort_override="high")
    check("effort_override='high' on a Low-classified call (an upgrade) is allowed",
          effort == "high")

    # The exact argv shape `codex exec` expects.
    args = build_codex_effort_args("high")
    check("build_codex_effort_args('high') produces the exact -m/-c argv shape",
          args == ["-m", PREFERRED_MODEL, "-c", "model_reasoning_effort=high"])

    # Simulates lane_b_call1.py's CODEX_MODEL debug override: swapping argv
    # index 1 (the model) must leave a single, still-validated -c effort arg
    # at index 3 untouched -- proves the model-only override can never carry
    # an unvalidated effort through alongside it.
    args = build_codex_effort_args("high")
    args[1] = "some-debug-model-slug"
    check("simulated WI_CODEX_MODEL override leaves exactly one validated -c effort arg",
          args[2] == "-c" and args[3] == "model_reasoning_effort=high" and len(args) == 4)

    print("")
    if fails:
        print(f"RESULT: {len(fails)} FAILED")
        return 1
    print("RESULT: all passed")
    return 0


if __name__ == "__main__":
    import sys
    sys.exit(_selftest())
