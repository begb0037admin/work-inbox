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

FALLBACK MODEL CHANGED, 10 Sep 2026, same day evening (Drew, Kevin's explicit
decision): gpt-5.5 -> gpt-5.6-terra, run at LOW effort specifically on the
fallback path regardless of workload_class (see resolve_model_effort()'s own
docstring below for the exact mechanics). Full trigger-semantics writeup is
in MODEL_POLICY.md's "Fallback trigger semantics" section -- read it before
changing anything about when luna_available gets set to False. Short version:
Terra is an AVAILABILITY fallback only (unavailable/unsupported/rate-limited/
technical-service-failure) -- never a quality-based swap. AUDITED same day,
before writing any of this: the only mechanism anywhere in work-inbox or
hris-dashboard that ever sets luna_available=False is the WI_CODEX_LUNA_
UNAVAILABLE env var, a static, manual, operator-flipped toggle read once at
process start (see lane_b_call1.py's CODEX_LUNA_AVAILABLE). No automatic
Luna-health detection and no content-based/quality-based model-swap logic
exists anywhere in either repo's calling code -- confirmed by reading every
retry loop in lane_b_call1.py and fetch_osm_report_connector.py in full.
Nothing needed to be removed or gated off as a result of this audit.
"""

from __future__ import annotations

# Luna is the estate-preferred model for scripted Codex connector calls
# (MODEL_POLICY.md "Preferred model"). Confirmed present in the failover
# CODEX_HOME's cached model list (models_cache.json, fetched 2026-09-10,
# client_version 0.151.0): slug "gpt-5.6-luna".
PREFERRED_MODEL = "gpt-5.6-luna"

# Fallback when Luna is unavailable (MODEL_POLICY.md Implementation point 5).
# CHANGED 10 Sep 2026, same day evening -- Kevin's explicit decision,
# superseding the gpt-5.5 choice made earlier the same day. gpt-5.6-terra
# confirmed present, supported_in_api: true, via models_cache.json on BOTH
# the primary (Edu) and failover (personal) CODEX_HOME identities on the
# real production host (fetched 10 Sep 2026, client_version 0.151.0),
# alongside gpt-5.6-luna and gpt-5.5 themselves.
#
# This is an AVAILABILITY fallback ONLY -- see MODEL_POLICY.md's "Fallback
# trigger semantics" section for the full, binding spec. Terra must never be
# selected because Luna's answer seemed weak, thin, or in need of more
# reasoning, and never because a quality/content check on Luna's response
# failed -- only because Luna itself is unavailable, unsupported, rate-
# limited, or fails for a technical/service reason. A quality concern is
# never grounds for this fallback; it must be surfaced/logged for human
# review instead, through a completely separate mechanism.
FALLBACK_MODEL = "gpt-5.6-terra"

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
    max, override or not, regardless of luna_available. It is ALSO rejected
    if it would be a DOWNGRADE below the workload_class's own mapped tier --
    this check is UNCONDITIONAL on luna_available (touchpoint-1 Codex review
    finding, 10 Sep 2026, on the ORIGINAL Luna-only version of this guard: an
    unvalidated override could otherwise defeat fail-closed-to-High for an
    unrecognised class, e.g. resolve_model_effort("typo-class",
    effort_override="low"), or silently understate a genuinely High-
    classified call; a SECOND touchpoint-1 finding, same day, on the Terra-
    fallback version of this docstring: an earlier draft gated this check on
    `luna_available`, which meant ANY effort_override passed alongside
    luna_available=False bypassed downgrade-rejection entirely -- e.g.
    resolve_model_effort("high", luna_available=False,
    effort_override="medium") silently returned ("gpt-5.6-terra", "medium")
    instead of raising, since the guard never even looked at it. Caught live
    by Codex before commit -- see codex_model_policy.py's git history /
    HANDOVER.md for this session's touchpoint-1 pass). Fixed: the guard is
    identical in EVERY case now, luna_available or not. An override that is
    EQUAL to or HIGHER than the mapped tier is allowed (never less scrutiny
    than the class implies, more is fine for a deliberate test).

    luna_available=False -- the Terra availability fallback (MODEL_POLICY.md
    "Fallback trigger semantics", added 10 Sep 2026, Kevin's explicit
    decision): forces LOW effort, REGARDLESS of workload_class, but ONLY
    when the caller did not also pass an explicit effort_override (falsy --
    None or ""). This is the ONE specific, intentional bypass of the
    workload_class -> effort tier mapping -- it exists because Terra is only
    ever invoked for an availability reason (Luna unavailable/unsupported/
    rate-limited/a technical-service failure), never because of anything
    about the workload itself, so the workload's own effort tier is not the
    right thing to run Terra at. It does NOT bypass the downgrade-rejection
    guard above -- that guard is checked FIRST, unconditionally, against
    whatever effort_override the caller actually passed; only once that has
    already passed (or no override was given at all) does this forced-low
    default apply. The hard ceiling (FORBIDDEN_EFFORTS) is also still
    enforced unconditionally -- an explicit effort_override of "xhigh"/"max"
    passed alongside luna_available=False is still rejected, not silently
    forced to low. No current production caller passes effort_override, so
    in practice every real fallback call resolves to low."""
    mapped_tier = (_WORKLOAD_EFFORT.get(workload_class) or "high")
    effort = (effort_override or mapped_tier).strip().lower()
    # Downgrade-rejection guard: unconditional on luna_available -- a caller-
    # supplied effort_override is validated identically whether the call is
    # running Luna or falling back to Terra. See this function's own
    # docstring above ("effort_override" paragraph) for the touchpoint-1
    # finding this fixed: an earlier draft gated this on luna_available,
    # which let a downgrade slip through unrejected on the fallback path.
    # The SEPARATE, deliberate forced-low default for a fallback call that
    # supplied NO override lives further down, after the ceiling check, and
    # is not this guard.
    if effort_override and effort not in FORBIDDEN_EFFORTS and effort in _TIER_RANK:
        if _TIER_RANK[effort] < _TIER_RANK[mapped_tier]:
            raise ModelPolicyViolation(
                f"refusing effort_override={effort!r} -- it is a DOWNGRADE below "
                f"workload_class {workload_class!r}'s own mapped tier {mapped_tier!r}. "
                f"effort_override may only raise scrutiny, never lower it; this exists "
                f"to stop an override from silently defeating the fail-closed-to-High "
                f"rule for an unrecognised class, or understating a genuinely High call. "
                f"This applies whether or not luna_available is False -- the Terra "
                f"fallback's own forced-low default only ever applies when no override "
                f"was supplied at all, it never re-validates or waives one that was."
            )
    if effort in FORBIDDEN_EFFORTS:
        raise ModelPolicyViolation(
            f"refusing to build a codex exec call with reasoning effort {effort!r} -- "
            f"MODEL_POLICY.md's hard ceiling is 'high'; xhigh/max are never permitted for "
            f"a scripted call, whether that call is running Luna or the Terra fallback. "
            f"If a workload genuinely needs more than high, that is a signal to stop and "
            f"get Kevin's explicit sign-off for that specific call, the same pattern as "
            f"Constitution Section 10's human-seat effort gate -- not to raise this "
            f"ceiling in code."
        )
    if effort not in ALLOWED_EFFORTS:
        raise ModelPolicyViolation(
            f"refusing to build a codex exec call with unrecognised reasoning effort "
            f"{effort!r} -- must be one of {sorted(ALLOWED_EFFORTS)}."
        )
    if not luna_available and not effort_override:
        # The Terra availability fallback itself: force low, regardless of
        # workload_class. `not effort_override` (falsy -- None or ""), not
        # `is None` -- matches the exact same truthiness this function
        # already uses two lines up to decide whether effort_override was
        # "really" supplied (`effort_override or mapped_tier`); an earlier
        # draft used `is None` here, which meant effort_override="" (falsy,
        # already treated as "no override" everywhere else in this function)
        # fell through this check and escaped the forced-low default,
        # silently returning FALLBACK_MODEL at the mapped_tier's effort
        # instead of low. Only reached once the hard-ceiling/allowed-effort
        # checks above have already passed, so this can never be a route
        # around the ceiling -- it only ever lowers an already-valid effort
        # to low. If a caller explicitly passed a truthy effort_override,
        # that value was already validated above (downgrade-rejection AND
        # ceiling both still enforced) and is honoured as-is -- this bypass
        # never overrides an explicit request, it only supplies the default
        # when the caller made none. No current production caller passes
        # effort_override, so in practice every fallback call resolves to
        # low.
        effort = "low"
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

    # luna_available=False (Terra availability fallback, 10 Sep 2026, Kevin's
    # explicit decision) -> FALLBACK_MODEL at LOW effort, REGARDLESS of
    # workload_class -- not "effort still honoured" (that was gpt-5.5's
    # behaviour, superseded). Every workload_class, including "high" (every
    # real current call site), must resolve to low on this path.
    for cls in ("low", "medium", "high"):
        model, effort = resolve_model_effort(cls, luna_available=False)
        check(f"luna_available=False, workload_class={cls!r} -> (FALLBACK_MODEL, 'low')",
              model == FALLBACK_MODEL and effort == "low")

    # An unrecognised class still fails closed to High for tier-mapping
    # purposes, but the Terra fallback still forces low regardless -- the
    # fail-closed-to-High rule and the fallback-forces-low rule are
    # independent; low is correct here, not a violation of fail-closed.
    model, effort = resolve_model_effort("nonsense-class-nobody-declared", luna_available=False)
    check("luna_available=False on an unrecognised workload_class still -> (FALLBACK_MODEL, 'low')",
          model == FALLBACK_MODEL and effort == "low")

    # (b) The downgrade-rejection guard is UNAFFECTED for every
    # luna_available=True (default) call -- proven by the existing
    # downgrade-rejection checks further down in this function, which pass
    # luna_available's default (True) and are otherwise untouched by this
    # change. See those checks below; not duplicated here.
    #
    # It is ALSO now unconditional on luna_available (fixed live during
    # touchpoint-1 Codex review, 10 Sep 2026 -- an earlier draft gated this
    # guard on luna_available, which let ANY caller-supplied effort_override
    # bypass downgrade-rejection entirely on the fallback path, e.g.
    # resolve_model_effort("high", luna_available=False,
    # effort_override="medium") silently returning ("gpt-5.6-terra",
    # "medium") instead of raising). These two checks are the regression
    # test for that fix -- an explicit downgrade is rejected identically
    # whether luna_available is True or False; only the NO-override default
    # differs (see the "regardless of workload_class" checks above).
    raised = False
    try:
        resolve_model_effort("high", luna_available=False, effort_override="medium")
    except ModelPolicyViolation:
        raised = True
    check("luna_available=False, effort_override='medium' on workload_class='high' still "
          "raises ModelPolicyViolation (downgrade-rejection not bypassed by the fallback "
          "path -- regression test for a bug caught live in touchpoint-1 review)", raised)

    model, effort = resolve_model_effort("low", luna_available=False, effort_override="high")
    check("luna_available=False, effort_override='high' on workload_class='low' (an upgrade, "
          "not a downgrade) is honoured as 'high', not silently forced to 'low' -- the forced-"
          "low default only applies when the caller supplied no override at all",
          model == FALLBACK_MODEL and effort == "high")

    # Regression test for a second bug caught live in the same touchpoint-1
    # pass: effort_override="" (empty string, falsy) must be treated exactly
    # like effort_override=None everywhere in this function, including the
    # forced-low default below -- an earlier draft checked
    # `effort_override is None` there, which meant "" fell through and
    # escaped the forced-low default even though `effort or mapped_tier`
    # two lines up already treats "" as "no override".
    model, effort = resolve_model_effort("high", luna_available=False, effort_override="")
    check("luna_available=False, effort_override='' (falsy, not None) still forces low, "
          "same as effort_override omitted entirely -- regression test for a bug caught "
          "live in touchpoint-1 review",
          model == FALLBACK_MODEL and effort == "low")

    # (c) xhigh/max are still hard-rejected through the fallback path too --
    # the ceiling is never bypassed by luna_available=False.
    for banned in ("xhigh", "max", "XHIGH", " Max "):
        raised = False
        try:
            resolve_model_effort("high", luna_available=False, effort_override=banned)
        except ModelPolicyViolation:
            raised = True
        check(f"luna_available=False, effort_override={banned!r} still raises "
              f"ModelPolicyViolation (ceiling not bypassed by the fallback path)", raised)

    # (d) There is no response-content / quality-signal parameter anywhere
    # on this function's surface, and no automatic trigger-detection code
    # was added by this change (audited 10 Sep 2026: the only thing that
    # ever sets luna_available=False anywhere in work-inbox or
    # hris-dashboard is the WI_CODEX_LUNA_UNAVAILABLE env var, a static
    # manual operator toggle read once at process start -- see
    # lane_b_call1.py's CODEX_LUNA_AVAILABLE). This test documents that: a
    # call shaped exactly like a caller reacting to a "weak"/short/thin
    # Luna response (i.e. an ordinary call with luna_available left at its
    # default) can NEVER select FALLBACK_MODEL -- the only way in is the
    # explicit, boolean luna_available=False argument, never anything
    # inferred from a response.
    model, effort = resolve_model_effort("high")
    check("a default call (simulating 'Luna answered, even a short/thin answer') "
          "-> PREFERRED_MODEL, never FALLBACK_MODEL -- content is never a trigger",
          model == PREFERRED_MODEL)

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
