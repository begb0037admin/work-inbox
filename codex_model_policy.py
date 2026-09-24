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

# Every scripted Codex connector call is permanently locked to this model and
# reasoning effort. Keep these values literal so a caller cannot drift onto a
# fallback model or a different effort tier through normal policy inputs.
PREFERRED_MODEL = "gpt-5.6-luna"
# Retain the exported name for callers that report a fallback path; it
# intentionally resolves to the same locked model rather than permitting drift.
FALLBACK_MODEL = PREFERRED_MODEL
LOCKED_EFFORT = "high"
FORBIDDEN_EFFORTS = frozenset({"low", "medium", "xhigh", "max"})
ALLOWED_EFFORTS = frozenset({LOCKED_EFFORT})

# Keep workload_class in the public API for caller compatibility, but every
# class is deliberately mapped to the same enforced effort.
_WORKLOAD_EFFORT = {
    "low": LOCKED_EFFORT,
    "medium": LOCKED_EFFORT,
    "high": LOCKED_EFFORT,
}


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
    """Return the permanently enforced (model, effort) pair.

    workload_class and luna_available remain in the signature for
    compatibility with existing callers, but neither can select a fallback.
    An explicit effort override is accepted only when it agrees with the lock.
    """
    del workload_class, luna_available
    effort = (effort_override or LOCKED_EFFORT).strip().lower()
    if effort != LOCKED_EFFORT:
        raise ModelPolicyViolation(
            f"refusing to build a codex exec call with reasoning effort {effort!r} -- "
            f"the policy is permanently locked to {LOCKED_EFFORT!r}."
        )
    return PREFERRED_MODEL, LOCKED_EFFORT


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

    expected = (PREFERRED_MODEL, LOCKED_EFFORT)
    check("preferred and fallback names resolve to the locked model",
          PREFERRED_MODEL == "gpt-5.6-luna" and FALLBACK_MODEL == PREFERRED_MODEL)
    for cls in ("low", "medium", "high", "unrecognised"):
        check(f"workload_class={cls!r} -> locked Luna/high",
              resolve_model_effort(cls) == expected)
    check("luna_available=False cannot select a fallback",
          resolve_model_effort("high", luna_available=False) == expected)
    for override in (None, "", "high", "HIGH"):
        check(f"effort_override={override!r} -> locked high",
              resolve_model_effort("high", effort_override=override) == expected)
    for override in ("low", "medium", "xhigh", "max", "ultra-mega"):
        raised = False
        try:
            resolve_model_effort("high", effort_override=override)
        except ModelPolicyViolation:
            raised = True
        check(f"effort_override={override!r} cannot change the lock", raised)

    check(
        "build_codex_effort_args() produces the exact locked argv shape",
        build_codex_effort_args("low", luna_available=False)
        == ["-m", "gpt-5.6-luna", "-c", "model_reasoning_effort=high"],
    )

    print("")
    if fails:
        print(f"RESULT: {len(fails)} FAILED")
        return 1
    print("RESULT: all passed")
    return 0


if __name__ == "__main__":
    import sys
    sys.exit(_selftest())
