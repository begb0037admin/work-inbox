import unittest
from unittest import mock

import lane_b_call1 as lane


IDENTITIES = [
    {"label": "edu", "CODEX_HOME": r"C:\edu"},
    {"label": "personal-uk", "CODEX_HOME": r"C:\uk"},
    {"label": "personal-com", "CODEX_HOME": r"C:\com"},
]


def failed_result(domain, label, status="codex_failed"):
    return {
        "domain": domain,
        "status": status,
        "served_by": label,
        "guard": {"seen": [], "unexpected": []},
        "tool_calls": [],
        "count": 0,
        "raw_items": [],
    }


class FailoverRingTests(unittest.TestCase):
    def test_explicit_failure_classifier_catches_exit_zero_limit_and_auth_errors(self):
        self.assertEqual(
            lane._explicit_connector_failure_reason("You've hit your usage limit", "", 0),
            "usage_limit",
        )
        self.assertEqual(
            lane._explicit_connector_failure_reason("TRIGGER_REAUTHENTICATION oauth_token_invalid_grant", "", 0),
            "authentication",
        )
        self.assertEqual(lane._explicit_connector_failure_reason("403 Forbidden", "", 1), "permission")

    def test_ring_order_and_success_identity(self):
        calls = []

        def fake_fetch(domain, prompt, **kwargs):
            label = kwargs["identity_label"]
            calls.append(label)
            if label != "personal-com":
                return failed_result(domain, label), [{"identity": label, "outcome": "codex_failed"}]
            return {**failed_result(domain, label), "status": "ok", "count": 1}, [
                {"identity": label, "outcome": "ok"}
            ]

        with mock.patch.object(lane, "available_identity_ring", return_value=IDENTITIES), \
             mock.patch.object(lane, "_fetch_domain_one_identity", side_effect=fake_fetch):
            result = lane.fetch_domain("calendar", "prompt", window_days=7, ts="ts", retries=3)

        self.assertEqual(calls, ["edu", "personal-uk", "personal-com"])
        self.assertEqual(result["served_by"], "personal-com")
        self.assertEqual(result["identity_ring"], ["edu", "personal-uk", "personal-com"])

    def test_falls_through_each_connector_failure_type(self):
        for reason in ("authentication", "usage_limit", "permission", "timeout"):
            calls = []

            def fake_fetch(domain, prompt, **kwargs):
                label = kwargs["identity_label"]
                calls.append(label)
                return None, [{"identity": label, "outcome": "codex_failed", "reason": reason}]

            with self.subTest(reason=reason), \
                 mock.patch.object(lane, "available_identity_ring", return_value=IDENTITIES), \
                 mock.patch.object(lane, "_fetch_domain_one_identity", side_effect=fake_fetch):
                result = lane.fetch_domain("teams", "prompt", window_days=7, ts="ts", retries=3)

            self.assertEqual(calls, ["edu", "personal-uk", "personal-com"])
            self.assertEqual(result["status"], "codex_failed")
            self.assertIsNone(result["served_by"])

    def test_missing_connector_falls_through_and_all_missing_fails_domain(self):
        calls = []

        def fake_fetch(domain, prompt, **kwargs):
            label = kwargs["identity_label"]
            calls.append(label)
            return failed_result(domain, label, "unavailable"), [
                {"identity": label, "outcome": "unavailable", "reason": "missing_connector"}
            ]

        with mock.patch.object(lane, "available_identity_ring", return_value=IDENTITIES), \
             mock.patch.object(lane, "_fetch_domain_one_identity", side_effect=fake_fetch):
            result = lane.fetch_domain("mail_inbox", "prompt", window_days=0, ts="ts", retries=3)

        self.assertEqual(calls, ["edu", "personal-uk", "personal-com"])
        self.assertEqual(result["status"], "unavailable")
        self.assertIsNone(result["served_by"])

    def test_next_call_starts_at_identity_one(self):
        calls = []

        def fake_fetch(domain, prompt, **kwargs):
            label = kwargs["identity_label"]
            calls.append(label)
            return {**failed_result(domain, label), "status": "ok"}, [
                {"identity": label, "outcome": "ok"}
            ]

        with mock.patch.object(lane, "available_identity_ring", return_value=IDENTITIES), \
             mock.patch.object(lane, "_fetch_domain_one_identity", side_effect=fake_fetch):
            lane.fetch_domain("calendar", "prompt", window_days=7, ts="ts1", retries=3)
            lane.fetch_domain("calendar", "prompt", window_days=7, ts="ts2", retries=3)

        self.assertEqual(calls, ["edu", "edu"])

    def test_missing_profile_is_logged_and_skipped(self):
        configured = [
            {"label": "missing", "CODEX_HOME": r"C:\missing"},
            {"label": "authenticated", "CODEX_HOME": r"C:\authenticated"},
        ]
        def profile_reason(identity):
            return "CODEX_HOME missing" if identity["label"] == "missing" else None

        with mock.patch.object(lane, "_log") as log, \
             mock.patch.object(lane, "load_identity_config", return_value=configured), \
             mock.patch.object(lane, "_profile_skip_reason", side_effect=profile_reason):
            available = lane.available_identity_ring()

        self.assertEqual([i["label"] for i in available], ["authenticated"])
        self.assertTrue(any("skipping missing" in call.args[0] for call in log.call_args_list))

    def test_identity_prompt_targets_configured_m365_account(self):
        calls = []
        identities = [{
            "label": "oxford",
            "CODEX_HOME": r"C:\oxford",
            "m365_account": "kevin.lelitte@admin.ox.ac.uk",
        }]

        def fake_fetch(domain, prompt, **kwargs):
            calls.append(prompt)
            return {**failed_result(domain, "oxford"), "status": "ok"}, []

        with mock.patch.object(lane, "available_identity_ring", return_value=identities), \
             mock.patch.object(lane, "_fetch_domain_one_identity", side_effect=fake_fetch):
            lane.fetch_domain("calendar", "BASE PROMPT", window_days=7, ts="ts", retries=1)

        self.assertEqual(len(calls), 1)
        self.assertTrue(calls[0].startswith(
            "TARGET MICROSOFT 365 ACCOUNT (SELECT THIS MAILBOX NOW): "
            "kevin.lelitte@admin.ox.ac.uk"
        ))
        self.assertIn("do not ask me which mailbox to use", calls[0])
        self.assertIn("If the connector offers an account choice, select Oxford", calls[0])
        self.assertIn("Do not use any other connected account", calls[0])

    def test_mail_and_calendar_prompts_share_explicit_oxford_target(self):
        identity = {
            "label": "personal-com",
            "CODEX_HOME": r"C:\com",
            "m365_account": "kevin.lelitte@admin.ox.ac.uk",
        }
        for prompt in (
            lane.build_mail_inbox_prompt("2026-09-25T00:00:00Z"),
            lane.build_mail_sent_prompt("2026-09-25T00:00:00Z"),
            lane.build_calendar_prompt("2026-09-25T00:00:00Z", "2026-10-02T00:00:00Z"),
        ):
            targeted = lane._prompt_for_identity(prompt, identity)
            self.assertLess(targeted.index("SELECT THIS MAILBOX NOW"), targeted.index("Using the Microsoft"))
            self.assertIn("kevin.lelitte@admin.ox.ac.uk", targeted)
            self.assertIn("choose this Oxford account yourself", targeted)

    def test_reported_wrong_account_is_a_mismatch(self):
        events = [{
            "type": "item.completed",
            "item": {"result": {"structured_content": {
                "mailbox": {"emailAddress": "kevin@lelitte.com"}
            }}}
        }]
        reason = lane._account_mismatch_reason(events, "kevin.lelitte@admin.ox.ac.uk")
        self.assertIn("kevin@lelitte.com", reason)

    def test_expected_reported_account_is_accepted(self):
        events = [{
            "type": "item.completed",
            "item": {"result": {"structured_content": {
                "mailbox": {"emailAddress": "kevin.lelitte@admin.ox.ac.uk"}
            }}}
        }]
        self.assertIsNone(lane._account_mismatch_reason(events, "kevin.lelitte@admin.ox.ac.uk"))


if __name__ == "__main__":
    unittest.main()
