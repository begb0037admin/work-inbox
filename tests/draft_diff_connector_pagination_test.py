"""Regression checks for bounded connector pagination in draft diff."""

import sys
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import draft_diff_connector as connector


def event(page, result):
    return {
        "type": "item.completed",
        "item": {
            "type": "mcp_tool_call",
            "status": "completed",
            "server": "codex_apps",
            "tool": "microsoft_outlook_email.list_messages",
            "arguments": {"folder_id": "drafts", "page": page},
            "result": {"structured_content": result},
        },
    }


class DraftDiffPaginationTests(unittest.TestCase):
    def test_accepts_advanced_next_list_messages_call_when_page_omits_token(self):
        events = [
            event(1, {"value": [{"id": "a"}]}),
            event(2, {"value": [{"id": "b"}]}),
        ]
        events[0]["item"]["arguments"] = {"folder_id": "drafts", "skip": 0}
        events[1]["item"]["arguments"] = {"folder_id": "drafts", "skip": 200}
        with mock.patch.object(connector._lb, "run_codex_json", return_value=(events, "raw")):
            rows = connector._list_folder_messages(
                "Drafts", extra_filter=None, top=200, max_total=500,
                tag="test", retries=1, log=lambda _: None,
            )
        self.assertEqual([row["id"] for row in rows], ["a", "b"])

    def test_prompt_matches_bridge_folder_path_and_oxford_target(self):
        prompt = connector._build_folder_prompt("Drafts", extra_filter=None, top=1000)
        targeted = connector._lb._prompt_for_identity(
            prompt,
            {"m365_account": "kevin.lelitte@admin.ox.ac.uk"},
        )
        self.assertIn("Use only the Oxford Microsoft 365 mailbox kevin.lelitte@admin.ox.ac.uk", targeted)
        self.assertIn("list_mail_folders once", prompt)
        self.assertIn("list_messages", prompt)
        self.assertNotIn("find_mail_folder", prompt)

    def test_uses_lane_b_ordered_ring_helper(self):
        identities = [
            {"label": "edu", "CODEX_HOME": r"C:\edu", "m365_account": "kevin.lelitte@admin.ox.ac.uk"},
            {"label": "personal-uk", "CODEX_HOME": r"C:\uk", "m365_account": "kevin.lelitte@admin.ox.ac.uk"},
            {"label": "personal-com", "CODEX_HOME": r"C:\com", "m365_account": "kevin.lelitte@admin.ox.ac.uk"},
        ]
        with mock.patch.object(connector._lb, "ordered_identity_ring", return_value=identities) as helper:
            with mock.patch.object(connector._lb, "run_codex_json", return_value=([], "raw")):
                with self.assertRaises(connector.ConnectorResultUnavailable):
                    connector._list_folder_messages(
                        "Drafts", extra_filter=None, top=2, max_total=10,
                        tag="test", retries=1, log=lambda _: None,
                    )
        helper.assert_called_once_with("mail_sent")

    def test_mail_prefers_personal_com_like_bridge_ring(self):
        identities = [
            {"label": "edu", "CODEX_HOME": r"C:\edu", "m365_account": "kevin.lelitte@admin.ox.ac.uk"},
            {"label": "personal-uk", "CODEX_HOME": r"C:\uk", "m365_account": "kevin.lelitte@admin.ox.ac.uk"},
            {"label": "personal-com", "CODEX_HOME": r"C:\com", "m365_account": "kevin.lelitte@admin.ox.ac.uk"},
        ]
        homes = []

        def run(_prompt, **kwargs):
            homes.append(kwargs["codex_home"])
            return ([event(1, {"results": [{"id": "draft"}], "has_more": False})], "raw")

        with mock.patch.object(
            connector._lb,
            "ordered_identity_ring",
            return_value=[identities[2], identities[0], identities[1]],
        ):
            with mock.patch.object(connector._lb, "run_codex_json", side_effect=run):
                rows = connector._list_folder_messages(
                    "Drafts", extra_filter=None, top=2, max_total=10,
                    tag="test", retries=1, log=lambda _: None,
                )
        self.assertEqual(homes, [r"C:\com"])
        self.assertEqual([row["id"] for row in rows], ["draft"])

    def test_extracts_payload_wrapped_events(self):
        wrapped = {"payload": event(1, {"results": [], "has_more": False})}
        calls = connector._lb.extract_tool_calls([wrapped])
        self.assertEqual([call["tool"] for call in calls], ["microsoft_outlook_email.list_messages"])

    def test_accepts_camel_case_structured_content_events(self):
        raw_event = event(1, {"value": [{"id": "camel"}]})
        raw_event["item"]["result"] = {"structuredContent": {"value": [{"id": "camel"}]}}
        with mock.patch.object(connector._lb, "run_codex_json", return_value=([raw_event], "raw")):
            rows = connector._list_folder_messages(
                "Drafts", extra_filter=None, top=2, max_total=10,
                tag="test", retries=1, log=lambda _: None,
            )
        self.assertEqual([row["id"] for row in rows], ["camel"])

    def test_collects_continuation_pages_and_deduplicates_boundary_ids(self):
        events = [
            event(1, {"value": [{"id": "a"}, {"id": "b"}], "next_link": "page-2"}),
            event(2, {"value": [{"id": "b"}, {"id": "c"}]}),
        ]
        with mock.patch.object(connector._lb, "run_codex_json", return_value=(events, "raw")):
            rows = connector._list_folder_messages(
                "Drafts", extra_filter=None, top=2, max_total=10,
                tag="test", retries=1, log=lambda _: None,
            )
        self.assertEqual([row["id"] for row in rows], ["a", "b", "c"])

    def test_requires_continuation_before_another_page(self):
        events = [
            event(1, {"value": [{"id": "a"}]}),
            event(2, {"value": [{"id": "b"}]}),
        ]
        with mock.patch.object(connector._lb, "run_codex_json", return_value=(events, "raw")):
            with self.assertRaises(connector.ConnectorResultIncomplete):
                connector._list_folder_messages(
                    "Drafts", extra_filter=None, top=1, max_total=10,
                    tag="test", retries=1, log=lambda _: None,
                )

    def test_refuses_a_continuation_left_after_page_bound(self):
        events = [event(1, {"value": [{"id": "a"}], "next_link": "page-2"})]
        with mock.patch.object(connector._lb, "run_codex_json", return_value=(events, "raw")):
            with self.assertRaises(connector.ConnectorResultIncomplete):
                connector._list_folder_messages(
                    "Drafts", extra_filter=None, top=1, max_total=10,
                    tag="test", retries=1, log=lambda _: None,
                )


if __name__ == "__main__":
    unittest.main()
