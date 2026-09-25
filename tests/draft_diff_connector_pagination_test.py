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
            "tool": "microsoft_outlook_email.search_messages",
            "arguments": {"from_index": (page - 1) * 2},
            "result": {"structured_content": result},
        },
    }


class DraftDiffPaginationTests(unittest.TestCase):
    def test_collects_continuation_pages_and_deduplicates_boundary_ids(self):
        events = [
            event(1, {"results": [{"id": "a"}, {"id": "b"}], "has_more": True, "next_from_index": 2}),
            event(2, {"results": [{"id": "b"}, {"id": "c"}], "has_more": False}),
        ]
        with mock.patch.object(connector._lb, "run_codex_json", return_value=(events, "raw")):
            rows = connector._list_folder_messages(
                "Drafts", extra_filter=None, top=2, max_total=10,
                tag="test", retries=1, log=lambda _: None,
            )
        self.assertEqual([row["id"] for row in rows], ["a", "b", "c"])

    def test_refuses_a_continuation_left_after_page_bound(self):
        events = [event(1, {"results": [{"id": "a"}], "has_more": True, "next_from_index": 2})]
        with mock.patch.object(connector._lb, "run_codex_json", return_value=(events, "raw")):
            with self.assertRaises(connector.ConnectorResultIncomplete):
                connector._list_folder_messages(
                    "Drafts", extra_filter=None, top=1, max_total=10,
                    tag="test", retries=1, log=lambda _: None,
                )


if __name__ == "__main__":
    unittest.main()
