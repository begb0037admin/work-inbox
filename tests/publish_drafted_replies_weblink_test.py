"""Focused regression checks for Drafted Replies OWA link publishing."""

import importlib.util
import os
import sys
import types
import unittest


TOOLS = os.path.join(os.path.dirname(__file__), "..", "tools")
sys.path.insert(0, os.path.abspath(TOOLS))
SPEC = importlib.util.spec_from_file_location(
    "publish_drafted_replies", os.path.join(TOOLS, "publish_drafted_replies.py")
)
publisher = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(publisher)


class DraftWeblinkTests(unittest.TestCase):
    def test_normalize_never_emits_com_and_rejects_non_owa_urls(self):
        base = {
            "subject": "Example", "sender_tier": "other", "draft_text": "Reply",
            "composed_at": "2026-09-24T10:00:00Z", "source_entry_id": "legacy-entry",
            "draft_id": "draft-1",
        }
        entry, reason = publisher.normalize_entry(base)
        self.assertIsNone(reason)
        self.assertEqual(entry["open_mode"], "none")
        self.assertNotIn("web_link", entry)

        entry, reason = publisher.normalize_entry({
            **base, "web_link": "https://outlook.office365.com/owa/?ItemID=abc"
        })
        self.assertIsNone(reason)
        self.assertEqual(entry["open_mode"], "web")
        self.assertEqual(entry["web_link"], "https://outlook.office365.com/owa/?ItemID=abc")

        entry, reason = publisher.normalize_entry({**base, "web_link": "https://example.com/owa"})
        self.assertIsNone(reason)
        self.assertEqual(entry["open_mode"], "none")

    def test_resolution_respects_per_run_cap_and_caches_valid_owa_result(self):
        calls = []
        original_module = sys.modules.get("lane_b_call1")
        original_load = publisher.load_weblink_cache
        original_save = publisher.save_weblink_cache
        original_limit = os.environ.get("WI_DRAFT_WEBLINK_MAX_RESOLVES")
        try:
            sys.modules["lane_b_call1"] = types.SimpleNamespace(
                resolve_mail_weblink_by_subject=lambda subject, received: calls.append(subject) or "https://outlook.office.com/owa/?ItemID=" + subject
            )
            cache = {}
            publisher.load_weblink_cache = lambda: cache
            publisher.save_weblink_cache = lambda value: cache.update(value)
            os.environ["WI_DRAFT_WEBLINK_MAX_RESOLVES"] = "2"
            entries = [{"draft_id": f"draft-{i}", "subject": str(i), "received": "", "open_mode": "none"} for i in range(4)]
            self.assertEqual(publisher.resolve_missing_weblinks(entries), 2)
            self.assertEqual(calls, ["0", "1"])
            self.assertEqual(entries[0]["open_mode"], "web")
            self.assertEqual(entries[2]["open_mode"], "none")
        finally:
            publisher.load_weblink_cache = original_load
            publisher.save_weblink_cache = original_save
            if original_module is None:
                sys.modules.pop("lane_b_call1", None)
            else:
                sys.modules["lane_b_call1"] = original_module
            if original_limit is None:
                os.environ.pop("WI_DRAFT_WEBLINK_MAX_RESOLVES", None)
            else:
                os.environ["WI_DRAFT_WEBLINK_MAX_RESOLVES"] = original_limit

    def test_resolution_uses_one_batch_resolver_for_current_drafts(self):
        calls = []
        original_module = sys.modules.get("lane_b_call1")
        original_load = publisher.load_weblink_cache
        original_save = publisher.save_weblink_cache
        try:
            sys.modules["lane_b_call1"] = types.SimpleNamespace(
                resolve_mail_weblinks_by_subjects=lambda subjects: calls.append(subjects) or [
                    {"subject": subject, "web_link": "https://outlook.office365.com/owa/?ItemID=" + str(i)}
                    for i, subject in enumerate(subjects)
                ]
            )
            cache = {}
            publisher.load_weblink_cache = lambda: cache
            publisher.save_weblink_cache = lambda value: cache.update(value)
            entries = [{
                "draft_id": f"draft-{i}", "subject": f"Subject {i}",
                "received": "", "open_mode": "none",
            } for i in range(4)]
            self.assertEqual(publisher.resolve_missing_weblinks(entries), 4)
            self.assertEqual(len(calls), 1)
            self.assertEqual(calls[0], [f"Subject {i}" for i in range(4)])
            self.assertTrue(all(entry["open_mode"] == "web" for entry in entries))
            self.assertEqual(len(cache), 4)
        finally:
            publisher.load_weblink_cache = original_load
            publisher.save_weblink_cache = original_save
            if original_module is None:
                sys.modules.pop("lane_b_call1", None)
            else:
                sys.modules["lane_b_call1"] = original_module

    def test_batch_resolution_matches_reply_prefix_to_original_subject(self):
        original_module = sys.modules.get("lane_b_call1")
        original_load = publisher.load_weblink_cache
        original_save = publisher.save_weblink_cache
        try:
            sys.modules["lane_b_call1"] = types.SimpleNamespace(
                resolve_mail_weblinks_by_subjects=lambda subjects: [
                    {
                        "subject": "RE: My Development Insight reports",
                        "web_link": "https://outlook.office365.com/owa/?ItemID=reply-prefix",
                    }
                ]
            )
            cache = {}
            publisher.load_weblink_cache = lambda: cache
            publisher.save_weblink_cache = lambda value: cache.update(value)
            entries = [
                {
                    "draft_id": "draft-re-prefix", "subject": "Re: My Development Insight reports",
                    "received": "", "open_mode": "none",
                },
                {
                    "draft_id": "draft-unmatched", "subject": "Unmatched subject",
                    "received": "", "open_mode": "none",
                },
            ]
            self.assertEqual(publisher.resolve_missing_weblinks(entries), 2)
            self.assertEqual(entries[0]["open_mode"], "web")
            self.assertIn("outlook.office365.com", entries[0]["web_link"])
        finally:
            publisher.load_weblink_cache = original_load
            publisher.save_weblink_cache = original_save
            if original_module is None:
                sys.modules.pop("lane_b_call1", None)
            else:
                sys.modules["lane_b_call1"] = original_module

    def test_batch_resolution_prefers_one_exact_subject_when_normalized_subject_is_ambiguous(self):
        original_module = sys.modules.get("lane_b_call1")
        original_load = publisher.load_weblink_cache
        original_save = publisher.save_weblink_cache
        try:
            sys.modules["lane_b_call1"] = types.SimpleNamespace(
                resolve_mail_weblinks_by_subjects=lambda subjects: [
                    {
                        "subject": "Re: My Development Insight reports",
                        "web_link": "https://outlook.office365.com/owa/?ItemID=exact",
                    },
                    {
                        "subject": "RE: My Development Insight reports",
                        "web_link": "https://outlook.office365.com/owa/?ItemID=other-case",
                    },
                ]
            )
            cache = {}
            publisher.load_weblink_cache = lambda: cache
            publisher.save_weblink_cache = lambda value: cache.update(value)
            entries = [{
                "draft_id": "draft-exact", "subject": "Re: My Development Insight reports",
                "received": "", "open_mode": "none",
            }, {
                "draft_id": "draft-ambiguous", "subject": "Fwd: My Development Insight reports",
                "received": "", "open_mode": "none",
            }]
            self.assertEqual(publisher.resolve_missing_weblinks(entries), 2)
            self.assertEqual(entries[0]["web_link"], "https://outlook.office365.com/owa/?ItemID=exact")
            self.assertEqual(entries[1]["open_mode"], "none")
        finally:
            publisher.load_weblink_cache = original_load
            publisher.save_weblink_cache = original_save
            if original_module is None:
                sys.modules.pop("lane_b_call1", None)
            else:
                sys.modules["lane_b_call1"] = original_module


if __name__ == "__main__":
    unittest.main()
