"""Import and smoke-run laptop guard/publisher entry points without live I/O."""

import base64
import importlib.util
import os
import sys
import types
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def load_module(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


class EntrypointDryRunTests(unittest.TestCase):
    def test_calendar_guard_dry_run_imports_and_does_not_launch_child(self):
        guard = load_module("lane_b_cal_guard_dry_run", ROOT / "lane_b_cal_guard.py")
        self.assertEqual(guard.main(["--dry-run"]), 0)
        self.assertGreaterEqual(guard.lb.IDENTITY_RING_MAX, 1)

    def test_publishers_dry_run_with_mocked_sources(self):
        drafted = load_module(
            "publish_drafted_replies_dry_run", ROOT / "tools" / "publish_drafted_replies.py"
        )
        drafted_source = {"content": base64.b64encode(b'{"entries": []}').decode("ascii")}
        drafted.gh_get = lambda *args: drafted_source
        old_resolves = os.environ.get("WI_DRAFT_WEBLINK_MAX_RESOLVES")
        os.environ["WI_DRAFT_WEBLINK_MAX_RESOLVES"] = "0"
        try:
            result = drafted.run("mock-token", dry_run=True)
        finally:
            if old_resolves is None:
                os.environ.pop("WI_DRAFT_WEBLINK_MAX_RESOLVES", None)
            else:
                os.environ["WI_DRAFT_WEBLINK_MAX_RESOLVES"] = old_resolves
        self.assertFalse(result["pushed"])
        self.assertEqual(result["entries_published"], 0)

        needs = load_module(
            "publish_needs_reply_dry_run", ROOT / "tools" / "publish_needs_reply.py"
        )
        needs_source = {
            "content": base64.b64encode(b'{"urgent": [], "needs": []}').decode("ascii")
        }
        needs.gh_get = lambda *args: needs_source
        dynamic = types.ModuleType("win32com.client.dynamic")
        dynamic.Dispatch = lambda name: types.SimpleNamespace(
            GetNamespace=lambda kind: types.SimpleNamespace()
        )
        client = types.ModuleType("win32com.client")
        client.dynamic = dynamic
        win32com = types.ModuleType("win32com")
        win32com.client = client
        originals = {name: sys.modules.get(name) for name in (
            "win32com", "win32com.client", "win32com.client.dynamic"
        )}
        sys.modules.update({"win32com": win32com, "win32com.client": client,
                            "win32com.client.dynamic": dynamic})
        try:
            result = needs.run("mock-token", dry_run=True)
        finally:
            for name, original in originals.items():
                if original is None:
                    sys.modules.pop(name, None)
                else:
                    sys.modules[name] = original
        self.assertFalse(result["pushed"])
        self.assertEqual(result["published"], 0)


if __name__ == "__main__":
    unittest.main()
