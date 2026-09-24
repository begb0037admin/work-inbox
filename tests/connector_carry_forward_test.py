import copy
import unittest
from datetime import datetime, timezone, timedelta

from connector_carry_forward import reconcile_domains


NOW = datetime(2026, 9, 24, 21, 45, tzinfo=timezone.utc)


def status(as_of, name="ok"):
    return {"status": name, "as_of": as_of.isoformat()}


class ConnectorCarryForwardTests(unittest.TestCase):
    def test_only_explicit_ok_is_fresh_for_every_non_success_status(self):
        previous = {
            "teams": [{"channel": "HR", "preview": "last good"}],
            "connector_status": {"teams": status(NOW - timedelta(hours=2))},
        }
        for failed_status in ("n/a", "missing", "error", "timeout", "unavailable", "halt"):
            with self.subTest(failed_status=failed_status):
                current = {
                    "teams": [{"channel": "HR", "preview": "must not be fresh"}],
                    "connector_status": {"teams": failed_status},
                }
                result, _, carried = reconcile_domains(
                    current, previous, {}, {"teams": failed_status},
                    enabled_domains=["teams"], now=NOW,
                )
                self.assertEqual(carried, ["teams"])
                self.assertEqual(result["teams"], previous["teams"])
                self.assertEqual(result["connector_status"]["teams"]["status"], "carried_forward")

    def test_failed_domain_keeps_previous_data_as_one_snapshot(self):
        previous = {
            "calToday": [{"time": "09:00", "title": "Monday catch-up"}],
            "calTomorrow": [{"time": "11:00", "title": "Old planning"}],
            "calDay2": [], "calDay3": [],
            "calFull": [
                {"date": "2026-09-21", "isToday": True, "items": [{"time": "09:00", "title": "Monday catch-up"}]},
                {"date": "2026-09-24", "isToday": False, "items": [{"time": "09:00", "title": "Daily catch-up"}]},
                {"date": "2026-09-25", "items": [{"time": "11:00", "title": "Planning"}]},
            ],
            "absences": ["A colleague - returns Friday"],
            "refreshed_at": "Monday 21 September · 16:00",
            "connector_status": {"calendar": status(NOW - timedelta(days=3))},
        }
        current = {
            "calToday": [], "calTomorrow": [], "calDay2": [], "calDay3": [],
            "calFull": [{"date": "2026-09-24", "items": []}], "absences": [],
            "connector_status": {"calendar": "quota_exceeded"},
        }

        result, _, carried = reconcile_domains(
            current, previous, {}, {"calendar": "quota_exceeded"},
            enabled_domains=["calendar"], now=NOW,
        )

        self.assertEqual(carried, ["calendar"])
        self.assertEqual(result["calToday"], [{"time": "09:00", "title": "Daily catch-up"}])
        self.assertEqual(result["calTomorrow"], [{"time": "11:00", "title": "Planning"}])
        self.assertEqual(result["calDay2"], [])
        self.assertEqual(result["calDay3"], [])
        self.assertNotIn("Monday catch-up", [item["title"] for item in result["calToday"]])
        self.assertFalse(result["calFull"][0]["isToday"])
        self.assertTrue(result["calFull"][1]["isToday"])
        self.assertEqual(result["absences"], previous["absences"])
        self.assertEqual(result["connector_status"]["calendar"]["status"], "carried_forward")

    def test_missing_or_undated_full_days_clear_stored_relative_buckets(self):
        previous = {
            "calToday": [{"time": "09:00", "title": "Stale Monday"}],
            "calTomorrow": [{"time": "11:00", "title": "Stale Tuesday"}],
            "calDay2": [{"time": "12:00", "title": "Stale Wednesday"}],
            "calDay3": [{"time": "13:00", "title": "Stale Thursday"}],
            "calFull": [{"items": [{"time": "09:00", "title": "No date"}]}],
            "connector_status": {"calendar": status(NOW - timedelta(days=1))},
        }
        current = {"calToday": [], "calTomorrow": [], "calDay2": [], "calDay3": [], "connector_status": {"calendar": "timeout"}}

        result, _, carried = reconcile_domains(
            current, previous, {}, {"calendar": "timeout"},
            enabled_domains=["calendar"], now=NOW,
        )

        self.assertEqual(carried, ["calendar"])
        for field in ("calToday", "calTomorrow", "calDay2", "calDay3"):
            self.assertEqual(result[field], [])

    def test_data_older_than_seven_days_is_not_carried(self):
        previous = {
            "teams": [{"channel": "HR", "preview": "old"}],
            "connector_status": {"teams": status(NOW - timedelta(days=8, minutes=1))},
        }
        current = {"teams": [], "connector_status": {"teams": "timeout"}}

        result, _, carried = reconcile_domains(
            current, previous, {}, {"teams": "timeout"},
            enabled_domains=["teams"], now=NOW,
        )

        self.assertEqual(carried, [])
        self.assertEqual(result["teams"], [])
        self.assertEqual(result["connector_status"]["teams"]["status"], "unavailable")
        self.assertEqual(result["connector_status"]["teams"]["reason"], "last_good_data_older_than_7_days")

    def test_successful_fetch_replaces_previous_domain(self):
        previous = {"teams": [{"channel": "Old", "preview": "old"}], "refreshed_at": "Monday 21 September · 16:00"}
        current = {"teams": [{"channel": "New", "preview": "new"}], "connector_status": {"teams": "ok"}}
        before = copy.deepcopy(current)

        result, cache, carried = reconcile_domains(
            current, previous, {}, {"teams": "ok"},
            {"teams": NOW.isoformat()}, enabled_domains=["teams"], now=NOW,
        )

        self.assertEqual(carried, [])
        self.assertEqual(result["teams"], before["teams"])
        self.assertEqual(result["connector_status"]["teams"]["status"], "ok")
        self.assertEqual(cache["domains"]["teams"]["data"]["teams"], before["teams"])


if __name__ == "__main__":
    unittest.main()
