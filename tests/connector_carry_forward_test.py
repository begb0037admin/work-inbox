import copy
import unittest
from datetime import datetime, timezone, timedelta

from connector_carry_forward import reconcile_domains


NOW = datetime(2026, 9, 24, 21, 45, tzinfo=timezone.utc)


def status(as_of, name="ok"):
    return {"status": name, "as_of": as_of.isoformat()}


class ConnectorCarryForwardTests(unittest.TestCase):
    def test_failed_domain_keeps_previous_data_as_one_snapshot(self):
        previous = {
            "calToday": [{"time": "09:00", "title": "Daily catch-up"}],
            "calTomorrow": [{"time": "11:00", "title": "Planning"}],
            "calDay2": [], "calDay3": [],
            "calFull": [
                {"date": "2026-09-24", "items": [{"time": "09:00", "title": "Daily catch-up"}]},
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
        self.assertEqual(result["calToday"], previous["calToday"])
        self.assertEqual(result["calTomorrow"], previous["calTomorrow"])
        self.assertEqual(result["absences"], previous["absences"])
        self.assertEqual(result["connector_status"]["calendar"]["status"], "carried_forward")

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
