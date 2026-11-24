import importlib.util
import os
from datetime import datetime, timezone
from pathlib import Path
from unittest import TestCase
from unittest.mock import patch
from urllib.parse import parse_qs, urlsplit

from fastapi import HTTPException

spec = importlib.util.spec_from_file_location(
    "m365_main", Path(__file__).resolve().parents[1] / "app" / "main.py"
)
main = importlib.util.module_from_spec(spec)
with patch.dict(os.environ, {"PYTHON_DOTENV_DISABLED": "1"}):
    spec.loader.exec_module(main)


class CalendarTests(TestCase):
    def setUp(self):
        self.api_key = self.enterContext(patch.object(main, "API_KEY", "test-key"))

    def test_calendar_window_ends_after_twelve_calendar_months(self):
        for now, expected in [
            (datetime(2026, 11, 21, 12, tzinfo=timezone.utc), "2027-11-01T00:00:00+00:00"),
            (datetime(2026, 12, 31, 12, tzinfo=timezone.utc), "2027-12-01T00:00:00+00:00"),
            (datetime(2028, 2, 29, 12, tzinfo=timezone.utc), "2029-02-01T00:00:00+00:00"),
        ]:
            class Clock(datetime):
                @classmethod
                def now(cls, tz=None):
                    return now

            with self.subTest(now=now), patch.object(main, "datetime", Clock), \
                    patch.object(main, "graph_get", return_value={"status": 200, "data": {"value": []}}) as get:
                self.assertEqual(main.calendar_upcoming("test-key")["events"], [])
                query = parse_qs(urlsplit(get.call_args.args[0]).query)
                self.assertEqual(query["startDateTime"], [now.isoformat()])
                self.assertEqual(query["endDateTime"], [expected])
                self.assertEqual(query["$top"], ["100"])

    def test_paginates_all_events_and_preserves_shape(self):
        event = {"subject": "Test", "start": {"dateTime": "2026-01-01T10:00:00"},
                 "end": {"dateTime": "2026-01-01T11:00:00"}, "location": {"displayName": "Office"}}
        next_link = "https://graph.microsoft.com/v1.0/me/calendarView?%24skiptoken=abc"
        with patch.object(main, "graph_get", side_effect=[
            {"status": 200, "data": {"value": [event] * 100, "@odata.nextLink": next_link}},
            {"status": 200, "data": {"value": [event]}},
        ]) as get:
            result = main.calendar_upcoming("test-key")
        self.assertEqual(result["count"], 101)
        self.assertEqual(len(result["events"]), 101)
        self.assertEqual(result["events"][0]["location"], "Office")
        self.assertEqual(get.call_args_list[1].args[0], "/me/calendarView?%24skiptoken=abc")

    def test_invalid_or_repeated_next_links_fail_closed(self):
        for link in [
            "https://example.test/v1.0/me/calendarView?token=x",
            "https://graph.microsoft.com/v1.0/me/messages?token=x",
            "https://graph.microsoft.com/v1.0/me/calendarView?token=x#fragment",
        ]:
            with self.subTest(link=link), patch.object(main, "graph_get", return_value={
                "status": 200, "data": {"value": [], "@odata.nextLink": link},
            }) as get:
                with self.assertRaises(HTTPException) as error:
                    main.calendar_upcoming("test-key")
                self.assertEqual(error.exception.status_code, 502)
                self.assertEqual(get.call_count, 1)

        link = "https://graph.microsoft.com/v1.0/me/calendarView?token=x"
        with patch.object(main, "graph_get", return_value={
            "status": 200, "data": {"value": [], "@odata.nextLink": link},
        }) as get:
            with self.assertRaises(HTTPException):
                main.calendar_upcoming("test-key")
            self.assertEqual(get.call_count, 2)

    def test_errors_on_later_pages_do_not_return_partial_events(self):
        link = "https://graph.microsoft.com/v1.0/me/calendarView?token=x"
        with patch.object(main, "graph_get", side_effect=[
            {"status": 200, "data": {"value": [], "@odata.nextLink": link}},
            {"status": 503, "data": {"error": {}}},
        ]):
            self.assertEqual(main.calendar_upcoming("test-key")["status"], 503)

    def test_requires_api_key_and_rejects_malformed_pages(self):
        with patch.object(main, "graph_get") as get:
            with self.assertRaises(HTTPException) as error:
                main.calendar_upcoming("invalid")
            self.assertEqual(error.exception.status_code, 401)
            get.assert_not_called()
        with patch.object(main, "graph_get", return_value={"status": 200, "data": {}}):
            with self.assertRaises(HTTPException) as error:
                main.calendar_upcoming("test-key")
            self.assertEqual(error.exception.status_code, 502)
