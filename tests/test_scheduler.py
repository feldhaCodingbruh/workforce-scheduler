import unittest
from copy import deepcopy
from unittest.mock import patch

import app


class AvailabilityParserTests(unittest.TestCase):
    def test_normalize_text_removes_lithuanian_diacritics(self):
        self.assertEqual(app.normalize_text("Rytinė norėčiau"), "rytine noreciau")
        self.assertEqual(app.normalize_text("Galiu šeštadienį"), "galiu sestadieni")

    def test_morning_preference_with_diacritics_is_recognized(self):
        parsed = app.parse_single_line_rule_based("Rytinė norėčiau")

        self.assertEqual(parsed["type"], "available")
        self.assertEqual(parsed["preference"], "morning")

    def test_common_lithuanian_availability_formats(self):
        cases = {
            "galiu nuo 16:40": ("from_time", "16:40", None),
            "iki 17": ("until_time", None, "17:00"),
            "10-18": ("time_range", "10:00", "18:00"),
            "negaliu (egzas)": ("unavailable", None, None),
        }

        for text, expected in cases.items():
            with self.subTest(text=text):
                parsed = app.parse_single_line_rule_based(text)
                self.assertEqual(
                    (parsed["type"], parsed["from_time"], parsed["until_time"]),
                    expected,
                )

    def test_split_availability_is_recognized(self):
        parsed = app.parse_single_line_rule_based("iki14/nuo19")

        self.assertEqual(parsed["type"], "split")
        self.assertEqual(parsed["until_time"], "14:00")
        self.assertEqual(parsed["second_from_time"], "19:00")

    def test_blank_lines_preserve_calendar_day_positions(self):
        lines = app.parse_availability_lines("galiu\n\nnegaliu\nnuo 17\n")

        self.assertEqual(lines, ["galiu", "", "negaliu", "nuo 17"])

    def test_blank_day_does_not_count_as_complete_input(self):
        self.assertEqual(app.get_worker_status(3, 3, 2), "Truksta 1 d.")
        self.assertEqual(app.get_worker_status(3, 3, 3), "Gerai")


class ScheduleRegressionTests(unittest.TestCase):
    def test_demand_above_template_creates_override_full_shift(self):
        shifts, warnings = app.build_requested_shifts("location-g", 2026, 6, 1, 3.0)

        self.assertEqual(len(shifts), 3)
        self.assertEqual([shift["slot_kind"] for shift in shifts], ["Pilna", "Pilna", "Pilna+"])
        self.assertEqual(shifts[2]["start"], "13:00")
        self.assertEqual(shifts[2]["end"], "21:30")
        self.assertTrue(shifts[2]["template_override"])
        self.assertIn("Papildomos pamainos kuriamos kaip override", warnings[0])

    def test_schedule_generation_assigns_override_full_shift(self):
        settings = {"year": 2026, "month": 6, "full_time_hours": 160}
        demand = {day: (3.0 if day == 1 else 0.0) for day in range(1, 31)}
        workers = [
            {"id": "a", "name": "Austeja", "etatas": "1.0", "availability_raw": "\n".join(["galiu"] * 30)},
            {"id": "b", "name": "Benas", "etatas": "1.0", "availability_raw": "\n".join(["galiu"] * 30)},
            {"id": "c", "name": "Lukas", "etatas": "1.0", "availability_raw": "\n".join(["galiu"] * 30)},
        ]
        runtime_workers = [
            app.build_worker_runtime(worker, settings, "location-g", demand)
            for worker in workers
        ]

        schedule, _ = app.generate_month_schedule(runtime_workers, settings, "location-g", demand)
        day_one = schedule[0]

        self.assertEqual(len(day_one["assignments"]), 3)
        self.assertEqual([item["slot_kind"] for item in day_one["assignments"]], ["Pilna", "Pilna", "Pilna+"])
        self.assertTrue(all(item["worker_name"] for item in day_one["assignments"]))
        self.assertFalse(any("Nera darbuotojo" in warning for warning in day_one["warnings"]))

    def test_monthly_target_uses_configured_full_time_hours(self):
        self.assertEqual(app.etatas_to_month_hours("0.5", 168), 84)
        self.assertEqual(app.etatas_to_month_hours("0.75", 168), 126)

    def test_worker_with_partial_availability_stays_in_summary(self):
        settings = {"year": 2026, "month": 6, "full_time_hours": 160}
        demand = {day: (1.0 if day == 1 else 0.0) for day in range(1, 31)}
        workers = [
            {"id": "a", "name": "Lukas", "etatas": "0.5", "availability_raw": "galiu"},
            {"id": "b", "name": "Benas", "etatas": "0.5", "availability_raw": "galiu"},
        ]
        runtime_workers = [
            app.build_worker_runtime(worker, settings, "location-d", demand)
            for worker in workers
        ]

        _, summary = app.generate_month_schedule(runtime_workers, settings, "location-d", demand)

        self.assertEqual({item["name"] for item in summary}, {"Lukas", "Benas"})


class WorkerEditingTests(unittest.TestCase):
    def setUp(self):
        self.location = app.app_data["locations"]["location-a"]
        self.original_location = deepcopy(self.location)

    def tearDown(self):
        self.location.clear()
        self.location.update(self.original_location)

    def test_worker_can_be_updated_without_writing_real_data(self):
        self.location["workers"] = [
            {"id": "worker-1", "name": "Old", "etatas": "0.5", "availability_raw": "galiu"}
        ]
        self.location["generated_schedule"] = [{"day": 1}]
        client = app.app.test_client()

        with patch.object(app, "save_app_data"):
            response = client.post(
                "/worker/worker-1/update",
                data={
                    "location_id": "location-a",
                    "name": "New",
                    "etatas": "0.75",
                    "availability": "galiu\n\nnegaliu",
                },
            )

        self.assertEqual(response.status_code, 302)
        self.assertEqual(self.location["workers"][0]["name"], "New")
        self.assertEqual(self.location["workers"][0]["etatas"], "0.75")
        self.assertEqual(self.location["workers"][0]["availability_raw"], "galiu\n\nnegaliu")
        self.assertEqual(self.location["generated_schedule"], [])


class WorkbookExportTests(unittest.TestCase):
    def test_compact_schedule_export_matches_manager_view(self):
        settings = {"year": 2026, "month": 6, "full_time_hours": 160}
        generated_schedule = [
            {
                "day": 12,
                "weekday_name": "Pn",
                "demand_label": "5",
                "assignments": [
                    {"slot_kind": "Pilna", "shift_time": "10:00-17:00", "worker_name": "Deividas"},
                    {"slot_kind": "Pilna", "shift_time": "13:00-21:30", "worker_name": "Liepa"},
                ],
                "warnings": [],
            },
            {
                "day": 13,
                "weekday_name": "St",
                "demand_label": "3",
                "assignments": [
                    {"slot_kind": "Pilna", "shift_time": "10:00-16:00", "worker_name": "Liepa"},
                    {"slot_kind": "Pilna", "shift_time": "13:30-21:30", "worker_name": None},
                ],
                "warnings": ["Nera darbuotojo Pilna 2"],
            },
        ]

        workbook = app.create_schedule_workbook(
            "Location A",
            settings,
            generated_schedule,
            [
                {
                    "name": "Deividas",
                    "etatas": "0.75",
                    "assigned_shifts": 1,
                    "assigned_hours": 7,
                    "target_hours": 120,
                    "hours_difference": -113,
                    "weekend_days": 0,
                    "closing_shifts": 0,
                }
            ],
        )
        sheet = workbook["Grafikas"]

        self.assertEqual(sheet["A1"].value, "5")
        self.assertEqual(sheet["B1"].value, "P")
        self.assertEqual(sheet["C1"].value, "birželio 12")
        self.assertEqual(sheet["D1"].value, "Deividas 10:00-17:00")
        self.assertEqual(sheet["E1"].value, "Liepa 13:00-21:30")
        self.assertEqual(sheet["D1"].fill.fgColor.rgb, "00F5A623")
        self.assertEqual(sheet["E2"].value, "13:30-21:30")
        self.assertEqual(sheet["E2"].fill.fgColor.rgb, "00B00000")
        self.assertIn("Ispejimai", workbook.sheetnames)
        self.assertEqual(sheet["D1"].font.name, "Arial")
        self.assertEqual(workbook["Ispejimai"]["B2"].font.name, "Arial")
        self.assertEqual(workbook["Suvestine"]["A2"].font.name, "Arial")


if __name__ == "__main__":
    unittest.main()
