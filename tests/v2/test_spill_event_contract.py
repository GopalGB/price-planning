import json
import unittest

from openpyxl import Workbook
from openpyxl.styles import Protection

from tests.v2.api_support import APIFixture


def spill_event_fixture(path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Spill"
    sheet.protection.sheet = True
    sheet["A1"] = 1
    sheet["A1"].protection = Protection(locked=False)
    sheet["A2"] = "Keep row two"
    sheet.column_dimensions["B"].protection = Protection(locked=False)
    sheet["B1"] = "=SEQUENCE(A1)"
    sheet["C1"] = "=IF(A1=1,,42)"
    workbook.save(path)


class SpillEventContractTests(unittest.TestCase):
    def setUp(self):
        self.fixture = APIFixture(spill_event_fixture)
        self.client, self.csrf, _ = self.fixture.signup(
            "spill-event@example.test", "Spill Event Owner"
        )
        _, self.copy = self.fixture.make_copy(self.client, self.csrf)

    def tearDown(self):
        self.fixture.close()

    def event_changes(self, revision):
        with self.fixture.database.connect() as connection:
            row = connection.execute(
                "SELECT payload_json FROM pp_events WHERE copy_id=? AND event_type='calc' AND input_revision=?",
                (self.copy["id"], revision),
            ).fetchone()
        return json.loads(row["payload_json"])["changes"]

    def edit_size(self, value, version):
        return self.client.post(f"/api/copies/{self.copy['id']}/ops", {
            "op_id": f"spill-size-{value}", "action": "edit",
            "cells": [{"sheet": 0, "r": 0, "c": 0, "v": value, "expected_version": version}],
        }, {"x-csrf-token": self.csrf})

    def test_spill_events_reclassify_and_restore_a_blank_input(self):
        initial = self.client.get(
            f"/api/copies/{self.copy['id']}/sheets/0/rows?start=1&count=1"
        ).json()["rows"]["1"]["1"]
        with self.fixture.database.connect() as connection:
            absent = connection.execute(
                "SELECT 1 FROM pp_results WHERE copy_id=? AND sheet=0 AND row=1 AND col=1",
                (self.copy["id"],),
            ).fetchone()
        self.assertEqual(initial, {"v": None, "s": 1, "k": "i", "ver": 0})
        self.assertIsNone(absent)

        self.assertEqual(self.edit_size(2, 0).status_code, 200)
        expanded = self.client.get(
            f"/api/copies/{self.copy['id']}/sheets/0/rows?start=1&count=1"
        ).json()["rows"]["1"]["1"]
        self.assertEqual(expanded, {
            "v": 2, "s": 1, "k": "f", "input": True,
            "input_value": None, "ver": 0,
        })
        self.assertIn(
            {"sheet": 0, "r": 1, "c": 1, "v": 2, "s": 1, "k": "f",
             "input": True, "input_value": None, "ver": 0},
            self.event_changes(1),
        )

        self.assertEqual(self.edit_size(1, 1).status_code, 200)
        restored = self.client.get(
            f"/api/copies/{self.copy['id']}/sheets/0/rows?start=1&count=1"
        ).json()["rows"]["1"]["1"]
        self.assertEqual(restored, initial)
        self.assertIn(
            {"sheet": 0, "r": 1, "c": 1, "v": None, "s": 1, "k": "i", "ver": 0},
            self.event_changes(2),
        )

    def test_real_null_formula_result_is_not_a_removal(self):
        from pp.calc_service import CalculationPublication

        with self.fixture.database.transaction() as connection:
            rows = connection.execute(
                "SELECT sheet,row,col,value_json FROM pp_results WHERE copy_id=?",
                (self.copy["id"],),
            ).fetchall()
            connection.execute(
                "UPDATE pp_copies SET input_revision=1 WHERE id=?", (self.copy["id"],)
            )
        results = [
            (item["sheet"], item["row"], item["col"],
             None if (item["sheet"], item["row"], item["col"]) == (0, 0, 2)
             else json.loads(item["value_json"]))
            for item in rows
        ]
        self.assertTrue(self.fixture.app.state.services.calculations.complete(
            CalculationPublication(self.copy["id"], 1, results)
        ))
        row = self.client.get(
            f"/api/copies/{self.copy['id']}/sheets/0/rows?start=0&count=1"
        ).json()["rows"]["0"]["2"]

        self.assertEqual(row, {"v": None, "s": 0, "k": "f"})
        self.assertIn(
            {"sheet": 0, "r": 0, "c": 2, "v": None, "s": 0, "k": "f"},
            self.event_changes(1),
        )


if __name__ == "__main__":
    unittest.main()
