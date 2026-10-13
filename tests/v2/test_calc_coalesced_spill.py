import json
import unittest

from openpyxl import Workbook
from openpyxl.styles import Protection

from pp.calc_service import CalculationPublication
from pp.engine.pool import CalculationRequest
from tests.v2.api_support import APIFixture


def coalesced_spill_inputs(path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Coalesced"
    sheet.protection.sheet = True
    sheet["B1"] = "=SEQUENCE(2)"
    sheet["B2"] = 7
    sheet["B2"].protection = Protection(locked=False)
    sheet["D1"] = 10
    sheet["D1"].protection = Protection(locked=False)
    workbook.save(path)


class CoalescedSpillPublicationTests(unittest.TestCase):
    def setUp(self):
        self.fixture = APIFixture(coalesced_spill_inputs)
        self.addCleanup(self.fixture.close)
        client, csrf, self.user = self.fixture.signup(
            "coalesced-spill@example.test", "Coalesced Owner"
        )
        _, copy = self.fixture.make_copy(client, csrf)
        self.copy_id = copy["id"]
        self.services = self.fixture.app.state.services

    def apply(self, op_id, cell):
        return self.services.operations.apply(
            self.copy_id,
            self.user["id"],
            {"op_id": op_id, "action": "edit", "cells": [cell]},
        )

    def calculate(self, revision):
        changes = self.services.operations.revision_changes(self.copy_id, revision)
        return self.fixture.pool.apply(
            CalculationRequest(self.copy_id, revision, self.fixture.source, changes)
        )

    def event(self, revision):
        with self.fixture.database.connect() as connection:
            row = connection.execute(
                "SELECT payload_json FROM pp_events WHERE copy_id=? "
                "AND event_type='calc' AND input_revision=?",
                (self.copy_id, revision),
            ).fetchone()
        return json.loads(row["payload_json"])

    def publish(self, revision, result):
        return self.services.calculations.complete(
            CalculationPublication(self.copy_id, revision, result["results"])
        )

    def assert_coalesced_event(self, payload):
        changes = {
            (item["sheet"], item["r"], item["c"]): item
            for item in payload["changes"]
        }
        self.assertIn((0, 1, 1), changes)
        self.assertEqual(payload["from_calculated_revision"], 1)
        self.assertEqual(
            {
                key: changes[(0, 1, 1)][key]
                for key in ("v", "k", "input", "input_value", "ver")
            },
            {"v": 2, "k": "f", "input": True, "input_value": None, "ver": 2},
        )

    def assert_stale_publication_refused(self, results):
        self.assertFalse(
            self.services.calculations.complete(
                CalculationPublication(self.copy_id, 2, results)
            )
        )
        with self.fixture.database.connect() as connection:
            copy = connection.execute(
                "SELECT calculated_revision FROM pp_copies WHERE id=?", (self.copy_id,)
            ).fetchone()
            job = connection.execute(
                "SELECT status FROM pp_calc_jobs WHERE copy_id=? AND revision=2",
                (self.copy_id,),
            ).fetchone()
            rows = connection.execute(
                "SELECT sheet,row,col,value_json FROM pp_overrides WHERE copy_id=?",
                (self.copy_id,),
            ).fetchall()
        overrides = {
            (row["sheet"], row["row"], row["col"]): json.loads(row["value_json"])
            for row in rows
        }
        self.assertEqual(overrides[(0, 1, 1)], None)
        self.assertEqual(overrides[(0, 0, 3)], 11)
        self.assertEqual((copy["calculated_revision"], job["status"]), (3, "superseded"))

    def test_newer_publication_republishes_inputs_touched_since_prior_calculation(self):
        self.apply(
            "clear-one",
            {"sheet": 0, "r": 1, "c": 1, "v": None, "expected_version": 0},
        )
        self.assertTrue(self.publish(1, self.calculate(1)))

        self.apply(
            "clear-two",
            {"sheet": 0, "r": 1, "c": 1, "v": None, "expected_version": 1},
        )
        self.apply(
            "edit-d1",
            {"sheet": 0, "r": 0, "c": 3, "v": 11, "expected_version": 0},
        )
        third = self.calculate(3)
        results = {item[:3]: item[3] for item in third["results"]}

        self.assertEqual(third["changes"], [])
        self.assertEqual(results[(0, 1, 1)], 2)
        self.assertTrue(self.publish(3, third))
        self.assert_coalesced_event(self.event(3))
        self.assert_stale_publication_refused(third["results"])


if __name__ == "__main__":
    unittest.main()
