import json
import unittest

from openpyxl import Workbook
from openpyxl.styles import Protection

from tests.v2.api_support import APIFixture


def occupied_input_spill_fixture(path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Spill"
    sheet.protection.sheet = True
    sheet["B1"] = "=SEQUENCE(2)"
    sheet["B2"] = 7
    sheet["B2"].protection = Protection(locked=False)
    workbook.save(path)


def implicit_input_spill_fixture(path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Implicit Spill"
    sheet.protection.sheet = True
    sheet["A1"] = 1
    sheet["A1"].protection = Protection(locked=False)
    sheet["A2"] = "bounds marker"
    sheet["B1"] = "=SEQUENCE(A1)"
    sheet.column_dimensions["B"].protection = Protection(locked=False)
    workbook.save(path)


class OccupiedInputSpillUndoTests(unittest.TestCase):
    def setUp(self):
        self.fixture = APIFixture(occupied_input_spill_fixture)
        self.client, self.csrf, self.user = self.fixture.signup(
            "spill-undo@example.test", "Spill Undo Owner"
        )
        _, self.copy = self.fixture.make_copy(self.client, self.csrf)

    def tearDown(self):
        self.fixture.close()

    def operation(self, op_id, value, version, action="edit"):
        return self.client.post(f"/api/copies/{self.copy['id']}/ops", {
            "op_id": op_id,
            "action": action,
            "cells": [{
                "sheet": 0, "r": 1, "c": 1,
                "v": value, "expected_version": version,
            }],
        }, {"x-csrf-token": self.csrf})

    def rows(self):
        return self.client.get(
            f"/api/copies/{self.copy['id']}/sheets/0/rows?start=0&count=2"
        ).json()

    def test_clear_spill_and_undo_restores_the_immutable_input(self):
        initial_input = self.rows()["rows"]["1"]["1"]
        self.assertTrue(self.fixture.snapshot.is_input(0, 1, 1))
        self.assertIn("0:1:1", self.fixture.snapshot.input_mask)
        self.assertEqual(
            (initial_input["v"], initial_input["k"], initial_input["ver"]),
            (7, "i", 0),
        )

        initial_anchor = self.rows()["rows"]["0"]["1"]

        cleared = self.operation("clear-blocker", None, 0)
        self.assertEqual(cleared.status_code, 200, cleared.body)
        spilled = self.rows()
        spill_input = spilled["rows"]["1"]["1"]
        detail = self.client.get(
            f"/api/copies/{self.copy['id']}/sheets/0/cell?r=1&c=1"
        ).json()
        with self.fixture.database.connect() as connection:
            result = connection.execute(
                "SELECT value_json FROM pp_results WHERE copy_id=? AND sheet=0 AND row=1 AND col=1",
                (self.copy["id"],),
            ).fetchone()
            override = connection.execute(
                "SELECT value_json,version,updated_by,updated_at FROM pp_overrides "
                "WHERE copy_id=? AND sheet=0 AND row=1 AND col=1",
                (self.copy["id"],),
            ).fetchone()
            calc_event = connection.execute(
                "SELECT payload_json FROM pp_events WHERE copy_id=? "
                "AND event_type='calc' AND input_revision=1",
                (self.copy["id"],),
            ).fetchone()

        undone = self.operation("undo-clear-blocker", 7, 1, "undo")
        self.assertEqual(undone.status_code, 200, undone.body)
        self.assertEqual(json.loads(result["value_json"]), 2)
        self.assertEqual(
            (json.loads(override["value_json"]), override["version"], override["updated_by"]),
            (None, 1, self.user["id"]),
        )
        self.assertTrue(override["updated_at"])
        self.assertEqual(
            (spill_input["v"], spill_input["k"], spill_input["ver"]),
            (2, "f", 1),
        )
        self.assertEqual(spill_input["by"], {
            "id": self.user["id"], "name": self.user["name"],
        })
        self.assertEqual(spill_input["at"], override["updated_at"])
        spill_event = next(
            item for item in json.loads(calc_event["payload_json"])["changes"]
            if (item["sheet"], item["r"], item["c"]) == (0, 1, 1)
        )
        self.assertEqual(
            {key: spill_event[key] for key in ("v", "s", "k", "ver", "by", "at")},
            {key: spill_input[key] for key in ("v", "s", "k", "ver", "by", "at")},
        )
        self.assertEqual((detail["v"], detail["k"]), (2, "f"))
        self.assertEqual(
            (spilled["input_revision"], spilled["calculated_revision"]), (1, 1)
        )

        restored = self.rows()
        restored_input = restored["rows"]["1"]["1"]
        self.assertEqual(restored["rows"]["0"]["1"], initial_anchor)
        self.assertEqual(
            (restored_input["v"], restored_input["k"], restored_input["ver"]),
            (7, "i", 2),
        )
        self.assertEqual(restored_input["by"], {
            "id": self.user["id"], "name": self.user["name"],
        })
        self.assertTrue(restored_input["at"])
        self.assertEqual(
            (restored["input_revision"], restored["calculated_revision"]), (2, 2)
        )
        with self.fixture.database.connect() as connection:
            current = connection.execute(
                "SELECT value_json,version,updated_by,updated_at FROM pp_overrides "
                "WHERE copy_id=? AND sheet=0 AND row=1 AND col=1",
                (self.copy["id"],),
            ).fetchone()
            spill = connection.execute(
                "SELECT 1 FROM pp_results WHERE copy_id=? AND sheet=0 AND row=1 AND col=1",
                (self.copy["id"],),
            ).fetchone()
        self.assertEqual(
            (json.loads(current["value_json"]), current["version"], current["updated_by"]),
            (7, 2, self.user["id"]),
        )
        self.assertEqual(current["updated_at"], restored_input["at"])
        self.assertIsNone(spill)

    def test_stale_undo_never_overwrites_a_concurrent_input(self):
        self.assertEqual(self.operation("clear-for-race", None, 0).status_code, 200)
        concurrent = self.operation("concurrent-input", 9, 1)
        stale = self.operation("stale-undo", 7, 1, "undo")
        current = self.rows()["rows"]["1"]["1"]

        self.assertEqual(concurrent.status_code, 200, concurrent.body)
        self.assertEqual((stale.status_code, stale.json()["error"]), (409, "conflict"))
        self.assertEqual((current["v"], current["k"], current["ver"]), (9, "i", 2))


class ImplicitInputSpillUndoTests(unittest.TestCase):
    def setUp(self):
        self.fixture = APIFixture(implicit_input_spill_fixture)
        self.client, self.csrf, _ = self.fixture.signup(
            "implicit-spill-undo@example.test", "Implicit Spill Owner"
        )
        _, self.copy = self.fixture.make_copy(self.client, self.csrf)

    def tearDown(self):
        self.fixture.close()

    def operation(self, op_id, row, col, value, version, action="edit"):
        return self.client.post(f"/api/copies/{self.copy['id']}/ops", {
            "op_id": op_id,
            "action": action,
            "cells": [{
                "sheet": 0, "r": row, "c": col,
                "v": value, "expected_version": version,
            }],
        }, {"x-csrf-token": self.csrf})

    def input_cell(self):
        return self.client.get(
            f"/api/copies/{self.copy['id']}/sheets/0/rows?start=1&count=1"
        ).json()["rows"]["1"]["1"]

    def test_implicit_original_input_restores_through_a_native_spill(self):
        self.assertTrue(self.fixture.snapshot.is_input(0, 1, 1))
        self.assertNotIn("0:1:1", self.fixture.snapshot.input_mask)
        self.assertEqual(self.input_cell(), {"v": None, "s": 1, "k": "i", "ver": 0})

        seeded = self.operation("seed-blocker", 1, 1, 7, 0)
        expanded = self.operation("expand-behind-blocker", 0, 0, 2, 0)
        cleared = self.operation("clear-implicit-blocker", 1, 1, None, 1)
        spilled = self.input_cell()
        restored = self.operation("undo-implicit-clear", 1, 1, 7, 2, "undo")
        final = self.input_cell()

        self.assertEqual(seeded.status_code, 200, seeded.body)
        self.assertEqual(expanded.status_code, 200, expanded.body)
        self.assertEqual(cleared.status_code, 200, cleared.body)
        self.assertEqual((spilled["v"], spilled["k"], spilled["ver"]), (2, "f", 2))
        self.assertEqual(restored.status_code, 200, restored.body)
        self.assertEqual((final["v"], final["k"], final["ver"]), (7, "i", 3))


if __name__ == "__main__":
    unittest.main()
