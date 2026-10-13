"""On a sheet without protection every non-formula cell is editable, as in Excel; formulas stay locked."""

import unittest

from tests.v2.api_support import APIFixture


class OpenSheetEditingTests(unittest.TestCase):
    def setUp(self):
        self.fixture = APIFixture()  # Inputs is protected; Lookup and Calc are not
        self.client, self.csrf, _ = self.fixture.signup("open@example.test", "Open Editor")
        _, copy = self.fixture.make_copy(self.client, self.csrf)
        self.copy_id = copy["id"]

    def tearDown(self):
        self.fixture.close()

    def edit(self, op_id, sheet, r, c, v):
        return self.client.post(f"/api/copies/{self.copy_id}/ops", {
            "op_id": op_id, "action": "edit",
            "cells": [{"sheet": sheet, "r": r, "c": c, "v": v, "expected_version": 0}],
        }, {"x-csrf-token": self.csrf})

    def rows(self, sheet, start, count):
        return self.client.get(
            f"/api/copies/{self.copy_id}/sheets/{sheet}/rows?start={start}&count={count}").json()

    def test_typed_value_on_an_unprotected_sheet_is_editable_and_recalculates(self):
        lookup = self.rows(1, 0, 3)["rows"]
        self.assertTrue(lookup["2"]["1"].get("open"))  # Lookup!B3, a typed value
        self.assertEqual(self.edit("open-value", 1, 2, 1, 100).status_code, 200)
        # Calc!B12 = SUMIFS(Lookup!B1:B10, Lookup!C1:C10, "A"); B3 sits in an "A" row (4.5 -> 100)
        calc = self.rows(2, 11, 1)["rows"]["11"]["1"]["v"]
        self.assertEqual(calc, 1.5 + 100 + 7.5 + 10.5 + 13.5)

    def test_blank_cell_on_an_unprotected_sheet_is_editable(self):
        # Calc!A5 is empty inside the sheet's used area
        self.assertEqual(self.edit("open-blank", 2, 4, 0, "note").status_code, 200)
        self.assertEqual(self.rows(2, 4, 1)["rows"]["4"]["0"]["v"], "note")

    def test_formulas_and_locked_cells_on_protected_sheets_stay_locked(self):
        formula = self.edit("formula", 2, 1, 1, 5)
        self.assertEqual((formula.status_code, formula.json()["error"]), (422, "locked"))
        self.assertNotIn("open", self.rows(2, 1, 1)["rows"]["1"]["1"])
        label = self.edit("label", 0, 1, 0, "x")  # Inputs!A2 is locked on a protected sheet
        self.assertEqual((label.status_code, label.json()["error"]), (422, "locked"))
        self.assertNotIn("open", self.rows(0, 1, 1)["rows"]["1"]["0"])


if __name__ == "__main__":
    unittest.main()
