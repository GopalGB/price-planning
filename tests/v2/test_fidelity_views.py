"""Behavioral sheet-view and final synthetic-fixture coverage."""

import json
import tempfile
import unittest
from pathlib import Path

from ops.fixtures import make_fidelity
from pp.template.importer import import_template
from tests.v2.api_support import APIFixture


class FidelityFixtureImportTests(unittest.TestCase):
    def test_fixture_imports_all_features_and_literal_sheet_views(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "fidelity.xlsx"
            make_fidelity(source)
            snapshot = import_template(source, Path(directory) / "data")
            meta = json.loads((snapshot / "meta.json").read_text())

        sheets = {sheet["name"]: sheet for sheet in meta["sheets"]}
        self.assertEqual(list(sheets), [
            "CF Rules", "Drawings", "Notes One", "Notes Two", "Tables",
            "View Demo", "Hidden Demo", "Very Hidden Demo",
        ])
        self.assertEqual(
            {rule["type"] for rule in sheets["CF Rules"]["conditional_formats"]},
            {"expression", "cellIs", "containsText", "duplicateValues",
             "beginsWith", "containsErrors"},
        )
        self.assertEqual(len(sheets["Drawings"]["drawings"]), 4)
        self.assertTrue(sheets["Notes One"]["has_notes"])
        self.assertTrue(sheets["Notes Two"]["has_notes"])
        self.assertEqual(len(sheets["Tables"]["tables"]), 2)
        self.assertEqual(len(sheets["Tables"]["filters"]), 3)

        view = sheets["View Demo"]
        self.assertEqual(view["tab_color"], "#107C41")
        self.assertFalse(view["show_gridlines"])
        self.assertEqual(view["zoom"], 1.25)
        self.assertEqual(view["scroll"], {"row": 5, "col": 3})
        self.assertEqual(view["freeze"], {
            "rows": 2, "cols": 1, "start_row": 5, "start_col": 3,
        })
        self.assertEqual(view["col_widths"]["1"], 18.0)
        self.assertEqual(view["row_heights"]["1"], 30.0)
        self.assertEqual(view["hidden_rows"], [4])
        self.assertEqual(view["hidden_cols"], [3])
        self.assertEqual(view["merges"], [[5, 1, 6, 2]])
        self.assertTrue(view["protected"])
        self.assertTrue(view["active"])
        self.assertEqual(sheets["Hidden Demo"]["hidden"], "hidden")
        self.assertEqual(sheets["Very Hidden Demo"]["hidden"], "veryHidden")


class FidelityConditionalFormatRowsTest(unittest.TestCase):
    def setUp(self):
        self.fixture = APIFixture(make_fidelity)
        self.owner, self.csrf, _ = self.fixture.signup("cf-rows@example.test")
        _, self.copy = self.fixture.make_copy(self.owner, self.csrf)

    def tearDown(self):
        self.fixture.close()

    def test_six_rule_types_render_literal_styles_on_the_fixture(self):
        response = self.owner.get(
            f"/api/copies/{self.copy['id']}/sheets/0/rows?start=0&count=3"
        )
        self.assertEqual(response.status_code, 200)
        rows = response.json()["rows"]
        styles = {
            (row, col): rows.get(str(row), {}).get(str(col), {}).get("style_override")
            for row, col in (
                (1, 0), (2, 0), (1, 1), (2, 1), (1, 2), (2, 2),
                (1, 3), (2, 3), (1, 4), (2, 4), (1, 5),
            )
        }
        self.assertEqual(styles, {
            (1, 0): {"fill": "#FFC7CE"},
            (2, 0): None,
            (1, 1): {"fill": "#D9EAD3"},
            (2, 1): None,
            (1, 2): {"fill": "#9C6500"},
            (2, 2): None,
            (1, 3): {"fill": "#EBF1DE"},
            (2, 3): {"fill": "#EBF1DE"},
            (1, 4): {"fill": "#CCFFFF"},
            (2, 4): None,
            (1, 5): {"bold": True, "color": "#9C0006"},
        })


class SheetViewAPIContractTests(unittest.TestCase):
    def setUp(self):
        self.fixture = APIFixture(make_fidelity)
        self.owner, self.csrf, _ = self.fixture.signup("views-owner@example.test")
        _, self.copy = self.fixture.make_copy(self.owner, self.csrf)

    def tearDown(self):
        self.fixture.close()

    def test_workbook_projects_view_geometry_and_input_only_protection(self):
        workbook = self.owner.get(
            f"/api/copies/{self.copy['id']}/workbook"
        ).json()
        sheets = {sheet["name"]: sheet for sheet in workbook["sheets"]}
        view = sheets["View Demo"]
        self.assertEqual(view["tab_color"], "#107C41")
        self.assertFalse(view["show_gridlines"])
        self.assertEqual(view["zoom"], 1.25)
        self.assertEqual(view["scroll"], {"row": 5, "col": 3})
        self.assertEqual(view["freeze"], {
            "rows": 2, "cols": 1, "start_row": 5, "start_col": 3,
        })
        self.assertEqual(view["col_widths"]["1"], 131)
        self.assertEqual(view["row_heights"]["1"], 40)
        self.assertEqual(view["merges"], [[5, 1, 6, 2]])
        self.assertTrue(view["active"])
        self.assertNotIn("tab_color", sheets["CF Rules"])

        sheet_index = view["index"]
        edit = self.owner.post(
            f"/api/copies/{self.copy['id']}/ops",
            {"op_id": "view-edit", "action": "edit",
             "cells": [{"sheet": sheet_index, "r": 1, "c": 1, "v": 24,
                        "expected_version": 0}]},
            {"x-csrf-token": self.csrf},
        )
        self.assertEqual(edit.status_code, 200, edit.body)
        blocked = self.owner.post(
            f"/api/copies/{self.copy['id']}/ops",
            {"op_id": "view-edit-2", "action": "edit",
             "cells": [{"sheet": sheet_index, "r": 1, "c": 0, "v": 99,
                        "expected_version": 0}]},
            {"x-csrf-token": self.csrf},
        )
        self.assertEqual(blocked.status_code, 422, blocked.body)

        response = self.owner.get(
            f"/api/copies/{self.copy['id']}/sheets/{sheet_index}/rows?start=0&count=4"
        )
        self.assertEqual(response.status_code, 200)
        rows = response.json()["rows"]
        self.assertEqual(rows["1"]["1"]["k"], "i")
        self.assertEqual(rows["1"]["1"]["v"], 24)
        self.assertEqual(rows["1"]["2"]["k"], "f")
        self.assertNotIn("0", rows["1"])


if __name__ == "__main__":
    unittest.main()
