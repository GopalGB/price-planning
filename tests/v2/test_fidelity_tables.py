"""Behavioral table and AutoFilter fidelity coverage."""

import json
import tempfile
import unittest
from pathlib import Path

from ops.fixtures import make_tables
from pp.template.importer import import_template
from tests.v2.api_support import APIFixture


class TableImporterTests(unittest.TestCase):
    def test_imports_two_tables_sheet_filter_flags_and_exact_dxfs(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "tables.xlsx"
            make_tables(source)
            snapshot = import_template(source, Path(directory) / "data")
            sheet = json.loads((snapshot / "meta.json").read_text())["sheets"][0]

        self.assertEqual(len(sheet["tables"]), 2)
        built_in, custom = sheet["tables"]
        self.assertEqual(built_in["range"], [0, 0, 259, 2])
        self.assertEqual(built_in["columns"], ["Region", "Price", "Status"])
        self.assertEqual(built_in["style_name"], "TableStyleMedium2")
        self.assertEqual(
            (built_in["header_row"], built_in["show_first_column"],
             built_in["show_row_stripes"], built_in["show_column_stripes"]),
            (True, True, True, False),
        )
        self.assertEqual(built_in["style"]["headerRow"]["style"], {
            "bold": True, "color": "#FFFFFF", "fill": "#4F81BD",
        })
        self.assertEqual(
            built_in["style"]["firstRowStripe"]["style"]["fill"],
            "#DCE6F2",
        )
        self.assertEqual(custom["style_name"], "CustomPlanning")
        self.assertEqual(custom["style"]["wholeTable"]["style"], {
            "color": "#333333",
            "nf": "$0.00",
            "border": {
                "t": {"s": 1, "cl": {"rgb": "#112233"}},
                "r": {"s": 1, "cl": {"rgb": "#112233"}},
                "b": {"s": 1, "cl": {"rgb": "#112233"}},
                "l": {"s": 1, "cl": {"rgb": "#112233"}},
                "vertical": {"s": 3, "cl": {"rgb": "#445566"}},
                "horizontal": {"s": 4, "cl": {"rgb": "#778899"}},
            },
        })
        self.assertEqual(
            [(item["source"], item["range"]) for item in sheet["filters"]],
            [("sheet", [0, 0, 259, 3]),
             ("table", [0, 0, 259, 2]),
             ("table", [0, 4, 3, 6])],
        )


class TableAPIContractTests(unittest.TestCase):
    def setUp(self):
        self.fixture = APIFixture(make_tables)
        self.owner, self.csrf, _ = self.fixture.signup("table-owner@example.test")
        _, self.copy = self.fixture.make_copy(self.owner, self.csrf)

    def tearDown(self):
        self.fixture.close()

    def rows(self, start, count):
        return self.owner.get(
            f"/api/copies/{self.copy['id']}/sheets/0/rows?start={start}&count={count}"
        )

    def test_rows_layer_table_styles_under_cf_with_exact_edges_and_sparse_blocks(self):
        workbook = self.owner.get(f"/api/copies/{self.copy['id']}/workbook").json()
        sheet = workbook["sheets"][0]
        self.assertEqual(len(sheet["tables"]), 2)
        self.assertEqual(len(sheet["filters"]), 3)
        self.assertNotIn("style", sheet["tables"][0])

        first = self.rows(0, 4)
        crossing = self.rows(255, 5)
        self.assertEqual((first.status_code, crossing.status_code), (200, 200))
        rows = first.json()["rows"]
        self.assertEqual(rows["0"]["0"]["style_override"], {
            "bold": True,
            "color": "#FFFFFF",
            "fill": "#4F81BD",
            "border": {
                "t": {"s": 1, "cl": {"rgb": "#95B3D7"}},
                "b": {"s": 1, "cl": {"rgb": "#95B3D7"}},
                "l": {"s": 1, "cl": {"rgb": "#95B3D7"}},
            },
        })
        self.assertEqual(rows["1"]["0"]["style_override"]["fill"], "#DCE6F2")
        self.assertTrue(rows["1"]["0"]["style_override"]["bold"])
        custom = rows["1"]["5"]
        self.assertEqual(custom["k"], "i")
        self.assertEqual(custom["style_override"], {
            "color": "#333333",
            "fill": "#FF0000",
            "nf": "$0.00",
            "border": {
                "r": {"s": 3, "cl": {"rgb": "#445566"}},
                "b": {"s": 4, "cl": {"rgb": "#778899"}},
            },
        })
        self.assertEqual(rows["3"]["6"]["style_override"]["border"], {
            "r": {"s": 1, "cl": {"rgb": "#112233"}},
            "b": {"s": 1, "cl": {"rgb": "#112233"}},
        })
        crossed = crossing.json()["rows"]
        self.assertEqual(set(crossed), {"255", "256", "257", "258", "259"})
        self.assertIsNone(crossed["257"]["0"]["v"])
        self.assertIn("style_override", crossed["257"]["0"])


if __name__ == "__main__":
    unittest.main()


class TableStyleBudgetTests(unittest.TestCase):
    def test_a_table_declared_out_to_the_last_column_is_bounded_per_request(self):
        import time
        from types import SimpleNamespace

        from pp.fidelity.tables import MAX_TABLE_STYLE_CELLS, table_style_overrides
        table = {"range": [0, 0, 399, 16_383], "header_row": True, "totals_row": False,
                 "show_row_stripes": False, "show_column_stripes": False,
                 "show_first_column": False, "show_last_column": False,
                 "style": {"wholeTable": {"style": {"bold": True}}}}
        started = time.monotonic()
        styles = table_style_overrides({"tables": [table]}, SimpleNamespace(start=0, count=200))
        self.assertLessEqual(len(styles), MAX_TABLE_STYLE_CELLS)
        self.assertEqual(styles[(199, 0)], {"bold": True})
        self.assertLess(time.monotonic() - started, 2)
