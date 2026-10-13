"""Empty cells made editable by unlocked row or column styles, found without scanning every column."""

import json
import tempfile
import unittest
from pathlib import Path

from tests.v2.support import TMP


def snapshot(sheet):
    from pp.template.importer import IMPORT_VERSION
    from pp.workbook_service import TemplateSnapshot

    TMP.mkdir(parents=True, exist_ok=True)
    directory = tempfile.TemporaryDirectory(dir=TMP)
    root = Path(directory.name)
    (root / "sheets" / "0").mkdir(parents=True)
    (root / "sheets" / "0" / "0.json").write_text(json.dumps({
        "1": {"3": {"v": 5, "s": 1, "k": "c"}, "4": {"v": 6, "s": 1, "k": "i"}},
        "2": {"7": {"v": "x", "s": 0, "k": "c"}},
    }))
    (root / "meta.json").write_text(json.dumps({
        "import_version": IMPORT_VERSION, "hash": "h", "name": "t", "sheets": [sheet]}))
    (root / "styles.json").write_text(json.dumps([{"locked": True}, {"locked": False}]))
    (root / "input_mask.json").write_text("[]")
    return directory, TemplateSnapshot(root)


def brute_force(snap, row):
    return {str(col): record for col in range(snap.sheets[0]["cols"])
            if (record := snap.cell(0, row, col))["k"] == "i"}


class ImplicitInputTests(unittest.TestCase):
    def check(self, sheet, row):
        directory, snap = snapshot(sheet)
        self.addCleanup(directory.cleanup)
        return snap.implicit_cells(0, row), brute_force(snap, row)

    def test_column_spans_match_the_cell_rules_with_first_range_winning(self):
        sheet = {"rows": 10, "cols": 60, "protected": True, "row_styles": {},
                 "col_styles": [[2, 5, 1], [4, 9, 0], [12, 39, 1]]}
        for row in (0, 1, 2):
            fast, slow = self.check(sheet, row)
            self.assertEqual(fast, slow)
        self.assertEqual(sorted(fast, key=int), ["2", "3", "4", "5", *map(str, range(12, 40))])

    def test_unprotected_sheet_keeps_only_stored_inputs(self):
        sheet = {"rows": 10, "cols": 40, "protected": False, "row_styles": {"1": 1}, "col_styles": [[0, 39, 1]]}
        fast, slow = self.check(sheet, 1)
        self.assertEqual(fast, slow)
        self.assertEqual(list(fast), ["4"])

    def test_unlocked_row_style_stops_at_the_used_width(self):
        sheet = {"rows": 10, "cols": 16384, "protected": True, "row_styles": {"1": 1}, "col_styles": []}
        fast, slow = self.check(sheet, 1)
        self.assertEqual(len(slow), 16383)  # every empty column of the row, plus the stored input
        # widest stored cell is column 7; column 3 holds a locked constant
        self.assertEqual(sorted(fast, key=int), ["0", "1", "2", "4", "5", "6", "7"])
        self.assertTrue(all(fast[col] == slow[col] for col in fast))

    def test_column_style_to_the_sheet_edge_stops_at_the_used_width(self):
        sheet = {"rows": 10, "cols": 16384, "protected": True, "row_styles": {},
                 "col_styles": [[0, 1, 0], [2, 16383, 1]]}
        fast, slow = self.check(sheet, 0)
        self.assertEqual(len(slow), 16382)
        self.assertEqual(sorted(fast, key=int), [str(col) for col in range(2, 8)])


if __name__ == "__main__":
    unittest.main()
