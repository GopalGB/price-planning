"""The engine's edit path on the Python calculator: publish what changed, revert, typed formulas."""

import tempfile
import unittest
from functools import partial
from pathlib import Path

from openpyxl import Workbook

from tests.v2.support import TMP
from tests.v2.test_engine_cached import _replace_formula_values


def workbook(directory):
    path = Path(directory) / "plan.xlsx"
    book = Workbook()
    calc = book.active
    calc.title = "Calc"
    data = book.create_sheet("Data Src")
    calc["A1"] = 2
    calc["B1"] = "=A1*2"
    calc["C1"] = "=SUM(B:B)"
    calc["D1"] = "='Data Src'!B2*10"
    calc["E1"] = "=7"
    data["B2"] = 3
    book.save(path)
    _replace_formula_values(path, {"xl/worksheets/sheet1.xml": {
        "B1": ("4", None), "C1": ("4", None), "D1": ("30", None), "E1": ("7", None)}})
    return path


class CalcHostTests(unittest.TestCase):
    def load(self):
        from pp.calc.host import CalcHost
        from pp.engine.excel import ExcelEngine

        TMP.mkdir(parents=True, exist_ok=True)
        directory = tempfile.TemporaryDirectory(dir=TMP)
        self.addCleanup(directory.cleanup)
        cache = Path(directory.name) / "cache"
        engine = ExcelEngine.load(workbook(directory.name), cache_dir=cache,
                                  host_factory=partial(CalcHost, cache_dir=cache))
        self.addCleanup(engine.close)
        return engine

    def test_edit_publishes_reachable_changes_and_reverting_removes_them(self):
        engine = self.load()
        self.assertEqual(engine.get(0, 0, 1), 4)
        self.assertEqual(engine.apply([(0, 0, 0, 5)]), [(0, 0, 1, 10), (0, 0, 2, 10)])
        engine.apply([(0, 0, 0, 2)])
        self.assertEqual(engine.overlay_snapshot(), [])

    def test_cross_sheet_edit_and_clear(self):
        engine = self.load()
        self.assertEqual(engine.apply([(1, 1, 1, 4)]), [(0, 0, 3, 40)])
        self.assertEqual(engine.apply([(0, 0, 0, None)]), [(0, 0, 1, 0), (0, 0, 2, 0)])

    def test_formula_typed_into_a_copy_calculates_and_feeds_its_readers(self):
        engine = self.load()
        changes = engine.apply([(0, 0, 0, {"formula": "='Data Src'!B2+1"})])
        self.assertEqual(changes, [(0, 0, 0, 4), (0, 0, 1, 8), (0, 0, 2, 8)])
        self.assertEqual(engine.apply([(1, 1, 1, 9)]), [(0, 0, 0, 10), (0, 0, 1, 20), (0, 0, 2, 20), (0, 0, 3, 90)])

    def test_typed_text_stays_text_and_errors_come_back_as_errors(self):
        engine = self.load()
        self.assertEqual(engine.apply([(0, 0, 0, "abc")]), [(0, 0, 1, {"error": "#VALUE!"}), (0, 0, 2, {"error": "#VALUE!"})])


if __name__ == "__main__":
    unittest.main()
