from pathlib import Path
import tempfile
import unittest

from openpyxl import Workbook

from ops.fixtures import make_semantic


ROOT = Path(__file__).resolve().parents[2]
TMP = ROOT / "results-private" / "tmp"


class FormualizerAdapterTests(unittest.TestCase):
    def test_apply_recalculates_semantic_dependent_with_real_engine(self):
        """Replacing the engine call with cached fixture values must fail."""
        from pp.engine.formualizer import FormualizerAdapter

        TMP.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            source = Path(temp) / "semantic.xlsm"
            make_semantic(source)
            engine = FormualizerAdapter.load(source)

            changes = engine.apply([(0, 1, 1, 12)])

            self.assertEqual(engine.get(2, 1, 1), 24)
            self.assertIn((2, 1, 1, 24), changes)
            self.assertEqual(engine.get(2, 1, 1), 24)

    def test_template_formula_map_answers_isformula_without_engine_semantics(self):
        """Dropping formula metadata must make ISFORMULA compatibility wrong."""
        from pp.engine.formualizer import FormualizerAdapter

        TMP.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            source = Path(temp) / "semantic.xlsm"
            make_semantic(source)
            engine = FormualizerAdapter.load(source)

            self.assertTrue(engine.is_formula(2, 1, 1))
            self.assertFalse(engine.is_formula(0, 1, 1))

    def test_formula_map_and_dependency_interface_are_not_dimension_capped(self):
        """Scanning only a fixed 256 by 64 window must miss this formula."""
        from pp.engine.formualizer import FormualizerAdapter

        TMP.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            source = Path(temp) / "wide.xlsm"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Wide"
            sheet.cell(300, 70, "=A1+1")
            workbook.save(source)

            engine = FormualizerAdapter.load(source)

            self.assertTrue(engine.is_formula(0, 299, 69))
            self.assertIn("'Wide'!BR300", engine.dependents(0, 0, 0))

    def test_required_compatibility_functions_have_excel_semantics(self):
        """Trusting the engine's broken ISFORMULA result must fail."""
        from pp.engine.formualizer import FormualizerAdapter

        TMP.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            source = Path(temp) / "compat.xlsm"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Compat"
            sheet["A1"] = 10
            sheet["A2"] = 20
            sheet["B1"] = "=A1*2"
            sheet["C1"] = "=ISFORMULA(B1)"
            sheet["C2"] = '=CELL("address",A1)'
            sheet["C3"] = '=INDIRECT("A2")'
            sheet["C4"] = "=OFFSET(A1,1,0)"
            sheet["C5"] = '=SWITCH(A1,10,"ten","other")'
            workbook.save(source)

            engine = FormualizerAdapter.load(source)

            self.assertIs(engine.get(0, 0, 2), True)
            self.assertEqual(engine.get(0, 1, 2), "$A$1")
            self.assertEqual(engine.get(0, 2, 2), 20)
            self.assertEqual(engine.get(0, 3, 2), 20)
            self.assertEqual(engine.get(0, 4, 2), "ten")


if __name__ == "__main__":
    unittest.main()
