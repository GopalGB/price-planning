import json
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Protection

from pp.export.xlsm import export_xlsm
from pp.template.importer import import_template
from tests.v2.test_template_importer import rewrite_member

ROOT = Path(__file__).resolve().parents[2]
TMP = ROOT / "results-private" / "tmp"
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def array_fixture(path, ref="$B$1:$B$1", formula_type="array"):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Calc"
    sheet.protection.sheet = True
    sheet["A1"] = 3
    sheet["A1"].protection = Protection(locked=False)
    sheet["B1"] = "=A1*2"
    sheet["B1"].protection = Protection(locked=False)
    sheet["C1"] = "=ISFORMULA(B1)"
    sheet["D1"] = '="array-text"'
    workbook.save(path)

    def rewrite(data):
        root = ET.fromstring(data)
        formula = root.find(f".//{NS}c[@r='B1']/{NS}f")
        formula.attrib.update({"t": formula_type})
        if ref is not None:
            formula.set("ref", ref)
        return ET.tostring(root, encoding="utf-8", xml_declaration=True)

    rewrite_member(path, "xl/worksheets/sheet1.xml", rewrite)


class SingleCellArrayTests(unittest.TestCase):
    def setUp(self):
        TMP.mkdir(parents=True, exist_ok=True)

    def test_single_cell_array_loads_and_recalculates_after_input_edit(self):
        from pp.engine.formualizer import FormualizerAdapter

        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            for ref in ("$B$1:$B$1", "B1", "B1:B1"):
                with self.subTest(ref=ref):
                    source = Path(temp) / f"array-{ref.replace(':', '-')}.xlsm"
                    array_fixture(source, ref=ref)
                    engine = FormualizerAdapter.load(source)

                    self.assertEqual(engine.get(0, 0, 1), 6)
                    changes = engine.apply([(0, 0, 0, 5)])

                    self.assertEqual(engine.get(0, 0, 1), 10)
                    self.assertIn((0, 0, 1, 10), changes)

    def test_array_anchor_is_formula_for_native_and_isformula_views(self):
        from pp.engine.formualizer import FormualizerAdapter

        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            for ref in ("$B$1:$B$1", "B1", "B1:B1"):
                with self.subTest(ref=ref):
                    source = Path(temp) / f"array-isformula-{ref.replace(':', '-')}.xlsm"
                    array_fixture(source, ref=ref)
                    engine = FormualizerAdapter.load(source)

                    self.assertTrue(engine.is_formula(0, 0, 1))
                    self.assertIs(engine.get(0, 0, 2), True)

    def test_protected_import_keeps_array_locked_and_ordinary_input_editable(self):
        from pp.workbook_service import TemplateSnapshot

        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            source = Path(temp) / "array-import.xlsm"
            destination = Path(temp) / "snapshot"
            array_fixture(source)
            snapshot = import_template(source, destination)

            mask = json.loads((snapshot / "input_mask.json").read_text())
            imported = TemplateSnapshot(snapshot)
            self.assertIn("0:0:0", mask)
            self.assertNotIn("0:0:1", mask)
            self.assertTrue(imported.is_input(0, 0, 0))
            self.assertFalse(imported.is_input(0, 0, 1))

    def test_export_preserves_array_ref_and_formula_text_after_input_edit(self):
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            for ref in ("$B$1:$B$1", "B1", "B1:B1"):
                with self.subTest(ref=ref):
                    source = Path(temp) / f"array-export-{ref.replace(':', '-')}.xlsm"
                    output = Path(temp) / f"exported-{ref.replace(':', '-')}.xlsm"
                    array_fixture(source, ref=ref)
                    export_xlsm(source, [(0, 0, 0, 5)], output)

                    with zipfile.ZipFile(output) as archive:
                        root = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
                    formula = root.find(f".//{NS}c[@r='B1']/{NS}f")
                    self.assertEqual(formula.get("t"), "array")
                    self.assertEqual(formula.get("ref"), ref)
                    self.assertEqual(formula.text, "A1*2")
                    self.assertEqual(root.findtext(f".//{NS}c[@r='A1']/{NS}v"), "5")
                    self.assertEqual(root.find(f".//{NS}c[@r='D1']/{NS}f").text,
                                     '"array-text"')

    def test_array_rejects_missing_malformed_foreign_and_multi_cell_refs(self):
        from pp.engine.formualizer import EngineError, FormualizerAdapter

        cases = [(None, "missing"), ("bad", "malformed"), ("A1:A1", "foreign"),
                 ("B1:C1", "multi-cell")]
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            for ref, label in cases:
                with self.subTest(label=label):
                    source = Path(temp) / f"{label}.xlsm"
                    array_fixture(source, ref=ref)
                    with self.assertRaisesRegex(EngineError, "array"):
                        FormualizerAdapter.load(source)

    def test_data_table_is_always_rejected(self):
        from pp.engine.formualizer import EngineError, FormualizerAdapter

        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            source = Path(temp) / "data-table.xlsm"
            array_fixture(source, formula_type="dataTable")
            with self.assertRaisesRegex(EngineError, "dataTable"):
                FormualizerAdapter.load(source)


if __name__ == "__main__":
    unittest.main()
