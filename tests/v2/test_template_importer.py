import json
import shutil
from pathlib import Path
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName

from ops.fixtures import make_semantic


ROOT = Path(__file__).resolve().parents[2]
TMP = ROOT / "results-private" / "tmp"


def rewrite_member(source, member, transform):
    replacement = source.with_name("rewritten.xlsm")
    with zipfile.ZipFile(source) as source_zip, zipfile.ZipFile(replacement, "w") as target_zip:
        for info in source_zip.infolist():
            data = source_zip.read(info.filename)
            if info.filename == member:
                data = transform(data)
            target_zip.writestr(info, data)
    shutil.move(replacement, source)


def shared_formula_fixture(path):
    workbook = Workbook()
    values = workbook.active
    values.title = "Values"
    values["A1"] = 3
    values["B1"] = 4
    formulas = workbook.create_sheet("Formulas")
    formulas["B2"] = "=Values!A1*2"
    formulas["C2"] = "=Values!B1*2"
    workbook.save(path)

    def shared_formula_xml(data):
        namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        qualified = lambda name: f"{{{namespace}}}{name}"
        root = ET.fromstring(data)
        row = root.find(f".//{qualified('row')}[@r='2']")
        cells = {cell.get("r"): cell for cell in row.findall(qualified("c"))}
        master = cells["B2"]
        follower = cells["C2"]
        master_formula = master.find(qualified("f"))
        follower_formula = follower.find(qualified("f"))
        master_formula.attrib.update({"t": "shared", "ref": "B2:C2", "si": "7"})
        follower_formula.attrib.clear()
        follower_formula.attrib.update({"t": "shared", "si": "7"})
        follower_formula.text = None
        row.remove(follower)
        row.insert(list(row).index(master), follower)
        return ET.tostring(root, encoding="utf-8", xml_declaration=True)

    rewrite_member(path, "xl/worksheets/sheet2.xml", shared_formula_xml)


def metadata_fixture(path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Visible"
    sheet.protection.sheet = True
    sheet.freeze_panes = "C3"
    sheet.column_dimensions["B"].width = 18
    sheet.column_dimensions["D"].hidden = True
    sheet.row_dimensions[4].height = 24
    sheet.row_dimensions[6].hidden = True
    sheet.merge_cells("A8:B9")
    sheet["B2"] = 10
    sheet["B2"].protection = Protection(locked=False)
    sheet["C2"] = "=B2*2"
    sheet["C2"].protection = Protection(locked=False)
    sheet["B2"].number_format = "0.00%"
    sheet["B2"].font = Font(bold=True, italic=True, color="FF112233", size=14)
    sheet["B2"].fill = PatternFill("solid", fgColor="FF445566")
    edge = Side(style="thin", color="FF000000")
    sheet["B2"].border = Border(top=edge, right=edge, bottom=edge, left=edge)
    sheet["B2"].alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
    validation = DataValidation(type="list", formula1='"A,B,C"')
    validation.add("B2:B4")
    sheet.add_data_validation(validation)
    hidden = workbook.create_sheet("Hidden")
    hidden.sheet_state = "veryHidden"
    workbook.defined_names.add(DefinedName("PlanInput", attr_text="'Visible'!$B$2"))
    sheet.conditional_formatting.add("B2", CellIsRule(operator="greaterThan", formula=["0"]))
    workbook.save(path)


class TemplateImporterTests(unittest.TestCase):
    def test_import_writes_atomic_snapshot_with_input_mask_and_formula_map(self):
        """Removing protection/formula parsing must make this contract fail."""
        from pp.template.importer import import_template

        TMP.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            workspace = Path(temp)
            source = workspace / "semantic.xlsm"
            make_semantic(source)

            snapshot = import_template(source, workspace / "data")

            meta = json.loads((snapshot / "meta.json").read_text(encoding="utf-8"))
            mask = json.loads((snapshot / "input_mask.json").read_text(encoding="utf-8"))
            blocks = json.loads((snapshot / "sheets" / "0" / "0.json").read_text(encoding="utf-8"))
            self.assertEqual([sheet["name"] for sheet in meta["sheets"]], ["Inputs", "Lookup", "Calc"])
            self.assertIn("0:1:1", mask)
            self.assertNotIn("2:1:1", mask)
            self.assertEqual(blocks["1"]["1"]["k"], "i")
            self.assertEqual(blocks["1"]["1"]["v"], 10)

    def test_failed_import_does_not_publish_partial_snapshot(self):
        """Publishing staging data after a malformed archive would be corruption."""
        from pp.template.importer import TemplateImportError, import_template

        TMP.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            workspace = Path(temp)
            source = workspace / "broken.xlsm"
            source.write_bytes(b"not a zip")

            with self.assertRaises(TemplateImportError):
                import_template(source, workspace / "data")

            templates = workspace / "data" / "templates"
            self.assertEqual(list(templates.iterdir()), [])

    def test_shared_formula_follower_before_master_is_expanded(self):
        """Resolving followers only when their XML is visited must fail."""
        from pp.template.importer import import_template

        TMP.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            workspace = Path(temp)
            source = workspace / "shared.xlsm"
            shared_formula_fixture(source)

            snapshot = import_template(source, workspace / "data")
            block = json.loads((snapshot / "sheets" / "1" / "0.json").read_text())

            self.assertEqual(block["1"]["1"]["formula"], "=Values!A1*2")
            self.assertEqual(block["1"]["2"]["formula"], "=Values!B1*2")

    def test_metadata_styles_protection_and_progress_match_api_contract(self):
        """Dropping render metadata or treating unlocked formulas as inputs must fail."""
        from pp.template.importer import import_template

        TMP.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            workspace = Path(temp)
            source = workspace / "metadata.xlsm"
            metadata_fixture(source)

            snapshot = import_template(source, workspace / "data")
            meta = json.loads((snapshot / "meta.json").read_text())
            styles = json.loads((snapshot / "styles.json").read_text())
            mask = json.loads((snapshot / "input_mask.json").read_text())
            progress = json.loads((snapshot / "progress.json").read_text())
            visible = meta["sheets"][0]
            hidden = meta["sheets"][1]
            block = json.loads((snapshot / "sheets" / "0" / "0.json").read_text())
            style = styles[block["1"]["1"]["s"]]

            self.assertEqual((visible["rows"], visible["cols"]), (9, 4))
            self.assertEqual(visible["freeze"], {"rows": 2, "cols": 2})
            self.assertGreater(visible["col_widths"]["1"], 0)
            self.assertGreater(visible["row_heights"]["3"], 0)
            self.assertEqual(visible["hidden_rows"], [5])
            self.assertEqual(visible["hidden_cols"], [3])
            self.assertEqual(visible["merges"], [[7, 0, 8, 1]])
            self.assertEqual(visible["validations"][0]["list"], ["A", "B", "C"])
            self.assertEqual(visible["validations"][0]["ranges"], [[1, 1, 3, 1]])
            self.assertEqual(hidden["hidden"], "veryHidden")
            self.assertEqual(meta["defined_names"][0]["name"], "PlanInput")
            self.assertEqual(progress, {"complete": True, "sheets_completed": 2, "sheets_total": 2})
            self.assertIn("0:1:1", mask)
            self.assertNotIn("0:1:2", mask)
            self.assertEqual(block["1"]["2"]["k"], "f")
            self.assertEqual(style["nf"], "0.00%")
            self.assertTrue(style["bold"])
            self.assertTrue(style["italic"])
            self.assertEqual(style["align"], "center")
            self.assertEqual(style["valign"], "top")
            self.assertTrue(style["wrap"])
            self.assertEqual(style["border"], {"t": "thin", "r": "thin", "b": "thin", "l": "thin"})

    def test_sheet_blocks_are_split_at_256_rows(self):
        """Keeping a single sheet dictionary instead of publishing row blocks must fail."""
        from pp.template.importer import import_template

        TMP.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            workspace = Path(temp)
            source = workspace / "blocks.xlsm"
            workbook = Workbook(write_only=True)
            sheet = workbook.create_sheet("Rows")
            for row in range(520):
                sheet.append([row])
            workbook.save(source)

            snapshot = import_template(source, workspace / "data")

            files = sorted(path.name for path in (snapshot / "sheets" / "0").glob("*.json"))
            self.assertEqual(files, ["0.json", "1.json", "2.json"])


if __name__ == "__main__":
    unittest.main()
