"""Native template load must not depend on per-cell formula registration.

Formualizer 0.11.1 expands a shared-formula follower by translating the master
text with the follower offset. That translation also shifts a cross-sheet
qualifier whose sheet name parses as a cell address, so ``S01!$A$1`` becomes
``T1!$A$1`` and the follower reads a sheet that does not exist. The template is
therefore handed to the engine with shared followers already expanded.
"""

import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

import formualizer

from ops.engine_scale_fixture import expected_formula_values, generate
from tests.v2.support import TMP

SEED = 20261011
SCALE = 0.0001
MASTER = 'SUMIF(S01!$D$1:$D$3,">=2",S01!$E$1:$E$3)+S01!$A$1'


def shared_workbook(path, master=MASTER):
    """Write a workbook whose shared family is a master plus two followers."""
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package'
        '.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats'
        '-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd'
        '.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006'
        '/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="S01" sheetId="1" r:id="rId1"/></sheets></workbook>')
    workbook_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006'
        '/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
    sheet = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
        '<row r="1"><c r="A1" t="n"><v>7</v></c><c r="B1" t="n"><v>3</v></c>'
        '<c r="D1" t="n"><v>1</v></c><c r="E1" t="n"><v>100</v></c></row>'
        f'<row r="2"><c r="A2" t="n"><f t="shared" ref="A2:C2" si="1">{_escape(master)}</f>'
        '<v>507</v></c>'
        '<c r="B2" t="n"><f t="shared" si="1"></f><v>507</v></c>'
        '<c r="C2" t="n"><f t="shared" si="1"/><v>507</v></c></row>'
        '</sheetData></worksheet>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", root_rels)
        archive.writestr("xl/workbook.xml", workbook)
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        archive.writestr("xl/worksheets/sheet1.xml", sheet)
    return path


def _escape(text):
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def library_shifts_the_sheet_qualifier(path):
    """Report whether the pinned engine still corrupts the cross-sheet prefix."""
    workbook = formualizer.Workbook.from_path(str(path))
    return workbook.get_formula("S01", 2, 2) != MASTER


class NativeSharedFormulaTests(unittest.TestCase):
    def test_shared_family_crosses_a_cell_address_sheet_name(self):
        with tempfile.TemporaryDirectory(dir=_tmp()) as temp:
            source = shared_workbook(Path(temp) / "shared.xlsx")

            self.assertTrue(library_shifts_the_sheet_qualifier(source))

    def test_engine_qualifier_defect_makes_every_follower_read_a_missing_sheet(self):
        with tempfile.TemporaryDirectory(dir=_tmp()) as temp:
            source = shared_workbook(Path(temp) / "shared.xlsx")
            workbook = formualizer.Workbook.from_path(str(source))

            value = workbook.evaluate_cell("S01", 2, 2)

            self.assertEqual(value, {"type": "Error", "kind": "Ref"})

    def test_prepared_template_gives_the_engine_explicit_follower_formulas(self):
        from pp.engine.native_load import prepared_template

        with tempfile.TemporaryDirectory(dir=_tmp()) as temp:
            source = shared_workbook(Path(temp) / "shared.xlsx")

            prepared = prepared_template(source, Path(temp) / "cache")
            workbook = formualizer.Workbook.from_path(str(prepared))

            self.assertEqual([workbook.get_formula("S01", 2, column) for column in (1, 2, 3)],
                             [MASTER, MASTER, MASTER])
            self.assertEqual(workbook.evaluate_cell("S01", 2, 2), 7)

    def test_expansion_leaves_no_shared_formula_in_the_prepared_package(self):
        from pp.engine.native_load import prepared_template

        with tempfile.TemporaryDirectory(dir=_tmp()) as temp:
            source = shared_workbook(Path(temp) / "shared.xlsx")

            prepared = prepared_template(source, Path(temp) / "cache")

            with zipfile.ZipFile(prepared) as archive:
                sheet = archive.read("xl/worksheets/sheet1.xml").decode()
            self.assertNotIn('t="shared"', sheet)
            self.assertIn("xl/workbook.xml", zipfile.ZipFile(prepared).namelist())

    def test_expansion_keeps_relative_and_absolute_reference_translation(self):
        from pp.engine.native_load import translate_shared

        master = 'IF(A1>2,"Yes, No",A2)+$B$1+C1+D1'
        self.assertEqual(translate_shared(master, "A1", "C3"),
                         'IF(C3>2,"Yes, No",C4)+$B$1+E3+F3')

    def test_expansion_does_not_translate_names_or_string_literals(self):
        from pp.engine.native_load import translate_shared

        master = 'LOG10(A1)+MyRange1+_xlfn.CONCAT("S01!$A$1",A1)'
        self.assertEqual(translate_shared(master, "A1", "B2"),
                         'LOG10(B2)+MyRange1+_xlfn.CONCAT("S01!$A$1",B2)')


    def test_a_cell_split_across_two_reads_is_still_expanded(self):
            from pp.engine import native_load
            from pp.engine.native_load import expand_shared_formulas

            with tempfile.TemporaryDirectory(dir=_tmp()) as temp:
                source = Path(temp) / "fixture.xlsx"
                manifest = generate(source, SCALE, SEED)
                target = Path(temp) / "prepared.xlsx"

                with mock.patch.object(native_load, "CHUNK_BYTES", 64), \
                        mock.patch.object(native_load, "TAIL_GUARD_BYTES", 64):
                    stats = expand_shared_formulas(source, target)

                self.assertGreater(stats.followers, 0)
                with zipfile.ZipFile(target) as archive:
                    sheets = "".join(archive.read(f"xl/worksheets/sheet{number}.xml").decode()
                                     for number in (1, 2, 3))
                self.assertNotIn('t="shared"', sheets)
                workbook = formualizer.Workbook.from_path(str(target))
                self.assertEqual(_mismatches(workbook, manifest), [])


class NativeTemplateParityTests(unittest.TestCase):
    def test_native_template_load_evaluates_every_shared_family_correctly(self):
        with tempfile.TemporaryDirectory(dir=_tmp()) as temp:
            source = Path(temp) / "fixture.xlsx"
            manifest = generate(source, SCALE, SEED)

            workbook = _native_load_without_registrations(source)

            mismatches = _mismatches(workbook, manifest)
            self.assertEqual(
                mismatches, [],
                f"native load produced {len(mismatches)} mismatched formula cells")

    def test_engine_load_registers_no_formula_to_correct_a_shared_family(self):
        from pp.engine.formualizer import FormualizerAdapter

        with tempfile.TemporaryDirectory(dir=_tmp()) as temp:
            source = Path(temp) / "fixture.xlsx"
            manifest = generate(source, SCALE, SEED)
            registered = []

            def record(workbook, names, registrations):
                registered.extend(registrations)

            with mock.patch("pp.engine.formualizer._set_native_formulas", side_effect=record):
                engine = FormualizerAdapter.load(source, cache_dir=Path(temp) / "cache")

            self.assertEqual(registered, [])
            self.assertEqual(_engine_mismatches(engine, manifest), [])


def _native_load_without_registrations(source):
    """Load the prepared template exactly as the engine does, with no writes."""
    from pp.engine.native_load import prepared_template

    prepared = prepared_template(source, source.parent / "cache")
    return formualizer.load_workbook(str(prepared))


def _mismatches(workbook, manifest):
    names = workbook.sheet_names
    found = []
    for sheet_index, row, col, value in expected_formula_values(manifest, 7, 3):
        actual = workbook.evaluate_cell(names[sheet_index], row + 1, col + 1)
        if not isinstance(actual, (int, float)) or abs(actual - value) > 1e-9:
            found.append((names[sheet_index], row + 1, col + 1, actual, value))
    return found


def _engine_mismatches(engine, manifest):
    return [(sheet, row + 1, col + 1, engine.get(sheet, row, col), value)
            for sheet, row, col, value in expected_formula_values(manifest, 7, 3)
            if not _numeric(engine.get(sheet, row, col), value)]


def _numeric(actual, expected):
    return isinstance(actual, (int, float)) and abs(actual - expected) <= 1e-9


def _tmp():
    TMP.mkdir(parents=True, exist_ok=True)
    return TMP


if __name__ == "__main__":
    unittest.main()