import json
import math
import subprocess
import sys
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile
from collections import Counter
from pathlib import Path

from openpyxl.utils.cell import coordinate_to_tuple

from ops.engine_scale_fixture import build_manifest, expected_formula_values


MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


class EngineScaleFixtureTests(unittest.TestCase):
    def test_baseline_manifest_is_exact_without_generating_workbook(self):
        manifest = build_manifest(1.0, 20261011)

        self.assertEqual(manifest["cell_count"], 5_200_000)
        self.assertEqual(manifest["formula_count"], 1_500_000)
        self.assertEqual(
            [(sheet["cells"], sheet["formulas"]) for sheet in manifest["sheets"][:3]],
            [(1_700_000, 100_000), (760_000, 650_000), (467_000, 364_000)],
        )
        self.assertEqual(sum(sheet["cells"] for sheet in manifest["sheets"][3:]), 2_273_000)
        self.assertEqual(sum(sheet["formulas"] for sheet in manifest["sheets"][3:]), 386_000)
        self.assertEqual(manifest["formula_encoding_counts"], {
            "array": 664_000, "normal": 156_000, "shared": 680_000,
        })
        self.assertEqual(manifest["ordinary_coordinate"], "B1")
        self.assertEqual(
            {kind: sum(sheet["formula_encodings"][kind] for sheet in manifest["sheets"])
             for kind in ("array", "normal", "shared")},
            manifest["formula_encoding_counts"],
        )

    def test_tiny_fixture_is_reproducible_valid_ooxml_and_matches_xml_counts(self):
        with tempfile.TemporaryDirectory() as tmp:
            first = Path(tmp) / "a.xlsx"
            second = Path(tmp) / "b.xlsx"
            for output in (first, second):
                subprocess.run(
                    [sys.executable, "-m", "ops.engine_scale_fixture", str(output),
                     "--scale", "0.0001", "--seed", "20261011"],
                    check=True, capture_output=True, text=True,
                )
            self.assertEqual(first.read_bytes(), second.read_bytes())
            with zipfile.ZipFile(first) as archive:
                self.assertIn("[Content_Types].xml", archive.namelist())
                self.assertIn("_rels/.rels", archive.namelist())
                self.assertIn("xl/_rels/workbook.xml.rels", archive.namelist())
                manifest = json.loads(archive.read("engine-scale-manifest.json"))
                self.assertEqual(len(manifest["sheets"]), 22)
                self.assertEqual(manifest["hidden_sheet_count"], 9)
                counts, encodings, function_uses, cached = self._count_sheet_xml(
                    archive, manifest
                )
                self.assertEqual(counts["cells"], manifest["cell_count"])
                self.assertEqual(counts["formulas"], manifest["formula_count"])
                self.assertEqual(encodings, manifest["formula_encoding_counts"])
                self.assertEqual(function_uses, manifest["function_use_counts"])
                self.assertEqual(cached, list(expected_formula_values(manifest, 7, 3)))
                self.assertIn("LEN", function_uses)
                self.assertTrue(all(count > 0 for count in manifest["function_use_counts"].values()))
                self.assertGreaterEqual(manifest["broad_expected_changed_count"],
                                        manifest["formula_count"] // 10)
                self.assertEqual(manifest["ordinary_expected_changed_count"], 1)

    def test_tiny_fixture_uses_verified_ooxml_function_prefixes(self):
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "fixture.xlsx"
            subprocess.run(
                [sys.executable, "-m", "ops.engine_scale_fixture", str(output),
                 "--scale", "0.0001", "--seed", "20261011"], check=True,
            )
            with zipfile.ZipFile(output) as archive:
                formulas = "".join(
                    element.text or ""
                    for number in range(1, 23)
                    for _, element in ET.iterparse(
                        archive.open(f"xl/worksheets/sheet{number}.xml")
                    )
                    if element.tag == f"{MAIN}f"
                )

        for prefix in ("_xlfn.XLOOKUP(", "_xlfn.LET(_xlpm.v,",
                       "_xlfn.SWITCH(", "_xlfn.NUMBERVALUE(", "_xlfn.CONCAT("):
            with self.subTest(prefix=prefix):
                self.assertIn(prefix, formulas)

    def test_array_and_shared_formula_nodes_are_real_ooxml_encodings(self):
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "fixture.xlsx"
            subprocess.run(
                [sys.executable, "-m", "ops.engine_scale_fixture", str(output),
                 "--scale", "0.0001", "--seed", "20261011"], check=True,
            )
            with zipfile.ZipFile(output) as archive:
                arrays, shared_masters, shared_followers = [], 0, 0
                for number in range(1, 23):
                    for _, element in ET.iterparse(archive.open(f"xl/worksheets/sheet{number}.xml")):
                        if element.tag != f"{MAIN}c":
                            continue
                        formula = element.find(f"{MAIN}f")
                        if formula is None:
                            continue
                        if formula.get("t") == "array":
                            arrays.append((element.get("r"), formula.get("ref")))
                        if formula.get("t") == "shared":
                            if formula.text:
                                shared_masters += 1
                            else:
                                shared_followers += 1
                self.assertTrue(arrays)
                self.assertTrue(all(cell == reference for cell, reference in arrays))
                self.assertGreater(shared_masters, 0)
                self.assertGreater(shared_followers, 0)

    def test_every_tiny_cached_formula_matches_native_evaluation(self):
        from pp.engine.formualizer import FormualizerAdapter

        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "fixture.xlsx"
            subprocess.run(
                [sys.executable, "-m", "ops.engine_scale_fixture", str(output),
                 "--scale", "0.0001", "--seed", "20261011"], check=True,
            )
            engine = FormualizerAdapter.load(output)
            anchors = sorted(engine.formula_anchors)
            cached = {coordinate: engine.get(*coordinate) for coordinate in anchors}

            engine.workbook.evaluate_cells([
                (engine.names[sheet], row + 1, col + 1) for sheet, row, col in anchors
            ])
            native = {
                coordinate: engine.workbook.get_value(
                    engine.names[coordinate[0]], coordinate[1] + 1, coordinate[2] + 1
                )
                for coordinate in anchors
            }

            self.assertEqual(len(anchors), 150)
            self.assertEqual(native, cached)

    def test_scale_validation_rejects_nonfinite_range_and_input_overlap(self):
        for scale in (math.nan, math.inf, -math.inf, -0.1, 0.0, 1.01, 0.00006):
            with self.subTest(scale=scale), self.assertRaisesRegex(ValueError, "scale"):
                build_manifest(scale, 1)

    def _count_sheet_xml(self, archive, manifest):
        counts, encodings, function_uses, cached = Counter(), Counter(), Counter(), []
        for number in range(1, 23):
            shared = {}
            with archive.open(f"xl/worksheets/sheet{number}.xml") as source:
                for _, element in ET.iterparse(source):
                    if element.tag != f"{MAIN}c":
                        continue
                    counts["cells"] += 1
                    formula = element.find(f"{MAIN}f")
                    if formula is None:
                        continue
                    counts["formulas"] += 1
                    encodings[formula.get("t") or "normal"] += 1
                    row, column = coordinate_to_tuple(element.get("r"))
                    cached.append((number - 1, row - 1, column - 1,
                                   float(element.findtext(f"{MAIN}v"))))
                    text = formula.text
                    if formula.get("t") == "shared":
                        shared_id = formula.get("si")
                        if text:
                            shared[shared_id] = text
                        else:
                            text = shared[shared_id]
                    if text:
                        for function in manifest["function_use_counts"]:
                            function_uses[function] += text.upper().count(f"{function}(")
        return dict(counts), dict(encodings), dict(function_uses), cached


if __name__ == "__main__":
    unittest.main()
