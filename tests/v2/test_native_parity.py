"""The native parity audit must pass on a healthy template and go red otherwise.

An audit that cannot report a failure is worse than no audit, so each check
below breaks the engine or the stored cache on purpose and requires the counts
to move.
"""

import json
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

from openpyxl import Workbook

from ops.engine_scale_fixture import generate
from ops.native_parity import audit
from tests.v2.support import TMP

SEED = 20261011
SCALE = 0.0001
UNRELATED = {"B1": ("999999", None)}


def _fixture(path):
    generate(path, SCALE, SEED)
    return path


def _cached(path, replacements):
    """Rewrite the value Excel stored for a cell without touching its formula."""
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    with zipfile.ZipFile(path) as source:
        members = {name: source.read(name) for name in source.namelist()}
    for member, changes in replacements.items():
        root = _parse(members[member])
        for cell in root.iter(f"{{{namespace}}}c"):
            if cell.get("r") not in changes:
                continue
            value, kind = changes[cell.get("r")]
            if kind is None:
                cell.attrib.pop("t", None)
            else:
                cell.set("t", kind)
            node = cell.find(f"{{{namespace}}}v")
            if node is None:
                node = _sub(cell, f"{{{namespace}}}v")
            node.text = value
        members[member] = _render(root)
    with zipfile.ZipFile(path, "w") as target:
        for name, data in members.items():
            target.writestr(name, data)
    return path


def _parse(data):
    from xml.etree import ElementTree as ET

    return ET.fromstring(data)


def _sub(parent, tag):
    from xml.etree import ElementTree as ET

    return ET.SubElement(parent, tag)


def _render(root):
    from xml.etree import ElementTree as ET

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _small_formula_workbook(path):
    """One hub input feeding a family, so an edit has a wide blast radius."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Calc"
    sheet["A1"] = 7
    sheet["B1"] = "=A1*2"
    sheet["C1"] = "=B1+1"
    sheet["D1"] = "=C1+A1"
    workbook.save(path)
    _cached(path, {"xl/worksheets/sheet1.xml": {"B1": ("14", None),
                                                "C1": ("15", None),
                                                "D1": ("22", None)}})
    return path


class _IgnoringRevert:
    """Workbook proxy that accepts the edit and silently drops the restore."""

    def __init__(self, workbook):
        self._workbook = workbook
        self._writes = 0

    def set_value(self, *arguments, **keywords):
        self._writes += 1
        if self._writes % 2 == 0:
            return None
        return self._workbook.set_value(*arguments, **keywords)

    def __getattr__(self, name):
        return getattr(self._workbook, name)


def _report(path, edits=0, probes=0):
    return audit(path, sample=200, edits=edits, probes=probes, seed=SEED,
                 cache_dir=path.parent / "cache")


class NativeParityAuditTests(unittest.TestCase):
    def test_audit_passes_when_the_engine_matches_the_stored_cache(self):
        with tempfile.TemporaryDirectory(dir=_tmp()) as temp:
            report = _report(_fixture(Path(temp) / "fixture.xlsx"))

            self.assertEqual(report["status"], "pass")
            self.assertEqual(report["mismatch_total"], 0)
            self.assertEqual(report["lazy_sample"]["matched"], 150)
            self.assertEqual(report["lazy_sample"]["mismatched"], 0)

    def test_audit_reports_a_wrong_stored_value_as_a_mismatch(self):
        with tempfile.TemporaryDirectory(dir=_tmp()) as temp:
            source = _fixture(Path(temp) / "fixture.xlsx")
            _cached(source, {"xl/worksheets/sheet2.xml": {"A2": ("1", None)}})

            report = _report(source)

            self.assertEqual(report["status"], "fail")
            self.assertEqual(report["lazy_sample"]["mismatched"], 1)
            self.assertGreater(report["mismatch_total"], 0)
            self.assertIn("SUMIF", report["lazy_sample"]["mismatches_by_function"])

    def test_audit_reports_a_formula_that_never_evaluates(self):
        with tempfile.TemporaryDirectory(dir=_tmp()) as temp:
            source = _fixture(Path(temp) / "fixture.xlsx")

            with mock.patch("formualizer.Workbook.evaluate_cell",
                            side_effect=RuntimeError("engine refused")):
                report = _report(source)

            self.assertEqual(report["status"], "fail")
            self.assertEqual(report["lazy_sample"]["evaluation_error"], 150)

    def test_audit_reports_an_input_that_does_not_restore_after_revert(self):
        with tempfile.TemporaryDirectory(dir=_tmp()) as temp:
            source = _small_formula_workbook(Path(temp) / "small.xlsx")
            report = _small_with_ignored_revert(source)

            self.assertEqual(report["status"], "fail")
            self.assertEqual(report["edit_revert"]["cycles"], 1)
            self.assertGreater(report["edit_revert"]["mismatched_after_revert"], 0)

    def test_audit_reports_nothing_when_revert_is_honoured(self):
        with tempfile.TemporaryDirectory(dir=_tmp()) as temp:
            source = _small_formula_workbook(Path(temp) / "small.xlsx")

            report = _report(source, edits=1, probes=4)

            self.assertEqual(report["edit_revert"]["mismatched_after_revert"], 0)
            self.assertGreater(report["edit_revert"]["cells_after_revert"], 0)

    def test_audit_output_carries_counts_only(self):
        with tempfile.TemporaryDirectory(dir=_tmp()) as temp:
            source = _small_formula_workbook(Path(temp) / "small.xlsx")

            text = json.dumps(_report(source, edits=1, probes=4))

            self.assertNotIn("B1", text.replace("VLOOKUP", ""))
            self.assertNotIn("=A1", text)
            for banned in ("sheet", "address", "cell", "value"):
                self.assertNotIn(f'"{banned}"', text)


def _small_with_ignored_revert(source):
    """Run the audit against an engine that drops the restoring write."""
    from ops.native_parity import FormulaReader, ParityAudit
    from pp.engine import formualizer as adapter
    from pp.engine.native_load import prepared_template

    prepared = prepared_template(source, source.parent / "cache")
    engine = adapter.FormualizerEngine.load(prepared)
    engine.workbook = _IgnoringRevert(engine.workbook)
    reader = FormulaReader(prepared, engine.names)
    result = ParityAudit(engine, reader, len(engine.formula_anchors), 1, 4, SEED).run()
    failures = (result["lazy_sample"]["mismatched"]
                + result["lazy_sample"]["evaluation_error"]
                + result["edit_revert"]["mismatched_after_revert"])
    return {**result, "status": "pass" if failures == 0 else "fail",
            "mismatch_total": failures}


def _tmp():
    TMP.mkdir(parents=True, exist_ok=True)
    return TMP


if __name__ == "__main__":
    unittest.main()