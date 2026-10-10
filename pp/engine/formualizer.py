"""Version-pinned Formualizer adapter with workbook compatibility shims."""

from __future__ import annotations

from dataclasses import dataclass
import importlib.metadata
from pathlib import Path
import re
import zipfile
import xml.etree.ElementTree as ET

import formualizer
from openpyxl.utils.cell import get_column_letter, quote_sheetname

from ops.typed import typed


NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RNS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
REF = re.compile(r"(?:(?P<sheet>'(?:[^']|'')+'|[^!(),]+)!)?\$?(?P<col>[A-Z]{1,3})\$?(?P<row>[1-9][0-9]*)")
ISFORMULA = re.compile(r"ISFORMULA\(\s*((?:'(?:[^']|'')+'|[^!(),]+)!)?\$?[A-Z]{1,3}\$?[1-9][0-9]*\s*\)", re.I)


class EngineError(Exception):
    """Calculation errors that may be safely exposed as an unfinished job."""


def _require_version() -> None:
    if importlib.metadata.version("formualizer") != "0.11.1":
        raise EngineError("unsupported Formualizer version")


def _sheet_parts(archive: zipfile.ZipFile) -> list[tuple[str, str]]:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {item.get("Id"): item.get("Target") for item in rels}
    parts = []
    for sheet in workbook.findall(f"{NS}sheets/{NS}sheet"):
        target = targets[sheet.get(f"{RNS}id")].lstrip("/")
        parts.append((sheet.get("name", ""), target if target.startswith("xl/") else f"xl/{target}"))
    return parts


def _column_number(letters: str) -> int:
    value = 0
    for letter in letters:
        value = value * 26 + ord(letter) - 64
    return value


def _formula_map(path: Path) -> tuple[frozenset[tuple[int, int, int]], dict[tuple[int, int, int], str], list[str]]:
    cells, formulas, names = set(), {}, []
    with zipfile.ZipFile(path) as archive:
        for sheet_index, (name, part) in enumerate(_sheet_parts(archive)):
            names.append(name)
            with archive.open(part) as stream:
                for _, element in ET.iterparse(stream, events=("end",)):
                    if element.tag != f"{NS}c":
                        continue
                    formula, reference = element.find(f"{NS}f"), element.get("r")
                    if formula is not None and reference:
                        match = REF.fullmatch(reference)
                        if match is not None:
                            key = (sheet_index, int(match.group("row")) - 1,
                                   _column_number(match.group("col")) - 1)
                            cells.add(key)
                            if formula.text:
                                formulas[key] = "=" + formula.text.lstrip("=")
                    element.clear()
    return frozenset(cells), formulas, names


def _target_key(reference: str, current_sheet: int, names: list[str]) -> tuple[int, int, int] | None:
    match = REF.fullmatch(reference.strip())
    if match is None:
        return None
    sheet_name = match.group("sheet")
    sheet = current_sheet
    if sheet_name:
        normalized = sheet_name.rstrip("!").strip("'").replace("''", "'")
        try:
            sheet = names.index(normalized)
        except ValueError:
            return None
    return sheet, int(match.group("row")) - 1, _column_number(match.group("col")) - 1


def _rewrite_isformula(formula: str, sheet: int, formula_cells, names) -> str:
    def replacement(match):
        body = match.group(0)
        reference = body[body.find("(") + 1:body.rfind(")")]
        target = _target_key(reference, sheet, names)
        if target is None:
            raise EngineError("unsupported ISFORMULA reference")
        return "TRUE" if target in formula_cells else "FALSE"
    return ISFORMULA.sub(replacement, formula)


def _json_value(value):
    kind, normalized = typed(value)
    return {"error": normalized} if kind == "error" else normalized


def _qualified(sheet: str, row: int, col: int) -> str:
    return f"{quote_sheetname(sheet)}!{get_column_letter(col + 1)}{row + 1}"


@dataclass
class FormualizerEngine:
    workbook: object
    formula_cells: frozenset[tuple[int, int, int]]

    @classmethod
    def load(cls, template_path: str | Path) -> "FormualizerEngine":
        _require_version()
        try:
            path = Path(template_path)
            formula_cells, formulas, names = _formula_map(path)
            workbook = formualizer.load_workbook(str(path))
            for (sheet, row, col), formula in formulas.items():
                if ISFORMULA.search(formula):
                    workbook.set_formula(names[sheet], row + 1, col + 1,
                                         _rewrite_isformula(formula, sheet, formula_cells, names))
            workbook.evaluate_all()
        except EngineError:
            raise
        except Exception as exc:
            raise EngineError("template calculation engine failed to load") from exc
        return cls(workbook, formula_cells)

    def apply(self, changes: list[tuple[int, int, int, object]]) -> list[tuple[int, int, int, object]]:
        try:
            for sheet, row, col, value in changes:
                self.workbook.set_value(self.workbook.sheet_names[sheet], row + 1, col + 1, value)
            return self.evaluate_dirty()
        except Exception as exc:
            raise EngineError("template calculation engine failed to apply changes") from exc

    def evaluate_dirty(self) -> list[tuple[int, int, int, object]]:
        self.workbook.evaluate_all()
        return [(sheet, row, col, self.get(sheet, row, col))
                for sheet, row, col in sorted(self.formula_cells)]

    def get(self, sheet: int, row: int, col: int):
        try:
            value = self.workbook.get_value(self.workbook.sheet_names[sheet], row + 1, col + 1)
            return _json_value(value)
        except Exception as exc:
            raise EngineError("template calculation engine failed to read a cell") from exc

    def dependents(self, sheet: int, row: int, col: int) -> list[str]:
        try:
            name = self.workbook.sheet_names[sheet]
            report = self.workbook.dependents(_qualified(name, row, col), max_results=100_000)
            return [_normalize_address(item.cell) for item in report.dependents]
        except Exception as exc:
            raise EngineError("template calculation engine failed to inspect dependents") from exc

    def is_formula(self, sheet: int, row: int, col: int) -> bool:
        return (sheet, row, col) in self.formula_cells


def _normalize_address(address: str) -> str:
    sheet, marker, cell = address.rpartition("!")
    if not marker:
        return address
    normalized = sheet.strip("'").replace("''", "'")
    return f"{quote_sheetname(normalized)}!{cell}"


FormualizerAdapter = FormualizerEngine
