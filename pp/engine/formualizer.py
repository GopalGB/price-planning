"""Version-pinned Formualizer adapter with workbook compatibility shims."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, time
import importlib.metadata
from pathlib import Path
import re
import zipfile
import xml.etree.ElementTree as ET

import formualizer
from openpyxl.formula.translate import Translator
from openpyxl.utils.cell import get_column_letter, quote_sheetname, range_to_tuple
from openpyxl.utils.datetime import MAC_EPOCH, WINDOWS_EPOCH, to_excel

from ops.typed import ENGINE_ERRORS, typed

from ..values import exactly_equal


NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RNS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
REF = re.compile(r"(?:(?P<sheet>'(?:[^']|'')+'|[^!(),]+)!)?\$?(?P<col>[A-Z]{1,3})\$?(?P<row>[1-9][0-9]*)", re.I)
ISFORMULA = re.compile(r"(?:_xlfn\.)?ISFORMULA\(\s*((?:'(?:[^']|'')+'|[^!(),]+)!)?\$?[A-Z]{1,3}\$?[1-9][0-9]*\s*\)", re.I)
ISFORMULA_CALL = re.compile(r"(?:_xlfn\.)?ISFORMULA\s*\(", re.I)
QUOTED_TEXT = re.compile(r'("(?:[^"]|"")*")')
ERROR_KINDS = {token: kind for kind, token in ENGINE_ERRORS.items()}
UNSUPPORTED_EXTENT_TYPES = {"array", "dataTable"}


class EngineError(Exception):
    """Calculation errors that may be safely exposed as an unfinished job."""


@dataclass(frozen=True)
class FormulaLookup:
    sheet: int
    formula_cells: frozenset[tuple[int, int, int]]
    names: list[str]


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


def _uses_1904_date_system(path: Path) -> bool:
    with zipfile.ZipFile(path) as archive:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    properties = workbook.find(f"{NS}workbookPr")
    return properties is not None and properties.get("date1904", "").lower() in {"1", "true"}


def _column_number(letters: str) -> int:
    value = 0
    for letter in letters.upper():
        value = value * 26 + ord(letter) - 64
    return value


def _formula_map(path: Path) -> tuple[frozenset[tuple[int, int, int]], dict[tuple[int, int, int], str], list[str]]:
    cells, formulas, names = set(), {}, []
    masters, followers = {}, []
    with zipfile.ZipFile(path) as archive:
        for sheet_index, (name, part) in enumerate(_sheet_parts(archive)):
            names.append(name)
            with archive.open(part) as stream:
                for element in _formula_cells(stream):
                    formula, reference = element.find(f"{NS}f"), element.get("r")
                    if formula is not None and reference:
                        _reject_unsupported_extent(formula, name, reference)
                        match = REF.fullmatch(reference)
                        if match is not None:
                            key = (sheet_index, int(match.group("row")) - 1,
                                   _column_number(match.group("col")) - 1)
                            cells.add(key)
                            if formula.text:
                                text = "=" + formula.text.lstrip("=")
                                formulas[key] = text
                                if formula.get("t") == "shared" and formula.get("si") is not None:
                                    masters[(sheet_index, formula.get("si"))] = key, text
                            elif formula.get("t") == "shared" and formula.get("si") is not None:
                                followers.append((key, formula.get("si")))
                    element.clear()
    _expand_shared_text(formulas, masters, followers)
    return frozenset(cells), formulas, names


def _reject_unsupported_extent(formula, sheet, reference):
    formula_type = formula.get("t")
    if formula_type in UNSUPPORTED_EXTENT_TYPES:
        extent = formula.get("ref") or reference
        raise EngineError(
            f"unsupported multi-cell formula type {formula_type} at {sheet}!{extent}"
        )


def _formula_cells(stream, parser=ET.iterparse):
    stack = []
    for event, element in parser(stream, events=("start", "end")):
        if event == "start":
            stack.append(element)
            continue
        is_cell = element.tag == f"{NS}c"
        if is_cell:
            yield element
        if is_cell and len(stack) > 1:
            stack[-2].remove(element)
        stack.pop()


def _expand_shared_text(formulas, masters, followers):
    for target, shared_id in followers:
        master = masters.get((target[0], shared_id))
        if master is None:
            raise EngineError("shared formula master is missing")
        origin, text = master
        origin_ref = f"{get_column_letter(origin[2] + 1)}{origin[1] + 1}"
        target_ref = f"{get_column_letter(target[2] + 1)}{target[1] + 1}"
        formulas[target] = Translator(text, origin=origin_ref).translate_formula(target_ref)


def _target_key(reference: str, current_sheet: int, names: list[str]) -> tuple[int, int, int] | None:
    match = REF.fullmatch(reference.strip())
    if match is None:
        return None
    sheet_name = match.group("sheet")
    sheet = current_sheet
    if sheet_name:
        normalized = sheet_name.rstrip("!").strip("'").replace("''", "'")
        sheet = next((index for index, name in enumerate(names)
                      if name.casefold() == normalized.casefold()), None)
        if sheet is None:
            return None
    return sheet, int(match.group("row")) - 1, _column_number(match.group("col")) - 1


def _rewrite_isformula(formula: str, lookup: FormulaLookup) -> str:
    def replacement(match):
        body = match.group(0)
        reference = body[body.find("(") + 1:body.rfind(")")]
        target = _target_key(reference, lookup.sheet, lookup.names)
        if target is None:
            raise EngineError("unsupported ISFORMULA reference")
        return "TRUE" if target in lookup.formula_cells else "FALSE"
    segments = QUOTED_TEXT.split(formula)
    for index in range(0, len(segments), 2):
        segments[index] = ISFORMULA.sub(replacement, segments[index])
        if ISFORMULA_CALL.search(segments[index]):
            raise EngineError("unsupported ISFORMULA reference")
    return "".join(segments)


def _json_value(value, date1904):
    if isinstance(value, (date, time)):
        return to_excel(value, MAC_EPOCH if date1904 else WINDOWS_EPOCH)
    if isinstance(value, str):
        return value
    kind, normalized = typed(value)
    return {"error": normalized} if kind == "error" else normalized


def _native_value(value):
    if not isinstance(value, dict) or set(value) != {"error"}:
        return value
    native_kind = ERROR_KINDS.get(value["error"])
    if native_kind is None:
        raise EngineError("unsupported workbook error value")
    return formualizer.LiteralValue.error(native_kind, "")


def _qualified(sheet: str, row: int, col: int) -> str:
    return f"{quote_sheetname(sheet)}!{get_column_letter(col + 1)}{row + 1}"


def _spill_cells(workbook, anchors, names):
    cells = set(anchors)
    for sheet, row, col in anchors:
        report = workbook.inspect_cell(_qualified(names[sheet], row, col))
        spill = report.cell.to_dict().get("spill")
        if not spill or spill.get("kind") != "Anchor" or not spill.get("extent"):
            continue
        spill_name, bounds = range_to_tuple(spill["extent"])
        if spill_name != names[sheet]:
            raise EngineError("formula spill extent crosses workbook sheets")
        min_col, min_row, max_col, max_row = bounds
        cells.update((sheet, spill_row - 1, spill_col - 1)
                     for spill_row in range(min_row, max_row + 1)
                     for spill_col in range(min_col, max_col + 1))
    return frozenset(cells)


@dataclass
class FormualizerEngine:
    workbook: object
    formula_anchors: frozenset[tuple[int, int, int]]
    formula_cells: frozenset[tuple[int, int, int]]
    formula_values: dict[tuple[int, int, int], object]
    date1904: bool

    @classmethod
    def load(cls, template_path: str | Path) -> "FormualizerEngine":
        _require_version()
        try:
            path = Path(template_path)
            formula_cells, formulas, names = _formula_map(path)
            date1904 = _uses_1904_date_system(path)
            workbook = formualizer.load_workbook(str(path))
            if list(workbook.sheet_names) != names:
                raise EngineError("template engine sheet names do not match workbook XML")
            for (sheet, row, col), formula in formulas.items():
                if ISFORMULA_CALL.search(formula):
                    workbook.set_formula(names[sheet], row + 1, col + 1,
                                         _rewrite_isformula(formula, FormulaLookup(sheet, formula_cells, names)))
            workbook.evaluate_all()
        except EngineError:
            raise
        except Exception as exc:
            raise EngineError("template calculation engine failed to load") from exc
        engine = cls(workbook, formula_cells, frozenset(), {}, date1904)
        engine.formula_cells = _spill_cells(workbook, formula_cells, names)
        engine.formula_values = {coordinate: engine.get(*coordinate)
                                 for coordinate in engine.formula_cells}
        return engine

    def apply(self, changes: list[tuple[int, int, int, object]]) -> list[tuple[int, int, int, object]]:
        try:
            for sheet, row, col, value in changes:
                self.workbook.set_value(self.workbook.sheet_names[sheet], row + 1, col + 1,
                                        _native_value(value))
            return self.evaluate_dirty()
        except Exception as exc:
            raise EngineError("template calculation engine failed to apply changes") from exc

    def evaluate_dirty(self) -> list[tuple[int, int, int, object]]:
        self.workbook.evaluate_all()
        names = list(self.workbook.sheet_names)
        current = _spill_cells(self.workbook, self.formula_anchors, names)
        values, changes = {}, []
        for sheet, row, col in sorted(self.formula_cells | current):
            coordinate = sheet, row, col
            value = self.get(*coordinate) if coordinate in current else None
            if coordinate not in self.formula_values or not exactly_equal(self.formula_values[coordinate], value):
                changes.append((*coordinate, value))
            if coordinate in current:
                values[coordinate] = value
        self.formula_cells, self.formula_values = current, values
        return changes

    def result_snapshot(self) -> list[tuple[int, int, int, object]]:
        return [(*coordinate, self.get(*coordinate)) for coordinate in sorted(self.formula_cells)]

    def get(self, sheet: int, row: int, col: int):
        try:
            value = self.workbook.get_value(self.workbook.sheet_names[sheet], row + 1, col + 1)
            return _json_value(value, self.date1904)
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
