"""Version-pinned Formualizer adapter with workbook compatibility shims."""

from __future__ import annotations

import hashlib
import importlib.metadata
import os
import pickle
import re
import xml.etree.ElementTree as ET
import zipfile
from collections import Counter
from collections.abc import Iterator, Set
from dataclasses import dataclass
from datetime import date, time
from pathlib import Path

import formualizer
from openpyxl.formula.translate import Translator
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils.cell import get_column_letter, quote_sheetname, range_to_tuple
from openpyxl.utils.cell import range_boundaries
from openpyxl.utils.datetime import MAC_EPOCH, WINDOWS_EPOCH, to_excel

from ops.typed import ENGINE_ERRORS, typed

from ..values import exactly_equal
from .native_load import prepared_template

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RNS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
REF = re.compile(r"(?:(?P<sheet>'(?:[^']|'')+'|[^!(),]+)!)?\$?(?P<col>[A-Z]{1,3})\$?(?P<row>[1-9][0-9]*)", re.IGNORECASE)
ISFORMULA = re.compile(r"(?:_xlfn\.)?ISFORMULA\(\s*((?:'(?:[^']|'')+'|[^!(),]+)!)?\$?[A-Z]{1,3}\$?[1-9][0-9]*\s*\)", re.IGNORECASE)
ISFORMULA_CALL = re.compile(r"(?:_xlfn\.)?ISFORMULA\s*\(", re.IGNORECASE)
QUOTED_TEXT = re.compile(r'("(?:[^"]|"")*")')
ERROR_KINDS = {token: kind for kind, token in ENGINE_ERRORS.items()}
UNSUPPORTED_EXTENT_TYPES = {"array", "dataTable"}
DYNAMIC_ROOT = re.compile(r"\b(?:INDIRECT|OFFSET|NOW|TODAY|RAND|RANDBETWEEN|CELL)\s*\(", re.I)
BINDING_ROOT = re.compile(r"\b(?:_xlfn\.)?(?:LET|LAMBDA)\s*\(", re.I)
POTENTIAL_SPILL = re.compile(
    r"\b(?:SEQUENCE|FILTER|UNIQUE|SORT(?:BY)?|RANDARRAY|TAKE|DROP|CHOOSECOLS|CHOOSEROWS|"
    r"TOCOL|TOROW|WRAPCOLS|WRAPROWS|EXPAND|VSTACK|HSTACK)\s*\(", re.I)
ABSENT = object()
MAX_ROWS, MAX_COLS = 1_048_576, 16_384


class EngineError(Exception):
    """Calculation errors that may be safely exposed as an unfinished job."""


@dataclass(frozen=True)
class FormulaLookup:
    sheet: int
    formula_cells: frozenset[tuple[int, int, int]]
    names: list[str]


@dataclass(frozen=True)
class FormulaOutputs(Set):
    anchors: frozenset[tuple[int, int, int]]
    spill_counts: Counter

    def __contains__(self, coordinate) -> bool:
        return coordinate in self.anchors or self.spill_counts.get(coordinate, 0) > 0

    def __iter__(self) -> Iterator[tuple[int, int, int]]:
        yield from self.anchors
        yield from (coordinate for coordinate, count in self.spill_counts.items()
                    if count > 0 and coordinate not in self.anchors)

    def __len__(self) -> int:
        return len(self.anchors) + sum(
            count > 0 and coordinate not in self.anchors
            for coordinate, count in self.spill_counts.items()
        )


TRACE_U32_MAX = 2**32 - 1


@dataclass(frozen=True)
class TraceBudget:
    nodes: int
    links: int
    range_members: int


@dataclass(frozen=True)
class VisibleBlocks:
    """The row ranges clients currently hold, one entry per sheet."""

    rows: tuple[tuple[int, int, int], ...] = ()

    @classmethod
    def of(cls, *ranges: tuple[int, int, int]) -> VisibleBlocks:
        return cls(tuple(_ordered(*item) for item in ranges))

    @classmethod
    def of_rows(cls, coordinates) -> VisibleBlocks:
        """Group cells into the smallest set of visible row blocks."""
        grouped: dict[int, list[int]] = {}
        for sheet, row, _col in coordinates:
            grouped.setdefault(sheet, []).append(row)
        return cls(tuple(_ordered(sheet, min(rows), max(rows))
                         for sheet, rows in grouped.items() if rows))

    def covers(self, coordinate: tuple[int, int, int]) -> bool:
        sheet, row, _col = coordinate
        return any(block[0] == sheet and block[1] <= row <= block[2] for block in self.rows)


def _ordered(sheet: int, first: int, last: int) -> tuple[int, int, int]:
    return (sheet, first, last) if first <= last else (sheet, last, first)


class _NoBlocks(VisibleBlocks):
    """An empty view hides nothing, so every affected cell is evaluated."""

    def covers(self, coordinate):
        return True


@dataclass(frozen=True)
class FormulaBounds:
    sheet: int
    names: list[str]
    dimensions: list[tuple[int, int]]


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


def _cached_formula_map(prepared: Path):
    """Reuse the formula map of an immutable prepared template across every copy.

    The map takes minutes on a large template. It sits beside the prepared
    package in the engine cache and is keyed by this module's source, so a code
    change can never reuse a map built by older parsing rules.
    """
    code = hashlib.sha256(Path(__file__).read_bytes()).hexdigest()[:12]
    path = prepared.with_name(f"{prepared.name}.map-{code}.pickle")
    try:
        with path.open("rb") as handle:
            return pickle.load(handle)
    except FileNotFoundError:
        pass
    except Exception:
        path.unlink(missing_ok=True)
    mapped = _formula_map(prepared)
    pending = path.with_name(f"{path.name}.{os.getpid()}.tmp")
    try:
        with pending.open("wb") as handle:
            pickle.dump(mapped, handle, protocol=pickle.HIGHEST_PROTOCOL)
        os.replace(pending, path)
    except OSError:
        pending.unlink(missing_ok=True)
    return mapped


def _formula_map(path: Path):
    cells, formulas, names, caches, dimensions = set(), {}, [], {}, []
    settled = set()
    masters, followers = {}, []
    with zipfile.ZipFile(path) as archive:
        sheet_parts = _sheet_parts(archive)
        defined_ranges = _defined_name_ranges(archive, [name for name, _ in sheet_parts])
        for sheet_index, (name, part) in enumerate(sheet_parts):
            names.append(name)
            max_row = max_col = 1
            with archive.open(part) as stream:
                for element in _formula_cells(stream):
                    formula, reference = element.find(f"{NS}f"), element.get("r")
                    match = REF.fullmatch(reference or "")
                    if match is not None:
                        max_row = max(max_row, int(match.group("row")))
                        max_col = max(max_col, _column_number(match.group("col")))
                    if formula is not None and reference:
                        _reject_unsupported_extent(formula, name, reference)
                        if match is not None:
                            key = (sheet_index, int(match.group("row")) - 1,
                                   _column_number(match.group("col")) - 1)
                            cells.add(key)
                            parsed = _cached_value(element)
                            if parsed is not None:
                                caches[key] = parsed
                                # Excel records a dynamic-array result that did not spill as an own-cell array.
                                if formula.get("t") == "array":
                                    settled.add(key)
                            if formula.text:
                                text = "=" + formula.text.lstrip("=")
                                formulas[key] = text
                                if formula.get("t") == "shared" and formula.get("si") is not None:
                                    masters[(sheet_index, formula.get("si"))] = key, text
                            elif formula.get("t") == "shared" and formula.get("si") is not None:
                                followers.append((key, formula.get("si")))
                    element.clear()
            dimensions.append((max_row, max_col))
    _expand_shared_text(formulas, masters, followers)
    _validate_isformula_forms(formulas, frozenset(cells), names)
    named_formulas, opaque_formulas = _expand_defined_names(formulas, names, defined_ranges)
    budget = _trace_budget(formulas, names, dimensions)
    return (frozenset(cells), formulas, names, caches,
            frozenset(key for key, _ in followers) | named_formulas,
            opaque_formulas, budget, frozenset(settled))


def _defined_name_ranges(archive, names):
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    ranges = {}
    for item in workbook.findall(f"{NS}definedNames/{NS}definedName"):
        name, expression = item.get("name"), (item.text or "").lstrip("=")
        if not name:
            continue
        local = item.get("localSheetId")
        if local is not None:
            try:
                local = int(local)
            except ValueError:
                continue
            if not 0 <= local < len(names):
                continue
        target = None
        if "!" in expression:
            try:
                sheet_name, _ = range_to_tuple(expression)
            except ValueError:
                pass
            else:
                if any(sheet_name.casefold() == candidate.casefold() for candidate in names):
                    target = expression
        ranges[(local, name.casefold())] = target
    return ranges


def _expand_defined_names(formulas, names, ranges):
    rewritten, opaque = set(), set()
    for coordinate, formula in list(formulas.items()):
        tokens, changed = Tokenizer(formula).items, False
        binding_formula = BINDING_ROOT.search(formula) is not None
        for token in tokens:
            if token.type != "OPERAND" or token.subtype != "RANGE":
                continue
            if "[" in token.value:
                raise EngineError("unsupported structured or external formula reference")
            if _static_reference(token.value, names):
                continue
            if ":" in token.value:
                raise EngineError("unsupported named range operator")
            target = _defined_name_target(token.value, coordinate[0], names, ranges)
            if target is not ABSENT and not binding_formula and target is not None:
                token.value, changed = target, True
            elif target is not ABSENT or "!" in token.value:
                opaque.add(coordinate)
        if changed:
            formulas[coordinate] = "=" + "".join(token.value for token in tokens)
            rewritten.add(coordinate)
    return frozenset(rewritten), frozenset(opaque)


def _defined_name_target(reference, current_sheet, names, ranges):
    sheet_name, marker, name = reference.rpartition("!")
    sheet = current_sheet
    if marker:
        normalized = sheet_name.strip("'").replace("''", "'")
        sheet = next((index for index, candidate in enumerate(names)
                      if candidate.casefold() == normalized.casefold()), None)
        if sheet is None:
            return ABSENT
    key = name.casefold()
    if (sheet, key) in ranges:
        return ranges[(sheet, key)]
    return ranges.get((None, key), ABSENT)


def _axis_within_excel(first, last, limit):
    if first is None or last is None:
        return first is None and last is None
    return 1 <= first <= last <= limit


def _static_reference(reference, names):
    sheet_name, marker, address = reference.rpartition("!")
    if marker:
        normalized = sheet_name.strip("'").replace("''", "'")
        if not any(name.casefold() == normalized.casefold() for name in names):
            return False
    try:
        min_col, min_row, max_col, max_row = range_boundaries(address)
    except ValueError:
        return False
    if ":" not in address and (min_col is None or min_row is None):
        return False
    columns_valid = _axis_within_excel(min_col, max_col, MAX_COLS)
    rows_valid = _axis_within_excel(min_row, max_row, MAX_ROWS)
    return columns_valid and rows_valid and (min_col is not None or min_row is not None)


def _trace_budget(formulas, names, dimensions):
    links, ranges = 0, {}
    for coordinate, formula in formulas.items():
        bounds = FormulaBounds(coordinate[0], names, dimensions)
        for token in Tokenizer(formula).items:
            if token.type != "OPERAND" or token.subtype != "RANGE":
                continue
            count, identity = _reference_shape(token.value, bounds)
            links += count
            if identity is not None:
                ranges[identity] = count
    nodes = len(formulas) + len(ranges)
    return TraceBudget(nodes, links, sum(ranges.values()))


def _reference_shape(reference, bounds):
    sheet_name, marker, address = reference.rpartition("!")
    sheet = bounds.sheet
    if marker:
        normalized = sheet_name.strip("'").replace("''", "'")
        sheet = next((index for index, name in enumerate(bounds.names)
                      if name.casefold() == normalized.casefold()), sheet)
    try:
        min_col, min_row, max_col, max_row = range_boundaries(address)
    except ValueError:
        return 1, None
    used_rows, used_cols = bounds.dimensions[sheet]
    min_col, max_col = min_col or 1, max_col or used_cols
    min_row, max_row = min_row or 1, max_row or used_rows
    count = max_col - min_col + 1
    count *= max_row - min_row + 1
    identity = (sheet, min_row, min_col, max_row, max_col) if ":" in address else None
    return count, identity


def _cached_value(element):
    value = element.find(f"{NS}v")
    if value is None:
        return None
    raw, kind = value.text or "", element.get("t")
    if not raw and kind is None:
        return None
    if kind == "e":
        return {"error": raw}
    if kind == "b":
        return raw == "1"
    if kind == "str":
        return raw
    try:
        return float(raw) if any(ch in raw for ch in ".eE") else int(raw)
    except ValueError:
        return raw


def _reject_unsupported_extent(formula, sheet, reference):
    formula_type = formula.get("t")
    if formula_type == "array" and _is_own_cell_array(formula.get("ref"), reference):
        return
    if formula_type in UNSUPPORTED_EXTENT_TYPES:
        extent = formula.get("ref") or "<missing>"
        raise EngineError(f"unsupported {formula_type} extent at {sheet}!{reference}: {extent}")


def _is_own_cell_array(extent, reference):
    if not extent:
        return False
    own = REF.fullmatch(reference)
    if own is None or own.group("sheet"):
        return False
    bounds = extent.split(":")
    if len(bounds) > 2:
        return False
    cells = [REF.fullmatch(item.strip()) for item in bounds]
    if any(match is None or match.group("sheet") for match in cells):
        return False
    return all((match.group("col").upper(), int(match.group("row"))) ==
               (own.group("col").upper(), int(own.group("row"))) for match in cells)


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


def _validate_isformula_forms(formulas, formula_cells, names):
    for (sheet, _, _), formula in formulas.items():
        if ISFORMULA_CALL.search(formula):
            _rewrite_isformula(formula, FormulaLookup(sheet, formula_cells, names))


def _formula_runs(registrations):
    run = []
    for coordinate in sorted(registrations):
        formula = registrations[coordinate]
        if run:
            previous = run[-1][0]
            contiguous = (coordinate[:2] == previous[:2]
                          and coordinate[2] == previous[2] + 1)
            if not contiguous:
                yield run
                run = []
        run.append((coordinate, formula))
    if run:
        yield run


def _set_native_formulas(workbook, names, registrations):
    for run in _formula_runs(registrations):
        (sheet, row, col), formula = run[0]
        if len(run) == 1:
            workbook.set_formula(names[sheet], row + 1, col + 1, formula)
            continue
        workbook.set_formulas_batch(
            names[sheet], row + 1, col + 1, [[item[1] for item in run]]
        )


def _json_value(value, date1904):
    if isinstance(value, (date, time)):
        return to_excel(value, MAC_EPOCH if date1904 else WINDOWS_EPOCH)
    if isinstance(value, dict) and set(value) == {"error"}:
        return value
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


def _spill_extent(workbook, coordinate, names):
    sheet, row, col = coordinate
    report = workbook.inspect_cell(_qualified(names[sheet], row, col))
    spill = report.cell.to_dict().get("spill")
    if not spill or spill.get("kind") != "Anchor" or not spill.get("extent"):
        return frozenset()
    spill_name, bounds = range_to_tuple(spill["extent"])
    if spill_name != names[sheet]:
        raise EngineError("formula spill extent crosses workbook sheets")
    min_col, min_row, max_col, max_row = bounds
    return frozenset((sheet, spill_row - 1, spill_col - 1)
                     for spill_row in range(min_row, max_row + 1)
                     for spill_col in range(min_col, max_col + 1))


def _spill_map(workbook, anchors, names):
    return {anchor: extent for anchor in anchors
            if (extent := _spill_extent(workbook, anchor, names))}


def _spill_counts(spills):
    counts = Counter()
    for extent in spills.values():
        counts.update(extent)
    return counts


@dataclass
class FormualizerEngine:
    workbook: object
    formula_anchors: frozenset[tuple[int, int, int]]
    formula_cells: FormulaOutputs
    date1904: bool
    names: list[str]
    cached_values: dict[tuple[int, int, int], object]
    dynamic_targets: frozenset[tuple[int, int, int]]
    spill_cells: dict[tuple[int, int, int], frozenset[tuple[int, int, int]]]
    spill_counts: Counter
    baseline_spill_cells: frozenset[tuple[int, int, int]]
    baseline_spills: dict[tuple[int, int, int], frozenset[tuple[int, int, int]]]
    baseline_values: dict[tuple[int, int, int], object]
    overlay_values: dict[tuple[int, int, int], object]
    overlay_before: dict[tuple[int, int, int], object]
    overlay_upserts: dict[tuple[int, int, int], object]
    overlay_removals: set[tuple[int, int, int]]
    trace_budget: TraceBudget
    visible: VisibleBlocks = _NoBlocks()
    stale: frozenset[tuple[int, int, int]] = frozenset()
    evaluated: frozenset[tuple[int, int, int]] = frozenset()

    @classmethod
    def load(cls, template_path: str | Path,
             recalculate: tuple[tuple[int, int, int], ...] = (),
             cache_dir: str | Path | None = None) -> FormualizerEngine:
        _require_version()
        try:
            path = Path(template_path)
            prepared = prepared_template(path, cache_dir)
            mapped = _cached_formula_map(prepared)
            (formula_cells, formulas, names, cached_values, native_formulas,
             opaque_formulas, trace_budget, settled) = mapped
            date1904 = _uses_1904_date_system(prepared)
            workbook = formualizer.load_workbook(str(prepared))
            if list(workbook.sheet_names) != names:
                raise EngineError("template engine sheet names do not match workbook XML")
            lookup = FormulaLookup(0, formula_cells, names)
            registrations = {}
            for coordinate, formula in formulas.items():
                if coordinate in native_formulas or ISFORMULA_CALL.search(formula):
                    sheet, _, _ = coordinate
                    if ISFORMULA_CALL.search(formula):
                        formula = _rewrite_isformula(
                            formula, FormulaLookup(sheet, lookup.formula_cells, lookup.names)
                        )
                    registrations[coordinate] = formula
            if registrations:
                _set_native_formulas(workbook, names, registrations)
            dynamic = frozenset(
                opaque_formulas | {
                    key for key, formula in formulas.items() if DYNAMIC_ROOT.search(formula)
                }
            )
            # Roots Excel saved as non-spilling with a value need no bootstrap evaluation.
            spill_roots = frozenset(key for key, formula in formulas.items()
                                    if POTENTIAL_SPILL.search(formula) and key not in settled)
            initial = sorted((formula_cells - cached_values.keys()) | spill_roots)
            if initial:
                workbook.evaluate_cells([(names[s], r + 1, c + 1) for s, r, c in initial])
                for coordinate in initial:
                    cached_values.pop(coordinate, None)
        except EngineError:
            raise
        except Exception as exc:
            raise EngineError("template calculation engine failed to load") from exc
        try:
            spills = _spill_map(workbook, spill_roots, names)
            spill_counts = _spill_counts(spills)
            outputs = FormulaOutputs(formula_cells, spill_counts)
            engine = cls(workbook, formula_cells, outputs, date1904, names, cached_values, dynamic,
                         spills, spill_counts, frozenset(spill_counts), dict(spills), {}, {}, {},
                         {}, set(), trace_budget)
            engine._close_bootstrap_spills(spill_roots)
            if recalculate:
                engine.recalculate(recalculate)
            return engine
        except EngineError:
            raise
        except Exception as exc:
            raise EngineError("template calculation engine failed to load") from exc

    def apply(self, changes: list[tuple[int, int, int, object]]) -> list[tuple[int, int, int, object]]:
        self._reset_overlay_delta()
        if not changes:
            return []
        try:
            edited = {(sheet, row, col) for sheet, row, col, _ in changes}
            for sheet, row, col, value in changes:
                self.workbook.set_value(self.workbook.sheet_names[sheet], row + 1, col + 1,
                                        _native_value(value))
            affected = self._dependent_targets(self._calculation_roots(edited))
            visible = {target for target in affected if self.visible.covers(target)}
            published = self._evaluate_clearing(visible)
            self._mark_stale(affected)
            return published
        except EngineError:
            raise
        except Exception as exc:
            raise EngineError("template calculation engine failed to apply changes") from exc

    def set_visible_blocks(self, blocks: VisibleBlocks) -> None:
        """Restrict eager recalculation to the blocks clients currently hold."""
        self.visible = blocks if blocks.rows else _NoBlocks()

    @property
    def stale_cells(self) -> frozenset[tuple[int, int, int]]:
        """Cells downstream of an edit whose value the engine has not produced."""
        return self.stale

    def recalculate_stale(self) -> list[tuple[int, int, int, object]]:
        """Calculate the stale cells that are now visible and publish them."""
        pending = {target for target in self.stale if self.visible.covers(target)}
        if not pending:
            return []
        self._reset_overlay_delta()
        published = self._evaluate_clearing(pending)
        self._mark_stale(pending)
        return published

    def _evaluate_clearing(self, targets):
        """Evaluate cells that are still marked stale, keeping them stale on failure.

        The marks are cleared first so the publish path captures the pre-edit
        value instead of resolving each cell on demand and reporting no change.
        """
        self.stale = self.stale - set(targets)
        try:
            return self.evaluate_dirty(targets)
        except BaseException:
            self.stale = self.stale | set(targets)
            raise

    def _mark_stale(self, affected) -> None:
        """Keep every affected cell that was not evaluated in this round."""
        pending = {target for target in affected if target not in self.evaluated}
        self.stale = frozenset((self.stale | pending) - self.evaluated)

    def _dependent_targets(self, roots, include_dynamic=True):
        # The native trace takes u32 budgets; a capped trace that truncates still fails closed below.
        node_budget = min(max(self.trace_budget.nodes + len(roots), 1), TRACE_U32_MAX)
        link_budget = min(max(self.trace_budget.links, node_budget), TRACE_U32_MAX)
        work_budget = node_budget + link_budget + 1
        native_roots = [(self.names[sheet], row + 1, col + 1) for sheet, row, col in roots]
        self.workbook.get_eval_plan(native_roots)
        trace = self.workbook.trace(
            [_qualified(self.names[sheet], row, col) for sheet, row, col in roots],
            direction=formualizer.TraceDirection.Dependents, max_depth=node_budget,
            max_nodes=node_budget, max_links=link_budget, max_work=work_budget,
            range_member_budget=min(max(self.trace_budget.range_members, 1), TRACE_U32_MAX),
            include_values=False)
        if trace.truncation.incomplete:
            raise EngineError("dependency trace truncated")
        targets = {target for address in trace.nodes.keys()
                   if (target := _target_key(_normalize_address(address), 0, self.names))
                   in self.formula_anchors}
        return targets | self.dynamic_targets if include_dynamic else targets

    def _close_bootstrap_spills(self, spill_roots) -> None:
        if not spill_roots:
            return
        roots = set(spill_roots)
        roots.update(*(self.spill_cells.get(root, frozenset()) for root in spill_roots))
        self.evaluate_dirty(self._dependent_targets(roots, include_dynamic=False))
        self.baseline_spill_cells = frozenset(self.spill_counts)
        self.baseline_spills = dict(self.spill_cells)
        self.baseline_values.clear()
        self.overlay_values.clear()
        self._reset_overlay_delta()

    def evaluate_dirty(self, targets) -> list[tuple[int, int, int, object]]:
        pending = {target for target in targets if target in self.formula_anchors}
        evaluated = set()
        if not pending:
            self.evaluated = frozenset()
            return []
        previous, touched = {}, set()
        for _ in range(max(len(self.formula_anchors) + 1, 1)):
            old_spills = {anchor: self.spill_cells.get(anchor, frozenset())
                          for anchor in pending}
            before = set(pending).union(*old_spills.values())
            self._capture_previous(before, previous)
            touched.update(before)
            self._evaluate_targets(pending)
            evaluated.update(before)
            changed_spills = self._refresh_spills(pending, old_spills)
            current = set(pending)
            current.update(*(self.spill_cells.get(anchor, frozenset()) for anchor in pending))
            touched.update(current)
            evaluated.update(current)
            if not changed_spills:
                break
            next_targets = self._dependent_targets(changed_spills, include_dynamic=False)
            if not next_targets:
                break
            pending = next_targets
        else:
            raise EngineError("spill dependency closure did not converge")
        self.evaluated = frozenset(evaluated)
        return self._publish_changes(previous, touched)

    def _capture_previous(self, coordinates, previous) -> None:
        self._remember_baseline(coordinates)
        for coordinate in coordinates:
            if coordinate not in previous:
                previous[coordinate] = (self.get(*coordinate)
                                        if coordinate in self.formula_cells else ABSENT)

    def _evaluate_targets(self, targets) -> None:
        for target in targets:
            self.cached_values.pop(target, None)
        # Formualizer 0.11.1 has no cache setter. Targeted evaluation may evaluate
        # cold native precedents, so every traced formula target loses its XML cache.
        native = [(self.names[s], r + 1, c + 1) for s, r, c in sorted(targets)]
        self.workbook.evaluate_cells(native)

    def _refresh_spills(self, targets, old_spills):
        changed = set()
        for anchor in targets:
            previous = old_spills[anchor]
            current = _spill_extent(self.workbook, anchor, self.names)
            self._replace_spill(anchor, current)
            if previous != current:
                changed.update(previous)
                changed.update(current)
        return changed

    def _publish_changes(self, previous, touched):
        values = {coordinate: self.get(*coordinate)
                  if coordinate in self.formula_cells else ABSENT for coordinate in touched}
        for coordinate, value in values.items():
            self._update_overlay(coordinate, value)
        return [(*coordinate, None if value is ABSENT else value)
                for coordinate, value in sorted(values.items())
                if _output_changed(previous.get(coordinate, ABSENT), value)]

    def adopt_baseline(self, source: FormualizerEngine):
        if self.formula_anchors != source.formula_anchors:
            raise EngineError("copy baseline formula topology changed")
        loaded_spills = dict(self.spill_cells)
        self.baseline_spill_cells = source.baseline_spill_cells
        self.baseline_spills = dict(source.baseline_spills)
        self.baseline_values = dict(source.baseline_values)
        self.spill_cells = dict(self.baseline_spills)
        self.spill_counts.clear()
        self.spill_counts.update(_spill_counts(self.spill_cells))
        anchors = loaded_spills.keys() | self.baseline_spills.keys()
        return frozenset(anchor for anchor in anchors
                         if loaded_spills.get(anchor) != self.baseline_spills.get(anchor))

    def recalculate(self, roots) -> list[tuple[int, int, int, object]]:
        self._reset_overlay_delta()
        return self.evaluate_dirty(self._dependent_targets(self._calculation_roots(roots)))

    def remember_baseline_for_roots(self, roots) -> None:
        targets = self._dependent_targets(self._calculation_roots(roots))
        cells = set(targets)
        cells.update(*(self.spill_cells.get(target, frozenset()) for target in targets))
        self._remember_baseline(cells)

    def _calculation_roots(self, edited):
        edited = set(edited)
        roots = edited | set(self.dynamic_targets)
        roots.update(anchor for anchor, extent in self.spill_cells.items() if edited & extent)
        return roots

    def _remember_baseline(self, coordinates) -> None:
        for coordinate in coordinates:
            if self._is_baseline_output(coordinate) and coordinate not in self.baseline_values:
                self.baseline_values[coordinate] = self.get(*coordinate)

    def _replace_spill(self, anchor, extent) -> None:
        previous = self.spill_cells.pop(anchor, frozenset())
        for coordinate in previous:
            remaining = self.spill_counts[coordinate] - 1
            if remaining:
                self.spill_counts[coordinate] = remaining
            else:
                self.spill_counts.pop(coordinate)
        if extent:
            self.spill_cells[anchor] = extent
            self.spill_counts.update(extent)

    def _is_baseline_output(self, coordinate) -> bool:
        return coordinate in self.formula_anchors or coordinate in self.baseline_spill_cells

    def _update_overlay(self, coordinate, value) -> None:
        self.overlay_before.setdefault(coordinate, self.overlay_values.get(coordinate, ABSENT))
        if value is ABSENT:
            if self._is_baseline_output(coordinate):
                self.overlay_values[coordinate] = None
            else:
                self.overlay_values.pop(coordinate, None)
        elif not self._is_baseline_output(coordinate):
            self.overlay_values[coordinate] = value
        else:
            if coordinate not in self.baseline_values:
                raise EngineError("formula baseline value is unavailable")
            if exactly_equal(self.baseline_values[coordinate], value):
                self.overlay_values.pop(coordinate, None)
            else:
                self.overlay_values[coordinate] = value
        self._record_overlay_delta(coordinate)

    def _record_overlay_delta(self, coordinate) -> None:
        before = self.overlay_before[coordinate]
        after = self.overlay_values.get(coordinate, ABSENT)
        if not _output_changed(before, after):
            self.overlay_upserts.pop(coordinate, None)
            self.overlay_removals.discard(coordinate)
        elif after is ABSENT:
            self.overlay_upserts.pop(coordinate, None)
            self.overlay_removals.add(coordinate)
        else:
            self.overlay_upserts[coordinate] = after
            self.overlay_removals.discard(coordinate)

    def _reset_overlay_delta(self) -> None:
        self.overlay_before.clear()
        self.overlay_upserts.clear()
        self.overlay_removals.clear()

    def result_snapshot(self) -> list[tuple[int, int, int, object]]:
        return [(*coordinate, self.get(*coordinate)) for coordinate in sorted(self.formula_cells)]

    def overlay_snapshot(self) -> list[tuple[int, int, int, object]]:
        return [(*coordinate, value) for coordinate, value in sorted(self.overlay_values.items())]

    def overlay_delta(self):
        upserts = [(*coordinate, value)
                   for coordinate, value in sorted(self.overlay_upserts.items())]
        return upserts, sorted(self.overlay_removals)

    def get(self, sheet: int, row: int, col: int):
        coordinate = (sheet, row, col)
        if coordinate in self.stale:
            self._resolve_stale(coordinate)
        try:
            if coordinate in self.cached_values:
                return _json_value(self.cached_values[coordinate], self.date1904)
            value = self.workbook.get_value(self.workbook.sheet_names[sheet], row + 1, col + 1)
            return _json_value(value, self.date1904)
        except Exception as exc:
            raise EngineError("template calculation engine failed to read a cell") from exc

    def _resolve_stale(self, coordinate):
        """Calculate a cell the last edit left stale before answering for it.

        Excel's stored cache still describes the pre-edit state, so the engine
        must produce the value rather than read the old one back. The mark is
        cleared first so the publish path can read the cell it is capturing.
        """
        self.stale = self.stale - {coordinate}
        try:
            return self.evaluate_dirty({coordinate})
        except BaseException:
            self.stale = self.stale | {coordinate}
            raise

    def _engine_value(self, coordinate):
        sheet, row, col = coordinate
        return _json_value(
            self.workbook.get_value(self.workbook.sheet_names[sheet], row + 1, col + 1),
            self.date1904)

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


def _output_changed(before, after):
    if before is ABSENT or after is ABSENT:
        return before is not after
    return not exactly_equal(before, after)
