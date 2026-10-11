"""Bounded conditional-format evaluation over current copy values."""

from __future__ import annotations

from collections import OrderedDict
from dataclasses import dataclass, field
import json
from pathlib import Path
import tempfile

from openpyxl import Workbook
from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import TokenizerError
from openpyxl.formula.translate import Translator, TranslatorError
from openpyxl.utils.cell import get_column_letter, range_boundaries, range_to_tuple
from openpyxl.workbook.defined_name import DefinedName

from pp.engine.formualizer import EngineError, FormualizerEngine


BLOCK_ROWS = 256
MAX_VALUE_BLOCKS = 2
FORMULA_BATCH_SIZE = 256
MAX_FORMULA_CELLS = 10_000
MAX_DUPLICATE_SCAN_CELLS = 100_000
MAX_OVERRIDE_RECORDS = 100_000
MAX_FORMULA_TARGETS = MAX_OVERRIDE_RECORDS
MAX_WARNINGS = 32
MAX_EXCEL_ROWS = 1_048_576
MAX_EXCEL_COLS = 16_384
WARNING_MESSAGE = "Some conditional formatting could not be applied."
ALIASABLE_FUNCTIONS = frozenset({"EXACT", "FALSE", "TRUE"})
_MISSING = object()
OPERATORS = {
    "equal": "=", "notEqual": "<>", "greaterThan": ">", "lessThan": "<",
    "greaterThanOrEqual": ">=", "lessThanOrEqual": "<=",
}


class ConditionalFormatError(RuntimeError):
    """A valid-looking rule cannot be evaluated safely within bounded resources."""

    def __init__(self, code: str, detail: str):
        self.code = code
        super().__init__(detail)


@dataclass(frozen=True)
class ConditionalFormatResult:
    styles: dict[tuple[int, int], dict]
    warnings: list[dict]


@dataclass(frozen=True)
class FormulaTask:
    rule_key: int
    coordinate: tuple[int, int]
    formula: str
    aliases: tuple[tuple[int, int], ...] = ()

    @property
    def coordinates(self):
        return (self.coordinate, *self.aliases)


@dataclass
class FormulaEvaluation:
    results: dict[int, list[tuple[tuple[int, int], object]]] = field(default_factory=dict)
    failures: dict[int, str] = field(default_factory=dict)


class CurrentCopyValues:
    """Read current results, inputs, then template values in two cached blocks."""

    def __init__(self, connection, snapshot, copy_id: str):
        self.connection, self.snapshot, self.copy_id = connection, snapshot, copy_id
        self.blocks: OrderedDict[tuple[int, int], tuple[dict, dict, dict]] = OrderedDict()
        self.extents: dict[int, tuple[int, int]] = {}

    def get(self, sheet: int, row: int, col: int):
        results, overrides, template = self._block(sheet, row // BLOCK_ROWS)
        coordinate = row, col
        if coordinate in results:
            return results[coordinate]
        if coordinate in overrides:
            return overrides[coordinate]
        record = template.get(str(row % BLOCK_ROWS), {}).get(str(col))
        return None if record is None or record["k"] == "f" else record["v"]

    def _block(self, sheet: int, block: int) -> tuple[dict, dict, dict]:
        key = sheet, block
        if key in self.blocks:
            self.blocks.move_to_end(key)
            return self.blocks[key]
        start = block * BLOCK_ROWS
        results = self._values("pp_results", sheet, start)
        overrides = self._values("pp_overrides", sheet, start)
        path = self.snapshot.path / "sheets" / str(sheet) / f"{block}.json"
        template = json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}
        self.blocks[key] = results, overrides, template
        while len(self.blocks) > MAX_VALUE_BLOCKS:
            self.blocks.popitem(last=False)
        return self.blocks[key]

    def _values(self, table: str, sheet: int, start: int) -> dict:
        rows = self.connection.execute(
            f"SELECT row,col,value_json FROM {table} WHERE copy_id=? AND sheet=? AND row>=? AND row<?",
            (self.copy_id, sheet, start, start + BLOCK_ROWS),
        ).fetchall()
        return {(item["row"], item["col"]): json.loads(item["value_json"]) for item in rows}

    def extent(self, sheet: int) -> tuple[int, int]:
        if sheet not in self.extents:
            row = self.connection.execute(
                """SELECT MAX(row) AS max_row, MAX(col) AS max_col FROM (
                     SELECT row,col FROM pp_results WHERE copy_id=? AND sheet=?
                     UNION ALL
                     SELECT row,col FROM pp_overrides WHERE copy_id=? AND sheet=?)""",
                (self.copy_id, sheet, self.copy_id, sheet),
            ).fetchone()
            info = self.snapshot.sheets[sheet]
            max_row = row["max_row"] if row["max_row"] is not None else -1
            max_col = row["max_col"] if row["max_col"] is not None else -1
            self.extents[sheet] = (
                max(info["rows"], max_row + 1),
                max(info["cols"], max_col + 1),
            )
        return self.extents[sheet]


@dataclass
class EvaluationState:
    snapshot: object
    values: CurrentCopyValues
    query: object
    output: dict[tuple[int, int], dict] = field(default_factory=dict)
    stopped: set[tuple[int, int]] = field(default_factory=set)
    warnings: list[dict] = field(default_factory=list)
    rule_before: dict[tuple[int, int], object] = field(default_factory=dict)
    rule_stops: set[tuple[int, int]] = field(default_factory=set)

    def run(self) -> ConditionalFormatResult:
        rules = sorted(self.snapshot.sheets[self.query.sheet].get("conditional_formats", []),
                       key=lambda rule: rule["priority"])
        indexed = list(enumerate(rules))
        formulas = self._formula_evaluation(indexed)
        for key, rule in indexed:
            if key not in formulas.failures:
                self._apply_atomic(rule, formulas.results.get(key))
        return ConditionalFormatResult(self.output, self.warnings)

    def _formula_evaluation(self, indexed) -> FormulaEvaluation:
        tasks, failures, total = [], {}, 0
        extent = self.values.extent(self.query.sheet)
        for key, rule in indexed:
            if rule["type"] not in {"expression", "cellIs"}:
                continue
            visible = _visible_bounds(rule["range"], self.query, extent)
            count = 0 if visible is None else _area(visible)
            if total + count > MAX_FORMULA_TARGETS:
                failures[key] = "formula_target_limit"
                self._warn(rule, failures[key])
                continue
            total += count
            try:
                tasks.extend(_formula_tasks(key, rule, _coordinates(visible) if visible else ()))
            except ConditionalFormatError as exc:
                failures[key] = exc.code
                self._warn(rule, exc.code)
        evaluated = PublicFormulaEvaluator(
            self.snapshot, self.values, self.query.sheet,
        ).evaluate(tasks)
        evaluated.failures.update(failures)
        for key, code in evaluated.failures.items():
            if key not in failures:
                self._warn(indexed[key][1], code)
            evaluated.results.pop(key, None)
        return evaluated

    def _apply_atomic(self, rule, formula_results) -> None:
        self.rule_before, self.rule_stops = {}, set()
        try:
            self.apply_rule(rule, formula_results)
        except ConditionalFormatError as exc:
            self._rollback()
            self._warn(rule, exc.code)

    def apply_rule(self, rule, formula_results=None) -> None:
        kind = rule["type"]
        if kind in {"expression", "cellIs"}:
            for coordinate, result in formula_results or ():
                if _formula_truth(result):
                    self.apply_style(rule, coordinate)
            return
        if kind == "duplicateValues":
            matches = _duplicate_matches(rule, self.values, self.query)
            for coordinate in matches:
                self.apply_style(rule, coordinate)
            return
        extent = self.values.extent(self.query.sheet)
        for coordinate in _targets(rule["range"], self.query, extent):
            if _direct_match(rule, self.values.get(self.query.sheet, *coordinate)):
                self.apply_style(rule, coordinate)

    def apply_style(self, rule, coordinate) -> None:
        if coordinate in self.stopped:
            return
        if coordinate not in self.output and len(self.output) >= MAX_OVERRIDE_RECORDS:
            raise ConditionalFormatError(
                "override_limit",
                f"conditional formatting exceeds {MAX_OVERRIDE_RECORDS} row override records",
            )
        if coordinate not in self.rule_before:
            self.rule_before[coordinate] = self.output.get(coordinate, _MISSING)
        self.output[coordinate] = _merge_dxf(self.output.get(coordinate, {}), rule["dxf"])
        if rule["stop_if_true"] and coordinate not in self.stopped:
            self.stopped.add(coordinate)
            self.rule_stops.add(coordinate)

    def _rollback(self) -> None:
        self.stopped.difference_update(self.rule_stops)
        for coordinate, previous in self.rule_before.items():
            if previous is _MISSING:
                self.output.pop(coordinate, None)
            else:
                self.output[coordinate] = previous

    def _warn(self, rule, code) -> None:
        warning = {"code": code, "message": WARNING_MESSAGE, "rule_id": _rule_id(rule)}
        if len(self.warnings) < MAX_WARNINGS - 1:
            self.warnings.append(warning)
        elif not self.warnings or self.warnings[-1]["code"] != "warnings_truncated":
            self.warnings.append({"code": "warnings_truncated", "message": WARNING_MESSAGE,
                                  "rule_id": "multiple"})


class PublicFormulaEvaluator:
    def __init__(self, snapshot, values, sheet: int):
        self.snapshot, self.values, self.sheet = snapshot, values, sheet

    def evaluate(self, tasks: list[FormulaTask]) -> FormulaEvaluation:
        evaluated = FormulaEvaluation()
        for batch in _chunks(tasks, FORMULA_BATCH_SIZE):
            batch = [task for task in batch if task.rule_key not in evaluated.failures]
            if batch:
                self._evaluate_group(batch, evaluated)
        for key in evaluated.failures:
            evaluated.results.pop(key, None)
        return evaluated

    def _evaluate_group(self, tasks, evaluated) -> None:
        tasks = [task for task in tasks if task.rule_key not in evaluated.failures]
        if not tasks:
            return
        try:
            values = self._evaluate_batch(tasks)
        except ConditionalFormatError as exc:
            groups = _task_groups(tasks)
            if len(groups) > 1:
                for group in groups:
                    self._evaluate_group(group, evaluated)
                return
            if len(tasks) > 1:
                middle = len(tasks) // 2
                for group in (tasks[:middle], tasks[middle:]):
                    self._evaluate_group(group, evaluated)
                return
            evaluated.failures[tasks[0].rule_key] = exc.code
            evaluated.results.pop(tasks[0].rule_key, None)
            return
        for task, value in zip(tasks, values):
            evaluated.results.setdefault(task.rule_key, []).extend(
                (coordinate, value) for coordinate in task.coordinates
            )

    def _evaluate_batch(self, tasks) -> list:
        book = _helper_book(self.snapshot)
        scan = ReferenceScan(self.snapshot, self.values)
        try:
            formulas = [_expand_defined_names(task.formula, self.sheet, self.snapshot)
                        for task in tasks]
            for formula in formulas:
                scan.collect(formula, self.sheet)
        except TokenizerError as exc:
            raise ConditionalFormatError("formula_invalid", str(exc)) from exc
        references = scan.output
        for sheet, row, col in sorted(references):
            _set_value(book.worksheets[sheet].cell(row + 1, col + 1), self.values.get(sheet, row, col))
        outputs = _scratch_coordinates(self.values, self.sheet, references, len(tasks))
        for formula, (_, row, col) in zip(formulas, outputs):
            book.worksheets[self.sheet].cell(row + 1, col + 1, formula)
        values = self._run(book, outputs)
        if any(isinstance(value, dict) and set(value) == {"error"} for value in values):
            raise ConditionalFormatError("formula_engine", "conditional formula returned an error")
        return values

    @staticmethod
    def _run(book, outputs) -> list:
        try:
            with tempfile.TemporaryDirectory() as directory:
                path = Path(directory) / "conditional-format.xlsx"
                book.save(path)
                engine = FormualizerEngine.load(path)
                return [engine.get(*coordinate) for coordinate in outputs]
        except EngineError as exc:
            raise ConditionalFormatError("formula_engine", str(exc)) from exc


def conditional_style_overrides(snapshot, values: CurrentCopyValues, query) -> ConditionalFormatResult:
    """Future calculation callers can supply any bounded current-value reader."""
    return EvaluationState(snapshot, values, query).run()


def _targets(bounds, query, extent):
    visible = _visible_bounds(bounds, query, extent)
    if visible is None:
        return
    yield from _coordinates(visible)


def _coordinates(bounds):
    row1, col1, row2, col2 = bounds
    for row in range(row1, row2 + 1):
        for col in range(col1, col2 + 1):
            yield row, col


def _chunks(items, size):
    batch = []
    for item in items:
        batch.append(item)
        if len(batch) == size:
            yield batch
            batch = []
    if batch:
        yield batch


def _task_groups(tasks):
    grouped = OrderedDict()
    for task in tasks:
        grouped.setdefault(task.rule_key, []).append(task)
    return list(grouped.values())


def _area(bounds):
    row1, col1, row2, col2 = bounds
    return (row2 - row1 + 1) * (col2 - col1 + 1)


def _formula_truth(value):
    return value is True or (type(value) in {int, float} and value != 0)


def _rule_id(rule):
    return str(rule.get("rule_id") or f"cf-{rule.get('priority', 'unknown')}-{rule['type']}")


def _direct_match(rule, value):
    text = None if value is None or isinstance(value, dict) else str(value)
    if rule["type"] == "containsErrors":
        return isinstance(value, dict) and set(value) == {"error"}
    if rule["type"] == "containsText":
        return text is not None and str(rule.get("text") or "").casefold() in text.casefold()
    if rule["type"] == "beginsWith":
        return text is not None and text.casefold().startswith(str(rule.get("text") or "").casefold())
    return False


def _merge_dxf(higher: dict, lower: dict) -> dict:
    merged = {**lower, **higher}
    if "border" in lower or "border" in higher:
        merged["border"] = {**lower.get("border", {}), **higher.get("border", {})}
    return merged


def _duplicate_matches(rule, values, query):
    extent = values.extent(query.sheet)
    visible = _visible_bounds(rule["range"], query, extent)
    if visible is None:
        return
    scanned = _clamped(rule["range"], extent)
    if scanned is None:
        return
    if _area(scanned) > MAX_DUPLICATE_SCAN_CELLS:
        raise ConditionalFormatError(
            "duplicate_scan_limit",
            f"duplicate rule scan exceeds {MAX_DUPLICATE_SCAN_CELLS} cells",
        )
    candidates = {_value_key(values.get(query.sheet, *coordinate))
                  for coordinate in _coordinates(visible)}
    candidates.discard(None)
    counts = dict.fromkeys(candidates, 0)
    for coordinate in _coordinates(scanned):
        key = _value_key(values.get(query.sheet, *coordinate))
        if key in counts:
            counts[key] += 1
    for coordinate in _coordinates(visible):
        key = _value_key(values.get(query.sheet, *coordinate))
        if key is not None and counts[key] > 1:
            yield coordinate


def _clamped(bounds, extent):
    rows, cols = extent
    row1, col1, row2, col2 = bounds
    clipped = max(0, row1), max(0, col1), min(row2, rows - 1), min(col2, cols - 1)
    return None if clipped[0] > clipped[2] or clipped[1] > clipped[3] else clipped


def _visible_bounds(bounds, query, extent):
    clipped = _clamped(bounds, extent)
    if clipped is None:
        return None
    row1, col1, row2, col2 = clipped
    visible = max(row1, query.start), col1, min(row2, query.start + query.count - 1), col2
    return None if visible[0] > visible[2] else visible


def _value_key(value):
    if value is None:
        return None
    if isinstance(value, bool):
        return "boolean", value
    if isinstance(value, (int, float)):
        return "number", float(value)
    if isinstance(value, str):
        return "text", value.casefold()
    if isinstance(value, dict) and set(value) == {"error"}:
        return "error", value["error"]
    return "value", json.dumps(value, sort_keys=True, separators=(",", ":"))


def _formula_tasks(rule_key, rule, coordinates):
    formulas = rule.get("formulas", [])
    origin = f"{get_column_letter(rule['top_left'][1] + 1)}{rule['top_left'][0] + 1}"
    try:
        translators = [Translator("=" + value.lstrip("="), origin=origin) for value in formulas]
        if not translators:
            raise ValueError("conditional rule is missing a formula")
        grouped, aliasable = OrderedDict(), {}
        for coordinate in coordinates:
            formula = _translated_formula(rule, translators, *coordinate)
            if formula not in aliasable:
                aliasable[formula] = _formula_aliasable(formula)
            key = formula, None if aliasable[formula] else coordinate
            grouped.setdefault(key, []).append(coordinate)
        return [FormulaTask(rule_key, targets[0], key[0], tuple(targets[1:]))
                for key, targets in grouped.items()]
    except (TokenizerError, TranslatorError, ValueError, IndexError) as exc:
        raise ConditionalFormatError("formula_invalid", str(exc)) from exc


def _formula_aliasable(formula):
    for token in Tokenizer(formula).items:
        if token.type == "FUNC" and token.subtype == "OPEN":
            name = token.value[:-1].upper().rsplit(".", 1)[-1]
            if name not in ALIASABLE_FUNCTIONS:
                return False
        if token.type == "OPERAND" and token.subtype == "RANGE":
            try:
                range_to_tuple(token.value) if "!" in token.value else range_boundaries(token.value)
            except ValueError:
                return False
    return True


def _translated_formula(rule, translators, row, col):
    target = f"{get_column_letter(col + 1)}{row + 1}"
    if rule["type"] == "expression":
        return _context_formula(translators[0].translate_formula(target), row, col)
    operands = [translator.translate_formula(target)[1:] for translator in translators]
    if rule.get("operator") in OPERATORS:
        return _context_formula(f"={target}{OPERATORS[rule['operator']]}({operands[0]})", row, col)
    if rule.get("operator") in {"between", "notBetween"} and len(operands) == 2:
        body = f"AND({target}>=({operands[0]}),{target}<=({operands[1]}))"
        formula = f"={body}" if rule["operator"] == "between" else f"=NOT({body})"
        return _context_formula(formula, row, col)
    raise ConditionalFormatError("formula_invalid",
                                 f"unsupported cellIs operator {rule.get('operator')}")


def _context_formula(formula: str, row: int, col: int) -> str:
    tokens, output, index = Tokenizer(formula).items, ["="], 0
    while index < len(tokens):
        token = tokens[index]
        name = token.value[:-1].upper() if token.type == "FUNC" and token.subtype == "OPEN" else ""
        empty = index + 1 < len(tokens) and tokens[index + 1].type == "FUNC" and tokens[index + 1].subtype == "CLOSE"
        if name in {"ROW", "COLUMN"} and empty:
            output.append(str(row + 1 if name == "ROW" else col + 1))
            index += 2
            continue
        output.append(token.value)
        index += 1
    return "".join(output)


def _helper_book(snapshot):
    book = Workbook()
    book.active.title = snapshot.sheets[0]["name"]
    for info in snapshot.sheets[1:]:
        book.create_sheet(info["name"])
    for item in snapshot.meta.get("defined_names", []):
        local = item["local_sheet_id"]
        name = DefinedName(item["name"], attr_text=item["formula"])
        name.hidden = item.get("hidden", False)
        target = book.defined_names if local is None else book.worksheets[local].defined_names
        target.add(name)
    return book


@dataclass
class ReferenceScan:
    snapshot: object
    values: CurrentCopyValues
    output: set[tuple[int, int, int]] = field(default_factory=set)
    resolving: set[tuple[int, str]] = field(default_factory=set)

    def collect(self, formula, current_sheet) -> None:
        for token in Tokenizer("=" + formula.lstrip("=")).items:
            if token.type != "OPERAND" or token.subtype != "RANGE":
                continue
            cells = self.range_cells(token.value, current_sheet)
            if cells is not None:
                self.output.update(cells)
                if len(self.output) > MAX_FORMULA_CELLS:
                    raise ConditionalFormatError(
                        "formula_reference_limit",
                        f"conditional formula exceeds {MAX_FORMULA_CELLS} referenced cells per batch",
                    )
                continue
            name = _defined_name(
                token.value, current_sheet, self.snapshot.meta.get("defined_names", []),
            )
            marker = current_sheet, name["name"].casefold()
            if marker in self.resolving:
                raise ConditionalFormatError(
                    "formula_invalid", f"circular conditional defined name {name['name']}",
                )
            self.resolving.add(marker)
            target = name["local_sheet_id"] if name["local_sheet_id"] is not None else current_sheet
            self.collect(name["formula"], target)
            self.resolving.remove(marker)

    def range_cells(self, reference, current_sheet):
        try:
            if "!" in reference:
                name, bounds = range_to_tuple(reference)
                name = name.replace("''", "'")
                sheet = next(index for index, item in enumerate(self.snapshot.sheets)
                             if item["name"].casefold() == name.casefold())
            else:
                bounds, sheet = range_boundaries(reference), current_sheet
        except (ValueError, StopIteration):
            return None
        min_col, min_row, max_col, max_row = bounds
        rows, cols = self.values.extent(sheet)
        min_col, min_row = min_col or 1, min_row or 1
        max_col, max_row = max_col or cols, max_row or rows
        cells = max(0, max_col - min_col + 1) * max(0, max_row - min_row + 1)
        if cells > MAX_FORMULA_CELLS:
            raise ConditionalFormatError(
                "formula_reference_limit",
                f"conditional formula exceeds {MAX_FORMULA_CELLS} referenced cells per batch",
            )
        return {(sheet, row - 1, col - 1)
                for row in range(min_row, max_row + 1) for col in range(min_col, max_col + 1)}


def _defined_name(value, current_sheet, names):
    match = _matching_name(value, current_sheet, names)
    if match is None:
        raise ConditionalFormatError(
            "formula_invalid", f"unsupported conditional formula reference {value}",
        )
    return match


def _matching_name(value, current_sheet, names):
    matching = [item for item in names if item["name"].casefold() == value.casefold()
                and item["local_sheet_id"] in {None, current_sheet}]
    matching.sort(key=lambda item: item["local_sheet_id"] is None)
    return matching[0] if matching else None


def _expand_defined_names(formula, current_sheet, snapshot):
    names = snapshot.meta.get("defined_names", [])

    def expand(text, sheet, resolving):
        output = []
        for token in Tokenizer("=" + text.lstrip("=")).items:
            name = _matching_name(token.value, sheet, names) if (
                token.type == "OPERAND" and token.subtype == "RANGE"
            ) else None
            if name is None:
                output.append(token.value)
                continue
            marker = sheet, name["name"].casefold()
            if marker in resolving:
                raise ConditionalFormatError("formula_invalid", "circular conditional defined name")
            target = name["local_sheet_id"] if name["local_sheet_id"] is not None else sheet
            output.append(f"({expand(name['formula'], target, resolving | {marker})})")
        return "".join(output)

    return "=" + expand(formula, current_sheet, set())


def _scratch_coordinates(values, sheet, references, count):
    rows, cols = values.extent(sheet)
    base_row = max([rows, *(row + 1 for _, row, _ in references)])
    base_col = max([cols, *(col + 1 for _, _, col in references)])
    width = MAX_EXCEL_COLS - base_col
    needed_rows = (count + width - 1) // width if width > 0 else MAX_EXCEL_ROWS + 1
    if base_row + needed_rows > MAX_EXCEL_ROWS or width <= 0:
        raise ConditionalFormatError("formula_scratch_limit",
                                     "conditional formula scratch cells are exhausted")
    return [(sheet, base_row + index // width, base_col + index % width)
            for index in range(count)]


def _set_value(cell, value):
    if isinstance(value, dict) and set(value) == {"error"}:
        cell.value, cell.data_type = value["error"], "e"
    elif isinstance(value, str):
        cell.value, cell.data_type = value, "s"
    else:
        cell.value = value
