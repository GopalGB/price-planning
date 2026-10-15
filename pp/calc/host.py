"""The Python calculator behind the engine's host interface, so a copy recalculates without Excel."""

from __future__ import annotations

from pathlib import Path

from ..engine.excel import ExcelHost
from ..engine.formualizer import EngineError
from ..engine.native_load import prepared_template
from .compile import Unsupported
from .engine import Book, Shared
from .model import Template
from .values import XlError

_SHARED = {}  # prepared path -> Shared, one per process


def shared_for(prepared: Path) -> Shared:
    found = _SHARED.get(prepared)
    if found is None:
        found = _SHARED[prepared] = Shared(Template.load(prepared))
    return found


def _to_calc(value):
    if isinstance(value, dict):
        raise EngineError("errors cannot be typed into input cells")
    if isinstance(value, int) and not isinstance(value, bool):
        return float(value)
    return value


def _to_json(value):
    return {"error": value.code} if isinstance(value, XlError) else value


class CalcHost(ExcelHost):
    def __init__(self, workbook: Path, sheet_names: list[str], cache_dir: Path | None = None):
        self.shared = shared_for(prepared_template(workbook, cache_dir))
        self.book = Book(self.shared)

    def write(self, sheet, cells):
        for row, col, value in cells:
            key = (sheet, row, col)
            self.book.inputs[key] = _to_calc(value)
            self.book.own.pop(key, None)
            self.book.results.pop(key, None)

    def write_formulas(self, sheet, cells):
        for row, col, text in cells:
            key = (sheet, row, col)
            try:
                function = self.shared.compiler.compile(text, key, False)
            except Unsupported as exc:
                raise EngineError(f"this formula cannot be calculated here yet ({exc})") from exc
            self.book.inputs.pop(key, None)
            self.book.own[key] = function
            self.book.stale.add(key)

    def calculate(self):
        pass  # values settle on demand when read

    def invalidate(self, cells):
        book = self.book
        book.stale.update(key for key in cells if book.is_formula(*key))
        book.areas.clear()
        book.area_cells = 0

    def read(self, sheet, top, left, bottom, right):
        keys = [(sheet, row, col) for row in range(top, bottom + 1) for col in range(left, right + 1)]
        found = self.book.values(keys)
        width = right - left + 1
        values = [_to_json(found[key]) for key in keys]
        return [values[index:index + width] for index in range(0, len(values), width)]

    def read_formulas(self, sheet, top, left, bottom, right):
        texts = self.shared.template.formula_text
        out = []
        for row in range(top, bottom + 1):
            line = []
            for col in range(left, right + 1):
                key = (sheet, row, col)
                text = texts.get(key) if self.book.is_formula(*key) and key not in self.book.own else None
                line.append(None if text is None else text if text.startswith("=") else "=" + text)
            out.append(line)
        return out

    def formula_areas(self, sheet):
        book = self.book
        keys = {key for key in (*self.shared.template.formulas, *book.own) if key[0] == sheet and book.is_formula(*key)}
        return [(row, col, row, col) for _, row, col in sorted(keys)]

    def used_range(self, sheet):
        return (0, 0, self.shared.heights[sheet] - 1, self.shared.widths[sheet] - 1)

    def run_macro(self, name, timeout):
        raise EngineError("template macros are not available in the Python calculator yet")

    def refresh_all(self, timeout):
        raise EngineError("data refresh is not available in the Python calculator yet")

    def pivot_ranges(self):
        return []

    def refresh_pivots(self, timeout):
        return []

    def set_slicer(self, cache, selected):
        raise EngineError("slicers are not available in the Python calculator yet")

    def memory_bytes(self):
        return 0

    def close(self):
        self.book = None
