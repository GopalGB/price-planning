"""One copy's calculator: Excel's saved results to start with, recalculation of only what an edit reaches.

Formula results are computed on demand: reading a cell whose result is stale evaluates its family's
function, which reads the cells it needs the same way, so dependencies settle first without an explicit
order. The template (formulas, families, compiled functions, saved results) is shared by every copy; a
copy holds only its own inputs and the results that differ from the template's.
"""

from __future__ import annotations

import sys
import threading

from .compile import Compiler, Unsupported
from .ops import RefArea
from .values import Area

sys.setrecursionlimit(1_000_000)
STACK_BYTES = 512 * 1024 * 1024  # deep dependency chains recurse once per link
AREA_CELLS = 10_000_000  # cached range cells per copy; past this the cache starts over


class Shared:
    """What every copy of one template shares."""

    def __init__(self, template):
        self.template = template
        heights = [1] * len(template.names)
        widths = [1] * len(template.names)
        for sheet, row, col in (*template.formulas, *template.constants):
            heights[sheet] = max(heights[sheet], row + 1)
            widths[sheet] = max(widths[sheet], col + 1)
        self.heights, self.widths = heights, widths
        self.compiler = Compiler(template.names)
        self.functions = {}
        self.unsupported = {}
        self.lock = threading.Lock()

    def function(self, family: int):
        found = self.functions.get(family)
        if found is None and family not in self.unsupported:
            anchor, text, array = self.template.families[family]
            try:
                found = self.compiler.compile(text, anchor, array)
            except Unsupported as exc:
                self.unsupported[family] = str(exc)
                return None
            self.functions[family] = found
        return found


class Book:
    def __init__(self, shared: Shared, inputs: dict | None = None):
        self.shared = shared
        self.template = shared.template
        self.inputs = dict(inputs or {})  # (sheet, row, col) -> typed value the user entered
        self.results = {}  # formula results that differ from the template's
        self.own = {}  # (sheet, row, col) -> compiled formula typed into this copy
        self.stale = set()
        self.areas = {}
        self.area_cells = 0

    # Reading

    def cell(self, sheet: int, row: int, col: int):
        key = (sheet, row, col)
        if key in self.inputs:
            return self.inputs[key]
        own = self.own.get(key)
        if own is not None:
            if key in self.stale:
                self.stale.discard(key)
                self.results[key] = own(self, row, col)
            return self.results.get(key)
        family = self.template.formulas.get(key)
        if family is None:
            return self.template.constants.get(key)
        if key in self.stale:
            return self._evaluate(key, family)
        found = self.results.get(key, _MISSING)
        return self.template.results.get(key) if found is _MISSING else found

    def area(self, sheet: int, top: int, left: int, bottom: int, right: int):
        key = (sheet, top, left, bottom, right)
        found = self.areas.get(key)
        if found is None:
            if bottom < top or right < left:
                return Area.scalar(None)
            if self.area_cells > AREA_CELLS:
                self.areas.clear()
                self.area_cells = 0
            cell = self.cell
            found = RefArea([[cell(sheet, r, c) for c in range(left, right + 1)] for r in range(top, bottom + 1)],
                            sheet, top, left)
            self.areas[key] = found
            self.area_cells += found.size
        return found

    def last_row(self, sheet: int) -> int:
        return self.shared.heights[sheet] - 1

    def last_col(self, sheet: int) -> int:
        return self.shared.widths[sheet] - 1

    def is_formula(self, sheet: int, row: int, col: int) -> bool:
        key = (sheet, row, col)
        return key in self.own or (key in self.template.formulas and key not in self.inputs)

    def _evaluate(self, key, family):
        self.stale.discard(key)  # a cycle reads the old value instead of recursing forever
        function = self.shared.function(family)
        if function is None:
            return self.template.results.get(key)
        value = function(self, key[1], key[2])
        if value == self.template.results.get(key, _MISSING) and type(value) is type(self.template.results.get(key)):
            self.results.pop(key, None)
        else:
            self.results[key] = value
        return value

    # Changing

    def recalculate(self, cells) -> dict:
        """Recompute the given formula cells (and whatever they read that is stale); returns new values."""
        self.stale.update(cells)
        self.areas.clear()
        self.area_cells = 0
        return self.values(cells)

    def values(self, cells) -> dict:
        """Current values of the given cells, settling whatever stale formulas they read."""
        out = {}
        error = []

        def run():
            try:
                for key in list(cells):
                    out[key] = self.cell(*key)
            except BaseException as exc:  # surfaced on the calling thread
                error.append(exc)

        threading.stack_size(STACK_BYTES)
        worker = threading.Thread(target=run)
        worker.start()
        worker.join()
        if error:
            raise error[0]
        return out


_MISSING = object()
