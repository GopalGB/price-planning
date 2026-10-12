"""Spawn-safe calculation worker entry point."""

import os
from dataclasses import dataclass
from pathlib import Path
from tempfile import TemporaryDirectory

from .excel import REFRESH_ALL
from .formualizer import EngineError, FormualizerAdapter
from ..export.xlsm import export_xlsm
from ..values import exactly_equal


MISSING = object()


@dataclass(frozen=True)
class WorkerBatch:
    changes: list[tuple]
    base_revision: int | None = None
    pivot_state: dict | None = None  # {"slicers": {cache: [items]}, "refreshed": bool}, applied after a fresh load


class WarmCopy:
    def __init__(self, copy_id: str, engine: str = "formualizer"):
        self.copy_id = copy_id
        self.kind = engine
        self.engine = None
        self.source = None
        self.baseline = {}
        self.overrides = {}
        self.cleared = frozenset()
        self.last_applied_revision = None

    def apply(self, revision: int, template_path: str, changes: list[tuple] | WorkerBatch) -> dict:
        batch = changes if isinstance(changes, WorkerBatch) else WorkerBatch(changes)
        calculated, reset = self._sync(template_path, batch.changes, batch.pivot_state)
        return self._result(revision, batch, reset, calculated)

    def run_macro(self, revision: int, template_path: str, changes: list[tuple], base_revision: int | None,
                  macro: str, timeout: float = 600, pivot_state: dict | None = None) -> dict:
        """Sync the desired state, run the template macro in Excel and fold what it wrote into the state."""
        if self.kind != "excel":
            raise EngineError("template macros need the Excel engine")
        batch = WorkerBatch(changes, base_revision, pivot_state)
        calculated, reset = self._sync(template_path, batch.changes, pivot_state)
        if macro == REFRESH_ALL:
            outcome = self.engine.refresh_data(timeout)
        else:
            outcome = self.engine.run_macro(macro, timeout)
        self._fold(outcome)
        result = self._result(revision, batch, reset, calculated + outcome["changes"])
        return {**result, "messages": outcome["messages"], "contents": outcome["contents"]}

    def run_slicer(self, revision: int, template_path: str, changes: list[tuple], base_revision: int | None,
                   cache: str, selected: list[str], pivot_state: dict | None = None) -> dict:
        """Select the items of one Excel slicer; the pivot output it changes is recorded like a macro's."""
        if self.kind != "excel":
            raise EngineError("slicers need the Excel engine")
        batch = WorkerBatch(changes, base_revision, pivot_state)
        calculated, reset = self._sync(template_path, batch.changes, pivot_state)
        outcome = self.engine.set_slicer(cache, selected)
        self._fold(outcome)
        result = self._result(revision, batch, reset, calculated + outcome["changes"])
        return {**result, "messages": outcome["messages"], "contents": outcome["contents"]}

    def _fold(self, outcome: dict) -> None:
        for coordinate, value in outcome["before"].items():
            if coordinate not in self.baseline:
                self.baseline[coordinate] = value
        for sheet, row, col, value in outcome["contents"]:
            self.overrides[(sheet, row, col)] = value
            if value is None:
                self.cleared = self.cleared | {(sheet, row, col)}

    def _sync(self, template_path: str, changes: list[tuple], pivot_state: dict | None = None) -> tuple[list, bool]:
        reset = self._load(template_path)
        desired = {(sheet, row, col): value for sheet, row, col, value in changes}
        cleared = frozenset(coordinate for coordinate, value in desired.items() if value is None)
        restored_blanks = frozenset(
            coordinate for coordinate in self.overrides.keys() - desired.keys()
            if self.baseline.get(coordinate, MISSING) is None
        )
        reload_clears = cleared | restored_blanks
        if not getattr(self.engine, "handles_clears", False) and reload_clears != self.cleared:
            self._reload_with_clears(template_path, reload_clears)
            self.overrides = {coordinate: None for coordinate in cleared}
            reset = True
        updates = []
        for coordinate in sorted(self.overrides.keys() | desired.keys()):
            if coordinate not in self.baseline:
                self.baseline[coordinate] = self._template_value(coordinate)
            before = self.overrides.get(coordinate, MISSING)
            after = desired.get(coordinate, MISSING)
            if _different(before, after):
                if after is MISSING and self.baseline[coordinate] is None:
                    continue
                value = self.baseline[coordinate] if after is MISSING else after
                updates.append((*coordinate, value))
        calculated = self.engine.apply(updates)
        self.overrides = desired
        self.cleared = cleared
        if reset and self.kind == "excel" and _has_pivot_state(pivot_state):
            calculated = calculated + self.engine.restore_pivots(pivot_state)["changes"]
        return calculated, reset

    def _result(self, revision: int, batch: WorkerBatch, reset: bool, changes: list, **extra) -> dict:
        delta = (not reset and batch.base_revision is not None
                 and batch.base_revision == self.last_applied_revision)
        if delta:
            results, removals = self.engine.overlay_delta()
        else:
            results, removals = self.engine.overlay_snapshot(), []
        self.last_applied_revision = revision
        return {"copy_id": self.copy_id, "revision": revision, "pid": os.getpid(),
                "changes": changes, "results": results, "removals": removals,
                "delta": delta, "base_revision": batch.base_revision, **extra}

    def close(self) -> None:
        """Release the engine; an Excel engine quits its private Excel."""
        engine, self.engine = self.engine, None
        if engine is not None and hasattr(engine, "close"):
            try:
                engine.close()
            except Exception:
                pass

    def engine_bytes(self) -> int:
        external = getattr(self.engine, "memory_bytes", None)
        return _engine_bytes() + (external() if external else 0)

    def _template_value(self, coordinate):
        formula_cells = getattr(self.engine, "formula_cells", ())
        formula_anchors = getattr(self.engine, "formula_anchors", ())
        if coordinate in formula_cells and coordinate not in formula_anchors:
            return None
        return self.engine.get(*coordinate)

    def _load(self, template_path: str) -> bool:
        if self.engine is None:
            self.source = template_path
            if self.kind == "excel":
                from .excel import ExcelEngine
                self.engine = ExcelEngine.load(Path(template_path))
            else:
                self.engine = FormualizerAdapter.load(Path(template_path))
            return True
        elif template_path != self.source:
            raise EngineError("copy template changed while its worker was active")
        return False

    def _reload_with_clears(self, template_path: str, cleared: frozenset) -> None:
        baseline = self.engine
        changed_clears = set(self.cleared | cleared)
        if changed_clears:
            baseline.remember_baseline_for_roots(changed_clears)
        for coordinate in cleared:
            if coordinate not in self.baseline:
                self.baseline[coordinate] = self._template_value(coordinate)
        with TemporaryDirectory(prefix="pp-calc-") as directory:
            source = Path(template_path)
            target = Path(directory) / source.name
            export_xlsm(source, [(*coordinate, None) for coordinate in sorted(cleared)], target)
            self.engine = FormualizerAdapter.load(target)
            changed_spills = self.engine.adopt_baseline(baseline)
            roots = set(cleared) | set(changed_spills)
            if roots:
                self.engine.recalculate(roots)


def _has_pivot_state(state) -> bool:
    return bool(state) and (bool(state.get("refreshed")) or bool(state.get("slicers")))


def _different(before, after):
    if before is MISSING or after is MISSING:
        return before is not after
    return not exactly_equal(before, after)


def _engine_bytes():
    """Peak size of this process, an upper bound the pool can budget against.

    Peak rather than resident, because the budget must not be relieved while
    the engine is still holding the memory it reached at its high-water mark.
    """
    import sys

    if os.name == "nt":  # no `resource` module on Windows
        import ctypes
        from ctypes import wintypes

        class Counters(ctypes.Structure):
            _fields_ = [("cb", wintypes.DWORD), ("PageFaultCount", wintypes.DWORD)] + [
                (name, ctypes.c_size_t) for name in (
                    "PeakWorkingSetSize", "WorkingSetSize", "QuotaPeakPagedPoolUsage", "QuotaPagedPoolUsage",
                    "QuotaPeakNonPagedPoolUsage", "QuotaNonPagedPoolUsage", "PagefileUsage", "PeakPagefileUsage")]

        counters = Counters(cb=ctypes.sizeof(Counters))
        kernel32 = ctypes.WinDLL("kernel32")
        kernel32.GetCurrentProcess.restype = wintypes.HANDLE
        ctypes.WinDLL("psapi").GetProcessMemoryInfo(
            kernel32.GetCurrentProcess(), ctypes.byref(counters), counters.cb)
        return int(counters.PeakWorkingSetSize)
    import resource

    peak = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    return int(peak if sys.platform == "darwin" else peak * 1024)


def _serve(copy: WarmCopy, request: tuple) -> dict:
    """Requests: (revision, path, changes[, base_revision[, pivot_state]]), ("macro", revision, path, changes,
    base, macro, timeout[, pivot_state]) or ("slicer", revision, path, changes, base, cache, selected[, pivot_state])."""
    if request[0] == "macro":
        _, revision, template_path, changes, base_revision, macro, timeout, *rest = request
        return copy.run_macro(revision, template_path, changes, base_revision, macro, timeout,
                              rest[0] if rest else None)
    if request[0] == "slicer":
        _, revision, template_path, changes, base_revision, cache, selected, *rest = request
        return copy.run_slicer(revision, template_path, changes, base_revision, cache, selected,
                               rest[0] if rest else None)
    if len(request) == 3:
        revision, template_path, changes = request
        return copy.apply(revision, template_path, WorkerBatch(changes))
    revision, template_path, changes, base_revision, *rest = request
    return copy.apply(revision, template_path, WorkerBatch(changes, base_revision, rest[0] if rest else None))


def worker_loop(connection, copy_id: str, engine: str = "formualizer") -> None:
    """Serve one copy serially inside its dedicated spawned process."""
    copy = WarmCopy(copy_id, engine)
    try:
        while True:
            request = connection.recv()
            if request is None:
                return
            try:
                result = _serve(copy, request)
                result["engine_bytes"] = copy.engine_bytes()
                connection.send((True, result))
            except Exception as exc:
                copy.close()
                copy = WarmCopy(copy_id, engine)
                connection.send((False, type(exc).__name__))
    finally:
        copy.close()
        connection.close()
