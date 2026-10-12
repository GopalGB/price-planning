"""Spawn-safe calculation worker entry point."""

import os
import resource
from dataclasses import dataclass
from pathlib import Path
from tempfile import TemporaryDirectory

from .formualizer import EngineError, FormualizerAdapter
from ..export.xlsm import export_xlsm
from ..values import exactly_equal


MISSING = object()


@dataclass(frozen=True)
class WorkerBatch:
    changes: list[tuple]
    base_revision: int | None = None


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
        reset = self._load(template_path)
        changes = batch.changes
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
        delta = (not reset and batch.base_revision is not None
                 and batch.base_revision == self.last_applied_revision)
        if delta:
            results, removals = self.engine.overlay_delta()
        else:
            results, removals = self.engine.overlay_snapshot(), []
        self.last_applied_revision = revision
        return {"copy_id": self.copy_id, "revision": revision, "pid": os.getpid(),
                "changes": calculated, "results": results, "removals": removals,
                "delta": delta, "base_revision": batch.base_revision}

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

    peak = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    return int(peak if sys.platform == "darwin" else peak * 1024)


def worker_loop(connection, copy_id: str, engine: str = "formualizer") -> None:
    """Serve one copy serially inside its dedicated spawned process."""
    copy = WarmCopy(copy_id, engine)
    try:
        while True:
            request = connection.recv()
            if request is None:
                return
            if len(request) == 3:
                revision, template_path, changes = request
                batch = WorkerBatch(changes)
            else:
                revision, template_path, changes, base_revision = request
                batch = WorkerBatch(changes, base_revision)
            try:
                result = copy.apply(revision, template_path, batch)
                result["engine_bytes"] = copy.engine_bytes()
                connection.send((True, result))
            except Exception as exc:
                copy.close()
                copy = WarmCopy(copy_id, engine)
                connection.send((False, type(exc).__name__))
    finally:
        copy.close()
        connection.close()
