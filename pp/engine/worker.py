"""Spawn-safe calculation worker entry point."""

import json
import logging
import os
import pickle
import shutil
import time
from dataclasses import dataclass
from pathlib import Path
import tempfile
from tempfile import TemporaryDirectory

from .excel import REFRESH_ALL
from .formualizer import EngineError, FormualizerAdapter
from ..export.xlsm import export_xlsm
from ..values import exactly_equal


MISSING = object()
LOG = logging.getLogger(__name__)
RESUME_VERSION = 1
RESUME_IDLE_S = 20  # save a resume point once a copy has had no request for this long
RESUME_MAX_AGE_S = 600  # and at the latest this long after its last unsaved change


@dataclass(frozen=True)
class WorkerBatch:
    changes: list[tuple]
    base_revision: int | None = None
    pivot_state: dict | None = None  # {"slicers": {cache: [items]}, "refreshed": bool}, applied after a fresh load


class WarmCopy:
    def __init__(self, copy_id: str, engine: str = "formualizer", resume: str | Path | None = None):
        self.copy_id = copy_id
        self.kind = engine
        self.engine = None
        self.source = None
        self.baseline = {}
        self.overrides = {}
        self.cleared = frozenset()
        self.last_applied_revision = None
        # Excel copies save a resume point here, so a restart opens the copy as it was instead of
        # replaying every edit on the template (a checkpoint, as databases keep beside their log)
        self.resume = Path(resume) / copy_id if resume and engine == "excel" else None
        self.unsaved_since = None

    def apply(self, revision: int, template_path: str, changes: list[tuple] | WorkerBatch) -> dict:
        batch = changes if isinstance(changes, WorkerBatch) else WorkerBatch(changes)
        started = time.monotonic()
        calculated, reset = self._sync(template_path, batch.changes, batch.pivot_state)
        LOG.info("Calculated copy %s revision %s in %.1fs (%d values changed)", self.copy_id, revision,
                 time.monotonic() - started, len(calculated))
        return self._result(revision, batch, reset, calculated)

    def run_macro(self, revision: int, template_path: str, changes: list[tuple], base_revision: int | None,
                  macro: str, timeout: float = 600, pivot_state: dict | None = None) -> dict:
        """Sync the desired state, run the template macro in Excel and fold what it wrote into the state."""
        if self.kind != "excel":
            raise EngineError("template macros need the Excel engine")
        batch = WorkerBatch(changes, base_revision, pivot_state)
        calculated, reset = self._sync(template_path, batch.changes, pivot_state)
        started = time.monotonic()
        if macro == REFRESH_ALL:
            outcome = self.engine.refresh_data(timeout)
        else:
            outcome = self.engine.run_macro(macro, timeout)
        LOG.info("Ran %s on copy %s in %.1fs", macro, self.copy_id, time.monotonic() - started)
        self._fold(outcome)
        self._changed()
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
        self._changed()
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
        if updates:
            self._changed()
        self.overrides = desired
        self.cleared = cleared
        if reset and self.kind == "excel" and _has_pivot_state(pivot_state):
            calculated = calculated + self.engine.restore_pivots(pivot_state)["changes"]
            self._changed()
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

    def _changed(self) -> None:
        if self.resume is not None and self.unsaved_since is None:
            self.unsaved_since = time.monotonic()

    def resume_due(self) -> float | None:
        """Seconds to wait for another request before saving a resume point, or None if nothing is unsaved."""
        if self.unsaved_since is None or self.engine is None:
            return None
        return max(0.0, min(RESUME_IDLE_S, self.unsaved_since + RESUME_MAX_AGE_S - time.monotonic()))

    def save_resume_point(self) -> None:
        """Save the copy as Excel holds it and the state that goes with it, then point the manifest at both.

        The manifest is replaced last and atomically, so a crash mid-save leaves the previous point usable.
        """
        save = getattr(getattr(self.engine, "host", None), "save_copy", None)
        if self.resume is None or save is None or self.unsaved_since is None:
            return
        self.unsaved_since = None
        point = self.resume / f"{time.time_ns()}"
        point.mkdir(parents=True)
        try:
            save(point / Path(self.source).name)
            state = {"engine": self.engine.checkpoint_state(), "baseline": self.baseline,
                     "overrides": self.overrides, "cleared": self.cleared}
            with open(point / "state.pickle", "wb") as handle:
                pickle.dump(state, handle, protocol=pickle.HIGHEST_PROTOCOL)
            manifest = {"version": RESUME_VERSION, "template": _template_stamp(self.source), "point": point.name}
            pending = self.resume / "manifest.json.tmp"
            pending.write_text(json.dumps(manifest))
            os.replace(pending, self.resume / "manifest.json")
        except Exception:
            shutil.rmtree(point, ignore_errors=True)
            raise
        for entry in self.resume.iterdir():
            if entry.is_dir() and entry != point:
                shutil.rmtree(entry, ignore_errors=True)

    def _resumed_engine(self, template_path: str):
        """The engine reopened from this copy's last resume point, or None to start from the template."""
        if self.resume is None:
            return None
        try:
            manifest = json.loads((self.resume / "manifest.json").read_text())
        except FileNotFoundError:
            return None
        except Exception:
            LOG.exception("Copy %s resume point is unreadable; starting from the template", self.copy_id)
            return None
        if manifest.get("version") != RESUME_VERSION or manifest.get("template") != _template_stamp(template_path):
            return None
        from .excel import ExcelEngine

        point = self.resume / manifest["point"]
        try:
            # written only by this server's workers into its own data folder
            with open(point / "state.pickle", "rb") as handle:
                state = pickle.load(handle)
            started = time.monotonic()
            engine = ExcelEngine.load(Path(template_path), workbook=point / Path(template_path).name,
                                      state=state["engine"])
        except Exception:
            LOG.exception("Copy %s could not resume; starting from the template", self.copy_id)
            return None
        self.baseline, self.overrides, self.cleared = state["baseline"], state["overrides"], state["cleared"]
        LOG.info("Copy %s resumed from its saved point in %.1fs", self.copy_id, time.monotonic() - started)
        return engine

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
                self.engine = self._resumed_engine(template_path) or ExcelEngine.load(Path(template_path))
            elif self.kind == "python":
                from ..calc.host import CalcHost
                from .excel import ExcelEngine
                self.engine = ExcelEngine.load(Path(template_path), host_factory=CalcHost)
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


def _template_stamp(path) -> list:
    """Which template a resume point was made from: a replaced template file starts copies afresh."""
    found = Path(path).stat()
    return [str(path), found.st_size, found.st_mtime_ns]


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


def worker_loop(connection, copy_id: str, engine: str = "formualizer", spool: str | None = None,
                resume: str | None = None) -> None:
    """Serve one copy serially inside its dedicated spawned process."""
    if not logging.getLogger().handlers:  # a spawned process starts without the server's logging
        logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s %(message)s")
    if spool:
        tempfile.tempdir = tempfile.gettempdir()  # Python temp files and the engine cache stay where they are
        os.environ.update(TMPDIR=spool, TMP=spool, TEMP=spool)  # only the engine's own files move
    copy = WarmCopy(copy_id, engine, resume)
    try:
        while True:
            wait = copy.resume_due()
            if wait is not None and not connection.poll(wait):
                _save(copy)
                continue
            request = connection.recv()
            if request is None:  # evicted or shutting down: the pool waits only seconds, too short to save
                return
            try:
                result = _serve(copy, request)
                result["engine_bytes"] = copy.engine_bytes()
                connection.send((True, result))
            except Exception as exc:
                copy.close()
                copy = WarmCopy(copy_id, engine, resume)
                connection.send((False, type(exc).__name__))
    finally:
        copy.close()
        connection.close()


def _save(copy: WarmCopy) -> None:
    try:
        copy.save_resume_point()
    except Exception:
        LOG.exception("Saving copy %s resume point failed", copy.copy_id)
