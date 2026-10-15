"""Copy-owned spawned calculation workers with bounded LRU eviction."""

from __future__ import annotations

from collections import OrderedDict
from dataclasses import dataclass
import multiprocessing
import os
from pathlib import Path
import shutil
import tempfile
import threading

from .formualizer import EngineError, MacroTimeout
from .worker import _has_pivot_state, worker_loop


def spool_root() -> Path:
    return Path(tempfile.gettempdir()) / "pp-engine-spool"


def _alive(pid: int) -> bool:
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    return True


def _sweep_spools(root: Path) -> None:
    """Remove spool folders whose server process is gone (killed before its pool closed)."""
    # shortcut: Windows sweeps nothing, since os.kill(pid, 0) sends Ctrl+C there; add a process lookup if
    # killed servers start filling the office disk
    if os.name == "nt" or not root.is_dir():
        return
    for entry in root.iterdir():
        owner = entry.name.split("-", 1)[0]
        if owner.isdigit() and int(owner) != os.getpid() and not _alive(int(owner)):
            shutil.rmtree(entry, ignore_errors=True)


class _WorkerUnavailable(EngineError):
    """A transient worker transport failure that permits one exact replay."""


@dataclass
class MemoryBudget:
    """Track the resident size of the engines the pool is holding."""

    limit_bytes: int | None = None
    used_bytes: int = 0

    def admit(self, size_bytes: int) -> bool:
        """Add a newly loaded engine, refusing it when the pool is already full."""
        if self.limit_bytes is None:
            self.used_bytes += size_bytes
            return True
        if self.used_bytes + size_bytes > self.limit_bytes:
            return False
        self.used_bytes += size_bytes
        return True

    def resize(self, previous_bytes: int, size_bytes: int) -> None:
        """Re-measure an engine that has already been admitted."""
        self.used_bytes = max(self.used_bytes - previous_bytes + size_bytes, 0)

    def release(self, size_bytes: int) -> None:
        self.used_bytes = max(self.used_bytes - size_bytes, 0)


@dataclass
class Protection:
    """Copies a client has open, which must survive any eviction sweep."""

    copies: set = None

    def __post_init__(self):
        if self.copies is None:
            self.copies = set()

    def protect(self, copy_id: str) -> None:
        self.copies.add(copy_id)

    def release(self, copy_id: str) -> None:
        self.copies.discard(copy_id)

    def can_evict(self, copy_id: str) -> bool:
        return copy_id not in self.copies


def _eviction_candidates(pool, needed_bytes, limit_bytes):
    """Return the idle, unprotected copies to drop, least recently used first.

    Only copies that are neither busy nor held by an editor are offered, so a
    connected client never pays for a reload it did not ask for. An unbounded
    budget never evicts.
    """
    if limit_bytes is None or needed_bytes <= 0:
        return []
    victims, freed = [], 0
    for copy_id in pool.order():
        if pool.is_busy(copy_id) or pool.has_editors(copy_id):
            continue
        victims.append(copy_id)
        freed += pool.resident_bytes(copy_id)
        if freed >= needed_bytes:
            break
    return victims


@dataclass(frozen=True)
class CalculationRequest:
    copy_id: str
    revision: int
    template_path: Path
    changes: list[tuple[int, int, int, object]]
    base_revision: int | None = None
    pivot_state: dict | None = None


@dataclass(frozen=True)
class MacroRequest:
    copy_id: str
    revision: int
    template_path: Path
    changes: list[tuple[int, int, int, object]]
    base_revision: int | None
    macro: str
    timeout: float = 600
    pivot_state: dict | None = None


@dataclass(frozen=True)
class SlicerRequest:
    copy_id: str
    revision: int
    template_path: Path
    changes: list[tuple[int, int, int, object]]
    base_revision: int | None
    cache: str
    selected: list[str]
    pivot_state: dict | None = None


class _CopyWorker:
    def __init__(self, context, copy_id: str, engine: str = "formualizer", spool: Path | None = None,
                 resume: Path | None = None):
        parent, child = context.Pipe()
        self.connection = parent
        self.process = context.Process(target=worker_loop, args=(child, copy_id, engine, spool and str(spool),
                                                                 resume and str(resume)), daemon=True)
        self.process.start()
        child.close()
        self.busy = False
        self.resident_bytes = 0
        self.accounted_bytes = 0

    def calculate(self, request: CalculationRequest) -> dict:
        return self._exchange((request.revision, str(request.template_path), request.changes,
                               request.base_revision, request.pivot_state))

    def macro(self, request: MacroRequest) -> dict:
        return self._exchange(("macro", request.revision, str(request.template_path), request.changes,
                               request.base_revision, request.macro, request.timeout, request.pivot_state))

    def slicer(self, request: SlicerRequest) -> dict:
        return self._exchange(("slicer", request.revision, str(request.template_path), request.changes,
                               request.base_revision, request.cache, request.selected, request.pivot_state))

    def _exchange(self, message: tuple) -> dict:
        if not self.process.is_alive():
            raise _WorkerUnavailable("calculation worker is unavailable")
        try:
            self.connection.send(message)
            success, payload = self.connection.recv()
        except (BrokenPipeError, EOFError, OSError) as exc:
            raise _WorkerUnavailable("calculation worker transport failed") from exc
        if not success:
            if payload == "MacroTimeout":
                raise MacroTimeout("The macro ran too long and was stopped.")
            raise EngineError(f"calculation worker failed: {payload}")
        self.resident_bytes = payload.get("engine_bytes") or 0
        return payload

    def close(self) -> None:
        if self.process.is_alive():
            try:
                self.connection.send(None)
            except (BrokenPipeError, EOFError, OSError):
                pass
        self.connection.close()
        self.process.join(timeout=5)
        if self.process.is_alive():
            self.process.terminate()
            self.process.join(timeout=5)


class EnginePool:
    def __init__(self, max_workers: int, memory_budget_mb: int | None = None,
                 engine: str = "formualizer", resume: Path | None = None):
        if max_workers < 1:
            raise ValueError("max_workers must be positive")
        if memory_budget_mb is not None and memory_budget_mb < 1:
            raise ValueError("memory_budget_mb must be positive")
        self._context = multiprocessing.get_context("spawn")
        self._max_workers = max_workers
        self._engine = engine
        self._resume = resume  # where Excel copies keep resume points across restarts
        self._workers: OrderedDict[str, _CopyWorker] = OrderedDict()
        self._condition = threading.Condition()
        self._closed = False
        self.memory_budget = MemoryBudget(
            limit_bytes=None if memory_budget_mb is None else memory_budget_mb * 1024 * 1024)
        self.protection = Protection()
        # formualizer spools big loads into TMPDIR and leaves them when a worker is killed;
        # workers spool here instead and close() removes the folder
        root = spool_root()
        _sweep_spools(root)
        root.mkdir(parents=True, exist_ok=True)
        self.spool = Path(tempfile.mkdtemp(prefix=f"{os.getpid()}-", dir=root))

    @property
    def kind(self) -> str:
        return self._engine

    def order(self):
        with self._condition:
            return list(self._workers)

    def is_busy(self, copy_id: str) -> bool:
        worker = self._workers.get(copy_id)
        return worker is None or worker.busy

    def has_editors(self, copy_id: str) -> bool:
        return not self.protection.can_evict(copy_id)

    def resident_bytes(self, copy_id: str) -> int:
        worker = self._workers.get(copy_id)
        return worker.resident_bytes if worker is not None else 0

    def protect(self, copy_id: str) -> None:
        """Hold a copy resident for as long as a client has it open."""
        self.protection.protect(copy_id)

    def release_protection(self, copy_id: str) -> None:
        self.protection.release(copy_id)

    def _reclaim(self, needed: int) -> None:
        """Free idle, unprotected engines until the next engine fits the budget."""
        limit = self.memory_budget.limit_bytes
        if limit is None:
            return
        for copy_id in _eviction_candidates(self, self.memory_budget.used_bytes + needed,
                                            limit):
            worker = self._workers.get(copy_id)
            if worker is None:
                continue
            del self._workers[copy_id]
            self.memory_budget.release(worker.accounted_bytes)
            worker.close()

    def _acquire(self, copy_id: str) -> _CopyWorker:
        with self._condition:
            while True:
                if self._closed:
                    raise EngineError("calculation pool is closed")
                worker = self._workers.get(copy_id)
                if worker is not None and not worker.busy and not worker.process.is_alive():
                    del self._workers[copy_id]
                    self.memory_budget.release(worker.accounted_bytes)
                    worker.close()
                    worker = None
                if worker is not None and not worker.busy:
                    worker.busy = True
                    self._workers.move_to_end(copy_id)
                    return worker
                if worker is None and len(self._workers) < self._max_workers:
                    self._reclaim(0)
                    return self._create(copy_id)
                victim = next(((key, item) for key, item in self._workers.items()
                               if not item.busy and key != copy_id), None)
                if worker is None and victim is not None:
                    key, evicted = victim
                    del self._workers[key]
                    self.memory_budget.release(evicted.accounted_bytes)
                    evicted.close()
                    return self._create(copy_id)
                self._condition.wait()

    def _create(self, copy_id: str) -> _CopyWorker:
        worker = _CopyWorker(self._context, copy_id, self._engine, self.spool, self._resume)
        worker.busy = True
        self._workers[copy_id] = worker
        return worker

    def apply(self, request: CalculationRequest) -> dict:
        if not request.changes and not _has_pivot_state(request.pivot_state) and request.copy_id not in self._workers:
            # nothing typed, nothing refreshed, no engine open: the copy is the template, whose saved results
            # need no engine, so a new copy never waits for Excel to open (nor evicts another copy's Excel)
            return {"copy_id": request.copy_id, "revision": request.revision, "changes": [], "results": [],
                    "removals": [], "delta": False, "base_revision": request.base_revision}
        return self._run(request, _CopyWorker.calculate)

    def warm(self, request: CalculationRequest) -> dict:
        """Like apply, but opens the copy's engine even when it holds nothing to calculate."""
        return self._run(request, _CopyWorker.calculate)

    def run_macro(self, request: MacroRequest) -> dict:
        return self._run(request, _CopyWorker.macro)

    def run_slicer(self, request: SlicerRequest) -> dict:
        return self._run(request, _CopyWorker.slicer)

    def _run(self, request, call) -> dict:
        for attempt in range(2):
            worker = self._acquire(request.copy_id)
            try:
                result = call(worker, request)
                self._account(worker)
                return result
            except _WorkerUnavailable:
                self._discard(request.copy_id, worker)
                if attempt:
                    raise
            finally:
                self._release(request.copy_id, worker)
        raise EngineError("calculation worker retry failed")

    def _account(self, worker):
        """Charge the worker for the size it reports, replacing its last reading."""
        with self._condition:
            self.memory_budget.release(worker.accounted_bytes)
            worker.accounted_bytes = 0
            if worker.resident_bytes and self.memory_budget.admit(worker.resident_bytes):
                worker.accounted_bytes = worker.resident_bytes

    def _discard(self, copy_id, worker):
        with self._condition:
            if self._workers.get(copy_id) is worker:
                del self._workers[copy_id]
                self.memory_budget.release(worker.accounted_bytes)
            self._condition.notify_all()
        worker.close()

    def _release(self, copy_id, worker):
        with self._condition:
            if self._workers.get(copy_id) is worker:
                worker.busy = False
            self._condition.notify_all()

    def close(self) -> None:
        with self._condition:
            self._closed = True
            workers = list(self._workers.values())
            self._workers.clear()
            for worker in workers:
                self.memory_budget.release(worker.accounted_bytes)
            self._condition.notify_all()
        for worker in workers:
            worker.close()
        shutil.rmtree(self.spool, ignore_errors=True)
