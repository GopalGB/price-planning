"""Copy-owned spawned calculation workers with bounded LRU eviction."""

from __future__ import annotations

from collections import OrderedDict
from dataclasses import dataclass
import multiprocessing
from pathlib import Path
import threading

from .formualizer import EngineError
from .worker import worker_loop


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


class _CopyWorker:
    def __init__(self, context, copy_id: str):
        parent, child = context.Pipe()
        self.connection = parent
        self.process = context.Process(target=worker_loop, args=(child, copy_id), daemon=True)
        self.process.start()
        child.close()
        self.busy = False
        self.resident_bytes = 0
        self.accounted_bytes = 0

    def calculate(self, request: CalculationRequest) -> dict:
        if not self.process.is_alive():
            raise _WorkerUnavailable("calculation worker is unavailable")
        try:
            self.connection.send((request.revision, str(request.template_path), request.changes,
                                  request.base_revision))
            success, payload = self.connection.recv()
        except (BrokenPipeError, EOFError, OSError) as exc:
            raise _WorkerUnavailable("calculation worker transport failed") from exc
        if not success:
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
    def __init__(self, max_workers: int, memory_budget_mb: int | None = None):
        if max_workers < 1:
            raise ValueError("max_workers must be positive")
        if memory_budget_mb is not None and memory_budget_mb < 1:
            raise ValueError("memory_budget_mb must be positive")
        self._context = multiprocessing.get_context("spawn")
        self._max_workers = max_workers
        self._workers: OrderedDict[str, _CopyWorker] = OrderedDict()
        self._condition = threading.Condition()
        self._closed = False
        self.memory_budget = MemoryBudget(
            limit_bytes=None if memory_budget_mb is None else memory_budget_mb * 1024 * 1024)
        self.protection = Protection()

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
        worker = _CopyWorker(self._context, copy_id)
        worker.busy = True
        self._workers[copy_id] = worker
        return worker

    def apply(self, request: CalculationRequest) -> dict:
        for attempt in range(2):
            worker = self._acquire(request.copy_id)
            try:
                result = worker.calculate(request)
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
