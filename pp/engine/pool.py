"""Copy-owned spawned calculation workers with bounded LRU eviction."""

from __future__ import annotations

from collections import OrderedDict
import multiprocessing
from pathlib import Path
import threading

from .formualizer import EngineError
from .worker import worker_loop


class _CopyWorker:
    def __init__(self, context, copy_id: str):
        parent, child = context.Pipe()
        self.connection = parent
        self.process = context.Process(target=worker_loop, args=(child, copy_id), daemon=True)
        self.process.start()
        child.close()
        self.busy = False

    def calculate(self, revision: int, template_path: Path, changes: list[tuple]) -> dict:
        if not self.process.is_alive():
            raise EngineError("calculation worker is unavailable")
        self.connection.send((revision, str(template_path), changes))
        success, payload = self.connection.recv()
        if not success:
            raise EngineError(f"calculation worker failed: {payload}")
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
    def __init__(self, max_workers: int):
        if max_workers < 1:
            raise ValueError("max_workers must be positive")
        self._context = multiprocessing.get_context("spawn")
        self._max_workers = max_workers
        self._workers: OrderedDict[str, _CopyWorker] = OrderedDict()
        self._condition = threading.Condition()
        self._closed = False

    def _acquire(self, copy_id: str) -> _CopyWorker:
        with self._condition:
            while True:
                if self._closed:
                    raise EngineError("calculation pool is closed")
                worker = self._workers.get(copy_id)
                if worker is not None and not worker.busy:
                    worker.busy = True
                    self._workers.move_to_end(copy_id)
                    return worker
                if worker is None and len(self._workers) < self._max_workers:
                    return self._create(copy_id)
                victim = next(((key, item) for key, item in self._workers.items()
                               if not item.busy and key != copy_id), None)
                if worker is None and victim is not None:
                    key, evicted = victim
                    del self._workers[key]
                    evicted.close()
                    return self._create(copy_id)
                self._condition.wait()

    def _create(self, copy_id: str) -> _CopyWorker:
        worker = _CopyWorker(self._context, copy_id)
        worker.busy = True
        self._workers[copy_id] = worker
        return worker

    def apply(self, copy_id: str, revision: int, template_path: Path,
              changes: list[tuple[int, int, int, object]]) -> dict:
        worker = self._acquire(copy_id)
        try:
            return worker.calculate(revision, template_path, changes)
        finally:
            with self._condition:
                worker.busy = False
                self._condition.notify_all()

    def close(self) -> None:
        with self._condition:
            self._closed = True
            workers = list(self._workers.values())
            self._workers.clear()
            self._condition.notify_all()
        for worker in workers:
            worker.close()
