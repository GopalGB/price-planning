"""Spawn-safe calculation worker entry point."""

import os
from pathlib import Path

from .formualizer import EngineError, FormualizerAdapter
from ..values import exactly_equal


MISSING = object()


class WarmCopy:
    def __init__(self, copy_id: str):
        self.copy_id = copy_id
        self.engine = None
        self.source = None
        self.baseline = {}
        self.overrides = {}

    def apply(self, revision: int, template_path: str, changes: list[tuple]) -> dict:
        self._load(template_path)
        desired = {(sheet, row, col): value for sheet, row, col, value in changes}
        updates = []
        for coordinate in sorted(self.overrides.keys() | desired.keys()):
            if coordinate not in self.baseline:
                self.baseline[coordinate] = self.engine.get(*coordinate)
            before = self.overrides.get(coordinate, MISSING)
            after = desired.get(coordinate, MISSING)
            if _different(before, after):
                value = self.baseline[coordinate] if after is MISSING else after
                updates.append((*coordinate, value))
        calculated = self.engine.apply(updates)
        self.overrides = desired
        return {"copy_id": self.copy_id, "revision": revision, "pid": os.getpid(),
                "changes": calculated, "results": self.engine.result_snapshot()}

    def _load(self, template_path: str) -> None:
        if self.engine is None:
            self.source = template_path
            self.engine = FormualizerAdapter.load(Path(template_path))
        elif template_path != self.source:
            raise EngineError("copy template changed while its worker was active")


def _different(before, after):
    if before is MISSING or after is MISSING:
        return before is not after
    return not exactly_equal(before, after)


def worker_loop(connection, copy_id: str) -> None:
    """Serve one copy serially inside its dedicated spawned process."""
    copy = WarmCopy(copy_id)
    try:
        while True:
            request = connection.recv()
            if request is None:
                return
            revision, template_path, changes = request
            try:
                connection.send((True, copy.apply(revision, template_path, changes)))
            except Exception as exc:
                copy = WarmCopy(copy_id)
                connection.send((False, type(exc).__name__))
    finally:
        connection.close()
