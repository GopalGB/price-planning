"""Spawn-safe calculation worker entry point."""

import os
from pathlib import Path
from tempfile import TemporaryDirectory

from .formualizer import EngineError, FormualizerAdapter
from ..export.xlsm import export_xlsm
from ..values import exactly_equal


MISSING = object()


class WarmCopy:
    def __init__(self, copy_id: str):
        self.copy_id = copy_id
        self.engine = None
        self.source = None
        self.baseline = {}
        self.overrides = {}
        self.cleared = frozenset()

    def apply(self, revision: int, template_path: str, changes: list[tuple]) -> dict:
        self._load(template_path)
        desired = {(sheet, row, col): value for sheet, row, col, value in changes}
        cleared = frozenset(coordinate for coordinate, value in desired.items() if value is None)
        if cleared != self.cleared:
            self._reload_with_clears(template_path, cleared)
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

    def _reload_with_clears(self, template_path: str, cleared: frozenset) -> None:
        for coordinate in cleared:
            if coordinate not in self.baseline:
                self.baseline[coordinate] = self.engine.get(*coordinate)
        with TemporaryDirectory(prefix="pp-calc-") as directory:
            source = Path(template_path)
            target = Path(directory) / source.name
            export_xlsm(source, [(*coordinate, None) for coordinate in sorted(cleared)], target)
            self.engine = FormualizerAdapter.load(target)
        self.overrides = {coordinate: None for coordinate in cleared}
        self.cleared = cleared


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
