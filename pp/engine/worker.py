"""Spawn-safe calculation worker entry point."""

import os
from pathlib import Path

from .formualizer import FormualizerAdapter


def apply_revision(copy_id: str, revision: int, template_path: str, changes: list[tuple[int, int, int, object]]) -> dict:
    engine = FormualizerAdapter.load(Path(template_path))
    return {"copy_id": copy_id, "revision": revision, "pid": os.getpid(), "changes": engine.apply(changes)}


def worker_loop(connection, copy_id: str) -> None:
    """Serve one copy serially inside its dedicated spawned process."""
    try:
        while True:
            request = connection.recv()
            if request is None:
                return
            revision, template_path, changes = request
            try:
                connection.send((True, apply_revision(copy_id, revision, template_path, changes)))
            except Exception as exc:
                connection.send((False, type(exc).__name__))
    finally:
        connection.close()
