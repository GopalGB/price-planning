"""Engine boundary with zero-based workbook coordinates."""

from __future__ import annotations

from pathlib import Path
from typing import Protocol, runtime_checkable


@runtime_checkable
class Engine(Protocol):
    @classmethod
    def load(cls, template_path: str | Path) -> Engine: ...

    def apply(self, changes: list[tuple[int, int, int, object]]) -> list[tuple[int, int, int, object]]: ...

    def get(self, sheet: int, r: int, c: int) -> object: ...
