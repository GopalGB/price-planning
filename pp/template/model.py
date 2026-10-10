"""Small immutable helpers used by the template importer."""

from dataclasses import dataclass


@dataclass(frozen=True)
class Cell:
    value: object
    style: int
    kind: str
    formula: str | None = None
