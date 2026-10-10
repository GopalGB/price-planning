"""Read-only access to imported template snapshots."""

from __future__ import annotations

from functools import lru_cache
import json
from pathlib import Path


class TemplateSnapshot:
    def __init__(self, path: str | Path):
        self.path = Path(path)
        self.meta = _read_json(self.path / "meta.json")
        self.styles = _read_json(self.path / "styles.json")
        self.input_mask = frozenset(_read_json(self.path / "input_mask.json"))
        self.hash = self.meta["hash"]
        self.name = self.meta["name"]
        self.sheets = self.meta["sheets"]

    @lru_cache(maxsize=64)
    def block(self, sheet: int, block: int) -> dict:
        path = self.path / "sheets" / str(sheet) / f"{block}.json"
        return _read_json(path) if path.exists() else {}

    def cell(self, sheet: int, row: int, col: int) -> dict:
        return self.block(sheet, row // 256).get(str(row % 256), {}).get(
            str(col), {"v": None, "s": 0, "k": "c"}
        )

    def is_input(self, sheet: int, row: int, col: int) -> bool:
        return f"{sheet}:{row}:{col}" in self.input_mask

    def validation(self, sheet: int, row: int, col: int) -> dict | None:
        for validation in self.sheets[sheet].get("validations", []):
            if any(r1 <= row <= r2 and c1 <= col <= c2
                   for r1, c1, r2, c2 in validation.get("ranges", [])):
                return validation
        return None


def _read_json(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))
