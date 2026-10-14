"""Copy extras: tabs and calculation columns a user added in Excel, kept with that copy only.

An upload with extras is imported once with the template importer (Excel's saved values), and the
extras are grafted into the template package for downloads and the Excel engine. ``CopySnapshot``
puts the template and the extras behind the snapshot interface, so the grid, edits and exports
read extras like any other cell: extra tabs get sheet indexes after the template's, extra columns
sit right of the template's own columns. Values the app cannot calculate keep Excel's saved results.
"""

from __future__ import annotations

import json
import shutil
import secrets
import threading
from collections import OrderedDict
from dataclasses import dataclass
from datetime import datetime, timezone
from functools import lru_cache
from pathlib import Path

from .errors import AppError
from .workbook_service import TemplateSnapshot, _read_json

EXTRAS_FILE = "extras.json"
GRAFTED = "grafted.xlsm"
# Extra-sheet features the app cannot evaluate per copy render as plain cells
_DROPPED_SHEET_KEYS = ("pivots", "excel_pivots", "slicers", "tables", "filters", "validations")


@dataclass(frozen=True)
class ExtrasPlan:
    sheets: tuple[str, ...]  # extra tab names, in the file's order
    columns: dict  # template sheet name -> first extra column (0-based)
    order: tuple[str, ...]  # every tab of the file, in the file's order

    def empty(self):
        return not self.sheets and not self.columns


def build_extras(upload_path, plan: ExtrasPlan, template: TemplateSnapshot, work_dir) -> Path:
    """Import the file and graft its extras into the template package; returns the extras directory."""
    from .export.extras_graft import GraftPlan, graft_extras
    from .template.importer import import_template

    target = Path(work_dir) / f"extras-{secrets.token_hex(8)}"
    target.mkdir(parents=True)
    try:
        imported = import_template(upload_path, target, "cached")
        _prune(imported, plan)
        graft_extras(template.source_path, upload_path, GraftPlan(plan.sheets, dict(plan.columns)),
                     target / GRAFTED)
        (target / EXTRAS_FILE).write_text(json.dumps({
            "snapshot": str(imported.relative_to(target)), "template": template.hash,
            "sheets": list(plan.sheets), "columns": plan.columns, "order": list(plan.order),
            "template_sheets": len(template.sheets),
            "column_indexes": {str(_sheet_index(template, name)): first for name, first in plan.columns.items()},
        }), encoding="utf-8")
        return target
    except BaseException:
        shutil.rmtree(target, ignore_errors=True)
        raise


def _prune(snapshot: Path, plan: ExtrasPlan) -> None:
    """Keep only what the copy shows: its own tabs, and its own columns of template sheets."""
    meta = _read_json(snapshot / "meta.json")
    tabs = {name.casefold() for name in plan.sheets}
    firsts = {name.casefold(): first for name, first in plan.columns.items()}
    for sheet in meta["sheets"]:
        name, index = sheet["name"].casefold(), str(sheet["index"])
        if name in tabs:
            continue
        for folder in ("sheets", "notes"):
            directory = snapshot / folder / index
            if name in firsts and folder == "sheets" and directory.is_dir():
                for path in directory.glob("*.json"):
                    rows = {row: kept for row, cells in _read_json(path).items()
                            if (kept := {col: cell for col, cell in cells.items() if int(col) >= firsts[name]})}
                    path.write_text(json.dumps(rows, separators=(",", ":")), encoding="utf-8")
            else:
                shutil.rmtree(directory, ignore_errors=True)
    (snapshot / "input_mask.json").write_text("[]", encoding="utf-8")  # inputs come from the template


def _sheet_index(template, name):
    return next(item["index"] for item in template.sheets if item["name"].casefold() == name.casefold())


def attach(connection, copy_id, built: Path | None, data_dir) -> Path | None:
    """Make ``built`` the copy's extras (None removes them); returns the replaced directory to delete
    after the transaction commits. Values the copy held inside old or new extras give way to the file's."""
    old = connection.execute("SELECT path FROM pp_copy_extras WHERE copy_id=?", (copy_id,)).fetchone()
    regions = [_regions(Path(old["path"]))] if old else []
    if built is not None:
        final = Path(data_dir) / "extras" / copy_id / built.name
        final.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(built), final)
        regions.append(_regions(final))
        connection.execute(
            "INSERT INTO pp_copy_extras(copy_id, path, created_at) VALUES(?,?,?) "
            "ON CONFLICT(copy_id) DO UPDATE SET path=excluded.path, created_at=excluded.created_at",
            (copy_id, str(final), datetime.now(timezone.utc).isoformat()),
        )
    elif old:
        connection.execute("DELETE FROM pp_copy_extras WHERE copy_id=?", (copy_id,))
    for first_sheet, columns in regions:
        for table in ("pp_overrides", "pp_results", "pp_copy_baselines"):
            connection.execute(f"DELETE FROM {table} WHERE copy_id=? AND sheet>=?",  # nosec B608 # fixed names
                               (copy_id, first_sheet))
            for sheet, first in columns.items():
                connection.execute(f"DELETE FROM {table} WHERE copy_id=? AND sheet=? AND col>=?",  # nosec B608
                                   (copy_id, sheet, first))
    if old is None:
        return None
    shared = connection.execute("SELECT 1 FROM pp_copy_extras WHERE path=? AND copy_id<>?",
                                (old["path"], copy_id)).fetchone()
    return None if shared else Path(old["path"])  # a fork still reads the folder: keep it


def current_plan(connection, copy_id) -> dict | None:
    row = connection.execute("SELECT path FROM pp_copy_extras WHERE copy_id=?", (copy_id,)).fetchone()
    return _read_json(Path(row["path"]) / EXTRAS_FILE) if row else None


def _regions(directory: Path):
    """(first extra sheet index, {template sheet index: first extra column}) of stored extras."""
    info = _read_json(directory / EXTRAS_FILE)
    return info["template_sheets"], {int(key): value for key, value in info["column_indexes"].items()}


class CopySnapshot(TemplateSnapshot):
    """The template plus one copy's extras, behind the snapshot interface."""

    def __init__(self, template: TemplateSnapshot, directory: Path):  # pylint: disable=super-init-not-called
        self.__dict__.update(template.__dict__)
        info = _read_json(directory / EXTRAS_FILE)
        self.template, self.directory = template, directory
        self.extras = TemplateSnapshot(directory / info["snapshot"])
        self.extras_id = directory.name
        self.grafted_path = directory / GRAFTED
        self._pivot_areas = {}
        count, offset = len(template.sheets), len(template.styles)
        self.style_offset = offset
        # The grid finds styles by id, so the extras' ids move with their indexes
        self.styles = template.styles + [{**style, "id": number + offset}
                                         for number, style in enumerate(self.extras.styles)]
        upload_index = {item["name"].casefold(): item["index"] for item in self.extras.sheets}
        names = {item["name"].casefold(): item["index"] for item in template.sheets}
        # template sheet index -> (first extra column, its index in the file)
        self.extra_columns = {names[name.casefold()]: (first, upload_index[name.casefold()])
                              for name, first in info["columns"].items()}
        self.sources = [upload_index[name.casefold()] for name in info["sheets"]]
        self.order = tuple(info["order"])
        sheets = [self._widened(item) for item in template.sheets]
        for number, source in enumerate(self.sources):
            sheets.append(self._extra_sheet(self.extras.sheets[source], count + number))
        self.sheets = sheets

    def _widened(self, item):
        extra = self.extra_columns.get(item["index"])
        if extra is None:
            return item
        first, source = extra
        upload = self.extras.sheets[source]
        widths = {key: value for key, value in upload["col_widths"].items() if int(key) >= first}
        return {**item, "cols": max(item["cols"], upload["cols"]),
                "rows": max(item["rows"], upload["rows"]),
                "col_widths": {**item["col_widths"], **widths}}

    def _extra_sheet(self, item, index):
        meta = {key: value for key, value in item.items() if key not in _DROPPED_SHEET_KEYS}
        offset = self.style_offset
        meta.update({
            "index": index, "protected": False, "active": False, "extra": True,
            "has_conditional_formats": False, "conditional_formats": [],
            "col_styles": [[first, last, style + offset] for first, last, style in item.get("col_styles", [])],
            "row_styles": {key: style + offset for key, style in item.get("row_styles", {}).items()},
            "pivots": [], "excel_pivots": [], "slicers": [], "tables": [], "filters": [], "validations": [],
        })
        return meta

    def extra_sheet(self, sheet: int) -> bool:
        return sheet >= len(self.template.sheets)

    def extra_column(self, sheet: int, col: int) -> bool:
        extra = self.extra_columns.get(sheet)
        return extra is not None and col >= extra[0]

    @lru_cache(maxsize=64)
    def block(self, sheet: int, block: int) -> dict:
        count = len(self.template.sheets)
        if sheet >= count:
            return self._shifted(self.extras.block(self.sources[sheet - count], block), None)
        base = self.template.block(sheet, block)
        extra = self.extra_columns.get(sheet)
        if extra is None:
            return base
        added = self._shifted(self.extras.block(extra[1], block), extra[0])
        merged = {row: dict(cells) for row, cells in base.items()}
        for row, cells in added.items():
            kept = {col: record for col, record in merged.get(row, {}).items() if int(col) < extra[0]}
            merged[row] = {**kept, **cells}
        return merged

    def _shifted(self, rows, first):
        """Extras cells with style indexes moved past the template's."""
        output = {}
        for row, cells in rows.items():
            kept = {}
            for col, record in cells.items():
                if first is not None and int(col) < first:
                    continue
                kept[col] = {**record, "s": record["s"] + self.style_offset,
                             **({} if record["k"] == "f" else {"k": "c"})}
            if kept:
                output[row] = kept
        return output

    @lru_cache(maxsize=64)
    def note_block(self, sheet: int, block: int) -> dict:
        count = len(self.template.sheets)
        if sheet >= count:
            return self.extras.note_block(self.sources[sheet - count], block)
        return self.template.note_block(sheet, block)

    def extra_cell(self, sheet: int, row: int, col: int) -> bool:
        return self.extra_sheet(sheet) or self.extra_column(sheet, col)

    def open_cell(self, sheet: int, row: int, col: int) -> bool:
        """The user's own tabs and columns are editable like an unprotected sheet: all but formulas."""
        if self.extra_cell(sheet, row, col):
            return 0 <= row < 1_048_576 and 0 <= col < 16_384 and self.cell(sheet, row, col)["k"] != "f"
        return super().open_cell(sheet, row, col)

    def implicit_cells(self, sheet: int, row: int) -> dict:
        """Empty cells of the user's own columns, so the grid lets them be typed into on a protected sheet."""
        cells = super().implicit_cells(sheet, row)
        extra = self.extra_columns.get(sheet)
        if extra is not None and row < self.sheets[sheet]["rows"]:
            stored = self.block(sheet, row // 256).get(str(row % 256), {})
            for col in range(extra[0], self.sheets[sheet]["cols"]):
                if str(col) not in stored and str(col) not in cells:
                    cells[str(col)] = {"v": None, "s": self._default_style(sheet, row, col), "k": "c"}
        return cells

    def input_flags(self, sheet: int, row: int, col: int, record: dict) -> tuple[bool, bool]:
        if self.extra_cell(sheet, row, col):
            return False, record["k"] != "f"
        return super().input_flags(sheet, row, col, record)

    @lru_cache(maxsize=64)
    def _used_cols(self, sheet: int) -> int:
        count = len(self.template.sheets)
        if sheet >= count:
            return self.extras._used_cols(self.sources[sheet - count])
        used = self.template._used_cols(sheet)
        return max(used, self.sheets[sheet]["cols"]) if sheet in self.extra_columns else used

    def validation(self, sheet: int, row: int, col: int) -> dict | None:
        if self.extra_sheet(sheet) or self.extra_column(sheet, col):
            return None
        return self.template.validation(sheet, row, col)


class ExtrasCache:
    """CopySnapshot per extras directory; extras never change in place, a new upload makes a new one."""

    def __init__(self, limit=16):
        self.limit = limit
        self._items = OrderedDict()
        self._lock = threading.Lock()

    def get(self, template: TemplateSnapshot, path: str) -> TemplateSnapshot:
        with self._lock:
            snapshot = self._items.get(path)
            if snapshot is not None:
                self._items.move_to_end(path)
                return snapshot
        directory = Path(path)
        if not (directory / EXTRAS_FILE).exists():
            raise AppError(500, "extras_missing", "This copy's own tabs and columns are unavailable")
        try:
            snapshot = CopySnapshot(template, directory)
        except AppError as exc:
            if exc.code != "template_reimport_required":
                raise
            snapshot = CopySnapshot(template, _reimport(directory))
        with self._lock:
            self._items[path] = snapshot
            while len(self._items) > self.limit:
                self._items.popitem(last=False)
        return snapshot


def _reimport(directory: Path) -> Path:
    """After an app update changes the import format, rebuild the extras snapshot from the stored file."""
    from .template.importer import import_template

    info = _read_json(directory / EXTRAS_FILE)
    old = directory / info["snapshot"]
    staged = directory / "reimport-source.xlsm"
    shutil.copyfile(old / "source.xlsm", staged)  # the rebuild replaces the folder holding the file
    try:
        imported = import_template(staged, directory, "cached")
    finally:
        staged.unlink(missing_ok=True)
    info["snapshot"] = str(imported.relative_to(directory))
    _prune(imported, ExtrasPlan(tuple(info["sheets"]), info["columns"], tuple(info["order"])))
    staged = directory / f"{EXTRAS_FILE}.tmp"
    staged.write_text(json.dumps(info), encoding="utf-8")
    staged.replace(directory / EXTRAS_FILE)
    return directory
