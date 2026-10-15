"""Read-only access to imported template snapshots."""

from __future__ import annotations

from dataclasses import dataclass
from copy import deepcopy
from functools import lru_cache
import json
from pathlib import Path

from .authorization import effective_role
from .errors import AppError
from .fidelity.conditional_formatting import (
    CurrentCopyValues, conditional_style_overrides,
)
from .fidelity.tables import merge_style_overrides, table_style_overrides
from .pivot_state import EMPTY_STATE as EMPTY_PIVOT_STATE, read_state
from .row_cache import RowBlockCache
from .template.importer import IMPORT_VERSION

SLICER_MESSAGE = ("Slicers need Microsoft Excel, and this server is running without it (values mode). "
                  "They work once Excel is set up for the app on the server.")
MACRO_MESSAGE = ("Macros need Microsoft Excel, and this server is running without it (values mode), so "
                 "buttons and recalculation are off. They work once Excel is set up for the app on the server.")


class TemplateSnapshot:
    def __init__(self, path: str | Path):
        self.path = Path(path)
        self.meta = _read_json(self.path / "meta.json")
        if self.meta.get("import_version") != IMPORT_VERSION:
            raise AppError(
                409,
                "template_reimport_required",
                "The copy template must be reimported before it can be opened",
            )
        self.styles = _read_json(self.path / "styles.json")
        self.input_mask = frozenset(_read_json(self.path / "input_mask.json"))
        buttons = self.path / "buttons.json"
        self.buttons = _read_json(buttons) if buttons.exists() else []
        drawings = self.path / "drawings.json"
        self.drawings = _read_json(drawings) if drawings.exists() else []
        self.refreshable = bool(self.meta.get("refreshable", False))
        self.hash = self.meta["hash"]
        self.results = self.meta.get("results", "engine")
        self.name = self.meta["name"]
        self.sheets = self.meta["sheets"]
        self.source_path = self.path / "source.xlsm"
        self._pivot_areas = {}

    @lru_cache(maxsize=64)
    def block(self, sheet: int, block: int) -> dict:
        path = self.path / "sheets" / str(sheet) / f"{block}.json"
        return _read_json(path) if path.exists() else {}

    @lru_cache(maxsize=64)
    def note_block(self, sheet: int, block: int) -> dict:
        path = self.path / "notes" / str(sheet) / f"{block}.json"
        return _read_json(path) if path.exists() else {}

    def cell(self, sheet: int, row: int, col: int) -> dict:
        cells = self.block(sheet, row // 256).get(str(row % 256), {})
        if str(col) in cells:
            return cells[str(col)]
        style = self._default_style(sheet, row, col)
        kind = "i" if self._implicit_input((sheet, row, col), style) else "c"
        return {"v": None, "s": style, "k": kind}

    def is_input(self, sheet: int, row: int, col: int) -> bool:
        return self.template_input(sheet, row, col) or self.open_cell(sheet, row, col)

    def input_flags(self, sheet: int, row: int, col: int, record: dict) -> tuple[bool, bool]:
        """(template input, editable) for a cell whose template record the caller already holds."""
        if self._pivot_output(sheet, row, col):
            return False, False
        template = f"{sheet}:{row}:{col}" in self.input_mask or record["k"] == "i"
        editable = template or (0 <= row < 1_048_576 and 0 <= col < 16_384 and record["k"] != "f"
                                and not self.sheets[sheet].get("protected", False))
        return template, editable

    def template_input(self, sheet: int, row: int, col: int) -> bool:
        if self._pivot_output(sheet, row, col):
            return False
        return f"{sheet}:{row}:{col}" in self.input_mask or self.cell(sheet, row, col)["k"] == "i"

    def open_cell(self, sheet: int, row: int, col: int) -> bool:
        """Excel lets anyone type into a sheet without protection; only formulas stay locked here."""
        if not (0 <= sheet < len(self.sheets) and 0 <= row < 1_048_576 and 0 <= col < 16_384):
            return False
        return (not self.sheets[sheet].get("protected", False) and self.cell(sheet, row, col)["k"] != "f"
                and not self._pivot_output(sheet, row, col))

    def _pivot_output(self, sheet: int, row: int, col: int) -> bool:
        if not 0 <= sheet < len(self.sheets):
            return False
        return any(first_row <= row <= last_row and first_col <= col <= last_col
                   for first_row, first_col, last_row, last_col in self._pivot_rectangles(sheet))

    def _pivot_rectangles(self, sheet: int) -> tuple:
        """Pivot bodies and their page-field rows; built once per sheet because uploads ask per cell."""
        areas = self._pivot_areas.get(sheet)
        if areas is None:
            outputs = [(pivot["location"], len(pivot.get("page_fields", [])))
                       for pivot in self.sheets[sheet].get("pivots", [])]
            outputs += [(pivot["location"], pivot["page_count"])  # Excel's own pivots (P6.4)
                        for pivot in self.sheets[sheet].get("excel_pivots", [])]
            areas = []
            for location, page_count in outputs:
                first_row, first_col, last_row, last_col = location["bounds"]
                areas.append((first_row, first_col, last_row, last_col))
                page_row = max(0, first_row - page_count - 1)
                if page_count:
                    areas.append((page_row, first_col, page_row + page_count - 1, first_col + 1))
            areas = self._pivot_areas[sheet] = tuple(areas)
        return areas

    def implicit_cells(self, sheet: int, row: int) -> dict:
        """Input cells of a row, including empty ones an unlocked row or column style makes editable."""
        if not 0 <= sheet < len(self.sheets) or not 0 <= row < self.sheets[sheet]["rows"]:
            return {}
        info = self.sheets[sheet]
        stored = self.block(sheet, row // 256).get(str(row % 256), {})
        cells = {col: record for col, record in stored.items()
                 if record["k"] == "i" and int(col) < info["cols"]}
        if not info.get("protected", False):
            return cells
        row_style = info.get("row_styles", {}).get(str(row))
        # shortcut: unlocked row or column styles stop at the used width, not column 16384 (2M empty cells
        # per window); typing far right of all data is refused there. Upgrade: send style spans to the client.
        width = min(info["cols"], self._used_cols(sheet))
        spans = ([(0, width - 1, row_style)] if row_style is not None else self._column_spans(sheet))
        for first, last, style in spans:
            last = min(last, width - 1)
            if self.styles[style].get("locked", True):
                continue
            cells.update((str(col), {"v": None, "s": style, "k": "i"})
                         for col in range(first, last + 1) if str(col) not in stored)
        return cells

    @lru_cache(maxsize=64)
    def _used_cols(self, sheet: int) -> int:
        """One past the rightmost stored cell, merge or styled column range that ends inside the sheet."""
        info, used = self.sheets[sheet], 0
        for path in (self.path / "sheets" / str(sheet)).glob("*.json"):
            for cells in _read_json(path).values():
                used = max([used, *(int(col) + 1 for col in cells)])
        used = max([used, *(merge[3] + 1 for merge in info.get("merges", []) if len(merge) == 4)])
        return max([used, *(last + 1 for _, last, _ in info.get("col_styles", []) if last < info["cols"] - 1)])

    @lru_cache(maxsize=64)
    def _column_spans(self, sheet: int) -> tuple:
        """Column style ranges clipped to the sheet; the first range listing a column wins, as in _default_style."""
        info, spans, covered = self.sheets[sheet], [], []
        for first, last, style in info.get("col_styles", []):
            pieces = [(max(first, 0), min(last, info["cols"] - 1))]
            for done_first, done_last in covered:
                pieces = [part for start, stop in pieces for part in (
                    (start, min(stop, done_first - 1)), (max(start, done_last + 1), stop))
                    if part[0] <= part[1]]
            spans += [(start, stop, style) for start, stop in pieces if start <= stop]
            covered.append((first, last))
        return tuple(spans)

    def _default_style(self, sheet, row, col):
        if not 0 <= sheet < len(self.sheets):
            return 0
        info = self.sheets[sheet]
        if str(row) in info.get("row_styles", {}):
            return info["row_styles"][str(row)]
        for first, last, style in info.get("col_styles", []):
            if first <= col <= last:
                return style
        return 0

    def _implicit_input(self, coordinate, style):
        sheet, row, col = coordinate
        if not 0 <= sheet < len(self.sheets):
            return False
        info = self.sheets[sheet]
        bounded = 0 <= row < info["rows"] and 0 <= col < info["cols"]
        return bounded and info.get("protected", False) and not self.styles[style].get("locked", True)

    def validation(self, sheet: int, row: int, col: int) -> dict | None:
        for validation in self.sheets[sheet].get("validations", []):
            if any(r1 <= row <= r2 and c1 <= col <= c2
                   for r1, c1, r2, c2 in validation.get("ranges", [])):
                return validation
        return None


class CurrentResults(dict):
    """Published results only hold cells that differ from the template; the rest are its saved values."""

    def __init__(self, snapshot, published):
        super().__init__(published)
        self.snapshot = snapshot

    def __contains__(self, key):
        return dict.__contains__(self, key) or self.snapshot.cell(*key)["k"] == "f"

    def __missing__(self, key):
        cell = self.snapshot.cell(*key)
        if cell["k"] != "f":
            raise KeyError(key)
        return cell["v"]


def calc_error(connection, copy_id: str, revision: int) -> str | None:
    """Why the copy's newest saved state could not be calculated, while it still waits."""
    row = connection.execute("SELECT last_error FROM pp_calc_jobs WHERE copy_id=? AND revision=? AND status='pending'",
                             (copy_id, revision)).fetchone()
    return row[0] if row is not None else None


def _calculated_revision(snapshot, copy):
    # Saved Excel results are current for every revision of a values-only template.
    return copy["input_revision"] if getattr(snapshot, "results", "engine") == "cached" else copy["calculated_revision"]


def _read_json(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))


@dataclass(frozen=True)
class RowQuery:
    copy_id: str
    user_id: str
    sheet: int
    start: int
    count: int


@dataclass(frozen=True)
class RowRender:
    snapshot: TemplateSnapshot
    query: RowQuery
    overrides: dict
    results: dict


@dataclass(frozen=True)
class CellQuery:
    copy_id: str
    user_id: str
    sheet: int
    row: int
    col: int


@dataclass(frozen=True)
class HistoryQuery:
    copy_id: str
    user_id: str
    sheet: int | None
    row: int | None
    col: int | None
    limit: int


class TemplateCatalog:
    def __init__(self, database, current):
        self.database = database
        self.current_hash = current.hash
        self._snapshots = {current.hash: current}
        self._extras = None

    def for_copy(self, connection, copy):
        """The copy's template, plus the tabs and columns its owner added in Excel, when it has any."""
        template = self.get(connection, copy["template_id"])
        row = connection.execute("SELECT path FROM pp_copy_extras WHERE copy_id=?", (copy["id"],)).fetchone()
        if row is None:
            return template
        if self._extras is None:
            from .extras import ExtrasCache
            self._extras = ExtrasCache()
        return self._extras.get(template, row["path"])

    def current(self, connection):
        return self.get(connection, self.current_hash)

    def get(self, connection, template_id):
        snapshot = self._snapshots.get(template_id)
        if snapshot is not None:
            return snapshot
        row = connection.execute("SELECT snapshot_path FROM pp_templates WHERE id=?", (template_id,)).fetchone()
        if row is None:
            raise AppError(500, "template_missing", "The copy template is unavailable")
        snapshot = TemplateSnapshot(row["snapshot_path"])
        if snapshot.hash != template_id:
            raise AppError(500, "template_invalid", "The copy template is invalid")
        self._snapshots[template_id] = snapshot
        return snapshot


class WorkbookService:
    def __init__(self, database, snapshot, row_cache=None):
        self.database = database
        self.templates = snapshot if isinstance(snapshot, TemplateCatalog) else TemplateCatalog(database, snapshot)
        self.row_cache = row_cache if row_cache is not None else RowBlockCache()

    def workbook(self, copy_id: str, user_id: str) -> dict:
        with self.database.connect() as connection:
            copy = self._copy(connection, copy_id, user_id)
            snapshot = self.templates.for_copy(connection, copy)
            extents = _result_extents(connection, copy_id)
            pivot_state = read_state(connection, copy_id)
            error = calc_error(connection, copy_id, copy["input_revision"])
        sheets = [{"index": item["index"], "name": item["name"], "hidden": item["hidden"],
                   "protected": bool(item.get("protected", False)),
                   "rows": _visible_extent(item, extents, 0),
                   "cols": _visible_extent(item, extents, 1),
                   "freeze": item["freeze"],
                   "default_row_height": _default_row_height(item),
                   "col_widths": {key: round(value * 7 + 5) for key, value in item["col_widths"].items()},
                   "row_heights": {key: round(value * 96 / 72) for key, value in item["row_heights"].items()},
                   "hidden_rows": item["hidden_rows"], "hidden_cols": item["hidden_cols"],
                   "merges": item["merges"], "pivots": deepcopy(item.get("pivots", [])),
                   "slicers": _slicers_with_state(item.get("slicers", []), pivot_state),
                   **({"tab_color": item["tab_color"]} if item.get("tab_color") else {}),
                   "show_gridlines": item.get("show_gridlines", True),
                   "zoom": item.get("zoom", 1.0),
                   "scroll": item.get("scroll") or {"row": 0, "col": 0},
                   "active": bool(item.get("active")),
                   "has_conditional_formats": item.get("has_conditional_formats", False),
                   **({"tables": _client_tables(item.get("tables", [])),
                       "filters": item.get("filters", [])}
                      if item.get("tables") or item.get("filters") else {}),
                   **({"extra": True} if item.get("extra") else {})}
                  for item in _tab_order(snapshot)]
        return {"template": {"hash": snapshot.hash, "name": snapshot.name},
                "input_revision": copy["input_revision"], "calculated_revision": _calculated_revision(snapshot, copy),
                "calc_error": error if _calculated_revision(snapshot, copy) < copy["input_revision"] else None,
                "calculation": snapshot.results, "sheets": sheets, "styles": snapshot.styles,
                "buttons": snapshot.buttons, "drawings": snapshot.drawings, "refreshable": snapshot.refreshable}

    def rows(self, query: RowQuery) -> dict:
        with self.database.read_transaction() as connection:
            copy = self._copy(connection, query.copy_id, query.user_id)
            snapshot = self.templates.for_copy(connection, copy)
            if (not 0 <= query.sheet < len(snapshot.sheets) or query.start < 0
                    or not 1 <= query.count <= 500):
                raise AppError(422, "invalid", "Sheet row range is invalid")
            # Access is settled first; only then may a body another reader rendered be reused.
            key = _row_cache_key(connection, copy, query)
            # A body shared by unedited copies carries the reading copy's own revisions
            revisions = {"calculated_revision": _calculated_revision(snapshot, copy),
                         "input_revision": copy["input_revision"]}
            cached = self.row_cache.get(key)
            if cached is not None:
                return {**cached, **revisions}
            render = RowRender(snapshot, query, _overrides(connection, query), _results(connection, query))
            conditional = conditional_style_overrides(
                snapshot, CurrentCopyValues(connection, snapshot, query.copy_id), query,
            )
            table_styles = table_style_overrides(snapshot.sheets[query.sheet], query)
            styles = merge_style_overrides(table_styles, conditional.styles)
            rows = self._row_data(connection, render, styles)
        body = {**revisions, "start": query.start, "rows": rows,
                "conditional_format_warnings": conditional.warnings}
        self.row_cache.put(key, body)
        return body

    def _row_data(self, connection, render, style_overrides=None):
        output = {}
        style_overrides = style_overrides or {}
        query = render.query
        override_rows = _group_rows(render.overrides)
        result_rows = _group_rows(render.results)
        style_rows = _group_rows(style_overrides)
        note_rows = _requested_notes(render.snapshot, query)
        open_sheet = not render.snapshot.sheets[query.sheet].get("protected", False)
        extra_cell = getattr(render.snapshot, "extra_cell", None)  # a copy's own columns are open cells
        for row in range(query.start, query.start + query.count):
            implicit = render.snapshot.implicit_cells(query.sheet, row)
            stored = render.snapshot.block(query.sheet, row // 256).get(str(row % 256), {})
            cells = {**implicit, **stored}
            styles = style_rows.get(row, {})
            notes = note_rows.get(row, {})
            coordinates = (set(cells) | {str(col) for col in override_rows.get(row, {})}
                           | {str(col) for col in result_rows.get(row, {})}
                           | {str(col) for col in styles | notes})
            rendered = {}
            for col_text in sorted(coordinates, key=int):
                col = int(col_text)
                coordinate = (query.sheet, row, col)
                record = dict(cells.get(col_text, render.snapshot.cell(*coordinate)))
                record.pop("formula", None)
                override = render.overrides.get((row, col))
                input_record = _input_record(connection, override) if override else None
                published = (row, col) in render.results
                if published:
                    record.update({"v": render.results[(row, col)], "k": "f"})
                record.update(_cell_fields(render.snapshot, coordinate, input_record, published))
                if not override and not published and record["k"] == "i":
                    record["ver"] = 0
                if record["k"] == "c" and (open_sheet or extra_cell and extra_cell(*coordinate)) \
                        and not render.snapshot._pivot_output(*coordinate):
                    record.update({"open": True, "ver": record.get("ver", 0)})
                if col in styles:
                    record["style_override"] = styles[col]
                if col in notes:
                    record["note"] = notes[col]
                rendered[col_text] = record
            if rendered:
                output[str(row)] = rendered
        return output

    def cell(self, query: CellQuery) -> dict:
        with self.database.read_transaction() as connection:
            copy = self._copy(connection, query.copy_id, query.user_id)
            snapshot = self.templates.for_copy(connection, copy)
            if min(query.sheet, query.row, query.col) < 0 or query.sheet >= len(snapshot.sheets):
                raise AppError(422, "invalid", "Cell coordinate is invalid")
            override = connection.execute(
                "SELECT * FROM pp_overrides WHERE copy_id=? AND sheet=? AND row=? AND col=?",
                (query.copy_id, query.sheet, query.row, query.col),
            ).fetchone()
            result = connection.execute(
                "SELECT value_json FROM pp_results WHERE copy_id=? AND sheet=? AND row=? AND col=?",
                (query.copy_id, query.sheet, query.row, query.col),
            ).fetchone()
            input_record = _input_record(connection, override) if override else None
        coordinate = (query.sheet, query.row, query.col)
        record = dict(snapshot.cell(*coordinate))
        if result:
            record["k"] = "f"
        response = {"v": json.loads(result["value_json"]) if result else record["v"], "s": record["s"],
                    "k": record["k"], "formula": record.get("formula"),
                    "validation": snapshot.validation(*coordinate)}
        response.update(_cell_fields(snapshot, coordinate, input_record, result is not None))
        if not override and not result and record["k"] == "i":
            response["ver"] = 0
        return response

    def history(self, query: HistoryQuery) -> dict:
        if type(query.limit) is not int or not 1 <= query.limit <= 500:
            raise AppError(422, "invalid", "History limit is invalid")
        with self.database.connect() as connection:
            copy = self._copy(connection, query.copy_id, query.user_id)
            rows = _history_rows(connection, copy, query)
        return {"entries": [_history_entry(item) for item in rows]}

    def macro_request(self, copy_id: str, user_id: str, button_id, op_id, engine_kind: str | None) -> str:
        """Authorize a button press and return the macro that button runs (clients never name macros)."""
        snapshot = self._press_access(copy_id, user_id, op_id)
        button = next((item for item in snapshot.buttons if item["id"] == button_id), None)
        if button is None:
            raise AppError(404, "unknown_button", "Macro button not found")
        _require_excel(snapshot, engine_kind)
        return button["macro"]

    def slicer_request(self, copy_id: str, user_id: str, op_id, cache, selected, engine_kind: str | None) -> list[str]:
        """Authorize a slicer change and return its validated item texts (clients never name pivots)."""
        snapshot = self._press_access(copy_id, user_id, op_id)
        slicer = next((item for sheet in snapshot.sheets for item in sheet.get("slicers", [])
                       if item.get("engine") == "excel" and item["cache"] == cache), None)
        if slicer is None:
            raise AppError(404, "unknown_slicer", "Slicer not found")
        _require_excel(snapshot, engine_kind, "slicers_unavailable", SLICER_MESSAGE)
        if not isinstance(selected, list) or not all(isinstance(text, str) for text in selected):
            raise AppError(422, "invalid", "Selected items must be a list of item texts")
        chosen = list(dict.fromkeys(selected))
        if not chosen:
            raise AppError(422, "empty_selection", "Select at least one item")
        captions = {item["caption"] for item in slicer["items"]}
        if any(text not in captions for text in chosen):
            raise AppError(422, "unknown_item", "Slicer item not found")
        return chosen

    def pivot_state(self, copy_id: str) -> dict:
        with self.database.connect() as connection:
            return read_state(connection, copy_id)

    def refresh_request(self, copy_id: str, user_id: str, op_id, engine_kind: str | None) -> None:
        """Authorize a data refresh of a template that has connections, on the Excel engine."""
        snapshot = self._press_access(copy_id, user_id, op_id)
        if not snapshot.refreshable:
            raise AppError(409, "refresh_unavailable", "This template has no data connections to refresh.")
        _require_excel(snapshot, engine_kind)

    def _press_access(self, copy_id: str, user_id: str, op_id):
        if not isinstance(op_id, str) or not op_id:
            raise AppError(422, "invalid", "Operation identifier is required")
        with self.database.connect() as connection:
            copy = self._copy(connection, copy_id, user_id)
            if effective_role(connection, copy_id, user_id) not in {"owner", "editor"}:
                raise AppError(403, "forbidden", "Editor access is required")
            return self.templates.for_copy(connection, copy)

    def source_and_overrides(self, copy_id: str, user_id: str):
        with self.database.connect() as connection:
            copy = self._copy(connection, copy_id, user_id)
            snapshot = self.templates.for_copy(connection, copy)
            template = connection.execute("SELECT source_path FROM pp_templates WHERE id=?",
                                          (copy["template_id"],)).fetchone()
            rows = connection.execute("SELECT sheet,row,col,value_json FROM pp_overrides WHERE copy_id=?",
                                      (copy_id,)).fetchall()
            _check_overrides(connection, snapshot, copy_id, rows)
        return _source(snapshot, template), {(item["sheet"], item["row"], item["col"]): json.loads(item["value_json"])
                                             for item in rows}

    def export_snapshot(self, copy_id: str, user_id: str):
        """Read inputs and only results calculated for the same copy revision."""
        with self.database.read_transaction() as connection:
            copy = self._copy(connection, copy_id, user_id)
            snapshot = self.templates.for_copy(connection, copy)
            template = connection.execute(
                "SELECT source_path FROM pp_templates WHERE id=?", (copy["template_id"],)
            ).fetchone()
            rows = connection.execute(
                "SELECT sheet,row,col,value_json FROM pp_overrides WHERE copy_id=?", (copy_id,)
            ).fetchall()
            _check_overrides(connection, snapshot, copy_id, rows)
            results = {}
            if _calculated_revision(snapshot, copy) == copy["input_revision"]:
                result_rows = connection.execute(
                    "SELECT sheet,row,col,value_json FROM pp_results WHERE copy_id=? AND calculated_revision=?",
                    (copy_id, copy["input_revision"]),
                ).fetchall()
                results = CurrentResults(snapshot, {
                    (item["sheet"], item["row"], item["col"]): json.loads(item["value_json"])
                    for item in result_rows})
        return {
            "source": _source(snapshot, template), "template_id": copy["template_id"],
            "order": getattr(snapshot, "order", None), "extras_id": getattr(snapshot, "extras_id", ""),
            "input_revision": copy["input_revision"],
            "calculated_revision": copy["calculated_revision"], "overrides": {
                (item["sheet"], item["row"], item["col"]): json.loads(item["value_json"]) for item in rows
            }, "results": results, "sheets": snapshot.sheets,
        }

    def calculation_source(self, copy_id: str) -> Path:
        with self.database.connect() as connection:
            copy = connection.execute(
                "SELECT * FROM pp_copies WHERE id=?", (copy_id,)
            ).fetchone()
            if copy is None:
                raise AppError(404, "not_found", "Copy not found")
            template = connection.execute(
                "SELECT source_path FROM pp_templates WHERE id=?", (copy["template_id"],)
            ).fetchone()
            if template is None:
                raise AppError(500, "template_missing", "The copy template is unavailable")
            return _source(self.templates.for_copy(connection, copy), template)

    def _copy(self, connection, copy_id, user_id):
        copy = connection.execute("SELECT * FROM pp_copies WHERE id=?", (copy_id,)).fetchone()
        if copy is None or effective_role(connection, copy_id, user_id) is None:
            raise AppError(404, "not_found", "Copy not found")
        return copy


def _tab_order(snapshot):
    """Sheets in the order the copy's file had them; a copy's own tabs keep their place among the template's."""
    order = getattr(snapshot, "order", None)
    if not order:
        return snapshot.sheets
    position = {name.casefold(): number for number, name in enumerate(order)}
    return sorted(snapshot.sheets, key=lambda item: (position.get(item["name"].casefold(), len(order)), item["index"]))


def _source(snapshot, template):
    """The package a copy is calculated and downloaded from: the template, with the copy's extras grafted in."""
    return getattr(snapshot, "grafted_path", None) or Path(template["source_path"])


def _overrides(connection, query):
    rows = connection.execute(
        "SELECT * FROM pp_overrides WHERE copy_id=? AND sheet=? AND row>=? AND row<?",
        (query.copy_id, query.sheet, query.start, query.start + query.count),
    ).fetchall()
    return {(item["row"], item["col"]): item for item in rows}


def _row_cache_key(connection, copy, query):
    """Every input the rendered body depends on, so one key change means one re-render.

    The template snapshot behind ``template_id`` is immutable while the process holds it, and
    conditional formatting reads current values from the whole copy, so an edit anywhere (which
    moves ``input_revision``) and a published calculation (which moves ``calculated_revision``)
    both invalidate every sheet. The copy's stored slicer selection is an input in its own right:
    Excel pivot output reaches the grid through published results, but the selection decides those
    results, so it is keyed by content rather than by the revision written next to it.
    """
    state = read_state(connection, query.copy_id)
    extras = connection.execute("SELECT path FROM pp_copy_extras WHERE copy_id=?", (query.copy_id,)).fetchone()
    if extras is not None:
        return (query.copy_id, copy["template_id"], extras["path"], query.sheet, query.start, query.count,
                copy["input_revision"], copy["calculated_revision"],
                json.dumps(state, sort_keys=True, separators=(",", ":")))
    if state == EMPTY_PIVOT_STATE and not _copy_has_values(connection, query.copy_id):
        # A copy with no edits, results or slicer choices renders exactly the template, so all
        # such copies share one body: 50 fresh copies cost one render per block, not 50
        return ("template", copy["template_id"], query.sheet, query.start, query.count)
    return (query.copy_id, copy["template_id"], query.sheet, query.start, query.count,
            copy["input_revision"], copy["calculated_revision"],
            json.dumps(state, sort_keys=True, separators=(",", ":")))


def _copy_has_values(connection, copy_id):
    return connection.execute(
        "SELECT EXISTS(SELECT 1 FROM pp_overrides WHERE copy_id=?) "
        "OR EXISTS(SELECT 1 FROM pp_results WHERE copy_id=?)", (copy_id, copy_id),
    ).fetchone()[0] == 1


def _group_rows(records):
    grouped = {}
    for (row, col), value in records.items():
        grouped.setdefault(row, {})[col] = value
    return grouped


def _client_tables(tables):
    return [{key: value for key, value in table.items() if key != "style"}
            for table in tables]


def _requested_notes(snapshot, query):
    grouped = {}
    end = query.start + query.count
    for block in range(query.start // 256, (end - 1) // 256 + 1):
        for offset, columns in snapshot.note_block(query.sheet, block).items():
            row = block * 256 + int(offset)
            if query.start <= row < end:
                grouped[row] = {int(col): note for col, note in columns.items()}
    return grouped


def _results(connection, query):
    rows = connection.execute(
        "SELECT row,col,value_json FROM pp_results WHERE copy_id=? AND sheet=? AND row>=? AND row<?",
        (query.copy_id, query.sheet, query.start, query.start + query.count),
    ).fetchall()
    return {(item["row"], item["col"]): json.loads(item["value_json"]) for item in rows}


def _history_rows(connection, copy, query, cutoff=None):
    rows = _local_history(connection, copy["id"], query, cutoff)
    if copy["parent_copy_id"]:
        parent = connection.execute(
            "SELECT * FROM pp_copies WHERE id=?", (copy["parent_copy_id"],),
        ).fetchone()
        inherited = _history_rows(connection, parent, query, copy["fork_revision"])
        marker = {"copy_id": parent["id"], "name": parent["name"], "revision": copy["fork_revision"]}
        for item in inherited:
            if item["origin"] is None:
                item["origin"] = marker
        rows.extend(inherited)
    return sorted(rows, key=lambda item: item["id"], reverse=True)[:query.limit]


def _local_history(connection, copy_id, query, cutoff):
    clauses, values = ["h.copy_id=?"], [copy_id]
    if cutoff is not None:
        clauses.append("h.revision<=?")
        values.append(cutoff)
    for name, value in (("sheet", query.sheet), ("row", query.row), ("col", query.col)):
        if value is not None:
            clauses.append(f"h.{name}=?")
            values.append(value)
    values.append(query.limit)
    rows = connection.execute(
        "SELECT h.*,u.id AS user_id,u.name FROM pp_history h JOIN pp_users u ON u.id=h.actor_id WHERE "  # nosec B608 # clause names are fixed (sheet,row,col); values are bound
        + " AND ".join(clauses) + " ORDER BY h.id DESC LIMIT ?", values,
    ).fetchall()
    return [{**dict(item), "origin": None} for item in rows]


def _history_entry(item):
    entry = {"op_id": item["op_id"], "revision": item["revision"], "at": item["at"],
             "by": {"id": item["user_id"], "name": item["name"]}, "sheet": item["sheet"],
             "r": item["row"], "c": item["col"], "before": json.loads(item["before_json"]),
             "after": json.loads(item["after_json"]), "action": item["action"]}
    if item["origin"] is not None:
        entry["origin"] = item["origin"]
    return entry


def _result_extents(connection, copy_id):
    rows = connection.execute(
        "SELECT sheet,MAX(row)+1 AS rows,MAX(col)+1 AS cols FROM pp_results "
        "WHERE copy_id=? GROUP BY sheet", (copy_id,),
    ).fetchall()
    return {row["sheet"]: (row["rows"], row["cols"]) for row in rows}


def _visible_extent(sheet, result_extents, axis):
    key, bound = (("rows", 2), ("cols", 3))[axis]
    marker = ("row", "column")[axis]
    values = [sheet[key], result_extents.get(sheet["index"], (0, 0))[axis]]
    values.extend(pivot["location"]["bounds"][bound] + 1
                  for pivot in sheet.get("pivots", []))
    for slicer in sheet.get("slicers", []):
        values.extend(slicer["anchor"][point][marker] + 1 for point in ("from", "to"))
    return max(values)


def _default_row_height(sheet):
    height = sheet.get("default_row_height")
    return round(height * 96 / 72) if height is not None else 28


def _input_record(connection, override):
    user = connection.execute("SELECT id,name FROM pp_users WHERE id=?", (override["updated_by"],)).fetchone()
    return {"v": json.loads(override["value_json"]), "k": "i", "ver": override["version"],
            "by": {"id": user["id"], "name": user["name"]}, "at": override["updated_at"]}


def _formula(value):
    return value["formula"] if isinstance(value, dict) and "formula" in value else None


def _cell_fields(snapshot, coordinate, input_record, published):
    """What an override adds to a cell record. A formula is never sent as a value: its published result is."""
    if published:
        return _spill_input_metadata(snapshot, coordinate, input_record)
    if input_record is None:
        return {}
    value = input_record["v"]
    formula = _formula(value)
    if not snapshot.is_input(*coordinate):
        return {} if formula is not None else {"v": value}
    if formula is None:
        return {"v": value, "k": "i", **{key: input_record[key] for key in ("ver", "by", "at")}}
    return {"v": snapshot.cell(*coordinate)["v"], "k": "i", "input_value": formula, "input_formula": True,
            **{key: input_record[key] for key in ("ver", "by", "at")}}


def _spill_input_metadata(snapshot, coordinate, input_record):
    if not snapshot.template_input(*coordinate):
        return {}
    value = input_record["v"] if input_record else snapshot.cell(*coordinate)["v"]
    formula = _formula(value)
    metadata = {"input": True, "input_value": value if formula is None else formula,
                "ver": input_record["ver"] if input_record else 0}
    if formula is not None:
        metadata["input_formula"] = True
    if input_record:
        metadata.update({key: input_record[key] for key in ("by", "at")})
    return metadata


def _require_excel(snapshot, engine_kind: str | None, code: str = "macros_unavailable",
                   message: str = MACRO_MESSAGE) -> None:
    if engine_kind != "excel" or snapshot.results == "cached":
        raise AppError(409, code, message)


def _slicers_with_state(slicers, state):
    """Excel slicers show the copy's stored selection; the rest keep the template's own."""
    output = []
    for slicer in deepcopy(slicers):
        if slicer.get("engine") == "excel" and slicer["cache"] in state["slicers"]:
            selected = set(state["slicers"][slicer["cache"]])
            slicer["items"] = [{**item, "selected": item["caption"] in selected} for item in slicer["items"]]
        output.append(slicer)
    return output


def _check_overrides(connection, snapshot, copy_id, rows):
    """Inputs may hold overrides; any other cell only when a template macro wrote it."""
    for item in rows:
        coordinate = (item["sheet"], item["row"], item["col"])
        if not snapshot.is_input(*coordinate) and not _macro_wrote(connection, copy_id, coordinate):
            raise AppError(422, "locked", "This cell is locked")


def _macro_wrote(connection, copy_id, coordinate):
    """The latest history action for a cell, walking up the fork chain (forks hold no history of their own)."""
    while copy_id is not None:
        latest = connection.execute(
            "SELECT action FROM pp_history WHERE copy_id=? AND sheet=? AND row=? AND col=? ORDER BY id DESC LIMIT 1",
            (copy_id, *coordinate),
        ).fetchone()
        if latest is not None:
            return latest["action"] in {"macro", "refresh", "slicer"}
        parent = connection.execute("SELECT parent_copy_id FROM pp_copies WHERE id=?", (copy_id,)).fetchone()
        copy_id = parent["parent_copy_id"] if parent is not None else None
    return False
