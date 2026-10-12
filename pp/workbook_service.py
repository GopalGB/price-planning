"""Read-only access to imported template snapshots."""

from __future__ import annotations

from dataclasses import dataclass
from functools import lru_cache
import json
from pathlib import Path

from .authorization import effective_role
from .errors import AppError
from .template.importer import IMPORT_VERSION


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
        self.hash = self.meta["hash"]
        self.results = self.meta.get("results", "engine")
        self.name = self.meta["name"]
        self.sheets = self.meta["sheets"]
        self.source_path = self.path / "source.xlsm"

    @lru_cache(maxsize=64)
    def block(self, sheet: int, block: int) -> dict:
        path = self.path / "sheets" / str(sheet) / f"{block}.json"
        return _read_json(path) if path.exists() else {}

    def cell(self, sheet: int, row: int, col: int) -> dict:
        cells = self.block(sheet, row // 256).get(str(row % 256), {})
        if str(col) in cells:
            return cells[str(col)]
        style = self._default_style(sheet, row, col)
        kind = "i" if self._implicit_input((sheet, row, col), style) else "c"
        return {"v": None, "s": style, "k": kind}

    def is_input(self, sheet: int, row: int, col: int) -> bool:
        return f"{sheet}:{row}:{col}" in self.input_mask or self.cell(sheet, row, col)["k"] == "i"

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


def _calculated_revision(snapshot, copy):
    # Saved Excel results are current for every revision of a values-only template.
    return copy["input_revision"] if snapshot.results == "cached" else copy["calculated_revision"]


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
    def __init__(self, database, snapshot):
        self.database = database
        self.templates = snapshot if isinstance(snapshot, TemplateCatalog) else TemplateCatalog(database, snapshot)

    def workbook(self, copy_id: str, user_id: str) -> dict:
        with self.database.connect() as connection:
            copy = self._copy(connection, copy_id, user_id)
            snapshot = self.templates.get(connection, copy["template_id"])
            extents = _result_extents(connection, copy_id)
        sheets = [{"index": item["index"], "name": item["name"], "hidden": item["hidden"],
                   "rows": max(item["rows"], extents.get(item["index"], (0, 0))[0]),
                   "cols": max(item["cols"], extents.get(item["index"], (0, 0))[1]),
                   "freeze": item["freeze"],
                   "col_widths": {key: round(value * 7 + 5) for key, value in item["col_widths"].items()},
                   "row_heights": {key: round(value * 96 / 72) for key, value in item["row_heights"].items()},
                   "hidden_rows": item["hidden_rows"], "hidden_cols": item["hidden_cols"],
                   "merges": item["merges"]} for item in snapshot.sheets]
        return {"template": {"hash": snapshot.hash, "name": snapshot.name},
                "input_revision": copy["input_revision"], "calculated_revision": _calculated_revision(snapshot, copy),
                "calculation": snapshot.results, "sheets": sheets, "styles": snapshot.styles,
                "buttons": snapshot.buttons}

    def rows(self, query: RowQuery) -> dict:
        with self.database.read_transaction() as connection:
            copy = self._copy(connection, query.copy_id, query.user_id)
            snapshot = self.templates.get(connection, copy["template_id"])
            if (not 0 <= query.sheet < len(snapshot.sheets) or query.start < 0
                    or not 1 <= query.count <= 500):
                raise AppError(422, "invalid", "Sheet row range is invalid")
            render = RowRender(snapshot, query, _overrides(connection, query), _results(connection, query))
            rows = self._row_data(connection, render)
        return {"calculated_revision": _calculated_revision(snapshot, copy), "input_revision": copy["input_revision"],
                "start": query.start, "rows": rows}

    def _row_data(self, connection, render):
        output = {}
        query = render.query
        for row in range(query.start, query.start + query.count):
            implicit = render.snapshot.implicit_cells(query.sheet, row)
            stored = render.snapshot.block(query.sheet, row // 256).get(str(row % 256), {})
            cells = {**implicit, **stored}
            coordinates = (set(cells) | {str(col) for r, col in render.overrides if r == row}
                           | {str(col) for r, col in render.results if r == row})
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
                rendered[col_text] = record
            if rendered:
                output[str(row)] = rendered
        return output

    def cell(self, query: CellQuery) -> dict:
        with self.database.read_transaction() as connection:
            copy = self._copy(connection, query.copy_id, query.user_id)
            snapshot = self.templates.get(connection, copy["template_id"])
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
        if not isinstance(op_id, str) or not op_id:
            raise AppError(422, "invalid", "Operation identifier is required")
        with self.database.connect() as connection:
            copy = self._copy(connection, copy_id, user_id)
            if effective_role(connection, copy_id, user_id) not in {"owner", "editor"}:
                raise AppError(403, "forbidden", "Editor access is required")
            snapshot = self.templates.get(connection, copy["template_id"])
        button = next((item for item in snapshot.buttons if item["id"] == button_id), None)
        if button is None:
            raise AppError(404, "unknown_button", "Macro button not found")
        if engine_kind != "excel" or snapshot.results == "cached":
            raise AppError(409, "macros_unavailable",
                           "Macros run on the office server, where the app uses Microsoft Excel.")
        return button["macro"]

    def source_and_overrides(self, copy_id: str, user_id: str):
        with self.database.connect() as connection:
            copy = self._copy(connection, copy_id, user_id)
            snapshot = self.templates.get(connection, copy["template_id"])
            template = connection.execute("SELECT source_path FROM pp_templates WHERE id=?",
                                          (copy["template_id"],)).fetchone()
            rows = connection.execute("SELECT sheet,row,col,value_json FROM pp_overrides WHERE copy_id=?",
                                      (copy_id,)).fetchall()
            _check_overrides(connection, snapshot, copy_id, rows)
        return Path(template["source_path"]), {(item["sheet"], item["row"], item["col"]): json.loads(item["value_json"])
                                                 for item in rows}

    def export_snapshot(self, copy_id: str, user_id: str):
        """Read inputs and only results calculated for the same copy revision."""
        with self.database.read_transaction() as connection:
            copy = self._copy(connection, copy_id, user_id)
            snapshot = self.templates.get(connection, copy["template_id"])
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
            "source": Path(template["source_path"]), "template_id": copy["template_id"],
            "input_revision": copy["input_revision"],
            "calculated_revision": copy["calculated_revision"], "overrides": {
                (item["sheet"], item["row"], item["col"]): json.loads(item["value_json"]) for item in rows
            }, "results": results, "sheets": snapshot.sheets,
        }

    def calculation_source(self, copy_id: str) -> Path:
        with self.database.connect() as connection:
            copy = connection.execute(
                "SELECT template_id FROM pp_copies WHERE id=?", (copy_id,)
            ).fetchone()
            if copy is None:
                raise AppError(404, "not_found", "Copy not found")
            template = connection.execute(
                "SELECT source_path FROM pp_templates WHERE id=?", (copy["template_id"],)
            ).fetchone()
        if template is None:
            raise AppError(500, "template_missing", "The copy template is unavailable")
        return Path(template["source_path"])

    def _copy(self, connection, copy_id, user_id):
        copy = connection.execute("SELECT * FROM pp_copies WHERE id=?", (copy_id,)).fetchone()
        if copy is None or effective_role(connection, copy_id, user_id) is None:
            raise AppError(404, "not_found", "Copy not found")
        return copy


def _overrides(connection, query):
    rows = connection.execute(
        "SELECT * FROM pp_overrides WHERE copy_id=? AND sheet=? AND row>=? AND row<?",
        (query.copy_id, query.sheet, query.start, query.start + query.count),
    ).fetchall()
    return {(item["row"], item["col"]): item for item in rows}


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
        "SELECT h.*,u.id AS user_id,u.name FROM pp_history h JOIN pp_users u ON u.id=h.actor_id WHERE "
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
    if not snapshot.is_input(*coordinate):
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
            return latest["action"] == "macro"
        parent = connection.execute("SELECT parent_copy_id FROM pp_copies WHERE id=?", (copy_id,)).fetchone()
        copy_id = parent["parent_copy_id"] if parent is not None else None
    return False
