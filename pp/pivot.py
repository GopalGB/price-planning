"""Transactional refresh of imported PivotTable metadata."""

from __future__ import annotations

from collections import Counter, OrderedDict, defaultdict
from copy import deepcopy
from dataclasses import dataclass
import json
import math
import re
import threading

from openpyxl.utils.datetime import from_ISO8601, to_excel

from .authorization import effective_role
from .errors import AppError, ErrorPayload
from .workbook_service import TemplateCatalog


RESULT_CACHE_LIMIT = 32
SOURCE_CACHE_LIMIT = 2
INCREMENTAL_CHANGE_LIMIT = 128
CELL_RANGE = re.compile(r"\$?([A-Z]{1,3})\$?([1-9][0-9]*)(?::\$?([A-Z]{1,3})\$?([1-9][0-9]*))?")


@dataclass(frozen=True)
class PivotQuery:
    copy_id: str
    user_id: str
    sheet: int
    name: str
    options: dict


@dataclass(frozen=True)
class SourceState:
    rows: tuple
    fields: tuple
    positions: dict
    first_row: int
    first_col: int
    sheet: int
    metadata: dict
    counts: dict


class PivotService:
    """Build a pivot grid from one authorized copy's current calculated values."""

    def __init__(self, database, snapshot):
        self.database = database
        self._sources = OrderedDict()
        self._results = OrderedDict()
        self._latest = OrderedDict()
        self._incremental_refreshes = 0
        self._lock = threading.RLock()
        current = None if isinstance(snapshot, TemplateCatalog) else snapshot
        self.templates = snapshot if current is None else TemplateCatalog(database, current)
        if current is not None:
            self._prepare(current)

    def refresh(self, query: PivotQuery) -> dict:
        with self.database.read_transaction() as connection:
            copy = self._copy(connection, query)
            snapshot = self.templates.get(connection, copy["template_id"])
            pivot, slicers = _metadata(snapshot, query.sheet, query.name)
            options = _options(pivot, slicers, query.options or {})
            key = _result_key(snapshot, copy, query, options)
            cached = self._cache_get(self._results, key)
            if cached is not None:
                return deepcopy(cached)
            source = self._source(snapshot, pivot, slicers)
            overlays = _overlays(connection, query.copy_id, copy, source)
            latest_key = _latest_key(snapshot, query, options)
            previous = self._cache_get(self._latest, latest_key)
        result = _incremental(pivot, source, overlays, options, previous, copy)
        if result is None:
            result = _refresh(pivot, slicers, source, overlays, options, copy, query.sheet)
        else:
            with self._lock:
                self._incremental_refreshes += 1
        self._cache_put(self._results, key, result, RESULT_CACHE_LIMIT)
        latest = {"overlays": overlays, "result": result}
        self._cache_put(self._latest, latest_key, latest, RESULT_CACHE_LIMIT)
        return deepcopy(result)

    def cache_info(self):
        with self._lock:
            return {"sources": len(self._sources), "results": len(self._results),
                    "incremental_refreshes": self._incremental_refreshes}

    def _copy(self, connection, query):
        role = effective_role(connection, query.copy_id, query.user_id)
        copy = connection.execute("SELECT * FROM pp_copies WHERE id=?", (query.copy_id,)).fetchone()
        if copy is None or role is None:
            raise AppError(404, "not_found", "Copy not found")
        if copy["calculated_revision"] < copy["input_revision"]:
            extra = {"input_revision": copy["input_revision"],
                     "calculated_revision": copy["calculated_revision"]}
            raise AppError(409, "calculation_pending", ErrorPayload(
                "Dashboard calculations are pending", extra,
            ))
        return dict(copy)

    def _source(self, snapshot, pivot, slicers):
        fields = _used_fields(pivot, slicers)
        key = (snapshot.hash, pivot["source"]["sheet"], pivot["source"]["ref"], fields)
        cached = self._cache_get(self._sources, key)
        if cached is not None:
            return cached
        source = _read_source(snapshot, pivot, fields)
        self._cache_put(self._sources, key, source, SOURCE_CACHE_LIMIT)
        return source

    def _prepare(self, snapshot):
        for sheet in snapshot.sheets:
            for pivot in sheet.get("pivots", []):
                slicers = [item for item in sheet.get("slicers", [])
                           if item["pivot_cache_id"] == pivot["cache_id"]]
                self._source(snapshot, pivot, slicers)

    def _cache_get(self, cache, key):
        with self._lock:
            value = cache.get(key)
            if value is not None:
                cache.move_to_end(key)
            return value

    def _cache_put(self, cache, key, value, limit):
        with self._lock:
            cache[key] = value
            cache.move_to_end(key)
            while len(cache) > limit:
                cache.popitem(last=False)


def _metadata(snapshot, sheet, name):
    if not isinstance(sheet, int) or not 0 <= sheet < len(snapshot.sheets):
        raise AppError(422, "invalid", "Pivot sheet is invalid")
    pivots = [item for item in snapshot.sheets[sheet].get("pivots", []) if item["name"] == name]
    if len(pivots) != 1:
        raise AppError(404, "not_found", "Pivot table not found")
    pivot = pivots[0]
    slicers = [item for item in snapshot.sheets[sheet].get("slicers", [])
               if item["pivot_cache_id"] == pivot["cache_id"]]
    return pivot, slicers


def _used_fields(pivot, slicers):
    indexes = set(pivot["row_fields"] + [item for item in pivot["column_fields"] if item >= 0])
    indexes.update(item["field"] for item in pivot["page_fields"])
    indexes.update(item["field"] for item in pivot["data_fields"])
    indexes.update(item["field"] for item in slicers)
    indexes.update(field["sort"]["rank_by"] for field in pivot["fields"]
                   if field["sort"]["rank_by"] is not None)
    return tuple(sorted(indexes))


def _read_source(snapshot, pivot, fields):
    source = pivot["source"]
    matches = [item["index"] for item in snapshot.sheets if item["name"] == source["sheet"]]
    if len(matches) != 1:
        raise AppError(500, "pivot_source_invalid", "Pivot source worksheet is unavailable")
    row1, col1, row2, _ = _bounds(source["ref"])
    metadata = {field: {_shared_key(item): index for index, item in enumerate(
        pivot["fields"][field]["shared_items"])} for field in fields}
    rows, counts = _snapshot_rows(snapshot, matches[0], row1 + 1, row2, col1, fields, metadata)
    return SourceState(rows, fields, {field: pos for pos, field in enumerate(fields)},
                       row1 + 1, col1, matches[0], metadata, counts)


def _snapshot_rows(snapshot, sheet, first_row, last_row, first_col, fields, metadata):
    output = []
    counts = {field: Counter() for field in fields if metadata[field]}
    count_positions = [(position, field) for position, field in enumerate(fields) if field in counts]
    for block_number in range(first_row // 256, last_row // 256 + 1):
        block = snapshot.block(sheet, block_number)
        start, stop = max(first_row, block_number * 256), min(last_row + 1, (block_number + 1) * 256)
        for row_number in range(start, stop):
            stored = block.get(str(row_number % 256), {})
            row = tuple(stored.get(str(first_col + field), {}).get("v") for field in fields)
            output.append(row)
            for position, field in count_positions:
                counts[field][_dimension_from_value(metadata, field, row[position])] += 1
    return tuple(output), counts


def _bounds(reference):
    match = CELL_RANGE.fullmatch(reference)
    if match is None:
        raise AppError(500, "pivot_source_invalid", "Pivot source range is invalid")
    col1, row1, col2, row2 = match.groups()
    return int(row1) - 1, _column(col1), int(row2 or row1) - 1, _column(col2 or col1)


def _column(text):
    result = 0
    for character in text:
        result = result * 26 + ord(character) - 64
    return result - 1


def _overlays(connection, copy_id, copy, source):
    limits = (copy_id, source.sheet, source.first_row,
              source.first_row + len(source.rows) - 1, source.first_col,
              source.first_col + max(source.fields))
    values = {}
    query = ("SELECT row,col,value_json FROM pp_overrides WHERE copy_id=? AND sheet=? "
             "AND row BETWEEN ? AND ? AND col BETWEEN ? AND ?")
    for item in connection.execute(query, limits):
        values[(item["row"], item["col"])] = json.loads(item["value_json"])
    query = ("SELECT row,col,value_json FROM pp_results WHERE copy_id=? AND sheet=? "
             "AND row BETWEEN ? AND ? AND col BETWEEN ? AND ? AND calculated_revision=?")
    for item in connection.execute(query, (*limits, copy["calculated_revision"])):
        values[(item["row"], item["col"])] = json.loads(item["value_json"])
    return values


def _options(pivot, slicers, raw):
    if not isinstance(raw, dict) or set(raw) - {"page", "slicers", "collapsed"}:
        raise AppError(422, "invalid", "Pivot options are invalid")
    page = _filter_options(raw.get("page"), _default_page(pivot), pivot, "page")
    slicer = _filter_options(raw.get("slicers"), _default_slicers(slicers), pivot, "slicers")
    collapsed = _collapsed_options(raw.get("collapsed", []), pivot)
    return {"page": page, "slicers": slicer, "collapsed": collapsed}


def _default_page(pivot):
    return [{"field": item["field"], "items": [item["cache_index"]]
             if item["cache_index"] is not None else None} for item in pivot["page_fields"]]


def _default_slicers(slicers):
    return [{"name": item["name"], "field": item["field"],
             "items": [entry["cache_index"] for entry in item["items"] if entry["selected"]]}
            for item in slicers]


def _filter_options(raw, defaults, pivot, label):
    if raw is None:
        return defaults
    if not isinstance(raw, list):
        raise AppError(422, "invalid", f"Pivot {label} filters are invalid")
    explicit_fields = {item.get("field") for item in raw if isinstance(item, dict)}
    result = [item for item in defaults if item["field"] not in explicit_fields]
    for item in raw:
        result.append(_one_filter(item, pivot, label))
    return result


def _one_filter(item, pivot, label):
    if not isinstance(item, dict) or not isinstance(item.get("field"), int):
        raise AppError(422, "invalid", f"Pivot {label} filter is invalid")
    field_index = item["field"]
    if not 0 <= field_index < len(pivot["fields"]):
        raise AppError(422, "invalid", f"Pivot {label} field is invalid")
    selected = item.get("items")
    count = len(pivot["fields"][field_index]["shared_items"])
    if selected is not None and (not isinstance(selected, list) or any(
            not _valid_selection(value, count) for value in selected)):
        raise AppError(422, "invalid", f"Pivot {label} items are invalid")
    output = {"field": field_index, "items": selected}
    if label == "slicers":
        output["name"] = str(item.get("name", ""))
    return output


def _valid_selection(value, count):
    if isinstance(value, int) and not isinstance(value, bool):
        return 0 <= value < count
    if not isinstance(value, dict) or set(value) != {"kind", "value"}:
        return False
    kind, item = value["kind"], value["value"]
    if kind == "blank":
        return item is None
    if kind == "boolean":
        return isinstance(item, bool)
    if kind == "number":
        return isinstance(item, (int, float)) and not isinstance(item, bool) and math.isfinite(item)
    return kind in {"string", "date", "error"} and isinstance(item, str)


def _collapsed_options(raw, pivot):
    if not isinstance(raw, list):
        raise AppError(422, "invalid", "Pivot collapsed groups are invalid")
    valid_fields = set(pivot["row_fields"][:-1])
    output = []
    for item in raw:
        if not isinstance(item, dict) or item.get("field") not in valid_fields:
            raise AppError(422, "invalid", "Pivot collapsed group is invalid")
        field = pivot["fields"][item["field"]]
        value = item.get("item")
        if isinstance(value, bool) or not isinstance(value, int) or not 0 <= value < len(field["shared_items"]):
            raise AppError(422, "invalid", "Pivot collapsed item is invalid")
        output.append({"field": item["field"], "item": value})
    return output


def _result_key(snapshot, copy, query, options):
    encoded = json.dumps(options, sort_keys=True, separators=(",", ":"))
    return (snapshot.hash, query.copy_id, copy["input_revision"], copy["calculated_revision"],
            query.sheet, query.name, encoded)


def _latest_key(snapshot, query, options):
    encoded = json.dumps(options, sort_keys=True, separators=(",", ":"))
    return snapshot.hash, query.copy_id, query.sheet, query.name, encoded


def _incremental(pivot, source, overlays, options, previous, copy):
    if previous is None:
        return None
    old_overlays = previous["overlays"]
    missing = object()
    changed = [coordinate for coordinate in old_overlays.keys() | overlays.keys()
               if old_overlays.get(coordinate, missing) != overlays.get(coordinate, missing)]
    if not changed or len(changed) > INCREMENTAL_CHANGE_LIMIT:
        return None
    data_fields = {item["field"] for item in pivot["data_fields"]}
    sort_fields = {field["sort"]["rank_by"] for field in pivot["fields"]
                   if field["sort"]["data_source"]}
    if any(coordinate[1] - source.first_col not in data_fields - sort_fields
           for coordinate in changed):
        return None
    result = deepcopy(previous["result"])
    for coordinate in changed:
        if not _apply_measure_delta(result, pivot, source, old_overlays,
                                    overlays, options, coordinate):
            return None
    result["pivot"]["input_revision"] = copy["input_revision"]
    result["pivot"]["calculated_revision"] = copy["calculated_revision"]
    return result


def _apply_measure_delta(result, pivot, source, old, new, options, coordinate):
    row_number, col_number = coordinate
    field = col_number - source.first_col
    old_value = old.get(coordinate, _source_value(source, row_number, field))
    new_value = new.get(coordinate, _source_value(source, row_number, field))
    old_number, new_number = _summable(old_value), _summable(new_value)
    if old_number is None or new_number is None:
        return False
    current = _source_row(source, row_number, new)
    if not _matches(current, _clauses(pivot, options)):
        return True
    row_key = tuple(_dimension_value(current, current[1], item) for item in pivot["row_fields"])
    col_fields = [item for item in pivot["column_fields"] if item >= 0]
    column_key = tuple(_dimension_value(current, current[1], item) for item in col_fields)
    return _update_cells(result, pivot, field, row_key, column_key, new_number - old_number)


def _source_value(source, row_number, field):
    offset = row_number - source.first_row
    return source.rows[offset][source.positions[field]]


def _source_row(source, row_number, overlays):
    offset = row_number - source.first_row
    stored = source.rows[offset]
    values = tuple(overlays.get((row_number, source.first_col + field), stored[position])
                   for position, field in enumerate(source.fields))
    return source, values


def _summable(value):
    if isinstance(value, dict) and set(value) == {"error"}:
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value):
        return value
    return 0


def _update_cells(result, pivot, field, row_key, column_key, delta):
    rows = {item["r"] for item in result["rows"]
            if item["kind"] == "grand_total" or _row_path(pivot, item) == row_key[:len(item["path"])]}
    columns = {item["c"] for item in result["columns"]
               if item["data_field"] == field
               and (item["grand_total"] or _column_path(pivot, item) == column_key)}
    for cell in result["cells"]:
        if cell["r"] in rows and cell["c"] in columns and cell.get("data_field") == field:
            if isinstance(cell["v"], dict):
                return False
            cell["v"] = (cell["v"] or 0) + delta
    return True


def _row_path(pivot, row):
    return tuple(_selection_key(pivot["fields"][field], selector)
                 for field, selector in zip(pivot["row_fields"], row["path"]))


def _column_path(pivot, column):
    fields = [field for field in pivot["column_fields"] if field >= 0]
    return tuple(_selection_key(pivot["fields"][field], selector)
                 for field, selector in zip(fields, column["path"]))


def _refresh(pivot, slicers, source, overlays, options, copy, sheet):
    rows = _current_rows(source, overlays)
    clauses = _clauses(pivot, options)
    visible = [row for row in rows if _matches(row, clauses)]
    cube = _aggregate(pivot, visible)
    columns = _columns(pivot, cube)
    output_rows = _rows(pivot, cube, columns, options["collapsed"])
    cells = _cells(pivot, output_rows, columns, cube, options)
    return {
        "pivot": _pivot_summary(pivot, copy, sheet),
        "filters": _filter_summary(pivot, source, overlays, options),
        "columns": [_public_column(item) for item in columns],
        "rows": [_public_row(item) for item in output_rows], "cells": cells,
        "slicers": _slicer_states(pivot, slicers, rows, source, overlays, options),
    }


def _current_rows(source, overlays):
    changes = defaultdict(dict)
    for (row, col), value in overlays.items():
        field = col - source.first_col
        if field in source.positions:
            changes[row - source.first_row][field] = value
    output = []
    for offset, stored in enumerate(source.rows):
        values = stored
        if offset in changes:
            mutable = list(stored)
            for field, value in changes[offset].items():
                mutable[source.positions[field]] = value
            values = tuple(mutable)
        output.append((source, values))
    return output


def _value(rows, row, field):
    source, _ = rows
    return row[source.positions[field]]


def _clauses(pivot, options):
    clauses = []
    for origin in ("page", "slicers"):
        for item in options[origin]:
            if item["items"] is not None:
                field = pivot["fields"][item["field"]]
                selected = frozenset(_selection_key(field, value) for value in item["items"])
                clauses.append((origin, item.get("name", ""), item["field"], selected))
    for field_index in pivot["row_fields"] + [item for item in pivot["column_fields"] if item >= 0]:
        hidden = {item["cache_index"] for item in pivot["fields"][field_index]["items"] if item["hidden"]}
        if hidden:
            clauses.append(("hidden", "", field_index, frozenset(hidden)))
    return clauses


def _matches(rows, clauses, excluded=None):
    source, row = rows
    for origin, name, field, selected in clauses:
        if excluded == (origin, name, field):
            continue
        index = _dimension_value(rows, row, field)
        if origin == "hidden" and index in selected:
            return False
        if origin != "hidden" and index not in selected:
            return False
    return True


def _shared_key(item):
    value = item["value"]
    if item["kind"] == "missing":
        return ("blank", None)
    if item["kind"] == "boolean":
        return ("boolean", value)
    if item["kind"] == "number":
        return ("number", float(value))
    if item["kind"] == "error":
        return ("error", value["error"])
    if item["kind"] == "date":
        try:
            return ("number", float(to_excel(from_ISO8601(str(value)))))
        except (TypeError, ValueError):
            return ("text", str(value))
    return ("text", str(value))


def _selection_key(field, value):
    if isinstance(value, int):
        return value
    selector = ("text" if value["kind"] == "string" else value["kind"], value["value"])
    for index, shared in enumerate(field["shared_items"]):
        if _selector_from_shared(shared) == value:
            return index
    return ("value", selector)


def _selector_from_shared(shared):
    kind = {"missing": "blank", "boolean": "boolean", "number": "number",
            "string": "string", "date": "date", "error": "error"}[shared["kind"]]
    value = shared["value"]
    if kind == "error":
        value = value["error"]
    return {"kind": kind, "value": value}


def _selector(field, item):
    if isinstance(item, int):
        return _selector_from_shared(field["shared_items"][item])
    kind, value = item[1]
    return {"kind": "string" if kind == "text" else kind, "value": value}


def _value_key(value):
    if value is None:
        return ("blank", None)
    if isinstance(value, bool):
        return ("boolean", value)
    if isinstance(value, (int, float)) and math.isfinite(value):
        return ("number", float(value))
    if isinstance(value, dict) and set(value) == {"error"}:
        return ("error", value["error"])
    return ("text", str(value))


def _aggregate(pivot, rows):
    groups, row_keys, column_keys = {}, set(), set()
    row_fields = pivot["row_fields"]
    column_fields = [field for field in pivot["column_fields"] if field >= 0]
    for current in rows:
        source, row = current
        row_key = tuple(_dimension_value(current, row, field) for field in row_fields)
        column_key = tuple(_dimension_value(current, row, field) for field in column_fields)
        values = [_value(current, row, item["field"]) for item in pivot["data_fields"]]
        groups[(row_key, column_key)] = _add_values(groups.get((row_key, column_key)), values)
        row_keys.add(row_key)
        column_keys.add(column_key)
    return {"groups": groups, "row_keys": row_keys, "column_keys": column_keys}


def _dimension_value(current, row, field):
    source, _ = current
    return _dimension_from_value(source.metadata, field, _value(current, row, field))


def _dimension_from_value(metadata, field, value):
    index = metadata[field].get(_value_key(value))
    return index if index is not None else ("value", _value_key(value))


def _add_values(total, values):
    output = [0] * len(values) if total is None else list(total)
    for index, value in enumerate(values):
        if isinstance(output[index], dict):
            continue
        if isinstance(value, dict) and set(value) == {"error"}:
            output[index] = deepcopy(value)
        elif isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value):
            output[index] += value
    return tuple(output)


def _columns(pivot, cube):
    start_col = pivot["location"]["bounds"][1]
    column_fields = [field for field in pivot["column_fields"] if field >= 0]
    column_keys = _sort_column_keys(pivot, cube["column_keys"], column_fields)
    data_order = _data_order(pivot)
    result, offset = [], 1
    for column_key in column_keys or [()]:
        for data_index in data_order:
            data = pivot["data_fields"][data_index]
            prefix = _column_caption(pivot, column_fields, column_key)
            caption = f"{prefix} {data['name']}" if prefix else data["name"]
            result.append({"c": start_col + offset, "field": column_fields[-1]
                           if column_fields else -2, "data_field": data["field"],
                           "caption": caption, "number_format": data["number_format"],
                           "grand_total": False,
                           "path": [_selector(pivot["fields"][field], item)
                                    for field, item in zip(column_fields, column_key)],
                           "_column_key": column_key, "_data_index": data_index})
            offset += 1
    if column_fields and pivot["grand_totals"]["columns"]:
        for data_index in data_order:
            data = pivot["data_fields"][data_index]
            result.append({"c": start_col + offset, "field": -2,
                           "data_field": data["field"],
                           "caption": f"{pivot['captions']['grand_total']} {data['name']}",
                           "number_format": data["number_format"], "grand_total": True,
                           "path": [], "_column_key": None, "_data_index": data_index})
            offset += 1
    return result


def _data_order(pivot):
    result = [item["data_index"] for item in pivot["column_items"]
              if item["data_index"] is not None
              and 0 <= item["data_index"] < len(pivot["data_fields"])]
    result.extend(index for index in range(len(pivot["data_fields"])) if index not in result)
    return result


def _sort_column_keys(pivot, keys, fields):
    return sorted(keys, key=lambda key: tuple(
        _label_key(_caption(pivot["fields"][field], item))
        for field, item in zip(fields, key)
    ))


def _column_caption(pivot, fields, key):
    return " / ".join(_caption(pivot["fields"][field], item)
                      for field, item in zip(fields, key))


def _rows(pivot, cube, columns, collapsed):
    start_row = pivot["location"]["bounds"][0] + 1
    collapsed_set = {(item["field"], item["item"]) for item in collapsed}
    records = []
    if pivot["row_fields"]:
        _append_level(records, pivot, cube, (), 0, collapsed_set)
    if pivot["grand_totals"]["rows"]:
        records.append(_row_record("grand_total", 0, None, None,
                                   pivot["captions"]["grand_total"], False, False, ()))
    for offset, item in enumerate(records):
        item["r"] = start_row + offset
    return records


def _append_level(records, pivot, cube, prefix, level, collapsed):
    field_index = pivot["row_fields"][level]
    values = {key[level] for key in cube["row_keys"] if key[:level] == prefix}
    for item in _sort_items(pivot, cube, prefix, level, values):
        current = (*prefix, item)
        leaf = level == len(pivot["row_fields"]) - 1
        is_collapsed = (field_index, item) in collapsed
        kind = "item" if leaf else "subtotal"
        caption = _caption(pivot["fields"][field_index], item)
        if not leaf:
            caption = f"{caption} Total"
        record = _row_record(kind, level, field_index, item, caption,
                             not leaf, is_collapsed, current)
        record["selector"] = _selector(pivot["fields"][field_index], item)
        record["path"] = [_selector(pivot["fields"][field], value)
                          for field, value in zip(pivot["row_fields"], current)]
        if not isinstance(item, int):
            record["item"] = None
        subtotal_top = pivot["fields"][field_index]["subtotal_top"]
        if leaf or subtotal_top:
            records.append(record)
        if not leaf and not is_collapsed:
            _append_level(records, pivot, cube, current, level + 1, collapsed)
        if not leaf and not subtotal_top:
            records.append(record)


def _row_record(kind, level, field, item, caption, collapsible, collapsed, key):
    return {"r": 0, "kind": kind, "level": level, "field": field, "item": item,
            "caption": caption, "collapsible": collapsible, "collapsed": collapsed,
            "path": [], "_row_key": key}


def _sort_items(pivot, cube, prefix, level, values):
    field = pivot["fields"][pivot["row_fields"][level]]
    sort = field["sort"]
    if sort["data_source"] and sort["rank_by"] is not None:
        measure = next((index for index, item in enumerate(pivot["data_fields"])
                        if item["field"] == sort["rank_by"]), None)
        if measure is not None:
            return _data_sorted(values, pivot, cube, prefix, measure, sort["type"])
    if sort["type"] == "manual":
        rank = _manual_rank(pivot, field, level)
        return sorted(values, key=lambda value: (rank.get(value, len(rank)),
                                                  _label_key(_caption(field, value))))
    ordered = sorted(values, key=lambda value: _label_key(_caption(field, value)))
    blanks = [value for value in ordered if _is_blank(field, value)]
    data = [value for value in ordered if not _is_blank(field, value)]
    if sort["type"] == "descending":
        data.reverse()
    return data + blanks


def _manual_rank(pivot, field, level):
    order = []
    if level == 0:
        order.extend(item["indexes"][0] for item in pivot["row_items"]
                     if item["type"] == "data" and item["indexes"])
    order.extend(item["cache_index"] for item in field["items"]
                 if item["kind"] == "data" and item["cache_index"] not in order)
    return {value: index for index, value in enumerate(order)}


def _data_sorted(values, pivot, cube, prefix, measure, direction):
    def score(value):
        total = _group_value(cube, (*prefix, value), (), measure)
        return total if isinstance(total, (int, float)) else float("-inf")
    field = pivot["fields"][pivot["row_fields"][len(prefix)]]
    blanks = [value for value in values if _is_blank(field, value)]
    data = [value for value in values if not _is_blank(field, value)]
    return sorted(data, key=score, reverse=direction == "descending") + blanks


def _label_key(value):
    if value == "(blank)":
        return (2, "")
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return (0, value)
    return (1, str(value).casefold(), str(value))


def _is_blank(field, item):
    return isinstance(item, int) and field["shared_items"][item]["kind"] == "missing"


def _caption(field, item):
    if isinstance(item, tuple):
        value = item[1][1]
        return "(blank)" if item[1][0] == "blank" else value
    shared = field["shared_items"][item]
    if shared["kind"] == "missing":
        return "(blank)"
    if shared["kind"] == "boolean":
        return "TRUE" if shared["value"] else "FALSE"
    if shared["kind"] == "error":
        return shared["value"]["error"]
    return shared["value"]


def _group_value(cube, row_prefix, column_key, data_index):
    total = None
    for (row_key, current_column), values in cube["groups"].items():
        if row_key[:len(row_prefix)] != row_prefix:
            continue
        if column_key is not None and current_column != column_key:
            continue
        total = _add_values(total, [values[data_index]])
    return total[0] if total is not None else None


def _cells(pivot, rows, columns, cube, options):
    start_row, start_col = pivot["location"]["bounds"][:2]
    cells = _page_cells(pivot, options, start_row, start_col)
    cells.append(_cell(start_row, start_col, pivot["captions"]["row"], "General",
                       "pivot_header", "header"))
    cells.extend(_cell(start_row, item["c"], item["caption"], "General",
                       "pivot_header", "header", item["data_field"]) for item in columns)
    for row in rows:
        style = f"pivot_{row['kind']}"
        cells.append(_cell(row["r"], start_col, row["caption"], "General", style, row["kind"]))
        for column in columns:
            value = _group_value(cube, row["_row_key"], column["_column_key"], column["_data_index"])
            cells.append(_cell(row["r"], column["c"], value, column["number_format"],
                               style, row["kind"], column["data_field"]))
    return cells


def _page_cells(pivot, options, start_row, start_col):
    cells = []
    filters = defaultdict(list)
    for item in options["page"]:
        filters[item["field"]].append(item["items"])
    first = max(0, start_row - len(pivot["page_fields"]) - 1)
    for offset, page in enumerate(pivot["page_fields"]):
        field = pivot["fields"][page["field"]]
        selected = _intersection(filters[page["field"]], field)
        caption = _selection_caption(pivot["fields"][page["field"]], selected)
        cells.append(_cell(first + offset, start_col, page["caption"], "General",
                           "pivot_page", "page"))
        cells.append(_cell(first + offset, start_col + 1, caption, "General",
                           "pivot_page", "page"))
    return cells


def _intersection(items, field):
    selected = None
    for value in items:
        if value is not None:
            keys = {_selection_key(field, item) for item in value}
            selected = keys if selected is None else selected.intersection(keys)
    return selected


def _selection_caption(field, selected):
    if selected is None:
        return "(All)"
    if len(selected) == 1:
        return _caption(field, next(iter(selected)))
    return "(Multiple Items)" if selected else "(None)"


def _cell(row, col, value, number_format, style, row_kind, data_field=None):
    output = {"r": row, "c": col, "v": value, "number_format": number_format,
              "style": style, "row_kind": row_kind, "locked": True}
    if data_field is not None:
        output["data_field"] = data_field
    return output


def _pivot_summary(pivot, copy, sheet):
    return {"name": pivot["name"], "sheet": sheet, "location": deepcopy(pivot["location"]),
            "captions": deepcopy(pivot["captions"]), "style": deepcopy(pivot["style"]),
            "input_revision": copy["input_revision"],
            "calculated_revision": copy["calculated_revision"]}


def _slicer_states(pivot, slicers, rows, source, overlays, options):
    clauses = _clauses(pivot, options)
    output = []
    for slicer in slicers:
        field = pivot["fields"][slicer["field"]]
        selected = _slicer_selection(slicer, field, options)
        items = []
        excluded = ("slicers", slicer["name"], slicer["field"])
        imported = {item["cache_index"]: item for item in slicer["items"]}
        current_keys = set(_current_counts(source, overlays, slicer["field"]))
        available_keys = {_dimension_value(current, current[1], slicer["field"])
                          for current in rows if _matches(current, clauses, excluded)}
        keys = list(imported) + sorted(
            (item for item in current_keys if not isinstance(item, int)),
            key=lambda item: _label_key(_caption(field, item)),
        )
        for key in keys:
            source_item = imported.get(key, {"no_data": False})
            available = key in available_keys
            items.append({"cache_index": key if isinstance(key, int) else None,
                          "selector": _selector(field, key), "caption": _caption(field, key),
                          "selected": key in selected, "available": available,
                          "no_data": source_item["no_data"] or not available})
        output.append({"name": slicer["name"], "field": slicer["field"],
                       "caption": slicer["caption"], "style": slicer["style"],
                       "column_count": slicer["column_count"], "items": items})
    return output


def _slicer_selection(slicer, field, options):
    matching = [item["items"] for item in options["slicers"]
                if item.get("name") == slicer["name"]]
    selected = _intersection(matching, field)
    if selected is None:
        return {item["cache_index"] for item in slicer["items"]}
    return selected


def _filter_summary(pivot, source, overlays, options):
    result = deepcopy(options)
    for item in result["page"]:
        field = pivot["fields"][item["field"]]
        keys = set(_current_counts(source, overlays, item["field"]))
        item["available_items"] = [{"cache_index": key if isinstance(key, int) else None,
                                    "selector": _selector(field, key),
                                    "caption": _caption(field, key)}
                                   for key in sorted(keys, key=lambda key: _label_key(_caption(field, key)))]
    return result


def _current_counts(source, overlays, field):
    counts = Counter(source.counts.get(field, {}))
    column = source.first_col + field
    for (row, col), value in overlays.items():
        if col != column:
            continue
        old = _dimension_from_value(source.metadata, field, _source_value(source, row, field))
        new = _dimension_from_value(source.metadata, field, value)
        if old != new:
            counts[old] -= 1
            counts[new] += 1
    return {key: count for key, count in counts.items() if count > 0}


def _public_column(item):
    return {key: value for key, value in item.items() if not key.startswith("_")}


def _public_row(item):
    return {key: value for key, value in item.items() if not key.startswith("_")}
