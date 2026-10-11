"""Authorized personal dashboard view persistence."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
import json
import math

from .authorization import authorized_category_copies, effective_role
from .errors import AppError
from .workbook_service import TemplateCatalog


COPY_SLICER = "__pp_copy__"


@dataclass(frozen=True)
class DashboardViewQuery:
    copy_id: str
    user_id: str
    sheet: int
    pivot: str


class DashboardViewService:
    def __init__(self, database, snapshots):
        self.database = database
        self.templates = snapshots if isinstance(snapshots, TemplateCatalog) else TemplateCatalog(
            database, snapshots
        )

    def get(self, query: DashboardViewQuery) -> dict:
        with self.database.read_transaction() as connection:
            pivot, slicers = self._metadata(connection, query)
            row = connection.execute(
                "SELECT view_json,updated_at FROM pp_dashboard_views "
                "WHERE copy_id=? AND user_id=? AND sheet=? AND pivot_name=?",
                (query.copy_id, query.user_id, query.sheet, query.pivot),
            ).fetchone()
        if row is None:
            return _response(query.pivot, _default_view(pivot, slicers), None)
        try:
            view = _normalize(json.loads(row["view_json"]), pivot, slicers)
        except (AppError, json.JSONDecodeError):
            return _response(query.pivot, _default_view(pivot, slicers), None)
        return _response(query.pivot, view, row["updated_at"])

    def update(self, query: DashboardViewQuery, raw: dict, session_check) -> dict:
        with self.database.transaction() as connection:
            session_check(connection)
            pivot, slicers = self._metadata(connection, query)
            view = _normalize(raw, pivot, slicers)
            updated_at = datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")
            connection.execute(
                "INSERT INTO pp_dashboard_views(copy_id,user_id,sheet,pivot_name,view_json,updated_at) "
                "VALUES(?,?,?,?,?,?) ON CONFLICT(copy_id,user_id,sheet,pivot_name) DO UPDATE SET "
                "view_json=excluded.view_json,updated_at=excluded.updated_at",
                (query.copy_id, query.user_id, query.sheet, query.pivot,
                 json.dumps(view, sort_keys=True, separators=(",", ":")), updated_at),
            )
        return _response(query.pivot, view, updated_at)

    def get_category(self, query: DashboardViewQuery) -> dict:
        with self.database.read_transaction() as connection:
            category_id, pivot, slicers, copies = self._category_metadata(connection, query)
            row = connection.execute(
                "SELECT view_json,updated_at FROM pp_category_dashboard_views "
                "WHERE category_id=? AND user_id=? AND sheet=? AND pivot_name=?",
                (category_id, query.user_id, query.sheet, query.pivot),
            ).fetchone()
        if row is None:
            return _response(query.pivot, _default_category_view(pivot, slicers), None)
        try:
            view = _normalize_category(
                json.loads(row["view_json"]), pivot, slicers, copies, strict=False
            )
        except (AppError, json.JSONDecodeError):
            view = _default_category_view(pivot, slicers)
        return _response(query.pivot, view, row["updated_at"])

    def update_category(self, query: DashboardViewQuery, raw: dict, session_check) -> dict:
        with self.database.transaction() as connection:
            session_check(connection)
            category_id, pivot, slicers, copies = self._category_metadata(connection, query)
            view = _normalize_category(raw, pivot, slicers, copies, strict=True)
            updated_at = datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")
            connection.execute(
                "INSERT INTO pp_category_dashboard_views"
                "(category_id,user_id,sheet,pivot_name,view_json,updated_at) "
                "VALUES(?,?,?,?,?,?) ON CONFLICT(category_id,user_id,sheet,pivot_name) "
                "DO UPDATE SET view_json=excluded.view_json,updated_at=excluded.updated_at",
                (category_id, query.user_id, query.sheet, query.pivot,
                 json.dumps(view, sort_keys=True, separators=(",", ":")), updated_at),
            )
        return _response(query.pivot, view, updated_at)

    def _metadata(self, connection, query):
        copy = connection.execute(
            "SELECT * FROM pp_copies WHERE id=?", (query.copy_id,)
        ).fetchone()
        if copy is None or effective_role(connection, query.copy_id, query.user_id) is None:
            raise AppError(404, "not_found", "Copy not found")
        snapshot = self.templates.get(connection, copy["template_id"])
        if not isinstance(query.sheet, int) or not 0 <= query.sheet < len(snapshot.sheets):
            raise AppError(422, "invalid", "Pivot sheet is invalid")
        pivots = [item for item in snapshot.sheets[query.sheet].get("pivots", [])
                  if item["name"] == query.pivot]
        if len(pivots) != 1:
            raise AppError(404, "not_found", "Pivot table not found")
        pivot = pivots[0]
        slicers = [item for item in snapshot.sheets[query.sheet].get("slicers", [])
                   if item["pivot_cache_id"] == pivot["cache_id"]]
        return pivot, slicers

    def _category_metadata(self, connection, query):
        base, copies = authorized_category_copies(connection, query.copy_id, query.user_id)
        if base is None:
            raise AppError(404, "not_found", "Category dashboard not found")
        if any(copy["template_id"] != base["template_id"] for copy in copies):
            raise AppError(
                409, "incompatible_templates", "Category copies use incompatible templates"
            )
        snapshot = self.templates.get(connection, base["template_id"])
        if not isinstance(query.sheet, int) or not 0 <= query.sheet < len(snapshot.sheets):
            raise AppError(422, "invalid", "Pivot sheet is invalid")
        pivots = [item for item in snapshot.sheets[query.sheet].get("pivots", [])
                  if item["name"] == query.pivot]
        if len(pivots) != 1:
            raise AppError(404, "not_found", "Pivot table not found")
        pivot = pivots[0]
        slicers = [item for item in snapshot.sheets[query.sheet].get("slicers", [])
                   if item["pivot_cache_id"] == pivot["cache_id"]]
        return base["category_id"], pivot, slicers, copies


def _response(name, view, updated_at):
    return {"pivot": name, **view, "updated_at": updated_at}


def _default_view(pivot, slicers):
    page = [{"field": item["field"], "items": [_shared_selector(
        pivot["fields"][item["field"]]["shared_items"][item["cache_index"]]
    )] if item["cache_index"] is not None else None} for item in pivot["page_fields"]]
    filters = []
    for slicer in slicers:
        selected = [item for item in slicer["items"] if item["selected"]]
        items = None if len(selected) == len(slicer["items"]) else [
            _shared_selector(pivot["fields"][slicer["field"]]["shared_items"][item["cache_index"]])
            for item in selected
        ]
        filters.append({"name": slicer["name"], "field": slicer["field"], "items": items})
    return {"page": page, "slicers": filters, "collapsed": []}


def _default_category_view(pivot, slicers):
    view = _default_view(pivot, slicers)
    view["slicers"].append({"name": COPY_SLICER, "field": -1, "items": None})
    return view


def _normalize(raw, pivot, slicers):
    if not isinstance(raw, dict) or set(raw) != {"page", "slicers", "collapsed"}:
        raise AppError(422, "invalid", "Dashboard view is invalid")
    defaults = _default_view(pivot, slicers)
    page_fields = {item["field"] for item in pivot["page_fields"]}
    pages = _complete(defaults["page"], _filters(raw["page"], pivot, page_fields, None),
                      "field")
    slicer_map = {item["name"]: item["field"] for item in slicers}
    filters = _complete(
        defaults["slicers"],
        _filters(raw["slicers"], pivot, set(slicer_map.values()), slicer_map),
        "name",
    )
    collapsed = _collapsed(raw["collapsed"], pivot)
    return {"page": pages, "slicers": filters, "collapsed": collapsed}


def _normalize_category(raw, pivot, slicers, copies, strict):
    if not isinstance(raw, dict) or set(raw) != {"page", "slicers", "collapsed"}:
        raise AppError(422, "invalid", "Dashboard view is invalid")
    if not isinstance(raw["slicers"], list):
        raise AppError(422, "invalid", "Dashboard filters are invalid")
    copy_filters = [item for item in raw["slicers"]
                    if isinstance(item, dict) and item.get("name") == COPY_SLICER]
    imported = [item for item in raw["slicers"]
                if not (isinstance(item, dict) and item.get("name") == COPY_SLICER)]
    if len(copy_filters) > 1:
        raise AppError(422, "invalid", "Category copy filter is invalid")
    normalized = _normalize({**raw, "slicers": imported}, pivot, slicers)
    copy_filter = copy_filters[0] if copy_filters else {
        "name": COPY_SLICER, "field": -1, "items": None
    }
    normalized["slicers"].append(_normalize_copy_filter(copy_filter, copies, strict))
    return normalized


def _normalize_copy_filter(raw, copies, strict):
    if (not isinstance(raw, dict) or set(raw) != {"name", "field", "items"}
            or raw["name"] != COPY_SLICER or raw["field"] != -1):
        raise AppError(422, "invalid", "Category copy filter is invalid")
    items = raw["items"]
    if items is None:
        return {"name": COPY_SLICER, "field": -1, "items": None}
    if not isinstance(items, list):
        raise AppError(422, "invalid", "Category copy filter is invalid")
    allowed, selected = {copy["id"] for copy in copies}, []
    for item in items:
        valid = (isinstance(item, dict) and set(item) == {"kind", "value"}
                 and item["kind"] == "string" and isinstance(item["value"], str))
        if not valid or (strict and item["value"] not in allowed):
            raise AppError(422, "invalid", "Category copy filter is invalid")
        if item["value"] in allowed and item["value"] not in selected:
            selected.append(item["value"])
    if not strict and items and not selected:
        return {"name": COPY_SLICER, "field": -1, "items": None}
    if strict and len(selected) != len(items):
        raise AppError(422, "invalid", "Category copy filter is invalid")
    return {"name": COPY_SLICER, "field": -1,
            "items": [{"kind": "string", "value": value} for value in selected]}


def _filters(raw, pivot, allowed_fields, slicers):
    if not isinstance(raw, list):
        raise AppError(422, "invalid", "Dashboard filters are invalid")
    result, seen = [], set()
    for item in raw:
        keys = {"field", "items"} | ({"name"} if slicers is not None else set())
        if not isinstance(item, dict) or set(item) != keys:
            raise AppError(422, "invalid", "Dashboard filter is invalid")
        field_index = item["field"]
        if (isinstance(field_index, bool) or not isinstance(field_index, int)
                or field_index not in allowed_fields):
            raise AppError(422, "invalid", "Dashboard filter is invalid")
        name = item.get("name")
        if (slicers is not None and (not isinstance(name, str)
                                    or slicers.get(name) != field_index)):
            raise AppError(422, "invalid", "Dashboard slicer is invalid")
        identity = name if slicers is not None else field_index
        if identity in seen:
            raise AppError(422, "invalid", "Dashboard filter appears more than once")
        seen.add(identity)
        field = pivot["fields"][item["field"]]
        values = item["items"]
        if values is not None and not isinstance(values, list):
            raise AppError(422, "invalid", "Dashboard filter items are invalid")
        normalized = None if values is None else [_selector(value, field) for value in values]
        result.append({key: item[key] for key in keys - {"items"}} | {"items": normalized})
    return result


def _complete(defaults, explicit, key):
    selected = {item[key]: item for item in explicit}
    return [selected.get(item[key], item) for item in defaults]


def _collapsed(raw, pivot):
    if not isinstance(raw, list):
        raise AppError(422, "invalid", "Dashboard collapsed groups are invalid")
    allowed = set(pivot["row_fields"][:-1])
    result = []
    for item in raw:
        if not isinstance(item, dict) or set(item) != {"field", "selector"}:
            raise AppError(422, "invalid", "Dashboard collapsed group is invalid")
        field_index = item["field"]
        if (isinstance(field_index, bool) or not isinstance(field_index, int)
                or field_index not in allowed):
            raise AppError(422, "invalid", "Dashboard collapsed group is invalid")
        result.append({"field": field_index,
                       "selector": _selector(item["selector"], pivot["fields"][field_index])})
    return result


def _selector(value, field):
    if isinstance(value, int) and not isinstance(value, bool):
        if 0 <= value < len(field["shared_items"]):
            return _shared_selector(field["shared_items"][value])
        raise AppError(422, "invalid", "Dashboard selector index is invalid")
    if not isinstance(value, dict) or set(value) != {"kind", "value"}:
        raise AppError(422, "invalid", "Dashboard selector is invalid")
    kind, item = value["kind"], value["value"]
    if not isinstance(kind, str):
        raise AppError(422, "invalid", "Dashboard selector is invalid")
    valid = ((kind == "blank" and item is None)
             or (kind == "boolean" and isinstance(item, bool))
             or (kind == "number" and _finite_number(item))
             or (kind in {"string", "date", "error"} and isinstance(item, str)
                 and len(item) <= 32767))
    if not valid:
        raise AppError(422, "invalid", "Dashboard selector is invalid")
    return {"kind": kind, "value": item}


def _finite_number(value):
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return False
    try:
        return math.isfinite(value)
    except OverflowError:
        return False


def _shared_selector(shared):
    kind = {"missing": "blank", "boolean": "boolean", "number": "number",
            "string": "string", "date": "date", "error": "error"}[shared["kind"]]
    value = shared["value"]
    if kind == "error":
        value = value["error"]
    return {"kind": kind, "value": value}
