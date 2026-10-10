"""Atomic, idempotent edit operation pipeline."""

from __future__ import annotations

import hashlib
import json
from datetime import datetime, timezone

from .authorization import effective_role
from .errors import AppError
from .values import validate_value


EDIT_ROLES = {"owner", "editor"}
ACTIONS = {"edit", "paste", "fill", "clear", "undo"}
FORMULA_PREFIXES = ("=", "+", "-", "@")


class OperationService:
    def __init__(self, database, snapshot):
        self.database = database
        self.snapshot = snapshot

    def apply(self, copy_id: str, actor_id: str, payload: dict) -> dict:
        payload_hash = _payload_hash(payload)
        cells, op_id, action = _shape(payload)
        with self.database.transaction() as connection:
            copy = connection.execute("SELECT * FROM pp_copies WHERE id=?", (copy_id,)).fetchone()
            if copy is None or effective_role(connection, copy_id, actor_id) not in EDIT_ROLES:
                raise AppError(403, "forbidden", "Editor access is required")
            prior = connection.execute(
                "SELECT * FROM pp_ops WHERE copy_id=? AND op_id=?", (copy_id, op_id)
            ).fetchone()
            if prior is not None:
                if prior["actor_id"] == actor_id and prior["payload_hash"] == payload_hash:
                    return json.loads(prior["ack_json"])
                raise AppError(409, "op_id_reused", "This operation identifier was already used")
            prepared = self._prepare(connection, copy_id, cells)
            revision = copy["input_revision"] + 1
            ack = {"op_id": op_id, "input_revision": revision, "accepted": len(prepared)}
            self._commit(connection, copy, actor_id, op_id, action, payload_hash, ack, prepared)
            return ack

    def _prepare(self, connection, copy_id: str, cells: list[dict]) -> list[dict]:
        prepared, conflicts, seen = [], [], set()
        for cell in cells:
            coordinate = _coordinate(cell)
            if coordinate in seen:
                raise AppError(422, "invalid", "A cell appears more than once")
            seen.add(coordinate)
            sheet, row, col = coordinate
            value = _value(cell.get("v"))
            if not self.snapshot.is_input(sheet, row, col):
                raise AppError(422, "locked", "This cell is locked", {"cells": [_public_cell(cell)]})
            _validate_rule(self.snapshot.validation(sheet, row, col), value, cell)
            current = connection.execute(
                "SELECT * FROM pp_overrides WHERE copy_id=? AND sheet=? AND row=? AND col=?",
                (copy_id, sheet, row, col),
            ).fetchone()
            version = current["version"] if current else 0
            if cell.get("expected_version") != version:
                conflicts.append(_conflict(connection, current, cell, version, self.snapshot))
                continue
            before = json.loads(current["value_json"]) if current else self.snapshot.cell(sheet, row, col)["v"]
            prepared.append({"sheet": sheet, "row": row, "col": col, "value": value,
                             "before": before, "version": version + 1})
        if conflicts:
            raise AppError(409, "conflict", "One or more cells changed", {"cells": conflicts})
        return prepared

    def _commit(self, connection, copy, actor_id, op_id, action, payload_hash, ack, cells):
        now, revision = _now(), ack["input_revision"]
        actor = connection.execute("SELECT id,name FROM pp_users WHERE id=?", (actor_id,)).fetchone()
        for cell in cells:
            values = (copy["id"], cell["sheet"], cell["row"], cell["col"], _json(cell["value"]),
                      cell["version"], actor_id, now)
            connection.execute(
                "INSERT INTO pp_overrides(copy_id,sheet,row,col,value_json,version,updated_by,updated_at) VALUES(?,?,?,?,?,?,?,?) "
                "ON CONFLICT(copy_id,sheet,row,col) DO UPDATE SET value_json=excluded.value_json,version=excluded.version,updated_by=excluded.updated_by,updated_at=excluded.updated_at",
                values,
            )
            connection.execute(
                "INSERT INTO pp_history(copy_id,op_id,revision,at,actor_id,action,sheet,row,col,before_json,after_json,version) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",
                (copy["id"], op_id, revision, now, actor_id, action, cell["sheet"], cell["row"],
                 cell["col"], _json(cell["before"]), _json(cell["value"]), cell["version"]),
            )
        event = {"type": "op", "op_id": op_id, "input_revision": revision,
                 "by": {"id": actor["id"], "name": actor["name"]}, "action": action,
                 "cells": [{"sheet": item["sheet"], "r": item["row"], "c": item["col"],
                            "v": item["value"], "ver": item["version"]} for item in cells]}
        connection.execute("UPDATE pp_copies SET input_revision=?,updated_at=? WHERE id=?",
                           (revision, now, copy["id"]))
        connection.execute("INSERT INTO pp_calc_jobs(copy_id,revision,status,created_at,updated_at) VALUES(?,?,'pending',?,?)",
                           (copy["id"], revision, now, now))
        connection.execute("INSERT INTO pp_events(copy_id,event_type,input_revision,payload_json,created_at) VALUES(?,'op',?,?,?)",
                           (copy["id"], revision, _json(event), now))
        connection.execute(
            "INSERT INTO pp_ops(copy_id,op_id,actor_id,payload_hash,ack_json,action,revision,created_at) VALUES(?,?,?,?,?,?,?,?)",
            (copy["id"], op_id, actor_id, payload_hash, _json(ack), action, revision, now),
        )

    def revision_changes(self, copy_id: str, revision: int) -> list[tuple[int, int, int, object]]:
        with self.database.connect() as connection:
            rows = connection.execute(
                "SELECT h.sheet,h.row,h.col,h.after_json FROM pp_history h JOIN "
                "(SELECT sheet,row,col,MAX(revision) AS revision FROM pp_history WHERE copy_id=? AND revision<=? GROUP BY sheet,row,col) latest "
                "ON latest.sheet=h.sheet AND latest.row=h.row AND latest.col=h.col AND latest.revision=h.revision "
                "WHERE h.copy_id=? ORDER BY h.sheet,h.row,h.col",
                (copy_id, revision, copy_id),
            ).fetchall()
            inherited = connection.execute(
                "SELECT sheet,row,col,value_json FROM pp_overrides WHERE copy_id=? AND origin_copy_id IS NOT NULL",
                (copy_id,),
            ).fetchall()
        values = {(row["sheet"], row["row"], row["col"]): json.loads(row["value_json"]) for row in inherited}
        values.update({(row["sheet"], row["row"], row["col"]): json.loads(row["after_json"]) for row in rows})
        return [(*coordinate, value) for coordinate, value in sorted(values.items())]


def _shape(payload):
    if not isinstance(payload, dict):
        raise AppError(422, "invalid", "Operation body must be an object")
    op_id, action, cells = payload.get("op_id"), payload.get("action"), payload.get("cells")
    if not isinstance(op_id, str) or not op_id or len(op_id) > 128:
        raise AppError(422, "invalid", "op_id is required")
    if action not in ACTIONS:
        raise AppError(422, "invalid", "Unknown edit action")
    if not isinstance(cells, list) or not 1 <= len(cells) <= 10_000:
        raise AppError(422, "invalid", "Submit between 1 and 10,000 cells")
    return cells, op_id, action


def _coordinate(cell):
    if not isinstance(cell, dict):
        raise AppError(422, "invalid", "Each cell must be an object")
    values = cell.get("sheet"), cell.get("r"), cell.get("c")
    if any(type(value) is not int or value < 0 for value in values):
        raise AppError(422, "invalid", "Cell coordinates must be non-negative integers")
    if type(cell.get("expected_version")) is not int or cell["expected_version"] < 0:
        raise AppError(422, "invalid", "expected_version must be a non-negative integer")
    return values


def _value(value):
    try:
        value = validate_value(value)
    except ValueError as exc:
        raise AppError(422, "invalid", str(exc)) from exc
    if isinstance(value, str) and value.startswith(FORMULA_PREFIXES):
        raise AppError(422, "invalid", "Formulas can't be typed into input cells")
    if isinstance(value, str) and len(value) > 32_767:
        raise AppError(422, "invalid", "Cell text is too long")
    return value


def _validate_rule(rule, value, cell):
    if rule is None or value is None:
        return
    if rule.get("type") == "list" and rule.get("list") and value not in rule["list"]:
        raise AppError(422, "invalid", "Value is outside the validation list", {"cells": [_public_cell(cell)]})
    if rule.get("type") in {"whole", "decimal"} and not isinstance(value, (int, float)):
        raise AppError(422, "invalid", "A numeric value is required", {"cells": [_public_cell(cell)]})


def _conflict(connection, current, cell, version, snapshot):
    user = None
    if current is not None:
        row = connection.execute("SELECT id,name FROM pp_users WHERE id=?", (current["updated_by"],)).fetchone()
        user = {"id": row["id"], "name": row["name"]}
    value = json.loads(current["value_json"]) if current else snapshot.cell(cell["sheet"], cell["r"], cell["c"])["v"]
    return {**_public_cell(cell), "current_version": version, "current_v": value, "by": user}


def _public_cell(cell):
    return {key: cell.get(key) for key in ("sheet", "r", "c")}


def _payload_hash(payload):
    try:
        encoded = json.dumps(payload, sort_keys=True, separators=(",", ":"), allow_nan=False).encode()
    except (TypeError, ValueError) as exc:
        raise AppError(422, "invalid", "Operation contains an unsupported value") from exc
    return hashlib.sha256(encoded).hexdigest()


def _json(value):
    return json.dumps(value, separators=(",", ":"), sort_keys=True, allow_nan=False)


def _now():
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")
