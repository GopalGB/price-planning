"""Atomic, idempotent edit operation pipeline."""

from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass
from datetime import datetime, timezone

from .authorization import effective_role
from .errors import AppError, ErrorPayload
from .values import validate_value
from .workbook_service import TemplateCatalog


EDIT_ROLES = {"owner", "editor"}
ACTIONS = {"edit", "paste", "fill", "clear", "undo"}
FORMULA_PREFIXES = ("=", "+", "-", "@")


@dataclass(frozen=True)
class OperationCommit:
    copy: object
    actor_id: str
    op_id: str
    action: str
    payload_hash: str
    ack: dict
    cells: list[dict]


@dataclass(frozen=True)
class ConflictState:
    current: object
    cell: dict
    version: int
    snapshot: object


class OperationService:
    def __init__(self, database, snapshot):
        self.database = database
        self.templates = snapshot if isinstance(snapshot, TemplateCatalog) else TemplateCatalog(database, snapshot)

    def apply(self, copy_id: str, actor_id: str, payload: dict, session_check=None) -> dict:
        ack, _ = self.apply_with_status(copy_id, actor_id, payload, session_check)
        return ack

    def apply_with_status(self, copy_id: str, actor_id: str, payload: dict, session_check=None) -> tuple[dict, bool]:
        payload_hash = _payload_hash(payload)
        cells, op_id, action = _shape(payload)
        with self.database.transaction() as connection:
            _require_session(connection, actor_id, session_check)
            copy = connection.execute("SELECT * FROM pp_copies WHERE id=?", (copy_id,)).fetchone()
            if copy is None or effective_role(connection, copy_id, actor_id) not in EDIT_ROLES:
                raise AppError(403, "forbidden", "Editor access is required")
            prior = connection.execute(
                "SELECT * FROM pp_ops WHERE copy_id=? AND op_id=?", (copy_id, op_id)
            ).fetchone()
            if prior is not None:
                if prior["actor_id"] == actor_id and prior["payload_hash"] == payload_hash:
                    return json.loads(prior["ack_json"]), False
                raise AppError(409, "op_id_reused", "This operation identifier was already used")
            prepared = self._prepare(connection, copy, cells)
            revision = copy["input_revision"] + 1
            ack = {"op_id": op_id, "input_revision": revision, "accepted": len(prepared)}
            commit = OperationCommit(copy, actor_id, op_id, action, payload_hash, ack, prepared)
            self._commit(connection, commit)
            return ack, True

    def _prepare(self, connection, copy, cells: list[dict]) -> list[dict]:
        snapshot = self.templates.get(connection, copy["template_id"])
        prepared, conflicts, seen = [], [], set()
        for cell in cells:
            coordinate = _coordinate(cell)
            if coordinate in seen:
                raise AppError(422, "invalid", "A cell appears more than once")
            seen.add(coordinate)
            sheet, row, col = coordinate
            value = _value(cell.get("v"))
            formula_result = connection.execute(
                "SELECT 1 FROM pp_results WHERE copy_id=? AND sheet=? AND row=? AND col=?",
                (copy["id"], sheet, row, col),
            ).fetchone()
            if formula_result is not None:
                raise AppError(422, "locked", ErrorPayload(
                    "This cell is locked", {"cells": [_public_cell(cell)]}
                ))
            if (sheet >= len(snapshot.sheets) or row >= snapshot.sheets[sheet]["rows"]
                    or col >= snapshot.sheets[sheet]["cols"]):
                raise AppError(422, "invalid", "Cell coordinate is outside the workbook")
            if not snapshot.is_input(sheet, row, col):
                raise AppError(422, "locked", ErrorPayload(
                    "This cell is locked", {"cells": [_public_cell(cell)]}
                ))
            _validate_rule(snapshot.validation(sheet, row, col), value, cell)
            current = connection.execute(
                "SELECT * FROM pp_overrides WHERE copy_id=? AND sheet=? AND row=? AND col=?",
                (copy["id"], sheet, row, col),
            ).fetchone()
            version = current["version"] if current else 0
            if cell.get("expected_version") != version:
                conflicts.append(_conflict(connection, ConflictState(current, cell, version, snapshot)))
                continue
            before = json.loads(current["value_json"]) if current else snapshot.cell(sheet, row, col)["v"]
            prepared.append({"sheet": sheet, "row": row, "col": col, "value": value,
                             "before": before, "version": version + 1})
        if conflicts:
            raise AppError(409, "conflict", ErrorPayload("One or more cells changed", {"cells": conflicts}))
        return prepared

    def _commit(self, connection, commit):
        now, revision = _now(), commit.ack["input_revision"]
        actor = connection.execute("SELECT id,name FROM pp_users WHERE id=?", (commit.actor_id,)).fetchone()
        for cell in commit.cells:
            values = (commit.copy["id"], cell["sheet"], cell["row"], cell["col"], _json(cell["value"]),
                      cell["version"], commit.actor_id, now)
            connection.execute(
                "INSERT INTO pp_overrides(copy_id,sheet,row,col,value_json,version,updated_by,updated_at) VALUES(?,?,?,?,?,?,?,?) "
                "ON CONFLICT(copy_id,sheet,row,col) DO UPDATE SET value_json=excluded.value_json,"
                "version=excluded.version,updated_by=excluded.updated_by,updated_at=excluded.updated_at,"
                "origin_copy_id=NULL,origin_revision=NULL",
                values,
            )
            connection.execute(
                "INSERT INTO pp_history(copy_id,op_id,revision,at,actor_id,action,sheet,row,col,before_json,after_json,version) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",
                (commit.copy["id"], commit.op_id, revision, now, commit.actor_id, commit.action, cell["sheet"], cell["row"],
                 cell["col"], _json(cell["before"]), _json(cell["value"]), cell["version"]),
            )
        event = {"type": "op", "op_id": commit.op_id, "input_revision": revision, "at": now,
                 "by": {"id": actor["id"], "name": actor["name"]}, "action": commit.action,
                 "cells": [{"sheet": item["sheet"], "r": item["row"], "c": item["col"],
                            "v": item["value"], "ver": item["version"]} for item in commit.cells]}
        connection.execute("UPDATE pp_copies SET input_revision=?,updated_at=? WHERE id=?",
                           (revision, now, commit.copy["id"]))
        connection.execute("INSERT INTO pp_calc_jobs(copy_id,revision,status,created_at,updated_at) VALUES(?,?,'pending',?,?)",
                           (commit.copy["id"], revision, now, now))
        connection.execute("INSERT INTO pp_events(copy_id,event_type,input_revision,payload_json,created_at) VALUES(?,'op',?,?,?)",
                           (commit.copy["id"], revision, _json(event), now))
        connection.execute(
            "INSERT INTO pp_ops(copy_id,op_id,actor_id,payload_hash,ack_json,action,revision,created_at) VALUES(?,?,?,?,?,?,?,?)",
            (commit.copy["id"], commit.op_id, commit.actor_id, commit.payload_hash,
             _json(commit.ack), commit.action, revision, now),
        )

    def revision_changes(self, copy_id: str, revision: int) -> list[tuple[int, int, int, object]]:
        with self.database.connect() as connection:
            connection.execute("BEGIN")
            copy = connection.execute(
                "SELECT input_revision FROM pp_copies WHERE id=?", (copy_id,)
            ).fetchone()
            rows = connection.execute(
                "SELECT h.sheet,h.row,h.col,h.after_json FROM pp_history h JOIN "
                "(SELECT sheet,row,col,MAX(revision) AS revision FROM pp_history WHERE copy_id=? AND revision<=? GROUP BY sheet,row,col) latest "
                "ON latest.sheet=h.sheet AND latest.row=h.row AND latest.col=h.col AND latest.revision=h.revision "
                "WHERE h.copy_id=? ORDER BY h.sheet,h.row,h.col",
                (copy_id, revision, copy_id),
            ).fetchall()
            baseline = connection.execute(
                "SELECT sheet,row,col,value_json FROM pp_copy_baselines WHERE copy_id=?",
                (copy_id,),
            ).fetchall()
            current = connection.execute(
                "SELECT sheet,row,col,value_json FROM pp_overrides WHERE copy_id=?",
                (copy_id,),
            ).fetchall() if copy is not None and revision == copy["input_revision"] else []
        values = {(row["sheet"], row["row"], row["col"]): json.loads(row["value_json"]) for row in baseline}
        values.update({(row["sheet"], row["row"], row["col"]): json.loads(row["after_json"]) for row in rows})
        values.update({(row["sheet"], row["row"], row["col"]): json.loads(row["value_json"]) for row in current})
        return [(*coordinate, value) for coordinate, value in sorted(values.items())]


def _require_session(connection, actor_id, session_check):
    if session_check is None:
        return
    user, _ = session_check(connection)
    if user["id"] != actor_id:
        raise AppError(401, "unauthorized", "Session user changed")


def _shape(payload):
    if not isinstance(payload, dict):
        raise AppError(422, "invalid", "Operation body must be an object")
    op_id, action, cells = payload.get("op_id"), payload.get("action"), payload.get("cells")
    if not isinstance(op_id, str) or not op_id or len(op_id) > 128:
        raise AppError(422, "invalid", "op_id is required")
    if not isinstance(action, str) or action not in ACTIONS:
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
    if not _blocking_validation(rule) or value is None:
        return
    kind = rule.get("type")
    if kind == "list":
        if "list" not in rule:
            _validation_error("invalid", "Reference validation is not supported", cell)
        if value not in rule["list"]:
            _validation_error("invalid", "Value is outside the validation list", cell)
        return
    if kind == "textLength":
        if not isinstance(value, str):
            _validation_error("invalid", "Text is required by validation", cell)
        _validate_static_rule(len(value), rule, cell)
        return
    if kind not in {"whole", "decimal", "date", "time"}:
        _validation_error("invalid", "Validation formula is not supported", cell)
    numeric = isinstance(value, (int, float)) and not isinstance(value, bool)
    integral = not isinstance(value, float) or value.is_integer()
    if not numeric or kind == "whole" and not integral:
        _validation_error("invalid", "A numeric value is required", cell)
    _validate_static_rule(value, rule, cell)


def _blocking_validation(rule):
    if rule is None or rule.get("type") in {None, "none"}:
        return False
    return rule.get("showErrorMessage") is True and rule.get("errorStyle", "stop") == "stop"


def _validate_static_rule(value, rule, cell):
    valid = _numeric_rule(value, rule)
    if valid is None:
        _validation_error("invalid", "Validation bounds are not static", cell)
    if not valid:
        _validation_error("invalid", "Value is outside the validation bounds", cell)


def _validation_error(code, message, cell):
    raise AppError(422, code, ErrorPayload(message, {"cells": [_public_cell(cell)]}))


def _numeric_rule(value, rule):
    lower, upper, operator = rule.get("min"), rule.get("max"), rule.get("operator", "between")
    if not _is_number(lower) or operator in {"between", "notBetween"} and not _is_number(upper):
        return None
    if operator == "between":
        return lower <= value <= upper
    if operator == "notBetween":
        return not lower <= value <= upper
    checks = {"equal": value == lower, "notEqual": value != lower,
              "greaterThan": value > lower, "lessThan": value < lower,
              "greaterThanOrEqual": value >= lower, "lessThanOrEqual": value <= lower}
    return checks.get(operator)


def _is_number(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def _conflict(connection, state):
    user = None
    if state.current is not None:
        row = connection.execute("SELECT id,name FROM pp_users WHERE id=?", (state.current["updated_by"],)).fetchone()
        user = {"id": row["id"], "name": row["name"]}
    cell = state.cell
    value = (json.loads(state.current["value_json"]) if state.current else
             state.snapshot.cell(cell["sheet"], cell["r"], cell["c"])["v"])
    return {**_public_cell(cell), "current_version": state.version, "current_v": value, "by": user}


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
