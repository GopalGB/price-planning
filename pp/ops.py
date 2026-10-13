"""Atomic, idempotent edit operation pipeline."""

from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass
from datetime import datetime, timezone

from .authorization import effective_role
from .errors import AppError, ErrorPayload
from .pivot_state import write_state
from .values import validate_value
from .workbook_service import TemplateCatalog


EDIT_ROLES = {"owner", "editor"}
ACTIONS = {"edit", "paste", "fill", "clear", "undo"}
UPLOAD_ACTION_PREFIX = "upload: "
UPLOAD_INVERSE_PREFIX = "upload-inverse: "
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


@dataclass(frozen=True)
class UploadOperationRequest:
    connection: object
    copy_id: str
    actor_id: str
    payload: dict
    session_check: object = None


@dataclass(frozen=True)
class OperationContext:
    connection: object
    copy_id: str
    actor_id: str
    payload: dict
    session_check: object
    trusted_action: str | None


@dataclass(frozen=True)
class InverseRequest:
    connection: object
    copy_id: str
    actor_id: str
    payload: dict


class OperationService:
    def __init__(self, database, snapshot):
        self.database = database
        self.templates = snapshot if isinstance(snapshot, TemplateCatalog) else TemplateCatalog(database, snapshot)

    def apply(self, copy_id: str, actor_id: str, payload: dict, session_check=None) -> dict:
        ack, _ = self.apply_with_status(copy_id, actor_id, payload, session_check)
        return ack

    def apply_with_status(self, copy_id: str, actor_id: str, payload: dict, session_check=None) -> tuple[dict, bool]:
        with self.database.transaction() as connection:
            ack, created, _ = self.apply_in_transaction(
                connection, copy_id, actor_id, payload, session_check
            )
            return ack, created

    def apply_in_transaction(self, connection, copy_id, actor_id, payload, session_check=None):
        return self._apply_transaction(OperationContext(
            connection, copy_id, actor_id, payload, session_check, None,
        ))

    def macro_replay(self, copy_id: str, actor_id: str, payload: dict) -> dict | None:
        """The stored ack when this user already committed this macro press; a reused op_id is refused."""
        op_id = _operation_id(payload)
        with self.database.connect() as connection:
            prior = connection.execute(
                "SELECT * FROM pp_ops WHERE copy_id=? AND op_id=?", (copy_id, op_id)
            ).fetchone()
        if prior is None:
            return None
        if prior["actor_id"] == actor_id and prior["payload_hash"] == _payload_hash(payload):
            return json.loads(prior["ack_json"])
        raise AppError(409, "op_id_reused", "This operation identifier was already used")

    def apply_macro(self, copy_id: str, actor_id: str, payload: dict, messages: list,
                    contents: list, expected_revision: int, action: str = "macro",
                    pivot_state: dict | None = None) -> dict | None:
        with self.database.transaction() as connection:
            return self.apply_macro_in_transaction(
                connection, copy_id, actor_id, payload, messages, contents, expected_revision, action, pivot_state)

    def apply_macro_in_transaction(self, connection, copy_id, actor_id, payload, messages, contents,
                                   expected_revision, action="macro", pivot_state=None):
        """Commit what a template macro wrote as one trusted operation (no version or input checks).

        ``pivot_state`` (slicers, refresh) is stored with the revision even when no cell changed.
        Returns None when nothing was written (no contents, or only blanks over template blanks).
        """
        op_id = _operation_id(payload)
        copy = connection.execute("SELECT * FROM pp_copies WHERE id=?", (copy_id,)).fetchone()
        if copy is None or effective_role(connection, copy_id, actor_id) not in EDIT_ROLES:
            raise AppError(403, "forbidden", "Editor access is required")
        if copy["input_revision"] + 1 != expected_revision:
            raise AppError(409, "conflict", "The copy changed while the macro ran")
        snapshot = self.templates.get(connection, copy["template_id"])
        cells = []
        for sheet, row, col, value in contents:
            current = connection.execute(
                "SELECT * FROM pp_overrides WHERE copy_id=? AND sheet=? AND row=? AND col=?",
                (copy_id, sheet, row, col),
            ).fetchone()
            before = json.loads(current["value_json"]) if current else snapshot.cell(sheet, row, col)["v"]
            if current is None and value is None and before is None:
                continue
            version = current["version"] if current else 0
            cells.append({"sheet": sheet, "row": row, "col": col, "value": value,
                          "before": before, "version": version + 1})
        if not cells and pivot_state is None:
            return None
        ack = {"op_id": op_id, "input_revision": expected_revision, "accepted": len(cells),
               "messages": list(messages)}
        self._commit(connection, OperationCommit(copy, actor_id, op_id, action, _payload_hash(payload), ack, cells))
        if pivot_state is not None:
            write_state(connection, copy_id, pivot_state, expected_revision)
        return ack

    def revisions(self, copy_id: str) -> dict:
        with self.database.connect() as connection:
            row = connection.execute(
                "SELECT input_revision,calculated_revision FROM pp_copies WHERE id=?", (copy_id,)
            ).fetchone()
        return {"input_revision": row["input_revision"], "calculated_revision": row["calculated_revision"]}

    def apply_upload_in_transaction(self, request):
        return self._apply_transaction(OperationContext(
            request.connection, request.copy_id, request.actor_id,
            request.payload, request.session_check, "upload",
        ))

    def _apply_transaction(self, request):
        payload_hash = _payload_hash(request.payload)
        op_id = _operation_id(request.payload)
        _require_session(request.connection, request.actor_id, request.session_check)
        copy = request.connection.execute(
            "SELECT * FROM pp_copies WHERE id=?", (request.copy_id,)
        ).fetchone()
        if copy is None or effective_role(
                request.connection, request.copy_id, request.actor_id) not in EDIT_ROLES:
            raise AppError(403, "forbidden", "Editor access is required")
        prior = request.connection.execute(
            "SELECT * FROM pp_ops WHERE copy_id=? AND op_id=?", (request.copy_id, op_id)
        ).fetchone()
        if prior is not None:
            if prior["actor_id"] == request.actor_id and prior["payload_hash"] == payload_hash:
                return json.loads(prior["ack_json"]), False, []
            raise AppError(409, "op_id_reused", "This operation identifier was already used")
        payload, trusted = _resolve_operation(InverseRequest(
            request.connection, request.copy_id, request.actor_id, request.payload,
        ), request.trusted_action)
        cells, _, action = _shape(payload, trusted)
        prepared = self._prepare(request.connection, copy, cells)
        revision = copy["input_revision"] + 1
        ack = {"op_id": op_id, "input_revision": revision, "accepted": len(prepared)}
        commit = OperationCommit(
            copy, request.actor_id, op_id, action, payload_hash, ack, prepared,
        )
        self._commit(request.connection, commit)
        return ack, True, prepared

    def _prepare(self, connection, copy, cells: list[dict]) -> list[dict]:
        # A values-only template (no Excel here) still saves typed values; its formulas keep Excel's saved results
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
            if formula_result is not None and not snapshot.template_input(sheet, row, col):
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
                 "cells": [_event_cell(item) for item in commit.cells]}
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


def _shape(payload, trusted_action=None):
    if not isinstance(payload, dict):
        raise AppError(422, "invalid", "Operation body must be an object")
    op_id, action, cells = payload.get("op_id"), payload.get("action"), payload.get("cells")
    if not isinstance(op_id, str) or not op_id or len(op_id) > 128:
        raise AppError(422, "invalid", "op_id is required")
    upload = isinstance(action, str) and action.startswith(UPLOAD_ACTION_PREFIX)
    inverse = isinstance(action, str) and action.startswith(UPLOAD_INVERSE_PREFIX)
    trusted_upload = trusted_action == "upload" and upload and len(action) <= 136
    trusted_inverse = trusted_action == "upload_inverse" and inverse and len(action) <= 160
    if not isinstance(action, str) or action not in ACTIONS and not trusted_upload and not trusted_inverse:
        raise AppError(422, "invalid", "Unknown edit action")
    if not isinstance(cells, list) or not cells:
        raise AppError(422, "invalid", "Submit at least one cell")
    if len(cells) > 10_000 and not (trusted_upload or trusted_inverse):
        raise AppError(422, "invalid", "Submit between 1 and 10,000 cells")
    return cells, op_id, action


def _operation_id(payload):
    if not isinstance(payload, dict):
        raise AppError(422, "invalid", "Operation body must be an object")
    op_id = payload.get("op_id")
    if not isinstance(op_id, str) or not op_id or len(op_id) > 128:
        raise AppError(422, "invalid", "op_id is required")
    return op_id


def _resolve_operation(request, trusted_action):
    source_id = request.payload.get("source_op_id") if isinstance(request.payload, dict) else None
    if source_id is None or trusted_action is not None:
        return request.payload, trusted_action
    if (request.payload.get("action") != "undo"
            or set(request.payload) != {"op_id", "action", "source_op_id"}
            or not isinstance(source_id, str) or not source_id):
        raise AppError(422, "invalid", "Upload undo reference is invalid")
    return _upload_inverse(request), "upload_inverse"


def _upload_inverse(request):
    source_id = request.payload["source_op_id"]
    source = request.connection.execute(
        "SELECT * FROM pp_ops WHERE copy_id=? AND op_id=?",
        (request.copy_id, source_id),
    ).fetchone()
    allowed = source and source["actor_id"] == request.actor_id and (
        source["action"].startswith(UPLOAD_ACTION_PREFIX)
        or source["action"].startswith(UPLOAD_INVERSE_PREFIX)
    )
    if not allowed:
        raise AppError(409, "invalid_lineage", "Upload undo lineage is invalid")
    inverse_action = f"{UPLOAD_INVERSE_PREFIX}{source_id}"
    replay = request.connection.execute(
        "SELECT 1 FROM pp_ops WHERE copy_id=? AND actor_id=? AND action=?",
        (request.copy_id, request.actor_id, inverse_action),
    ).fetchone()
    if replay is not None:
        raise AppError(409, "invalid_lineage", "Upload operation was already reversed")
    rows = _source_history(request.connection, request.copy_id, source)
    return {"op_id": request.payload["op_id"], "action": inverse_action, "cells": [
        {"sheet": row["sheet"], "r": row["row"], "c": row["col"],
         "v": json.loads(row["before_json"]), "expected_version": row["version"]}
        for row in rows
    ]}


def _source_history(connection, copy_id, source):
    rows = connection.execute(
        "SELECT * FROM pp_history WHERE copy_id=? AND op_id=? ORDER BY id",
        (copy_id, source["op_id"]),
    ).fetchall()
    try:
        accepted = json.loads(source["ack_json"])["accepted"]
    except (KeyError, TypeError, ValueError, json.JSONDecodeError) as exc:
        raise AppError(409, "invalid_lineage", "Upload undo lineage is invalid") from exc
    if (not rows or len(rows) != accepted
            or any(row["actor_id"] != source["actor_id"]
                   or row["action"] != source["action"] for row in rows)):
        raise AppError(409, "invalid_lineage", "Upload undo lineage is invalid")
    return rows


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


def _event_cell(item):
    """A formula is not a value: clients get the formula text and the result from the calc event."""
    formula = item["value"].get("formula") if isinstance(item["value"], dict) else None
    if formula is None:
        return {"sheet": item["sheet"], "r": item["row"], "c": item["col"], "v": item["value"], "ver": item["version"]}
    return {"sheet": item["sheet"], "r": item["row"], "c": item["col"], "v": None, "formula": formula,
            "ver": item["version"]}


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
