"""Copy creation, visibility, sharing, and fork service."""

from dataclasses import dataclass
from datetime import datetime, timezone
import uuid

from .authorization import effective_role
from .errors import AppError


@dataclass(frozen=True)
class ShareChange:
    actor_id: str
    copy_id: str
    email: object
    role: object


class CopyService:
    def __init__(self, database, template_id: str):
        self.database = database
        self.template_id = template_id

    def list(self, user_id: str, category_id: str | None = None) -> list[dict]:
        with self.database.connect() as connection:
            rows = connection.execute(
                "SELECT * FROM pp_copies WHERE (? IS NULL OR category_id=?) ORDER BY updated_at DESC,id",
                (category_id, category_id),
            ).fetchall()
            return [_copy(connection, row, role) for row in rows
                    if (role := effective_role(connection, row["id"], user_id)) is not None]

    def create(self, user_id: str, payload: dict, session_check=None) -> dict:
        with self.database.transaction() as connection:
            return self.create_in_transaction(connection, user_id, payload, session_check)

    def create_in_transaction(self, connection, user_id, payload, session_check=None):
        name, category_id, visibility, default_role = _copy_fields(payload)
        _require_session(connection, user_id, session_check)
        member = connection.execute(
            "SELECT 1 FROM pp_category_members WHERE category_id=? AND user_id=?", (category_id, user_id)
        ).fetchone()
        if member is None:
            raise AppError(403, "forbidden", "Category membership is required")
        copy_id, now = uuid.uuid4().hex, _now()
        connection.execute(
            "INSERT INTO pp_copies(id,name,category_id,owner_id,template_id,visibility,default_role,input_revision,calculated_revision,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
            (copy_id, name, category_id, user_id, self.template_id, visibility, default_role, 0, -1, now, now),
        )
        _queue_baseline(connection, copy_id, now)
        row = connection.execute("SELECT * FROM pp_copies WHERE id=?", (copy_id,)).fetchone()
        return _copy(connection, row, "owner")

    def fork(self, user_id: str, copy_id: str, name: object, session_check=None) -> dict:
        if not isinstance(name, str) or not name.strip():
            raise AppError(422, "invalid", "Fork name is required")
        with self.database.transaction() as connection:
            _require_session(connection, user_id, session_check)
            parent = connection.execute("SELECT * FROM pp_copies WHERE id=?", (copy_id,)).fetchone()
            if parent is None or effective_role(connection, copy_id, user_id) is None:
                raise AppError(404, "not_found", "Copy not found")
            fork_id, now = uuid.uuid4().hex, _now()
            connection.execute(
                "INSERT INTO pp_copies(id,name,category_id,owner_id,template_id,visibility,default_role,parent_copy_id,fork_revision,input_revision,calculated_revision,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (fork_id, name.strip()[:120], parent["category_id"], user_id, parent["template_id"], "private",
                 "viewer", copy_id, parent["input_revision"], 0, -1, now, now),
            )
            connection.execute(
                "INSERT INTO pp_overrides(copy_id,sheet,row,col,value_json,version,updated_by,updated_at,origin_copy_id,origin_revision) "
                "SELECT ?,sheet,row,col,value_json,version,updated_by,updated_at,?,? "
                "FROM pp_overrides WHERE copy_id=?",
                (fork_id, copy_id, parent["input_revision"], copy_id),
            )
            connection.execute(
                "INSERT INTO pp_copy_baselines(copy_id,sheet,row,col,value_json) "
                "SELECT ?,sheet,row,col,value_json FROM pp_overrides WHERE copy_id=?",
                (fork_id, copy_id),
            )
            _queue_baseline(connection, fork_id, now)
            row = connection.execute("SELECT * FROM pp_copies WHERE id=?", (fork_id,)).fetchone()
            return _copy(connection, row, "owner")

    def get(self, user_id: str, copy_id: str) -> dict:
        with self.database.connect() as connection:
            row = connection.execute("SELECT * FROM pp_copies WHERE id=?", (copy_id,)).fetchone()
            role = effective_role(connection, copy_id, user_id) if row is not None else None
            if role is None:
                raise AppError(404, "not_found", "Copy not found")
            return _copy(connection, row, role)

    def patch(self, user_id: str, copy_id: str, payload: dict, session_check=None) -> dict:
        allowed = {key: value for key, value in payload.items() if key in {"name", "visibility", "default_role"}}
        if not allowed:
            raise AppError(422, "invalid", "No copy fields were supplied")
        _validate_patch(allowed)
        with self.database.transaction() as connection:
            _require_session(connection, user_id, session_check)
            row = _owner_copy(connection, copy_id, user_id)
            values = {"name": row["name"], "visibility": row["visibility"], "default_role": row["default_role"], **allowed}
            connection.execute("UPDATE pp_copies SET name=?,visibility=?,default_role=?,updated_at=? WHERE id=?",
                               (values["name"], values["visibility"], values["default_role"], _now(), copy_id))
            row = connection.execute("SELECT * FROM pp_copies WHERE id=?", (copy_id,)).fetchone()
            return _copy(connection, row, "owner")

    def share(self, change: ShareChange, session_check=None) -> dict:
        if not isinstance(change.email, str) or not isinstance(change.role, (str, type(None))):
            raise AppError(422, "invalid", "Email and share role are invalid")
        if change.role not in {"editor", "viewer", None}:
            raise AppError(422, "invalid", "Email and share role are invalid")
        with self.database.transaction() as connection:
            _require_session(connection, change.actor_id, session_check)
            row = _owner_copy(connection, change.copy_id, change.actor_id)
            target = connection.execute(
                "SELECT id FROM pp_users WHERE email=?", (change.email.strip().lower(),)
            ).fetchone()
            if target is None:
                raise AppError(404, "not_found", "User not found")
            if change.role is None:
                connection.execute("DELETE FROM pp_copy_shares WHERE copy_id=? AND user_id=?",
                                   (change.copy_id, target["id"]))
            else:
                connection.execute(
                    "INSERT INTO pp_copy_shares(copy_id,user_id,role) VALUES(?,?,?) "
                    "ON CONFLICT(copy_id,user_id) DO UPDATE SET role=excluded.role",
                    (change.copy_id, target["id"], change.role),
                )
            return _copy(connection, row, "owner")


def _copy_fields(payload):
    if not isinstance(payload, dict):
        raise AppError(422, "invalid", "Copy body must be an object")
    values = payload.get("name"), payload.get("category_id"), payload.get("visibility"), payload.get("default_role")
    if not isinstance(values[0], str) or not values[0].strip() or not isinstance(values[1], str):
        raise AppError(422, "invalid", "Copy name and category are required")
    if not isinstance(values[2], str) or not isinstance(values[3], str):
        raise AppError(422, "invalid", "Copy visibility or role is invalid")
    if values[2] not in {"private", "category"} or values[3] not in {"editor", "viewer"}:
        raise AppError(422, "invalid", "Copy visibility or role is invalid")
    return values[0].strip()[:120], values[1], values[2], values[3]


def _validate_patch(payload):
    if "name" in payload and (not isinstance(payload["name"], str) or not payload["name"].strip()):
        raise AppError(422, "invalid", "Copy name is invalid")
    if "visibility" in payload and (not isinstance(payload["visibility"], str)
                                     or payload["visibility"] not in {"private", "category"}):
        raise AppError(422, "invalid", "Copy visibility is invalid")
    if "default_role" in payload and (not isinstance(payload["default_role"], str)
                                       or payload["default_role"] not in {"editor", "viewer"}):
        raise AppError(422, "invalid", "Copy role is invalid")


def _owner_copy(connection, copy_id, user_id):
    row = connection.execute("SELECT * FROM pp_copies WHERE id=?", (copy_id,)).fetchone()
    if row is None or row["owner_id"] != user_id:
        raise AppError(403, "forbidden", "Copy owner access is required")
    return row


def _require_session(connection, user_id, session_check):
    if session_check is None:
        return
    user, _ = session_check(connection)
    if user["id"] != user_id:
        raise AppError(401, "unauthorized", "Session user changed")


def _queue_baseline(connection, copy_id, now):
    connection.execute(
        "INSERT INTO pp_calc_jobs(copy_id,revision,status,created_at,updated_at) VALUES(?,0,'pending',?,?)",
        (copy_id, now, now),
    )


def _copy(connection, row, role):
    owner = connection.execute("SELECT id,name FROM pp_users WHERE id=?", (row["owner_id"],)).fetchone()
    members = connection.execute(
        "SELECT u.id,u.name,s.role FROM pp_copy_shares s JOIN pp_users u ON u.id=s.user_id WHERE s.copy_id=? ORDER BY u.name,u.id",
        (row["id"],),
    ).fetchall()
    parent = None
    if row["parent_copy_id"]:
        source = connection.execute("SELECT id,name FROM pp_copies WHERE id=?", (row["parent_copy_id"],)).fetchone()
        parent = {"id": source["id"], "name": source["name"], "revision": row["fork_revision"]}
    edit = connection.execute(
        "SELECT h.at,u.id,u.name FROM pp_history h JOIN pp_users u ON u.id=h.actor_id WHERE h.copy_id=? ORDER BY h.id DESC LIMIT 1",
        (row["id"],),
    ).fetchone()
    return {"id": row["id"], "name": row["name"], "category_id": row["category_id"],
            "owner": dict(owner), "visibility": row["visibility"], "default_role": row["default_role"],
            "my_role": role, "input_revision": row["input_revision"],
            "calculated_revision": row["calculated_revision"], "parent": parent,
            "last_edit": {"by": {"id": edit["id"], "name": edit["name"]}, "at": edit["at"]} if edit else None,
            "members": [dict(member) for member in members]}


def _now():
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")
