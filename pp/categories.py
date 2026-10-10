"""Category membership service."""

from dataclasses import dataclass
from datetime import datetime, timezone
import uuid

from .errors import AppError


@dataclass(frozen=True)
class MembershipChange:
    actor_id: str
    category_id: str
    email: object
    role: object


class CategoryService:
    def __init__(self, database):
        self.database = database

    def list(self, user_id: str) -> list[dict]:
        with self.database.connect() as connection:
            rows = connection.execute(
                "SELECT c.*,cm.role FROM pp_categories c JOIN pp_category_members cm ON cm.category_id=c.id WHERE cm.user_id=? ORDER BY c.name,c.id",
                (user_id,),
            ).fetchall()
            return [_category(connection, row, row["role"]) for row in rows]

    def create(self, user_id: str, name: object) -> dict:
        if not isinstance(name, str) or not name.strip() or len(name.strip()) > 120:
            raise AppError(422, "invalid", "Category name is required")
        category_id, now = uuid.uuid4().hex, _now()
        with self.database.transaction() as connection:
            connection.execute("INSERT INTO pp_categories(id,name,created_by,created_at) VALUES(?,?,?,?)",
                               (category_id, name.strip(), user_id, now))
            connection.execute("INSERT INTO pp_category_members(category_id,user_id,role) VALUES(?,?,'manager')",
                               (category_id, user_id))
            row = connection.execute("SELECT * FROM pp_categories WHERE id=?", (category_id,)).fetchone()
            return _category(connection, row, "manager")

    def update_member(self, change: MembershipChange) -> dict:
        if not isinstance(change.email, str) or not isinstance(change.role, (str, type(None))):
            raise AppError(422, "invalid", "Email and category role are invalid")
        if change.role not in {"manager", "member", None}:
            raise AppError(422, "invalid", "Email and category role are invalid")
        with self.database.transaction() as connection:
            _require_manager(connection, change.category_id, change.actor_id)
            target = connection.execute(
                "SELECT id FROM pp_users WHERE email=?", (change.email.strip().lower(),)
            ).fetchone()
            if target is None:
                raise AppError(404, "not_found", "User not found")
            if change.role is None:
                connection.execute("DELETE FROM pp_category_members WHERE category_id=? AND user_id=?",
                                   (change.category_id, target["id"]))
            else:
                connection.execute(
                    "INSERT INTO pp_category_members(category_id,user_id,role) VALUES(?,?,?) "
                    "ON CONFLICT(category_id,user_id) DO UPDATE SET role=excluded.role",
                    (change.category_id, target["id"], change.role),
                )
            row = connection.execute("SELECT * FROM pp_categories WHERE id=?", (change.category_id,)).fetchone()
            return _category(connection, row, "manager")


def _require_manager(connection, category_id, user_id):
    role = connection.execute(
        "SELECT role FROM pp_category_members WHERE category_id=? AND user_id=?", (category_id, user_id)
    ).fetchone()
    if role is None or role["role"] != "manager":
        raise AppError(403, "forbidden", "Category manager access is required")


def _category(connection, row, role):
    members = connection.execute(
        "SELECT u.id,u.name,u.email,cm.role FROM pp_category_members cm JOIN pp_users u ON u.id=cm.user_id WHERE cm.category_id=? ORDER BY u.name,u.id",
        (row["id"],),
    ).fetchall()
    return {"id": row["id"], "name": row["name"], "role": role,
            "members": [dict(member) for member in members]}


def _now():
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")
