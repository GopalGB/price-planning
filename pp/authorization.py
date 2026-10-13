"""Copy role calculation shared by HTTP, operations, and live channels."""


ROLE_RANK = {None: 0, "viewer": 1, "editor": 2, "owner": 3}


def effective_role(connection, copy_id: str, user_id: str) -> str | None:
    copy = connection.execute("SELECT * FROM pp_copies WHERE id=?", (copy_id,)).fetchone()
    if copy is None:
        return None
    roles = []
    if copy["owner_id"] == user_id:
        roles.append("owner")
    share = connection.execute(
        "SELECT role FROM pp_copy_shares WHERE copy_id=? AND user_id=?", (copy_id, user_id)
    ).fetchone()
    if share is not None:
        roles.append(share["role"])
    if copy["visibility"] == "category" and copy["category_id"]:
        member = connection.execute(
            "SELECT 1 FROM pp_category_members WHERE category_id=? AND user_id=?",
            (copy["category_id"], user_id),
        ).fetchone()
        if member is not None:
            roles.append(copy["default_role"])
    return max(roles, key=lambda role: ROLE_RANK[role]) if roles else None


def authorized_category_copies(connection, copy_id: str, user_id: str):
    """Return the category context without disclosing inaccessible copies."""
    base = connection.execute("SELECT * FROM pp_copies WHERE id=?", (copy_id,)).fetchone()
    if base is None or effective_role(connection, copy_id, user_id) is None:
        return None, []
    category_id = base["category_id"]
    member = connection.execute(
        "SELECT 1 FROM pp_category_members WHERE category_id=? AND user_id=?",
        (category_id, user_id),
    ).fetchone() if category_id else None
    if member is None:
        return None, []
    rows = connection.execute(
        "SELECT * FROM pp_copies WHERE category_id=? ORDER BY id", (category_id,)
    ).fetchall()
    allowed = [dict(row) for row in rows
               if effective_role(connection, row["id"], user_id) is not None]
    return dict(base), allowed
