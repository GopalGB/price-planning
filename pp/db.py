"""SQLite connection and transaction boundary for the v2 backend."""

from __future__ import annotations

from contextlib import contextmanager
import hashlib
from pathlib import Path
import sqlite3

from .schema import SCHEMA


class Database:
    def __init__(self, path: str | Path):
        self.path = Path(path)
        self._closed = False

    def _open(self) -> sqlite3.Connection:
        if self._closed:
            raise RuntimeError("database is closed")
        connection = sqlite3.connect(self.path, timeout=5)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys=ON")
        connection.execute("PRAGMA busy_timeout=5000")
        return connection

    def initialize(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with self._open() as connection:
            connection.execute("PRAGMA journal_mode=WAL")
            connection.executescript(SCHEMA)
            self._migrate_users(connection)

    @contextmanager
    def connect(self):
        connection = self._open()
        try:
            yield connection
        finally:
            connection.close()

    @contextmanager
    def transaction(self):
        connection = self._open()
        try:
            connection.execute("BEGIN IMMEDIATE")
            yield connection
            connection.commit()
        except Exception:
            connection.rollback()
            raise
        finally:
            connection.close()

    def _migrate_users(self, connection: sqlite3.Connection) -> None:
        if not _table_exists(connection, "users"):
            return
        columns = _columns(connection, "users")
        required = {"id", "email", "name", "password", "created"}
        if not required.issubset(columns):
            return
        connection.execute(
            "INSERT OR IGNORE INTO pp_users(id,email,name,password_hash,created_at) "
            "SELECT id,email,name,password,CAST(created AS TEXT) FROM users"
        )

    def migrate_legacy(self, template_id: str) -> None:
        with self.transaction() as connection:
            self._migrate_users(connection)
            if not _table_exists(connection, "scenarios"):
                return
            connection.execute(
                "INSERT OR IGNORE INTO pp_copies(id,name,owner_id,template_id,visibility,default_role,input_revision,calculated_revision,created_at,updated_at) "
                "SELECT id,name,owner,?,'private','viewer',revision,0,CAST(created AS TEXT),CAST(updated AS TEXT) FROM scenarios",
                (template_id,),
            )
            _migrate_members(connection)
            _migrate_edits(connection)

    def close(self) -> None:
        self._closed = True


def _table_exists(connection: sqlite3.Connection, name: str) -> bool:
    row = connection.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (name,)).fetchone()
    return row is not None


def _columns(connection: sqlite3.Connection, table: str) -> set[str]:
    return {row[1] for row in connection.execute(f"PRAGMA table_info({table})")}


def _migrate_members(connection: sqlite3.Connection) -> None:
    if not _table_exists(connection, "members"):
        return
    connection.execute(
        "INSERT OR IGNORE INTO pp_copy_shares(copy_id,user_id,role) "
        "SELECT scenario,user_id,CASE WHEN role='editor' THEN 'editor' ELSE 'viewer' END FROM members"
    )


def _migrate_edits(connection: sqlite3.Connection) -> None:
    if not _table_exists(connection, "edits"):
        return
    connection.execute(
        "INSERT OR IGNORE INTO pp_overrides(copy_id,sheet,row,col,value_json,version,updated_by,updated_at) "
        "SELECT scenario,sheet,r,c,value,version,actor,CAST(updated AS TEXT) FROM edits"
    )


def token_hash(token: str) -> str:
    return hashlib.sha256(token.encode("utf-8")).hexdigest()
