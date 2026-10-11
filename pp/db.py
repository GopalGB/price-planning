"""SQLite connection and transaction boundary for the v2 backend."""

from __future__ import annotations

from contextlib import contextmanager
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import sqlite3

from .schema import SCHEMA


LEGACY_TABLES = ("users", "scenarios", "edits", "events", "members")
LEGACY_FINGERPRINT_KEY = "legacy_migration_fingerprint"


class LegacyMigrationError(RuntimeError):
    """The legacy database changed after its one permitted v2 migration."""


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
            _assert_legacy_unchanged(connection)
            connection.execute(
                "INSERT OR IGNORE INTO pp_copy_baselines(copy_id,sheet,row,col,value_json) "
                "SELECT copy_id,sheet,row,col,value_json FROM pp_overrides WHERE origin_copy_id IS NOT NULL"
            )

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

    @contextmanager
    def read_transaction(self):
        connection = self._open()
        try:
            connection.execute("BEGIN")
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
            "SELECT id,email,name,password,strftime('%Y-%m-%dT%H:%M:%fZ',created,'unixepoch') FROM users"
        )
        connection.execute(
            "UPDATE pp_users SET created_at=(SELECT strftime('%Y-%m-%dT%H:%M:%fZ',u.created,'unixepoch') "
            "FROM users u WHERE u.id=pp_users.id) WHERE created_at NOT LIKE '%T%' "
            "AND id IN (SELECT id FROM users)"
        )

    def migrate_legacy(self, template_id: str) -> None:
        with self.transaction() as connection:
            fingerprint = _legacy_fingerprint(connection)
            if fingerprint is None:
                return
            stored = connection.execute(
                "SELECT value FROM pp_meta WHERE key=?", (LEGACY_FINGERPRINT_KEY,)
            ).fetchone()
            if stored is not None:
                _require_legacy_match(stored["value"], fingerprint)
                return
            self._migrate_users(connection)
            if _table_exists(connection, "scenarios"):
                _migrate_copies(connection, template_id)
                _migrate_members(connection)
                _migrate_edits(connection)
                _migrate_events(connection)
                _queue_migrated_jobs(connection)
            connection.execute(
                "INSERT INTO pp_meta(key,value) VALUES(?,?)",
                (LEGACY_FINGERPRINT_KEY, fingerprint),
            )

    def close(self) -> None:
        self._closed = True


def _table_exists(connection: sqlite3.Connection, name: str) -> bool:
    row = connection.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (name,)).fetchone()
    return row is not None


def _columns(connection: sqlite3.Connection, table: str) -> set[str]:
    return {row[1] for row in connection.execute(f"PRAGMA table_info({table})")}


def _assert_legacy_unchanged(connection):
    stored = connection.execute(
        "SELECT value FROM pp_meta WHERE key=?", (LEGACY_FINGERPRINT_KEY,)
    ).fetchone()
    if stored is not None:
        _require_legacy_match(stored["value"], _legacy_fingerprint(connection))


def _require_legacy_match(stored, current):
    if stored != current:
        raise LegacyMigrationError(
            "Legacy data changed after v2 migration; use an independent legacy data directory"
        )


def _legacy_fingerprint(connection):
    digest, found = hashlib.sha256(), False
    for table in LEGACY_TABLES:
        if not _table_exists(connection, table):
            continue
        found = True
        columns = [row[1] for row in connection.execute(f"PRAGMA table_info({table})")]
        schema = connection.execute(
            "SELECT sql FROM sqlite_master WHERE type='table' AND name=?", (table,)
        ).fetchone()[0]
        _hash_fingerprint_part(digest, f"{table}\0{schema}")
        quoted = [f'"{name.replace(chr(34), chr(34) * 2)}"' for name in columns]
        query = f'SELECT {",".join(quoted)} FROM "{table}" ORDER BY {",".join(quoted)}'
        for row in connection.execute(query):
            values = [_fingerprint_value(value) for value in row]
            _hash_fingerprint_part(digest, json.dumps(values, separators=(",", ":")))
    return digest.hexdigest() if found else None


def _fingerprint_value(value):
    return ["bytes", value.hex()] if isinstance(value, bytes) else [type(value).__name__, value]


def _hash_fingerprint_part(digest, value):
    encoded = value.encode("utf-8")
    digest.update(len(encoded).to_bytes(8, "big"))
    digest.update(encoded)


def _migrate_copies(connection, template_id):
    connection.execute(
        "INSERT OR IGNORE INTO pp_copies(id,name,owner_id,template_id,visibility,default_role,input_revision,calculated_revision,created_at,updated_at) "
        "SELECT id,name,owner,?,'private','viewer',revision,-1,"
        "strftime('%Y-%m-%dT%H:%M:%fZ',created,'unixepoch'),"
        "strftime('%Y-%m-%dT%H:%M:%fZ',updated,'unixepoch') FROM scenarios",
        (template_id,),
    )
    _repair_copy_times(connection)
    connection.execute(
        "UPDATE pp_copies SET calculated_revision=-1 WHERE id IN (SELECT id FROM scenarios) "
        "AND calculated_revision=0 AND NOT EXISTS "
        "(SELECT 1 FROM pp_results WHERE pp_results.copy_id=pp_copies.id)"
    )


def _queue_migrated_jobs(connection):
    connection.execute(
        "INSERT OR IGNORE INTO pp_calc_jobs(copy_id,revision,status,created_at,updated_at) "
        "SELECT id,input_revision,'pending',updated_at,updated_at FROM pp_copies "
        "WHERE id IN (SELECT id FROM scenarios) AND calculated_revision<input_revision"
    )
    connection.execute(
        "UPDATE pp_calc_jobs SET created_at=(SELECT updated_at FROM pp_copies c WHERE c.id=copy_id),"
        "updated_at=(SELECT updated_at FROM pp_copies c WHERE c.id=copy_id) "
        "WHERE copy_id IN (SELECT id FROM scenarios) "
        "AND (created_at NOT LIKE '%T%' OR updated_at NOT LIKE '%T%')"
    )


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
        "SELECT scenario,sheet-1,r-1,c-1,value,version,actor,"
        "strftime('%Y-%m-%dT%H:%M:%fZ',updated,'unixepoch') FROM edits "
        "WHERE sheet>=1 AND r>=1 AND c>=1"
    )
    connection.execute(
        "UPDATE pp_overrides SET updated_at=(SELECT strftime('%Y-%m-%dT%H:%M:%fZ',e.updated,'unixepoch') "
        "FROM edits e WHERE e.scenario=pp_overrides.copy_id AND e.sheet=pp_overrides.sheet+1 "
        "AND e.r=pp_overrides.row+1 AND e.c=pp_overrides.col+1) "
        "WHERE updated_at NOT LIKE '%T%' AND EXISTS "
        "(SELECT 1 FROM edits e WHERE e.scenario=pp_overrides.copy_id AND e.sheet=pp_overrides.sheet+1 "
        "AND e.r=pp_overrides.row+1 AND e.c=pp_overrides.col+1)"
    )


def _migrate_events(connection: sqlite3.Connection) -> None:
    if not _table_exists(connection, "events"):
        return
    for event in connection.execute("SELECT * FROM events ORDER BY created,id").fetchall():
        created_at = _legacy_iso(event["created"])
        _repair_event_times(connection, event, created_at)
        if connection.execute("SELECT 1 FROM pp_legacy_events WHERE event_id=?", (event["id"],)).fetchone():
            continue
        payload = _legacy_payload(event["payload"])
        user = connection.execute("SELECT id,name FROM pp_users WHERE id=?", (event["actor"],)).fetchone()
        cells = []
        for item in payload if isinstance(payload, list) else []:
            coordinate = _legacy_coordinate(item.get("cell"))
            sheet = _legacy_sheet(item.get("sheet"))
            if coordinate is None or sheet is None:
                continue
            row, col = coordinate
            before, after = item.get("before"), item.get("value")
            connection.execute(
                "INSERT INTO pp_history(copy_id,op_id,revision,at,actor_id,action,sheet,row,col,before_json,after_json,version) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",
                (event["scenario"], event["id"], event["revision"], created_at, event["actor"],
                 event["kind"], sheet, row, col, json.dumps(before), json.dumps(after),
                 int(item.get("version", 0))),
            )
            cells.append({"sheet": sheet, "r": row, "c": col,
                          "v": after, "ver": int(item.get("version", 0))})
        message = {"type": "op", "op_id": event["id"], "input_revision": event["revision"],
                   "by": {"id": user["id"], "name": user["name"]}, "action": event["kind"], "cells": cells}
        connection.execute(
            "INSERT INTO pp_events(copy_id,event_type,input_revision,payload_json,created_at) VALUES(?,'op',?,?,?)",
            (event["scenario"], event["revision"], json.dumps(message, separators=(",", ":")), created_at),
        )
        connection.execute("INSERT INTO pp_legacy_events(event_id) VALUES(?)", (event["id"],))


def _repair_copy_times(connection):
    connection.execute(
        "UPDATE pp_copies SET "
        "created_at=(SELECT strftime('%Y-%m-%dT%H:%M:%fZ',s.created,'unixepoch') FROM scenarios s WHERE s.id=pp_copies.id),"
        "updated_at=(SELECT strftime('%Y-%m-%dT%H:%M:%fZ',s.updated,'unixepoch') FROM scenarios s WHERE s.id=pp_copies.id) "
        "WHERE id IN (SELECT id FROM scenarios) AND (created_at NOT LIKE '%T%' OR updated_at NOT LIKE '%T%')"
    )


def _repair_event_times(connection, event, created_at):
    connection.execute("UPDATE pp_history SET at=? WHERE op_id=? AND at NOT LIKE '%T%'",
                       (created_at, event["id"]))
    rows = connection.execute(
        "SELECT sequence,payload_json,created_at FROM pp_events WHERE copy_id=? AND event_type='op' AND input_revision=?",
        (event["scenario"], event["revision"]),
    ).fetchall()
    for row in rows:
        payload = _legacy_payload(row["payload_json"])
        if isinstance(payload, dict) and payload.get("op_id") == event["id"] and "T" not in row["created_at"]:
            connection.execute("UPDATE pp_events SET created_at=? WHERE sequence=?", (created_at, row["sequence"]))


def _legacy_iso(value):
    return datetime.fromtimestamp(float(value), timezone.utc).isoformat().replace("+00:00", "Z")


def _legacy_payload(value):
    try:
        return json.loads(value)
    except (TypeError, json.JSONDecodeError):
        return []


def _legacy_coordinate(value):
    if not isinstance(value, str):
        return None
    letters, digits = "", ""
    for character in value.replace("$", ""):
        if character.isalpha() and not digits:
            letters += character.upper()
        elif character.isdigit():
            digits += character
        else:
            return None
    if not letters or not digits or int(digits) < 1:
        return None
    col = 0
    for letter in letters:
        col = col * 26 + ord(letter) - 64
    return int(digits) - 1, col - 1


def _legacy_sheet(value):
    return value - 1 if type(value) is int and value >= 1 else None


def token_hash(token: str) -> str:
    return hashlib.sha256(token.encode("utf-8")).hexdigest()
