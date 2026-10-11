import ast
from datetime import datetime, timezone
import hashlib
import sqlite3
import json
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace

from tests.v2.support import SYNTHETIC_LOGIN, TMP


def legacy_schema():
    source = Path(__file__).resolve().parents[2] / "app" / "db.py"
    tree = ast.parse(source.read_text(encoding="utf-8"))
    for node in tree.body:
        if isinstance(node, ast.Assign) and any(
                isinstance(target, ast.Name) and target.id == "SCHEMA" for target in node.targets):
            return ast.literal_eval(node.value)
    raise AssertionError("legacy SCHEMA literal not found")


def legacy_password(password):
    salt = bytes(range(16))
    digest = hashlib.scrypt(password.encode(), salt=salt, n=16384, r=8, p=1)
    return f"{salt.hex()}:{digest.hex()}"


def seed_legacy(path):
    with sqlite3.connect(path) as connection:
        connection.executescript(legacy_schema())
        users = [
            ("legacy-user", "legacy@example.test", "Legacy", legacy_password("password-a"), 0, 1.0),
            ("member-user", "member@example.test", "Member", legacy_password("password-b"), 0, 1.0),
        ]
        connection.executemany("INSERT INTO users VALUES(?,?,?,?,?,?)", users)
        connection.execute("INSERT INTO scenarios VALUES(?,?,?,?,?,?)",
                           ("legacy-copy", "Legacy Plan", "legacy-user", 1, 1.0, 2.0))
        connection.execute("INSERT INTO members VALUES(?,?,?)",
                           ("legacy-copy", "member-user", "editor"))
        connection.execute("INSERT INTO edits VALUES(?,?,?,?,?,?,?,?)",
                           ("legacy-copy", 1, 2, 2, "12", 1, "legacy-user", 2.0))


def v2_state(path):
    with sqlite3.connect(path) as connection:
        tables = [row[0] for row in connection.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name LIKE 'pp_%' ORDER BY name"
        )]
        return {table: sorted(connection.execute(f'SELECT * FROM "{table}"').fetchall(), key=repr)
                for table in tables}


def migrate_seeded(path):
    from pp.db import Database

    database = Database(path)
    database.initialize()
    with database.transaction() as connection:
        connection.execute("INSERT INTO pp_templates VALUES(?,?,?,?,?)",
                           ("template", "Synthetic", "source", "snapshot", "now"))
    database.migrate_legacy("template")
    database.close()


class DatabaseContractTests(unittest.TestCase):
    def test_additive_schema_enables_wal_foreign_keys_and_busy_timeout(self):
        """Dropping legacy tables or omitting SQLite safety pragmas must fail."""
        from pp.db import Database

        TMP.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            path = Path(temp) / "planning.db"
            with sqlite3.connect(path) as connection:
                connection.execute("CREATE TABLE users(id TEXT PRIMARY KEY)")
                connection.execute("INSERT INTO users VALUES('legacy')")

            database = Database(path)
            database.initialize()
            with database.connect() as connection:
                tables = {row[0] for row in connection.execute(
                    "SELECT name FROM sqlite_master WHERE type='table'"
                )}
                self.assertEqual(connection.execute("PRAGMA journal_mode").fetchone()[0], "wal")
                self.assertEqual(connection.execute("PRAGMA foreign_keys").fetchone()[0], 1)
                self.assertEqual(connection.execute("PRAGMA busy_timeout").fetchone()[0], 5000)
                self.assertEqual(connection.execute("SELECT id FROM users").fetchone()[0], "legacy")
                self.assertTrue({"pp_users", "pp_sessions", "pp_categories", "pp_copies",
                                 "pp_overrides", "pp_ops", "pp_history", "pp_calc_jobs",
                                 "pp_results", "pp_events"}.issubset(tables))
            database.close()

    def test_legacy_copies_edits_and_history_migrate_idempotently(self):
        """Keeping accounts but dropping legacy copy history must fail."""
        from pp.db import Database
        from pp.ops import OperationService

        TMP.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            path = Path(temp) / "planning.db"
            with sqlite3.connect(path) as connection:
                connection.executescript(legacy_schema())
                connection.execute("INSERT INTO users VALUES(?,?,?,?,?,?)",
                                   ("legacy-user", "legacy@example.test", "Legacy",
                                    legacy_password("correct-horse-battery"), 0, 1.0))
                connection.execute("INSERT INTO scenarios VALUES(?,?,?,?,?,?)",
                                   ("legacy-copy", "Legacy Plan", "legacy-user", 3, 1.0, 2.0))
                connection.execute("INSERT INTO edits VALUES(?,?,?,?,?,?,?,?)",
                                   ("legacy-copy", 1, 2, 2, "12", 1, "legacy-user", 2.0))
                connection.execute("INSERT INTO edits VALUES(?,?,?,?,?,?,?,?)",
                                   ("legacy-copy", 1, 2, 3, "7", 1, "legacy-user", 2.0))
                payload = [{"sheet": 1, "cell": "B2", "value": 12, "before": 10, "version": 1}]
                connection.execute("INSERT INTO events VALUES(?,?,?,?,?,?,?)",
                                   ("legacy-event", "legacy-copy", 3, "legacy-user", "edit", json.dumps(payload), 2.0))

            database = Database(path)
            database.initialize()
            with database.transaction() as connection:
                connection.execute("INSERT INTO pp_templates VALUES(?,?,?,?,?)",
                                   ("template", "Synthetic", "source", "snapshot", "now"))
            database.migrate_legacy("template")
            database.migrate_legacy("template")
            with database.connect() as connection:
                user_time = connection.execute(
                    "SELECT created_at FROM pp_users WHERE id='legacy-user'"
                ).fetchone()[0]
                copy = connection.execute(
                    "SELECT input_revision,calculated_revision,created_at,updated_at "
                    "FROM pp_copies WHERE id='legacy-copy'"
                ).fetchone()
                overrides = connection.execute(
                    "SELECT sheet,row,col,value_json,updated_at FROM pp_overrides ORDER BY sheet,row,col"
                ).fetchall()
                history = connection.execute(
                    "SELECT sheet,row,col,before_json,after_json,at FROM pp_history"
                ).fetchone()
                event_row = connection.execute(
                    "SELECT payload_json,created_at FROM pp_events WHERE copy_id='legacy-copy'"
                ).fetchone()
                event = json.loads(event_row["payload_json"])
                jobs = connection.execute(
                    "SELECT revision,status,created_at,updated_at FROM pp_calc_jobs WHERE copy_id='legacy-copy'"
                ).fetchall()
            changes = OperationService(database, SimpleNamespace(hash="template")).revision_changes(
                "legacy-copy", 3
            )
            self.assertEqual(tuple(copy[:2]), (3, -1))
            self.assertEqual([tuple(row[:4]) for row in overrides], [(0, 1, 1, "12"), (0, 1, 2, "7")])
            self.assertEqual(tuple(history[:5]), (0, 1, 1, "10", "12"))
            self.assertEqual(event["cells"], [{"sheet": 0, "r": 1, "c": 1, "v": 12, "ver": 1}])
            self.assertEqual([tuple(row[:2]) for row in jobs], [(3, "pending")])
            self.assertEqual(changes, [(0, 1, 1, 12), (0, 1, 2, 7)])
            timestamps = [user_time, copy["created_at"], copy["updated_at"], history["at"],
                          event_row["created_at"], jobs[0]["created_at"], jobs[0]["updated_at"]]
            timestamps.extend(row["updated_at"] for row in overrides)
            for timestamp in timestamps:
                with self.subTest(timestamp=timestamp):
                    parsed = datetime.fromisoformat(timestamp.replace("Z", "+00:00"))
                    self.assertEqual(parsed.tzinfo, timezone.utc)
            database.close()

    def test_migrated_legacy_account_logs_in_and_upgrades_only_v2_hash(self):
        """Copying a legacy hex scrypt hash without verification support must fail."""
        from pp.auth import AuthService, verify_password
        from pp.db import Database

        TMP.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            path = Path(temp) / "planning.db"
            password = SYNTHETIC_LOGIN
            encoded = legacy_password(password)
            with sqlite3.connect(path) as connection:
                connection.executescript(legacy_schema())
                connection.execute("INSERT INTO users VALUES(?,?,?,?,?,?)",
                                   ("legacy-user", "legacy@example.test", "Legacy", encoded, 0, 1.0))
            database = Database(path)
            database.initialize()
            database.migrate_legacy("template")

            user, token, csrf = AuthService(database, SimpleNamespace()).login(
                "legacy@example.test", password, "127.0.0.1"
            )

            with database.connect() as connection:
                upgraded = connection.execute(
                    "SELECT password_hash FROM pp_users WHERE id='legacy-user'"
                ).fetchone()[0]
                original = connection.execute("SELECT password FROM users WHERE id='legacy-user'").fetchone()[0]
            self.assertEqual(user["id"], "legacy-user")
            self.assertTrue(token and csrf)
            self.assertTrue(upgraded.startswith("scrypt$"))
            self.assertTrue(verify_password(password, upgraded))
            self.assertEqual(original, encoded)
            database.close()

    def test_changed_legacy_tables_abort_restart_before_v2_mutation(self):
        """Legacy edits, members, or users changed after migration must stop startup."""
        from pp.db import Database

        mutations = {
            "edit": "UPDATE edits SET value='99' WHERE scenario='legacy-copy'",
            "member": "UPDATE members SET role='viewer' WHERE scenario='legacy-copy'",
            "user": "UPDATE users SET name='Changed' WHERE id='legacy-user'",
        }
        TMP.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            for name, statement in mutations.items():
                with self.subTest(table=name):
                    path = Path(temp) / f"{name}.db"
                    seed_legacy(path)
                    migrate_seeded(path)
                    before = v2_state(path)
                    with sqlite3.connect(path) as connection:
                        connection.execute(statement)
                    restart = Database(path)
                    with self.assertRaisesRegex(RuntimeError, "Legacy data changed"):
                        restart.initialize()
                    self.assertEqual(v2_state(path), before)

    def test_v2_only_changes_do_not_invalidate_legacy_fingerprint(self):
        """The legacy guard must ignore normal v2 revisions and result publication."""
        from pp.db import Database

        TMP.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            path = Path(temp) / "planning.db"
            seed_legacy(path)
            migrate_seeded(path)
            with sqlite3.connect(path) as connection:
                connection.execute(
                    "UPDATE pp_copies SET calculated_revision=1 WHERE id='legacy-copy'"
                )

            restart = Database(path)
            restart.initialize()
            restart.migrate_legacy("template")

            with restart.connect() as connection:
                revision = connection.execute(
                    "SELECT calculated_revision FROM pp_copies WHERE id='legacy-copy'"
                ).fetchone()[0]
                fingerprint = connection.execute(
                    "SELECT value FROM pp_meta WHERE key='legacy_migration_fingerprint'"
                ).fetchone()[0]
            self.assertEqual(revision, 1)
            self.assertEqual(len(fingerprint), 64)
            restart.close()


if __name__ == "__main__":
    unittest.main()
