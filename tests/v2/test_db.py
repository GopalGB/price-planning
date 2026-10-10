import sqlite3
import tempfile
import unittest
from pathlib import Path

from tests.v2.support import TMP


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


if __name__ == "__main__":
    unittest.main()

