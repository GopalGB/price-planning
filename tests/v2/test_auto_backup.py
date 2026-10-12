"""Automatic daily backups of the live database while the v2 server runs."""

import contextlib
import importlib.util
import io
import json
import os
import sqlite3
import sys
import tempfile
import time
import unittest
from pathlib import Path
from types import ModuleType, SimpleNamespace
from unittest import mock

from pp import auto_backup
from pp.db import Database
from tests.v2.support import TMP

ROOT = Path(__file__).resolve().parents[2]
SERVER_SETTINGS = {
    "bind": "127.0.0.1",
    "port": 8000,
    "demo_mode": True,
    "allowed_origins": ["http://testserver"],
}


def live_database(root: Path) -> Path:
    """A real database in the app's own schema, so backups run against live tables."""
    path = root / "data" / "planning.db"
    database = Database(path)
    database.initialize()
    with database.transaction() as connection:
        connection.execute("INSERT OR REPLACE INTO pp_meta (key, value) VALUES ('probe', 'nightly')")
    return path


def meta(database: Path) -> dict:
    with sqlite3.connect(database) as connection:
        return dict(connection.execute("SELECT key, value FROM pp_meta"))


def backups(folder: Path) -> list:
    return sorted(folder.glob("planning-*.db"))


def wait_until(predicate, timeout=20.0) -> bool:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if predicate():
            return True
        time.sleep(0.05)
    return False


class PruneTests(unittest.TestCase):
    def folder_with_sets(self, root: Path, count: int) -> Path:
        folder = root / "backups"
        folder.mkdir()
        for index in range(count):
            stem = f"planning-2026010{index + 1}T000000000000Z"
            (folder / f"{stem}.db").write_bytes(b"database")
            (folder / f"{stem}.json").write_text(json.dumps({"created_at": stem}))
            templates = folder / f"{stem}.templates"
            templates.mkdir()
            (templates / "template-id").write_text("bundle")
        return folder

    def test_prune_keeps_the_newest_seven_of_nine_whole_sets(self):
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            folder = self.folder_with_sets(Path(temp), 9)
            partial = folder / "planning-20260110T120000000000Z.partial"
            partial.write_bytes(b"backup in progress")
            unrelated = folder / "notes.txt"
            unrelated.write_text("keep me", encoding="utf-8")

            removed = auto_backup.prune(folder)

            self.assertEqual([path.stem for path in removed],
                             ["planning-20260101T000000000000Z", "planning-20260102T000000000000Z"])
            kept = backups(folder)
            self.assertEqual([path.stem for path in kept],
                             [f"planning-2026010{day}T000000000000Z" for day in range(3, 10)])
            for database in kept:  # a kept set keeps its record and its template bundle
                self.assertTrue(database.with_suffix(".json").is_file())
                self.assertTrue(database.with_suffix(".templates").is_dir())
            for day in (1, 2):
                stem = f"planning-2026010{day}T000000000000Z"
                self.assertFalse((folder / f"{stem}.db").exists())
                self.assertFalse((folder / f"{stem}.json").exists())
                self.assertFalse((folder / f"{stem}.templates").exists())
            self.assertTrue(partial.is_file(), "a backup in progress must never be deleted")
            self.assertEqual(unrelated.read_text(encoding="utf-8"), "keep me")

    def test_prune_honours_keep_and_ignores_a_missing_folder(self):
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            folder = self.folder_with_sets(Path(temp), 3)
            self.assertEqual(len(auto_backup.prune(folder, keep=1)), 2)
            self.assertEqual(len(backups(folder)), 1)
            self.assertEqual(auto_backup.prune(folder / "absent"), [])


class AutoBackupTests(unittest.TestCase):
    def test_loop_backs_up_the_live_database_repeatedly_and_stops_promptly(self):
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            root = Path(temp)
            database, dest = live_database(root), root / "backups"
            auto = auto_backup.AutoBackup(database, dest, first_delay=0, interval=0.2)
            auto.start()
            try:
                self.assertTrue(wait_until(lambda: len(backups(dest)) >= 2),
                                "the daily loop must back up more than once")
            finally:
                started = time.monotonic()
                auto.stop(timeout=10)
                stopped_after = time.monotonic() - started
            self.assertLess(stopped_after, 2.0, "stop() must not wait for the next backup")
            self.assertFalse(auto.thread.is_alive())
            for made in backups(dest):
                self.assertEqual(meta(made), {"probe": "nightly"})
                self.assertTrue(made.with_suffix(".json").is_file())
            count = len(backups(dest))
            time.sleep(0.5)  # a stopped loop writes nothing more
            self.assertEqual(len(backups(dest)), count)

    def test_a_failed_backup_is_logged_and_the_next_run_still_happens(self):
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            root = Path(temp)
            database, dest = live_database(root), root / "backups"
            real, calls = auto_backup.backup, []

            def flaky(db_path, dest_dir):
                calls.append(db_path)
                if len(calls) == 1:
                    raise OSError("the backup folder is not writable")
                return real(db_path, dest_dir)

            with mock.patch.object(auto_backup, "backup", side_effect=flaky):
                auto = auto_backup.AutoBackup(database, dest, first_delay=0, interval=0.2)
                with self.assertLogs("pp.auto_backup", level="ERROR"):
                    auto.start()
                    try:
                        self.assertTrue(wait_until(lambda: len(backups(dest)) >= 1),
                                        "a failed backup must not stop the daily loop")
                    finally:
                        auto.stop(timeout=10)
            self.assertGreaterEqual(len(calls), 2)
            self.assertTrue(backups(dest))


class RunnerWiringTests(unittest.TestCase):
    def runner_module(self):
        runtime_module = ModuleType("pp.runtime")
        runtime_module.create_runtime = None
        spec = importlib.util.spec_from_file_location("test_auto_backup_runner", ROOT / "run_v2.py")
        module = importlib.util.module_from_spec(spec)
        with mock.patch.dict(sys.modules, {"pp.runtime": runtime_module}):
            spec.loader.exec_module(module)
        module._port_in_use = lambda port: False  # keep tests off whatever else listens here
        return module

    @mock.patch.dict(os.environ)
    def run_server(self, settings):
        for name in ("DATA_DIR", "PP_DATABASE_PATH", "PP_PORT"):
            os.environ.pop(name, None)  # another suite sets DATA_DIR process wide
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            runner, root = self.runner_module(), Path(temp)
            (root / "settings.json").write_text(json.dumps(settings), encoding="utf-8")
            runtime = SimpleNamespace(app=object(), config=runner.load_config(root), close=mock.Mock())
            started, stopped = [], []

            class FakeAutoBackup:
                def __init__(self, db_path, dest_dir, **kwargs):
                    self.db_path, self.dest_dir = Path(db_path), Path(dest_dir)

                def start(self):
                    started.append(self)
                    return self

                def stop(self, *args, **kwargs):
                    stopped.append(self)

            output = io.StringIO()
            with mock.patch.object(runner, "ROOT", root), mock.patch.object(
                runner, "AutoBackup", FakeAutoBackup
            ), mock.patch.object(runner, "create_runtime", return_value=runtime), mock.patch.object(
                runner.uvicorn, "run"
            ) as serve, contextlib.redirect_stdout(output):
                code = runner.main([])
            self.assertEqual(code, 0)
            serve.assert_called_once()
            runtime.close.assert_called_once_with()
            return root, started, stopped, output.getvalue()

    def test_the_server_backs_up_the_live_database_daily_and_stops_on_exit(self):
        root, started, stopped, output = self.run_server(SERVER_SETTINGS)
        self.assertEqual(len(started), 1)
        self.assertEqual(stopped, started, "the server must stop its backup thread on exit")
        self.assertEqual(started[0].db_path, root / "data" / "planning.db")
        self.assertEqual(started[0].dest_dir, root / "backups")
        self.assertEqual(output.strip().splitlines()[-1],
                         f"Automatic backup: daily into {root / 'backups'}, keeping the newest 7.")

    def test_settings_turn_the_daily_backup_off(self):
        root, started, stopped, output = self.run_server({**SERVER_SETTINGS, "auto_backup": False})
        self.assertEqual((started, stopped), ([], []))
        self.assertNotIn("Automatic backup", output)
        self.assertFalse((root / "backups").exists())


if __name__ == "__main__":
    unittest.main()