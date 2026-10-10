import hashlib
import io
import json
import os
from pathlib import Path
import shutil
import tempfile
import unittest
from unittest import mock
import zipfile
import xml.etree.ElementTree as ET

from openpyxl import load_workbook

from ops.fixtures import make_semantic
from pp import backup as bk
from tests.v2.asgi_client import ASGIClient
from tests.v2.support import BackendFixture, TMP


NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def _sha256(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def _settings(root, template):
    (root / "settings.json").write_text(json.dumps({
        "data_dir": "data",
        "database_path": "data/planning.db",
        "template_path": str(template),
        "demo_mode": True,
        "cookie_secure": False,
        "allowed_origins": ["http://testserver"],
        "max_workers": 1,
    }), encoding="utf-8")


def _make_h2(source, destination):
    shutil.copyfile(source, destination)
    workbook = load_workbook(destination)
    workbook["Calc"]["B2"] = "=Inputs!B2*3"
    workbook["Calc"].title = "Calc H2"
    workbook.save(destination)


def _exported_formula(response):
    with zipfile.ZipFile(io.BytesIO(response.body)) as archive:
        sheet = ET.fromstring(archive.read("xl/worksheets/sheet3.xml"))
    return sheet.findtext(f".//{NS}c[@r='B2']/{NS}f")


class TemplateBundleRuntimeTests(unittest.TestCase):
    @mock.patch.dict(os.environ)
    def test_restore_relocates_all_template_versions_and_recovers_old_copy(self):
        from pp.runtime import create_runtime

        for name in ("DATA_DIR", "PP_DATABASE_PATH", "PP_TEMPLATE_PATH"):
            os.environ.pop(name, None)
        TMP.mkdir(parents=True, exist_ok=True)
        temporary = tempfile.TemporaryDirectory(dir=TMP)
        self.addCleanup(temporary.cleanup)
        base = Path(temporary.name)
        original = base / "original"
        original.mkdir()
        h1 = original / "h1.xlsm"
        h2 = original / "h2.xlsm"
        make_semantic(h1)
        _make_h2(h1, h2)
        _settings(original, h1)

        first = create_runtime(original)
        client = ASGIClient(first.app)
        try:
            signup = client.post("/api/auth/signup", {
                "email": "portable@example.test",
                "name": "Portable",
                "password": "correct-horse-battery",
            })
            self.assertEqual(signup.status_code, 201)
            me = client.get("/api/me").json()
            csrf, user_id = me["csrf"], me["user"]["id"]
            category = client.post(
                "/api/categories", {"name": "Portable"}, {"x-csrf-token": csrf}
            ).json()["category"]
            old_copy = client.post("/api/copies", {
                "category_id": category["id"],
                "name": "H1 copy",
                "visibility": "private",
                "default_role": "viewer",
            }, {"x-csrf-token": csrf}).json()["copy"]
            h1_id = first.snapshot.hash
        finally:
            first.close()

        _settings(original, h2)
        second = create_runtime(original)
        client.app = second.app
        try:
            h2_id = second.snapshot.hash
            self.assertNotEqual(h1_id, h2_id)
            ack = second.app.state.services.operations.apply(old_copy["id"], user_id, {
                "op_id": "pending-after-backup",
                "action": "edit",
                "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 14, "expected_version": 0}],
            })
            self.assertEqual(ack["input_revision"], 1)
            backup_path = bk.backup(second.config.database_path, base / "backups")
        finally:
            second.close()

        restart = base / "restart"
        restart.mkdir()
        restart_h2 = restart / "h2.xlsm"
        shutil.copyfile(h2, restart_h2)
        original.rename(base / "unavailable-original")
        restored_db = bk.restore(backup_path, restart / "data")
        _settings(restart, restart_h2)

        restored = create_runtime(restart)
        client.app = restored.app
        try:
            workbook = client.get(f"/api/copies/{old_copy['id']}/workbook")
            rows = client.get(f"/api/copies/{old_copy['id']}/sheets/2/rows?start=1&count=1")
            exported = client.get(f"/api/copies/{old_copy['id']}/export.xlsm")
            self.assertEqual(workbook.status_code, 200)
            self.assertEqual(workbook.json()["sheets"][2]["name"], "Calc")
            self.assertEqual(rows.status_code, 200)
            self.assertEqual(rows.json()["rows"]["1"]["1"]["v"], 28)
            self.assertEqual(exported.status_code, 200)
            self.assertEqual(_exported_formula(exported), "Inputs!B2*2")
            with restored.database.connect() as connection:
                job = connection.execute(
                    "SELECT status FROM pp_calc_jobs WHERE copy_id=? AND revision=1",
                    (old_copy["id"],),
                ).fetchone()
                copy = connection.execute(
                    "SELECT calculated_revision FROM pp_copies WHERE id=?", (old_copy["id"],)
                ).fetchone()
                templates = connection.execute(
                    "SELECT id,source_path,snapshot_path FROM pp_templates ORDER BY id"
                ).fetchall()
            self.assertEqual((job["status"], copy["calculated_revision"]), ("complete", 1))
            self.assertEqual({row["id"] for row in templates}, {h1_id, h2_id})
            for row in templates:
                snapshot = restart / "data" / "templates" / row["id"]
                self.assertEqual(Path(row["snapshot_path"]), snapshot)
                self.assertEqual(Path(row["source_path"]), snapshot / "source.xlsm")
                self.assertEqual(_sha256(row["source_path"]), row["id"])
            self.assertEqual(restored_db, restart / "data" / "planning.db")

            record = json.loads(backup_path.with_suffix(".json").read_text(encoding="utf-8"))
            self.assertEqual(record["format_version"], 2)
            self.assertEqual({item["id"] for item in record["templates"]}, {h1_id, h2_id})
            bundle = backup_path.with_suffix(".templates")
            for item in record["templates"]:
                root = bundle / item["id"]
                self.assertEqual(_sha256(root / "source.xlsm"), item["id"])
                self.assertEqual(
                    {path: _sha256(root / path) for path in item["files"]}, item["files"]
                )

        finally:
            restored.close()


class TemplateBundleSafetyTests(unittest.TestCase):
    def setUp(self):
        self.fixture = BackendFixture()
        self.backup_path = bk.backup(self.fixture.database.path, self.fixture.root / "backups")
        self.record_path = self.backup_path.with_suffix(".json")
        self.bundle = self.backup_path.with_suffix(".templates")

    def tearDown(self):
        self.fixture.close()

    def test_missing_template_file_fails_closed(self):
        before = self.backup_path.read_bytes()
        (self.bundle / self.fixture.snapshot.hash / "source.xlsm").unlink()
        target = self.fixture.root / "missing-restore"
        with self.assertRaisesRegex(ValueError, "template"):
            bk.restore(self.backup_path, target)
        self.assertEqual(self.backup_path.read_bytes(), before)
        self.assertTrue(self.fixture.database.path.is_file())
        self.assertFalse((target / "planning.db").exists())

    def test_modified_template_file_fails_closed(self):
        path = self.bundle / self.fixture.snapshot.hash / "styles.json"
        path.write_text("{}", encoding="utf-8")
        target = self.fixture.root / "modified-restore"
        with self.assertRaisesRegex(ValueError, "checksum"):
            bk.restore(self.backup_path, target)
        self.assertTrue(self.fixture.database.path.is_file())
        self.assertFalse((target / "planning.db").exists())

    def test_unsafe_manifest_path_is_rejected_without_escape(self):
        record = json.loads(self.record_path.read_text(encoding="utf-8"))
        record["templates"][0]["files"]["../escape.json"] = hashlib.sha256(b"escape").hexdigest()
        self.record_path.write_text(json.dumps(record), encoding="utf-8")
        target = self.fixture.root / "unsafe-restore"
        with self.assertRaisesRegex(ValueError, "unsafe"):
            bk.restore(self.backup_path, target)
        self.assertFalse((self.fixture.root / "escape.json").exists())
        self.assertFalse((target / "planning.db").exists())

    def test_template_copy_failure_rolls_back_and_retry_succeeds(self):
        target = self.fixture.root / "copy-failure-restore"
        original_copy = bk.shutil.copyfile

        def fail_on_styles(source, destination):
            if Path(source).name == "styles.json":
                Path(destination).write_bytes(b"partial")
                raise OSError("synthetic template copy failure")
            return original_copy(source, destination)

        with mock.patch.object(bk.shutil, "copyfile", side_effect=fail_on_styles):
            with self.assertRaisesRegex(OSError, "synthetic template copy failure"):
                bk.restore(self.backup_path, target)
        self.assertEqual(list(target.iterdir()), [])
        restored = bk.restore(self.backup_path, target)
        self.assertTrue(restored.is_file())
        self.assertTrue((target / "templates" / self.fixture.snapshot.hash / "source.xlsm").is_file())

    def test_destination_race_does_not_replace_template_directory(self):
        target = self.fixture.root / "race-restore"
        original_copy = bk.shutil.copyfile
        raced = False

        def create_destination(source, destination):
            nonlocal raced
            result = original_copy(source, destination)
            if not raced:
                raced = True
                (target / "templates").mkdir()
                (target / "templates" / "racer.txt").write_text("keep", encoding="utf-8")
            return result

        with mock.patch.object(bk.shutil, "copyfile", side_effect=create_destination):
            with self.assertRaisesRegex(ValueError, "nothing was overwritten"):
                bk.restore(self.backup_path, target)
        self.assertEqual((target / "templates" / "racer.txt").read_text(encoding="utf-8"), "keep")
        self.assertFalse((target / "planning.db").exists())

    def test_legacy_sidecar_without_template_manifest_still_restores(self):
        record = json.loads(self.record_path.read_text(encoding="utf-8"))
        record.pop("format_version")
        record.pop("templates")
        self.record_path.write_text(json.dumps(record), encoding="utf-8")
        target = self.fixture.root / "legacy-restore"
        restored = bk.restore(self.backup_path, target)
        self.assertTrue(restored.is_file())
        self.assertEqual([path.name for path in target.iterdir()], ["planning.db"])


if __name__ == "__main__":
    unittest.main()
