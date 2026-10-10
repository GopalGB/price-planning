import json
import io
from pathlib import Path
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET

from openpyxl import load_workbook
from openpyxl.styles import Protection

from ops.fixtures import make_semantic
from tests.v2.asgi_client import ASGIClient
from tests.v2.support import TMP


NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def runtime_root():
    temporary = tempfile.TemporaryDirectory(dir=TMP)
    root = Path(temporary.name)
    template = root / "synthetic.xlsm"
    make_semantic(template)
    (root / "settings.json").write_text(json.dumps({
        "data_dir": "results-private/runtime", "database_path": "results-private/runtime/planning.db",
        "template_path": str(template), "demo_mode": True, "cookie_secure": False,
        "allowed_origins": ["http://testserver"], "max_workers": 1,
    }))
    return temporary, root


def replace_template(path):
    workbook = load_workbook(path)
    workbook["Inputs"]["B2"].protection = Protection(locked=True)
    workbook["Calc"]["B2"] = "=Inputs!B2*3"
    workbook["Calc"].title = "Calc V2"
    workbook.save(path)


def exported_formula(response):
    with zipfile.ZipFile(io.BytesIO(response.body)) as archive:
        sheet = ET.fromstring(archive.read("xl/worksheets/sheet3.xml"))
    return sheet.findtext(f".//{NS}c[@r='B2']/{NS}f")


class EndToEndContractTests(unittest.TestCase):
    def test_runtime_composes_import_auth_edit_engine_and_static_fallback(self):
        """A partial server that cannot execute one complete user flow must fail."""
        from pp.runtime import create_runtime

        TMP.mkdir(parents=True, exist_ok=True)
        temporary, root = runtime_root()
        runtime = create_runtime(root)
        try:
            client = ASGIClient(runtime.app)
            signup = client.post("/api/auth/signup", {
                "email": "e2e@example.test", "name": "E2E", "password": "correct-horse-battery"
            })
            self.assertEqual(signup.status_code, 201)
            csrf = client.get("/api/me").json()["csrf"]
            category = client.post("/api/categories", {"name": "E2E Category"},
                                   {"x-csrf-token": csrf}).json()["category"]
            copy = client.post("/api/copies", {
                "category_id": category["id"], "name": "E2E Plan",
                "visibility": "private", "default_role": "viewer",
            }, {"x-csrf-token": csrf}).json()["copy"]
            edit = client.post(f"/api/copies/{copy['id']}/ops", {
                "op_id": "e2e-op", "action": "edit",
                "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 13, "expected_version": 0}],
            }, {"x-csrf-token": csrf})
            self.assertEqual(edit.status_code, 200)
            rows = client.get(f"/api/copies/{copy['id']}/sheets/2/rows?start=1&count=1").json()
            self.assertEqual(rows["rows"]["1"]["1"]["v"], 26)
            self.assertEqual(client.get("/").status_code, 200)
        finally:
            runtime.close()
            temporary.cleanup()

    def test_restart_recovers_pending_exact_revision_job(self):
        """Committed input with an unfinished job must calculate on restart."""
        from pp.auth import hash_password
        from pp.ops import OperationService
        from pp.runtime import create_runtime

        TMP.mkdir(parents=True, exist_ok=True)
        temporary, root = runtime_root()
        first = create_runtime(root)
        try:
            with first.database.transaction() as connection:
                connection.execute(
                    "INSERT INTO pp_users(id,email,name,password_hash,created_at) VALUES(?,?,?,?,?)",
                    ("recovery-owner", "recover@example.test", "Recover", hash_password("correct-horse-battery"),
                     "2026-10-10T00:00:00Z"),
                )
                connection.execute(
                    "INSERT INTO pp_copies(id,name,owner_id,template_id,visibility,default_role,input_revision,calculated_revision,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?)",
                    ("recover-copy", "Recover", "recovery-owner", first.snapshot.hash, "private", "viewer", 0, 0,
                     "2026-10-10T00:00:00Z", "2026-10-10T00:00:00Z"),
                )
            OperationService(first.database, first.snapshot).apply("recover-copy", "recovery-owner", {
                "op_id": "pending", "action": "edit",
                "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 14, "expected_version": 0}],
            })
        finally:
            first.close()

        second = create_runtime(root)
        try:
            with second.database.connect() as connection:
                copy = connection.execute(
                    "SELECT input_revision,calculated_revision FROM pp_copies WHERE id='recover-copy'"
                ).fetchone()
                job = connection.execute(
                    "SELECT status FROM pp_calc_jobs WHERE copy_id='recover-copy' AND revision=1"
                ).fetchone()
            self.assertEqual(tuple(copy), (1, 1))
            self.assertEqual(job["status"], "complete")
        finally:
            second.close()
            temporary.cleanup()

    def test_replaced_template_path_keeps_old_copy_version_immutable(self):
        """Old copies must keep their source, mask, metadata, engine, and export after restart."""
        from pp.runtime import create_runtime

        TMP.mkdir(parents=True, exist_ok=True)
        temporary, root = runtime_root()
        first = create_runtime(root)
        client = ASGIClient(first.app)
        try:
            client.post("/api/auth/signup", {
                "email": "versions@example.test", "name": "Versions",
                "password": "correct-horse-battery",
            })
            csrf = client.get("/api/me").json()["csrf"]
            category = client.post("/api/categories", {"name": "Versions"},
                                   {"x-csrf-token": csrf}).json()["category"]
            old = client.post("/api/copies", {
                "category_id": category["id"], "name": "Old", "visibility": "private",
                "default_role": "viewer",
            }, {"x-csrf-token": csrf}).json()["copy"]
        finally:
            first.close()

        replace_template(root / "synthetic.xlsm")
        second = create_runtime(root)
        client.app = second.app
        try:
            old_edit = client.post(f"/api/copies/{old['id']}/ops", {
                "op_id": "old-v1", "action": "edit",
                "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 13, "expected_version": 0}],
            }, {"x-csrf-token": csrf})
            old_book = client.get(f"/api/copies/{old['id']}/workbook").json()
            old_cell = client.get(f"/api/copies/{old['id']}/sheets/2/cell?r=1&c=1").json()
            new = client.post("/api/copies", {
                "category_id": category["id"], "name": "New", "visibility": "private",
                "default_role": "viewer",
            }, {"x-csrf-token": csrf}).json()["copy"]
            new_book = client.get(f"/api/copies/{new['id']}/workbook").json()
            new_edit = client.post(f"/api/copies/{new['id']}/ops", {
                "op_id": "new-v2", "action": "edit",
                "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 13, "expected_version": 0}],
            }, {"x-csrf-token": csrf})
            self.assertEqual((old_edit.status_code, old_cell["v"], old_book["sheets"][2]["name"]),
                             (200, 26, "Calc"))
            self.assertEqual((new_edit.status_code, new_book["sheets"][2]["name"]), (422, "Calc V2"))
            self.assertEqual(exported_formula(client.get(f"/api/copies/{old['id']}/export.xlsm")), "Inputs!B2*2")
            self.assertEqual(exported_formula(client.get(f"/api/copies/{new['id']}/export.xlsm")), "Inputs!B2*3")
        finally:
            second.close()
            temporary.cleanup()


if __name__ == "__main__":
    unittest.main()
