import json
from pathlib import Path
import tempfile

from ops.fixtures import make_semantic


ROOT = Path(__file__).resolve().parents[2]
TMP = ROOT / "results-private" / "tmp"


class BackendFixture:
    def __init__(self):
        TMP.mkdir(parents=True, exist_ok=True)
        self._temporary = tempfile.TemporaryDirectory(dir=TMP)
        self.root = Path(self._temporary.name)
        self.source = self.root / "semantic.xlsm"
        make_semantic(self.source)

        from pp.db import Database
        from pp.template.importer import import_template
        from pp.workbook_service import TemplateSnapshot

        self.snapshot_path = import_template(self.source, self.root / "data")
        self.snapshot = TemplateSnapshot(self.snapshot_path)
        self.database = Database(self.root / "planning.db")
        self.database.initialize()
        self._seed()

    def _seed(self):
        with self.database.transaction() as connection:
            connection.execute(
                "INSERT INTO pp_users(id,email,name,password_hash,created_at) VALUES(?,?,?,?,?)",
                ("owner", "owner@example.test", "Owner", "x", "2026-10-10T00:00:00Z"),
            )
            connection.execute(
                "INSERT INTO pp_users(id,email,name,password_hash,created_at) VALUES(?,?,?,?,?)",
                ("editor", "editor@example.test", "Editor", "x", "2026-10-10T00:00:00Z"),
            )
            connection.execute(
                "INSERT INTO pp_users(id,email,name,password_hash,created_at) VALUES(?,?,?,?,?)",
                ("viewer", "viewer@example.test", "Viewer", "x", "2026-10-10T00:00:00Z"),
            )
            connection.execute(
                "INSERT INTO pp_templates(id,name,source_path,snapshot_path,created_at) VALUES(?,?,?,?,?)",
                (self.snapshot.hash, "Synthetic", str(self.source), str(self.snapshot_path), "2026-10-10T00:00:00Z"),
            )
            connection.execute(
                "INSERT INTO pp_copies(id,name,owner_id,template_id,visibility,default_role,input_revision,calculated_revision,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?)",
                ("copy-a", "Plan", "owner", self.snapshot.hash, "private", "viewer", 0, 0,
                 "2026-10-10T00:00:00Z", "2026-10-10T00:00:00Z"),
            )
            connection.execute(
                "INSERT INTO pp_copy_shares(copy_id,user_id,role) VALUES(?,?,?)",
                ("copy-a", "editor", "editor"),
            )
            connection.execute(
                "INSERT INTO pp_copy_shares(copy_id,user_id,role) VALUES(?,?,?)",
                ("copy-a", "viewer", "viewer"),
            )

    def count(self, table):
        with self.database.connect() as connection:
            return connection.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]

    def copy(self):
        with self.database.connect() as connection:
            row = connection.execute("SELECT * FROM pp_copies WHERE id='copy-a'").fetchone()
            return dict(row)

    def override(self, sheet=0, row=1, col=1):
        with self.database.connect() as connection:
            result = connection.execute(
                "SELECT * FROM pp_overrides WHERE copy_id='copy-a' AND sheet=? AND row=? AND col=?",
                (sheet, row, col),
            ).fetchone()
            if result is None:
                return None
            value = dict(result)
            value["value"] = json.loads(value.pop("value_json"))
            return value

    def close(self):
        self.database.close()
        self._temporary.cleanup()

