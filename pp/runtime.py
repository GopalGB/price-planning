"""Startup, recovery, and shutdown composition for the opt-in v2 server."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

from .config import Config, load_config
from .db import Database
from .engine.pool import CalculationRequest, EnginePool
from .ops import OperationService
from .recovery import recover_unfinished
from .server import ServiceContext, create_app
from .template.importer import import_template
from .workbook_service import TemplateSnapshot


@dataclass
class Runtime:
    config: Config
    database: Database
    snapshot: TemplateSnapshot
    pool: EnginePool
    app: object
    _closed: bool = False

    def close(self):
        if self._closed:
            return
        self.pool.close()
        self.database.close()
        self._closed = True


def create_runtime(root: str | Path) -> Runtime:
    root = Path(root)
    config = load_config(root)
    if config.template_path is None or not config.template_path.is_file():
        raise ValueError("template_path must name an existing .xlsm file")
    snapshot_path = import_template(config.template_path, config.data_dir)
    snapshot = TemplateSnapshot(snapshot_path)
    database = Database(config.database_path)
    database.initialize()
    _register_template(database, snapshot)
    database.migrate_legacy(snapshot.hash)
    pool = EnginePool(config.max_workers)
    try:
        operations = OperationService(database, snapshot)

        def calculate(copy_id, revision):
            source = _copy_source(database, copy_id)
            changes = operations.revision_changes(copy_id, revision)
            return pool.apply(CalculationRequest(copy_id, revision, source, changes))

        recover_unfinished(database, calculate, snapshot)
        app = create_app(ServiceContext(database, snapshot, config, pool))
        return Runtime(config, database, snapshot, pool, app)
    except Exception:
        pool.close()
        database.close()
        raise


def _register_template(database, snapshot):
    now = datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")
    with database.transaction() as connection:
        connection.execute(
            "INSERT INTO pp_templates(id,name,source_path,snapshot_path,created_at) VALUES(?,?,?,?,?) "
            "ON CONFLICT(id) DO UPDATE SET name=excluded.name,source_path=excluded.source_path,snapshot_path=excluded.snapshot_path",
            (snapshot.hash, snapshot.name, str(snapshot.source_path), str(snapshot.path), now),
        )


def _copy_source(database, copy_id):
    with database.connect() as connection:
        row = connection.execute(
            "SELECT t.source_path FROM pp_copies c JOIN pp_templates t ON t.id=c.template_id WHERE c.id=?",
            (copy_id,),
        ).fetchone()
    if row is None:
        raise ValueError("calculation copy is missing")
    return Path(row["source_path"])
