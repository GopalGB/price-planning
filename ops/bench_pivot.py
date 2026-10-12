"""Measure full synthetic pivot refreshes against the service boundary."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import platform
import resource
import sys
import tempfile
import time

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from ops.fixtures import make_dashboard  # noqa: E402
from pp.db import Database  # noqa: E402
from pp.pivot import PivotQuery, PivotService  # noqa: E402
from pp.template.importer import import_template  # noqa: E402
from pp.workbook_service import TemplateSnapshot  # noqa: E402


def _arguments():
    parser = argparse.ArgumentParser()
    parser.add_argument("--rows", type=int, default=100_000)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--work-dir", type=Path, default=Path("results-private/dashboard"))
    return parser.parse_args()


def _rss_mb():
    value = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    divisor = 1024 * 1024 if platform.system() == "Darwin" else 1024
    return round(value / divisor, 2)


def _timed(action):
    started = time.perf_counter()
    value = action()
    return value, round(time.perf_counter() - started, 6)


def _seed(database, snapshot, source):
    with database.transaction() as connection:
        connection.execute(
            "INSERT INTO pp_users(id,email,name,password_hash,created_at) VALUES(?,?,?,?,?)",
            ("owner", "owner@example.test", "Owner", "x", "2026-10-11T00:00:00Z"),
        )
        connection.execute(
            "INSERT INTO pp_templates(id,name,source_path,snapshot_path,created_at) VALUES(?,?,?,?,?)",
            (snapshot.hash, "Synthetic", str(source), str(snapshot.path), "2026-10-11T00:00:00Z"),
        )
        connection.execute(
            "INSERT INTO pp_copies(id,name,owner_id,template_id,visibility,default_role,"
            "input_revision,calculated_revision,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?)",
            ("copy-a", "Plan", "owner", snapshot.hash, "private", "viewer", 0, 0,
             "2026-10-11T00:00:00Z", "2026-10-11T00:00:00Z"),
        )
        connection.execute(
            "INSERT INTO pp_results(copy_id,sheet,row,col,value_json,calculated_revision) "
            "VALUES(?,?,?,?,?,?)", ("copy-a", 1, 1, 9, "4", 0),
        )


def _options(snapshot):
    sheet = snapshot.sheets[0]
    pivot = sheet["pivots"][0]
    return {
        "page": [{"field": item["field"], "items": None} for item in pivot["page_fields"]],
        "slicers": [{"name": item["name"], "field": item["field"], "items": None}
                    for item in sheet["slicers"]],
        "collapsed": [],
    }


def _edit(database):
    with database.transaction() as connection:
        connection.execute(
            "UPDATE pp_results SET value_json='10',calculated_revision=1 "
            "WHERE copy_id='copy-a' AND sheet=1 AND row=1 AND col=9",
        )
        connection.execute(
            "UPDATE pp_copies SET input_revision=1,calculated_revision=1 WHERE id='copy-a'",
        )


def _run(root, rows):
    timings, memory = {}, {"start": _rss_mb()}
    source = root / "dashboard.xlsm"
    _, timings["fixture_generation_seconds"] = _timed(lambda: make_dashboard(source, row_count=rows))
    snapshot_path, timings["snapshot_import_seconds"] = _timed(
        lambda: import_template(source, root / "snapshot")
    )
    snapshot = TemplateSnapshot(snapshot_path)
    database = Database(root / "planning.db")
    _, timings["database_setup_seconds"] = _timed(lambda: (database.initialize(),
                                                           _seed(database, snapshot, source)))
    service, timings["source_preparation_seconds"] = _timed(lambda: PivotService(database, snapshot))
    options = _options(snapshot)
    query = PivotQuery("copy-a", "owner", 0, "DashboardPivot", options)
    result, timings["cold_refresh_seconds"] = _timed(lambda: service.refresh(query))
    memory["after_cold"] = _rss_mb()
    _, timings["warm_refresh_seconds"] = _timed(lambda: service.refresh(query))
    filtered = {**options, "slicers": [{**options["slicers"][0], "items": [0]},
                                        *options["slicers"][1:]]}
    _, timings["filtered_refresh_seconds"] = _timed(
        lambda: service.refresh(PivotQuery("copy-a", "owner", 0, "DashboardPivot", filtered))
    )
    _edit(database)
    edited, timings["edited_refresh_seconds"] = _timed(lambda: service.refresh(query))
    _, timings["edited_warm_refresh_seconds"] = _timed(lambda: service.refresh(query))
    memory["after_edit"] = _rss_mb()
    database.close()
    return {"rows": rows, "timings": timings, "rss_mb": memory,
            "cache": service.cache_info(), "output_cells": len(result["cells"]),
            "edited_revision": edited["pivot"]["calculated_revision"]}


def main():
    args = _arguments()
    args.work_dir.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(dir=args.work_dir) as temporary:
        report = _run(Path(temporary), args.rows)
    encoded = json.dumps(report, indent=2, sort_keys=True)
    if args.output:
        args.output.write_text(encoded + "\n", encoding="utf-8")
    print(encoded)


if __name__ == "__main__":
    main()
