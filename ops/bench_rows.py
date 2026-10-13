"""Measure what opening a sheet tab costs, on a synthetic large sheet, twice.

    .venv/bin/python ops/bench_rows.py --rows 1200 --cols 24 --count 256

The workbook is generated into a temporary directory and deleted afterwards: nothing is written
into the repository and no company workbook is read. Only the first and second visit to one row
block are timed, so the number is the render plus the transfer the grid waits for.
"""
import argparse
import gzip
import json
from pathlib import Path
import sys
import tempfile
import time

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from ops.fixtures import make_large_grid  # noqa: E402
from pp.db import Database  # noqa: E402
from pp.template.importer import import_template  # noqa: E402
from pp.workbook_service import RowQuery, TemplateSnapshot, WorkbookService  # noqa: E402

OWNER = "bench-owner"


def build(root, rows, cols, sheets):
    """Import the generated workbook and give one copy an owner to read."""
    source = root / "large-grid.xlsx"
    make_large_grid(source, rows, cols, sheets)
    snapshot_path = import_template(source, root / "data")
    snapshot = TemplateSnapshot(snapshot_path)
    database = Database(root / "bench.db")
    database.initialize()
    with database.transaction() as connection:
        connection.execute(
            "INSERT INTO pp_users(id,email,name,password_hash,created_at) VALUES(?,?,?,?,?)",
            (OWNER, "bench@example.test", "Bench Owner", "x", "2026-10-13T00:00:00Z"),
        )
        connection.execute(
            "INSERT INTO pp_templates(id,name,source_path,snapshot_path,created_at) VALUES(?,?,?,?,?)",
            (snapshot.hash, "Large Grid", str(source), str(snapshot_path), "2026-10-13T00:00:00Z"),
        )
        connection.execute(
            "INSERT INTO pp_copies(id,name,owner_id,template_id,visibility,default_role,"
            "input_revision,calculated_revision,created_at,updated_at) "
            "VALUES(?,?,?,?,?,?,?,?,?,?)",
            ("bench-copy", "Bench", OWNER, snapshot.hash, "private", "viewer", 0, 0,
             "2026-10-13T00:00:00Z", "2026-10-13T00:00:00Z"),
        )
    return database, snapshot


def visit(service, sheet, count, label, report):
    started = time.perf_counter()
    body = service.rows(RowQuery("bench-copy", OWNER, sheet, 0, count))
    elapsed = time.perf_counter() - started
    payload = json.dumps(body).encode()
    report.append((label, elapsed, len(payload), len(gzip.compress(payload, 6))))
    return body


def source_bytes(snapshot):
    return sum(path.stat().st_size for path in snapshot.path.rglob("*.json")) // 1024


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--rows", type=int, default=1200)
    parser.add_argument("--cols", type=int, default=24)
    parser.add_argument("--count", type=int, default=256)
    parser.add_argument("--sheets", type=int, default=2)
    options = parser.parse_args()

    with tempfile.TemporaryDirectory(prefix="pp-bench-rows-") as directory:
        root = Path(directory)
        build_started = time.perf_counter()
        database, snapshot = build(root, options.rows, options.cols, options.sheets)
        cells = sum(sheet["rows"] * sheet["cols"] for sheet in snapshot.sheets)
        print(f"generated and imported {len(snapshot.sheets)} sheets of "
              f"{options.rows} x {options.cols} in {time.perf_counter() - build_started:.1f} s "
              f"({cells:,} cells, {source_bytes(snapshot):,} kB of snapshot)")
        service = WorkbookService(database, snapshot)
        report = []
        try:
            for sheet in range(options.sheets):
                visit(service, sheet, options.count, f"sheet {sheet} first visit", report)
                visit(service, sheet, options.count, f"sheet {sheet} second visit", report)
        finally:
            cache = service.row_cache
            database.close()
    print(f"cache holds {cache.entries} blocks in {cache.bytes_held / 1024:.0f} kB "
          f"({cache.hits} hits, {cache.misses} misses)")
    print(f"{'visit':<28}{'seconds':>10}{'json bytes':>14}{'gzip bytes':>14}")
    for label, elapsed, size, packed in report:
        print(f"{label:<28}{elapsed:>10.3f}{size:>14,}{packed:>14,}")


if __name__ == "__main__":
    main()