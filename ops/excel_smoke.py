"""Office-server check of the Excel engine on the configured template (counts and timings only).

    .venv\\Scripts\\python ops\\excel_smoke.py [runtime folder] [--edits 3]

Loads the template in a private hidden Excel, edits real input cells one at a time and proves the
dependency index missed nothing: every sheet is read in full before and after the edit, and each
cell Excel changed must be among the published changes. Each edit is then reverted and must leave
no overlay. Prints no cell values.
"""

import argparse
import random
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pp.config import load_config  # noqa: E402
from pp.engine.excel import ExcelEngine, from_com, warm_caches  # noqa: E402
from pp.template.importer import import_template  # noqa: E402
from pp.values import exactly_equal  # noqa: E402
from pp.workbook_service import TemplateSnapshot  # noqa: E402


def full_read(engine, snapshot):
    values = {}
    for index, sheet in enumerate(snapshot.sheets):
        rows, cols = sheet["rows"], sheet["cols"]
        for top in range(0, rows, 2048):
            bottom = min(top + 2047, rows - 1)
            block = engine.host.read(index, top, 0, bottom, cols - 1)
            for offset, line in enumerate(block):
                for col, value in enumerate(line):
                    if value is not None:
                        values[(index, top + offset, col)] = from_com(value)
    return values


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("root", nargs="?", default=".")
    parser.add_argument("--edits", type=int, default=3)
    args = parser.parse_args()
    config = load_config(Path(args.root))
    snapshot = TemplateSnapshot(import_template(config.template_path, config.data_dir, "saved"))
    start = time.time()
    warm_caches(snapshot.source_path)
    print(f"caches ready in {time.time() - start:.1f} s", flush=True)
    start = time.time()
    engine = ExcelEngine.load(snapshot.source_path)
    print(f"Excel opened in {time.time() - start:.1f} s; Excel peak memory "
          f"{engine.memory_bytes() // 2**20} MB", flush=True)
    failures = 0
    try:
        start = time.time()
        cells = len(full_read(engine, snapshot))
        print(f"full read of every sheet: {time.time() - start:.1f} s, {cells} non-empty cells", flush=True)
        inputs = sorted(snapshot.input_mask)
        random.seed(7)
        for key in random.sample(inputs, min(args.edits, len(inputs))):
            sheet, row, col = map(int, key.split(":"))
            original = engine.get(sheet, row, col)
            trial = (original + 1) if isinstance(original, (int, float)) and not isinstance(original, bool) else 1
            before = full_read(engine, snapshot)
            start = time.time()
            published = engine.apply([(sheet, row, col, trial)])
            elapsed = time.time() - start
            after = full_read(engine, snapshot)
            changed = {cell for cell in before.keys() | after.keys()
                       if not exactly_equal(before.get(cell), after.get(cell))} - {(sheet, row, col)}
            missed = changed - {item[:3] for item in published}
            engine.apply([(sheet, row, col, original)])
            leftover = len(engine.overlay_snapshot())
            failures += bool(missed) + bool(leftover)
            print(f"edit sheet {sheet}: {elapsed:.1f} s, published {len(published)}, "
                  f"Excel changed {len(changed)}, missed {len(missed)}, overlay after revert {leftover}",
                  flush=True)
        print(f"Excel peak memory {engine.memory_bytes() // 2**20} MB")
    finally:
        engine.close()
    print("PASS" if not failures else f"FAIL ({failures})")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
