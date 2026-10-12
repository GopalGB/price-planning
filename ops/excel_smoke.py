"""Office-server check of the Excel engine on the configured template (counts and timings only).

    .venv\\Scripts\\python ops\\excel_smoke.py [runtime folder] [--edits 3]
    .venv\\Scripts\\python ops\\excel_smoke.py [runtime folder] --macros

Loads the template in a private hidden Excel, edits real input cells one at a time and proves the
dependency index missed nothing: every sheet is read in full before and after the edit, and each
cell Excel changed must be among the published changes. Each edit is then reverted and must leave
no overlay. Prints no cell values.

--macros runs every template button on a fresh engine, then proves a second fresh engine that
applies only the recorded contents reads identical values everywhere. Prints counts and timings only.

--refresh runs Excel's RefreshAll on a fresh engine, then replays the recorded contents on a second
fresh engine and compares every value. Needs a template with data connections. Counts only.
"""

import argparse
import random
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pp.config import load_config  # noqa: E402
from pp.engine.excel import ExcelEngine, from_com, warm_caches  # noqa: E402
from pp.engine.formualizer import EngineError  # noqa: E402
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


def check_macros(snapshot, open_engine, report=print):
    """Each button on a fresh engine, then a fresh replay of its contents; returns the number of failures."""
    failures = 0
    for button in snapshot.buttons:
        try:
            outcome, expected = _run_on_fresh(snapshot, open_engine, button)
        except EngineError as exc:
            report(f"button {button['id']}: macro did not finish ({type(exc).__name__})")
            failures += 1
            continue
        report(f"button {button['id']}: {outcome['seconds']:.1f} s, messages {outcome['messages']}, "
               f"contents {outcome['contents']}, formula contents {outcome['formulas']}, "
               f"published {outcome['published']}, Excel peak {outcome['peak_mb']} MB")
        try:
            start = time.time()
            replay = open_engine(snapshot.source_path)
            try:
                replay.apply(outcome["values"])
                actual = full_read(replay, snapshot)
            finally:
                replay.close()
        except EngineError as exc:
            report(f"  replay did not finish ({type(exc).__name__})")
            failures += 1
            continue
        mismatched = sum(1 for cell in expected.keys() | actual.keys()
                         if not exactly_equal(expected.get(cell), actual.get(cell)))
        report(f"  replay: {time.time() - start:.1f} s, mismatched cells {mismatched}")
        failures += bool(mismatched)
    report("PASS" if not failures else f"FAIL ({failures})")
    return failures


def check_refresh(snapshot, open_engine, report=print):
    """RefreshAll on a fresh engine, then a fresh replay of its contents; returns the number of failures."""
    start = time.time()
    engine = open_engine(snapshot.source_path)
    try:
        outcome = engine.refresh_data()
        seconds = time.time() - start
        expected = full_read(engine, snapshot)
        peak_mb = engine.memory_bytes() // 2**20
    except EngineError as exc:
        report(f"refresh did not finish ({type(exc).__name__})")
        report("FAIL (1)")
        return 1
    finally:
        engine.close()
    report(f"refresh: {seconds:.1f} s, messages {len(outcome['messages'])}, "
           f"contents {len(outcome['contents'])}, published {len(outcome['changes'])}, "
           f"Excel peak {peak_mb} MB")
    try:
        start = time.time()
        replay = open_engine(snapshot.source_path)
        try:
            replay.apply(outcome["contents"])
            actual = full_read(replay, snapshot)
        finally:
            replay.close()
    except EngineError as exc:
        report(f"  replay did not finish ({type(exc).__name__})")
        report("FAIL (1)")
        return 1
    mismatched = sum(1 for cell in expected.keys() | actual.keys()
                     if not exactly_equal(expected.get(cell), actual.get(cell)))
    report(f"  replay: {time.time() - start:.1f} s, mismatched cells {mismatched}")
    failures = bool(mismatched)
    report("PASS" if not failures else f"FAIL ({failures})")
    return failures


def _run_on_fresh(snapshot, open_engine, button):
    start = time.time()
    engine = open_engine(snapshot.source_path)
    try:
        outcome = engine.run_macro(button["macro"])
        seconds = time.time() - start
        expected = full_read(engine, snapshot)
        peak_mb = engine.memory_bytes() // 2**20
    finally:
        engine.close()
    contents = outcome["contents"]
    summary = {"seconds": seconds, "messages": len(outcome["messages"]), "contents": len(contents),
               "formulas": sum(isinstance(value, dict) for *_, value in contents),
               "published": len(outcome["changes"]), "peak_mb": peak_mb, "values": contents}
    return summary, expected


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("root", nargs="?", default=".")
    parser.add_argument("--edits", type=int, default=3)
    parser.add_argument("--macros", action="store_true", help="run the template buttons instead of edits")
    parser.add_argument("--refresh", action="store_true", help="refresh the data connections instead of edits")
    args = parser.parse_args()
    config = load_config(Path(args.root))
    snapshot = TemplateSnapshot(import_template(config.template_path, config.data_dir, "saved"))
    start = time.time()
    warm_caches(snapshot.source_path)
    print(f"caches ready in {time.time() - start:.1f} s", flush=True)
    if args.macros:
        return 1 if check_macros(snapshot, ExcelEngine.load) else 0
    if args.refresh:
        if not snapshot.refreshable:
            print("This template has no data connections to refresh.")
            return 1
        return 1 if check_refresh(snapshot, ExcelEngine.load) else 0
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
