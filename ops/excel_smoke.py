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

--slicers selects one item (with data) of each Excel slicer and restores all items, on a fresh engine;
then a second fresh engine applies the final stored state and must read the same pivot cells.
Prints counts, timings and slicer numbers only, never item texts or values.
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
    for index in range(len(snapshot.sheets)):
        box = engine.host.used_range(index)  # content only: formatting spans 16384 columns on most tabs
        if box is None:
            continue
        first, left, last, right = box
        for top in range(first, last + 1, 2048):
            bottom = min(top + 2047, last)
            block = engine.host.read(index, top, left, bottom, right)
            for offset, line in enumerate(block):
                for col, value in enumerate(line):
                    if value is not None:
                        values[(index, top + offset, left + col)] = from_com(value)
    return values


def parse_cells(text):
    """'sheet:row:col=value,...' (zero-based), numbers as numbers, anything else as text."""
    cells = []
    for item in text.split(","):
        key, _, raw = item.partition("=")
        sheet, row, col = map(int, key.split(":"))
        try:
            value = float(raw)
            value = int(value) if value.is_integer() else value
        except ValueError:
            value = raw
        cells.append((sheet, row, col, value))
    return cells


def check_cells(engine, snapshot, cells, report=print):
    """Type several cells at once, as a user filling a row; every cell Excel changes must be published."""
    originals = [(sheet, row, col, engine.get(sheet, row, col)) for sheet, row, col, _ in cells]
    before = full_read(engine, snapshot)
    start = time.time()
    published = engine.apply(cells)
    elapsed = time.time() - start
    after = full_read(engine, snapshot)
    typed = {cell[:3] for cell in cells}
    changed = {cell for cell in before.keys() | after.keys()
               if not exactly_equal(before.get(cell), after.get(cell))} - typed
    missed = changed - {item[:3] for item in published}
    sheets = sorted({cell[0] for cell in changed})
    engine.apply(originals)
    leftover = len(engine.overlay_snapshot())
    report(f"typed {len(cells)} cells: {elapsed:.1f} s, published {len(published)}, Excel changed {len(changed)} "
           f"on sheets {sheets}, missed {len(missed)}, overlay after revert {leftover}")
    return bool(missed) + bool(leftover)


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


def _pivot_values(engine):
    return {cell: engine.get(*cell) for cell in engine.pivot_cells}


def check_slicers(snapshot, open_engine, report=print):
    """Select then restore each Excel slicer on a fresh engine; a fresh replay of the final state must match."""
    slicers = [item for sheet in snapshot.sheets for item in sheet.get("slicers", [])
               if item.get("engine") == "excel"]
    if not slicers:
        report("This template has no Excel slicers.")
        return 1
    failures = 0
    for number, slicer in enumerate(slicers, start=1):
        captions = [item["caption"] for item in slicer["items"]]
        pick = next((item["caption"] for item in slicer["items"] if not item["no_data"]), None)
        if pick is None:
            report(f"slicer {number}: no item with data")
            failures += 1
            continue
        engine = open_engine(snapshot.source_path)
        try:
            start = time.time()
            selected = engine.set_slicer(slicer["cache"], [pick])
            select_seconds = time.time() - start
            start = time.time()
            restored = engine.set_slicer(slicer["cache"], captions)
            restore_seconds = time.time() - start
            expected = _pivot_values(engine)
        except EngineError as exc:
            report(f"slicer {number}: did not finish ({type(exc).__name__})")
            failures += 1
            continue
        finally:
            engine.close()
        report(f"slicer {number}: select {select_seconds:.1f} s, published {len(selected['changes'])}; "
               f"restore {restore_seconds:.1f} s, published {len(restored['changes'])}")
        replay = open_engine(snapshot.source_path)
        try:
            start = time.time()
            replay.restore_pivots({"slicers": {slicer["cache"]: captions}, "refreshed": False})
            actual = _pivot_values(replay)
            seconds = time.time() - start
        except EngineError as exc:
            report(f"  replay did not finish ({type(exc).__name__})")
            failures += 1
            continue
        finally:
            replay.close()
        mismatched = sum(1 for cell in expected.keys() | actual.keys()
                         if not exactly_equal(expected.get(cell), actual.get(cell)))
        report(f"  replay: {seconds:.1f} s, pivot cells {len(expected)}, mismatched pivot cells {mismatched}")
        failures += bool(mismatched)
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
    parser.add_argument("--slicers", action="store_true", help="select and restore each Excel slicer")
    parser.add_argument("--cells", help="type these cells together, e.g. 4:22:1=GROUP,4:22:7=10 (zero-based)")
    args = parser.parse_args()
    config = load_config(Path(args.root))
    snapshot = TemplateSnapshot(import_template(config.template_path, config.data_dir, "saved"))
    start = time.time()
    warm_caches(snapshot.source_path)
    print(f"caches ready in {time.time() - start:.1f} s", flush=True)
    if args.macros:
        return 1 if check_macros(snapshot, ExcelEngine.load) else 0
    if args.slicers:
        return 1 if check_slicers(snapshot, ExcelEngine.load) else 0
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
        if args.cells:
            failures += check_cells(engine, snapshot, parse_cells(args.cells))
        inputs = [] if args.cells else sorted(snapshot.input_mask)
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
