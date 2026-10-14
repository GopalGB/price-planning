"""Counts-only native parity audit for the real price-plan template.

The controller runs this against the real workbook. It answers three questions
and prints only counts, so no company data can leave the machine:

  (a) lazy-evaluate a large random sample of formula cells after native load
      and compare each result with the value Excel cached in the file;
  (b) edit an input cell, then revert it, and confirm every evaluated affected
      cell returns to the value cached in the file;
  (c) report mismatches grouped by function name.

Cell values, cell addresses, formula text and defined names are never printed.

    .venv/bin/python -m ops.native_parity --template /path/to/template.xlsm \\
        --sample 20000 --output results-private/native-parity.json

Exit status is 0 when the audit passed and 1 when it did not.
"""

from __future__ import annotations

import argparse
import importlib.metadata
import json
import os
import random
import sys
import time
import zipfile
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path

from openpyxl.utils.cell import range_to_tuple

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from pp.engine import formualizer as adapter
from pp.engine.native_load import prepared_template

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
FUNCTIONS = ("IFERROR", "INDEX", "MATCH", "VLOOKUP", "XLOOKUP", "LET", "SWITCH", "SUMIF",
             "SUMIFS", "COUNTIFS", "OFFSET", "INDIRECT", "FILTER", "UNIQUE", "SORT", "SEQUENCE",
             "IF", "SUM", "ROUND", "AND", "OR", "NOT", "CONCAT", "TEXT", "LEN", "NUMBERVALUE",
             "COLUMN", "COLUMNS", "ROW", "ROWS", "NOW", "TODAY", "EOMONTH", "YEAR", "MONTH",
             "CHOOSE", "SUBSTITUTE", "TRIM", "LEFT", "RIGHT", "MID", "SEARCH", "FIND", "MAX",
             "MIN", "AVERAGE", "ABS", "POWER", "SQRT", "N", "T", "ISNUMBER", "ISBLANK",
             "ISERROR", "ISFORMULA", "NA", "AVERAGEIFS", "SUMPRODUCT")
EDIT_DELTA = 8


def _today():
    return datetime.now(timezone.utc).date().isoformat()


def _equivalent(actual, expected):
    """Compare two values already normalised to the adapter's JSON form."""
    if isinstance(actual, dict) or isinstance(expected, dict):
        return isinstance(actual, dict) and isinstance(expected, dict) and actual == expected
    if isinstance(actual, bool) or isinstance(expected, bool):
        return actual is expected
    if isinstance(actual, (int, float)) and isinstance(expected, (int, float)):
        return abs(actual - expected) <= 1e-9 * max(1.0, abs(expected))
    return actual == expected


def _function_names(text):
    upper = text.upper()
    return [name for name in FUNCTIONS if f"{name}(" in upper] or ["<no recognised function>"]


def _qualified(names, sheet, row, col):
    return f"{adapter.quote_sheetname(names[sheet])}!{adapter.get_column_letter(col + 1)}{row + 1}"


def _native(names, sheet, row, col):
    return names[sheet], row + 1, col + 1


MAX_PRECEDENT_CELLS = 256


def _precedent_cells(reference, current_sheet, names):
    """Expand one precedent reference into the cells it actually reads.

    Only text that parses as an A1 cell or range is used, so names, tables and
    external references are skipped rather than guessed at. A reference spanning
    more than a small block is a computed region rather than an input.
    """
    text = reference.resolved or reference.declared or reference.address
    if not isinstance(text, str):
        return ()
    try:
        sheet_name, (min_col, min_row, max_col, max_row) = range_to_tuple(text)
    except ValueError:
        return ()
    sheet = current_sheet
    if sheet_name:
        normalized = sheet_name.strip("'").replace("''", "")
        sheet = next((index for index, name in enumerate(names)
                      if name.casefold() == normalized.casefold()), None)
        if sheet is None:
            return ()
    if min_col is None or min_row is None:
        return ()
    if (max_col - min_col + 1) * (max_row - min_row + 1) > MAX_PRECEDENT_CELLS:
        return ()
    return [(sheet, row - 1, column - 1)
            for row in range(min_row, max_row + 1)
            for column in range(min_col, max_col + 1)]


class FormulaReader:
    """Read one cell's formula text on demand, so a clean audit stays cheap."""

    def __init__(self, path, names):
        self.path, self.names = Path(path), list(names)
        self._cache = {}

    def text_for(self, sheet, row, col):
        key = (sheet, row, col)
        if key not in self._cache:
            self._cache[key] = self._read(sheet, row, col)
        return self._cache[key]

    def _read(self, sheet, row, col):
        target = adapter.get_column_letter(col + 1) + str(row + 1)
        with zipfile.ZipFile(self.path) as archive:
            for name, part in adapter._sheet_parts(archive):
                if name != self.names[sheet]:
                    continue
                with archive.open(part) as stream:
                    for element in adapter._formula_cells(stream):
                        if element.get("r") != target:
                            continue
                        formula = element.find(f"{NS}f")
                        return (formula.text or "") if formula is not None else ""
        return ""


class ParityAudit:
    def __init__(self, engine, reader, sample, edits, probes, seed):
        self.engine, self.reader = engine, reader
        self.sample, self.edits, self.probes = sample, edits, probes
        self.generator = random.Random(seed)
        # Evaluating a cell drops its XML cache, so the values Excel wrote are
        # copied once and stay the oracle for every pass in this audit.
        self.baseline = dict(engine.cached_values)

    def run(self):
        sampled = self._sample()
        lazy = self._lazy_pass(sampled)
        edit = self._edit_pass(sampled)
        return {
            "lazy_sample": lazy["summary"],
            "edit_revert": edit,
            "sample_size": len(sampled),
            "pool_size": len(self.engine.formula_anchors),
        }

    def _sample(self):
        cells = sorted(self.engine.formula_anchors)
        if self.sample >= len(cells):
            return cells
        return self.generator.sample(cells, self.sample)

    def _lazy_pass(self, sampled):
        counts, by_function = Counter(), Counter()
        names = self.engine.names
        started = time.perf_counter()
        for sheet, row, col in sampled:
            counts["sampled"] += 1
            expected = self.baseline.get((sheet, row, col))
            try:
                actual = self.engine.workbook.evaluate_cell(*_native(names, sheet, row, col))
            except Exception:
                counts["evaluation_error"] += 1
                continue
            if expected is None:
                counts["uncached"] += 1
                continue
            if _equivalent(adapter._json_value(actual, self.engine.date1904), expected):
                counts["matched"] += 1
            else:
                counts["mismatched"] += 1
                self._attribute(by_function, self.reader.text_for(sheet, row, col))
        return {"summary": {"sampled": counts["sampled"], "matched": counts["matched"],
                            "mismatched": counts["mismatched"],
                            "evaluation_error": counts["evaluation_error"],
                            "uncached": counts["uncached"],
                            "seconds": round(time.perf_counter() - started, 3),
                            "mismatches_by_function": dict(sorted(by_function.items()))}}

    def _attribute(self, counter, text):
        for name in _function_names(text):
            counter[name] += 1

    def _edit_pass(self, sampled):
        counts, by_function = Counter(), Counter()
        names = self.engine.names
        started = time.perf_counter()
        for origin in sampled[:self.probes]:
            if counts["cycles"] >= self.edits:
                break
            sheet = origin[0]
            try:
                report = self.engine.workbook.precedents(_qualified(names, *origin))
            except Exception:
                counts["probe_error"] += 1
                continue
            counts["probed"] += 1
            if report.truncation.incomplete:
                counts["probe_truncated"] += 1
            for precedent in report.precedents:
                for key in _precedent_cells(precedent.reference, sheet, names):
                    if key in self.engine.formula_anchors:
                        continue
                    original = self.engine.get(*key)
                    if not isinstance(original, (int, float)) or isinstance(original, bool):
                        continue
                    counts["cycles"] += 1
                    counts.update(self._cycle(key, original, by_function))
                    break
                if counts["cycles"] >= self.edits:
                    break
        return {"cycles": counts["cycles"], "inputs_evaluated": counts["evaluated"],
                "cells_after_revert": counts["revert_checked"],
                "mismatched_after_revert": counts["revert_mismatched"],
                "probe_errors": counts["probe_error"],
                "probes_truncated": counts["probe_truncated"],
                "seconds": round(time.perf_counter() - started, 3),
                "mismatches_by_function": dict(sorted(by_function.items()))}

    def _cycle(self, key, original, by_function):
        """Edit one input, then revert it, and compare the affected cells."""
        counts = Counter()
        names = self.engine.names
        affected = self.engine._dependent_targets({key})
        baseline = {target: value
                    for target in affected
                    if target in self.engine.formula_anchors
                    and (value := self.baseline.get(target)) is not None}
        if not baseline:
            counts["no_baseline"] += 1
            return counts
        try:
            self.engine.workbook.set_value(*_native(names, *key), original + EDIT_DELTA)
            counts["evaluated"] += 1
            self.engine._evaluate_targets(affected)
            counts["revert_checked"] += len(baseline)
        finally:
            self.engine.workbook.set_value(*_native(names, *key), original)
            self.engine._evaluate_targets(affected)
            for target, expected in baseline.items():
                actual = adapter._json_value(
                    self.engine.workbook.get_value(*_native(names, *target)),
                    self.engine.date1904)
                if not _equivalent(actual, expected):
                    counts["revert_mismatched"] += 1
                    self._attribute(by_function, self.reader.text_for(*target))
        return counts


def audit(template, sample, edits, probes, seed, cache_dir=None):
    started = time.perf_counter()
    prepared = prepared_template(template, cache_dir)
    prepare_seconds = time.perf_counter() - started
    started = time.perf_counter()
    engine = adapter.FormualizerEngine.load(prepared)
    load_seconds = time.perf_counter() - started
    reader = FormulaReader(prepared, engine.names)
    report = ParityAudit(engine, reader, sample, edits, probes, seed).run()
    failures = (report["lazy_sample"]["mismatched"] + report["lazy_sample"]["evaluation_error"]
                + report["edit_revert"]["mismatched_after_revert"])
    return {
        "report_date": _today(),
        "audit": "native_parity",
        "counts_only": True,
        "engine_version": importlib.metadata.version("formualizer"),
        "workbook_bytes": Path(template).stat().st_size,
        "prepare_seconds": round(prepare_seconds, 3),
        "load_seconds": round(load_seconds, 3),
        "total_seconds": round(time.perf_counter() - started, 3),
        "status": "pass" if failures == 0 else "fail",
        "mismatch_total": failures,
        **report,
    }


def _arguments(argv):
    parser = argparse.ArgumentParser(allow_abbrev=False, description=__doc__)
    parser.add_argument("--template", type=Path, required=True)
    parser.add_argument("--sample", type=int, default=20_000)
    parser.add_argument("--edits", type=int, default=25)
    parser.add_argument("--probes", type=int, default=2_000)
    parser.add_argument("--seed", type=int, default=20261011)
    parser.add_argument("--cache-dir", type=Path)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args(argv)
    for name in ("sample", "probes"):
        if getattr(args, name) < 1:
            parser.error(f"{name} must be positive")
    if args.edits < 0:
        parser.error("edits must not be negative")
    return args


def main(argv=None):
    args = _arguments(argv)
    report = audit(args.template, args.sample, args.edits, args.probes, args.seed,
                   args.cache_dir)
    text = json.dumps(report, indent=2, sort_keys=True) + "\n"
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        pending = args.output.with_suffix(args.output.suffix + ".tmp")
        pending.write_text(text, encoding="utf-8")
        os.replace(pending, args.output)
    print(text, end="")
    return 0 if report["status"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())