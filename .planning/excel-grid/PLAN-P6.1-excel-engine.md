# P6.1b Excel calculation engine (2026-10-12)

Decision: Sir confirmed the office Windows server has Microsoft Excel. formualizer 0.11.1 loads the real
workbook in about 1 GB but any realistic edit evaluation passed 3.5 GB, so copies calculate in real Excel.

## Goal

Editing input cells in a copy recalculates exactly as Excel does, on the office server, within the 16 GB cap,
with every changed formula result published to the existing results pipeline (grid, dashboard, export).

## Design

1. `pp/engine/depindex.py`: reverse dependency index built from the prepared template (static references,
   ranges, whole rows/columns, defined names; OFFSET/INDIRECT/volatile formulas always affected). `cone(cells)`
   returns every formula an edit can reach (numpy prefix sums; measured 1-2 s for the largest cone).
   Cached beside the prepared package, keyed by code version, like the formula map.
2. `pp/engine/excel.py`: `ExcelEngine` with the worker-facing API (`apply`, `get`, `overlay_snapshot`,
   `overlay_delta`, `formula_cells`, `formula_anchors`, `handles_clears`, `close`). Excel recalculates
   natively; the engine reads back only the cone's row blocks in bulk and diffs them against the last
   known value (Excel's saved value until changed).
   Hosts: `ComExcelHost` (pywin32, one hidden Excel per worker, macros force-disabled, events off,
   manual calculation + `Calculate()` after each batch, own working copy of the template file) and
   `FormualizerHost` used by the Mac tests.
3. Settings `"engine": "excel" | "formualizer"` (default formualizer); the pool passes it to workers.
   Worker skips the formualizer-only clear/reload path when the engine handles clears.
4. Pool memory budget counts the Excel process (peak working set), and workers quit Excel and kill
   the Excel PID on exit.

## Acceptance

- Unit tests on Mac with `FormualizerHost`: edits, clears, chains, ranges, names, cross-sheet, overlay delta.
- COM value conversion (errors, dates, booleans, text) tested as pure functions.
- Office server (Sir runs it): real template, edit three input cells, compare against Excel desktop,
  time per edit, memory per copy. Until then values-only mode stays the Mac default.

## Later phases

Macros as actions (8 subs; 4 MsgBox and 1 InputBox need non-blocking handling), Power Query RefreshAll,
import (feat/v2-transfer, finishing on Space Bunny).
