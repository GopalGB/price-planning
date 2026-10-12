# P6.4 Dashboard: pivot table and slicers (2026-10-12)

Sir: the dashboard must work in place, like the template. Built on the Excel engine; the Mac and
values-only mode show the saved pivot and slicers read-only.

## Facts (structure only, from the real template)

- One pivot table, on the "Coupon" sheet: body A5:F14 plus 3 report filters above it, 5 data fields.
- Four slicers on that pivot (one slicers part, four slicer caches, all tabular, not OLAP). Item
  counts 10, 78, 52 and 769; one slicer has a single item selected, the others have all selected;
  many items are flagged "no data" (`nd="1"`). The pivot cache records are 27 MB.
- One more pivot table part exists in the package; every pivot is found through Excel at run time.

## Revision 2026-10-12 13:30 (after merging feat/v2-dashboard)

The earlier dashboard track (merged, ed352f2) already rebuilds pivots from workbook cells in Python:
in-grid slicers, report filters, personal views, category totals and the team overview. It cannot
rebuild the real template's pivot, whose cache source is the external Power Query connection, so
that pivot now imports as Excel's saved values (fff35f6) instead of failing the import. The decisions
below apply to such Excel-only pivots; the D2 UI reuses the merged slicer frames instead of new ones.

## Decisions

1. Pivot output is computed, like formula results. Cells inside any pivot table's `TableRange2` are
   never stored as overrides: `ExcelEngine` reads the pivot ranges at load (`ExcelHost.pivot_ranges`),
   adds those cells to the cells it publishes (baseline = the template's saved values), and drops
   them from macro/refresh `contents`. Reason: Excel refuses writes inside a pivot table, so a
   replayed refresh would otherwise fail on reload. This also fixes P6.3/P6.5 replay.
2. Pivot state per copy is persisted as actions, not cells: table `pp_pivot_state(copy_id,
   state_json, revision)` holds the slicer selections `{cache name: [selected item names]}` and
   whether the pivots were refreshed. The worker receives it with every request and, after a fresh
   load and the overrides, re-applies it in Excel (selections, then `PivotTable.RefreshTable` when
   refreshed), then publishes. A "Refresh data" run marks the copy refreshed.
3. Slicers in the UI: the importer reads each slicer (name, caption, cache, column count, anchor
   from the sheet drawing, items in cache order with selected and no-data flags; item text from the
   pivot cache shared items) into `slicers.json`; the grid draws them where Excel draws them: caption
   header, a grid of item buttons (`columnCount`), selected items highlighted, no-data items dimmed and
   sorted last like Excel, a "clear filter" button. Click selects one item; Ctrl/Cmd-click toggles
   (Excel's multi-select). Lists scroll inside the slicer box.
4. Changing a slicer: `POST /api/copies/<id>/slicers` `{op_id, cache, selected: [names]}`; editor
   role; runs in the copy's Excel (`SlicerCaches(name)`, pivot `ManualUpdate` on while items are set,
   then off), records the pivot output change through the macro diff path (pivot cells published as
   results), stores the new selection in `pp_pivot_state`, history action `slicer`. Viewers and the
   Mac/values-only mode see slicers read-only. Each user's view follows the copy (shared, like Excel's
   shared workbook), not per-user filters.
5. Not in this phase: totals across copies and a team overview (needs Sir's definition of which
   numbers roll up), charts and pictures (P6.2).

## Tracks

- D1 engine + flow: decision 1 (pivot ranges as computed cells, contents filter), decision 2
  (state table, worker re-apply, refresh marks refreshed), decision 4 server route and engine
  `set_slicer(cache, selected)`. Fake-host tests; `ops/excel_smoke.py --slicers` (counts only).
- D2 importer + UI: decision 3 and the slicer click flow against the D1 route.

## Acceptance

Mac: all suites green; slicers drawn on Coupon with the template's selections; read-only here.
Office server: `excel_smoke.py --slicers` PASS (set each slicer to one item and back, pivot cells
published, replay on a fresh engine equal); in the browser a slicer click updates the pivot.
