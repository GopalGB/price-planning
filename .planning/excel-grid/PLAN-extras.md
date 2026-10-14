# Copy extras: user tabs and calculation columns (2026-10-14)

Sir's order: "user can upload their version where extra tabs or new columns as calculations. So let them
upload also while uploading show which tabs or columns exist and which is new." Grill answers: keep the
extras in that copy and show them in the app (and in downloads); cells in extras are editable like inputs.

## Definitions
- Extra tab: a sheet in the upload whose name is not a template sheet name.
- Extra column: on a template sheet, an upload column at or right of the template's first empty column
  (`filled_cols` of the template identity) that holds a value or formula. Columns inserted in the middle
  keep the current refusal ("columns may have moved").

## Behaviour
1. Preview never refuses because of extras. The structure report marks each extra tab and extra column
   "kept in your copy". Extras missing from the file but present in the copy show as "removed".
2. Confirm stores the extras for that copy (replacing earlier extras). Values the copy held inside old
   extras are replaced by the file's.
3. Grid: extra tabs appear at their file position; extra columns appear right of the template's.
   Formula cells show the values Excel saved; other cells are editable. No validation rules apply.
4. Download: the template package plus the extras (styles, formulas, widths), tab order as in the app,
   Excel recalculates on open (fullCalcOnLoad).
5. Excel engine (office server): the calculation source is the same grafted package.

## Pieces
- `pp/export/extras_graft.py`: graft extra sheets and columns from an upload into a template package;
  reorder sheets. Synthetic tests only.
- `pp/extras.py`: import the upload (existing importer, cached results) into a per-copy extras snapshot,
  `CopySnapshot` wrapper (template + extras) used by workbook, rows, cell, ops, export, engine source.
- `pp/upload.py`: extras plan in preview, inspection skips extra columns, formula row map limited to
  template columns, store/move extras on confirm.
- `web/src/views.mjs`: "kept in your copy" / "removed" labels.

## Acceptance
- Upload with one extra tab (values + formulas) and one extra formula column: preview lists both as
  kept; after confirm the grid shows them, an extra value cell accepts an edit, formulas show cached
  values; download opens in openpyxl with the extras, edits and styles; re-upload of that download
  shows the extras as same.
- Real template: append-rows and insert-row repro still pass; an extras upload previews and confirms.
- Full Python and web suites green, content guard 0.
