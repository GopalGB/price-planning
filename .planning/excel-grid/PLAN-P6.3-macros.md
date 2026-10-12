# P6.3 Template macros (2026-10-12)

Sir: "macros and all things also should be working". SPEC section 2 had VBA macros out of scope for v1;
this phase brings them in on the Excel engine only.

## Facts (counts only, from the real template)

- The template's own buttons name their macros in the file: DrawingML shapes (`macro="[0]!name"`) and
  legacy VML form buttons (`<x:FmlaMacro>`). Five buttons on four sheets cover every user-facing macro.
- What the macros change, measured against the imported snapshot:
  - one rewrites 91,000 locked constant cells with new values;
  - one turns 29,000 editable input cells into formulas, keeps "X"/"XB" results as plain values and
    deletes its own button;
  - two rewrite formulas that equal the template's own formulas (no change in the app, where formula
    cells cannot be overwritten);
  - one rebuilds a report sheet of constants (values; its formatting is not carried).
- Message boxes: OK-only MsgBox calls; the one InputBox is commented out.

## Decisions

1. Run the template's real VBA inside the copy's private Excel (exact behaviour, no VBA logic in the
   repo). The private Python port in `~/PricePlanning-private/actions/` stays unused and uncommitted.
2. Only the template's buttons can be run (the button list is the allow-list). The server resolves a
   button id to its macro; clients never send macro names.
3. Excel opens with macros allowed (`AutomationSecurity = 1`) and events still off. Every workbook an
   engine opens is the admin-imported template, so only its VBA ever runs; uploads carry values only.
4. Dialogs never block: while a macro runs, a watcher thread presses OK on every dialog window owned by
   that Excel process and records its text. The texts are shown to the user as the macro's messages.
   A macro that runs past the time limit (default 600 s) gets its Excel killed; the copy reloads.
5. What a macro changed is found by Excel itself, not by guessing: before and after the run the engine
   reads every sheet's used range (values) and its formula areas (`SpecialCells(xlCellTypeFormulas)`).
   - cells that became formulas (were not formulas before) -> formula overrides `{"formula": "=..."}`
     (text read with `Formula2`);
   - non-formula cells whose value changed -> value overrides (any cell, not only inputs);
   - formula cells whose result changed -> published results, as for an edit.
   Known limit: a macro that replaces an existing formula with a different formula is not tracked
   (none of this template's macros does; checked against the formula map).
6. Persistence reuses the desired-state model: macro outputs are stored in `pp_overrides` (attributed
   to the user who pressed the button, history action `macro`) and replayed by the worker like inputs.
   `ExcelEngine.apply` accepts `{"formula": ...}` values (written with `Formula2`).
7. Formula overrides are not in the template's dependency index. The engine keeps a small second index
   over them (same `DependencyIndex.build`, override formulas only) and computes the cone as a fixpoint
   over both indexes, so edits that reach an override formula re-read it and its dependents.
8. While a macro job is pending for a copy, edits to that copy get 409 `macro_running`; a server restart
   marks a pending macro job failed (the user presses the button again) instead of replaying it.
9. Formualizer and values-only deployments return 409 `macros_unavailable` with the reason.

## Tracks

- M1 (importer + API + UI): extract buttons into the snapshot (`buttons.json`: id, sheet, from/to
  cell, label, macro), include them in the workbook payload, draw them over the grid where Excel draws
  them, POST `/api/copies/<id>/macros` with `{button, op_id}`; until M3 lands the endpoint answers 409
  `macros_unavailable`. Messages and errors show in a dialog.
- M2 (engine): host methods `run_macro`, `formula_areas`, `read_formulas`, `write_formulas`,
  `used_range`; `ComExcelHost` implementation incl. the dialog watcher and time limit;
  `ExcelEngine.run_macro(macro) -> {messages, changes, contents}`; formula overrides in `apply`;
  second index + fixpoint cone. Tests on a fake host that runs Python callables as macros.
- M3 (flow): worker/pool request kind `macro`, ops/server job and history, override display for
  non-input cells, export writes formula overrides as formulas, recovery rule, `ops/excel_smoke.py
  --macros` (runs each button on a private engine, then proves a fresh engine replaying the recorded
  overrides reads identical values everywhere; prints counts and timings only).

## Acceptance

- Mac: all tests green; buttons appear on the right sheets at the right cells; pressing one on the
  Mac explains that macros run on the office server.
- Office server (Sir): `excel_smoke.py --macros` prints PASS for every button; in the browser, pressing
  each button shows Excel's message and the changed cells, and they survive a server restart.
