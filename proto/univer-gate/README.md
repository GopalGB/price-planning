# Univer gate (synthetic only)

P0 Task 7 prototype for SPEC 3.5. Source prepared on 2026-10-10 with no network,
npm install, browser run or Git writes. The pinned plan starting point is
`@univerjs/presets` 1.0.4, Vite 6.3.5 and Playwright 1.55.0. Package availability,
exports and browser behavior are UNVERIFIED. This prototype is not a P0 pass.

## Run (Claude, outside the restricted executor sandbox)

Only run when free disk exceeds 5 GB. From this directory:

```sh
npm install
npx playwright install chromium
npm run build
npm run gate
```

Commit the generated package-lock.json after verifying pins. Do not commit node_modules
or dist. This package intentionally has no CI job: the plan calls for a local,
memory-guarded gate. The runner uses the installed Vite binary, never an npx download.

The report is `.planning/excel-grid/results/univer-gate-<UTC-date>.json`, relative
to the repo root. A failure returns exit 1 and retains aggregate measurements.
Missing heap readings, page errors, API failures or missing checks cannot pass.
The preview and browser are closed on success and failure. Raw diagnostics belong
in terminal output or results-private, never in the public results directory.

## What is measured

Each sheet runs separately in a fresh browser context in both full and windowed modes:
4088 x 312 and 100000 x 17. The 200-row window explicitly clears its predecessor.
Snapshot stored-row counts check eviction as well as populated cells. Heap is measured
before operations, after operations and after four window replacements. All readings
must be positive and below 1500 MB. requestAnimationFrame samples during wheel scrolling
must reach 60 fps; this is a frame-cadence proxy, not a compositor/rendering trace.
Do not lower thresholds to obtain green.

`window.gate` exposes loadWindow(sheetId, firstRow, rowCount), heapMB(),
formulaCellCount(), pasteTsv(ref, tsv), undo(), plus measurement helpers.
Only column B is accepted by the scripted input helpers. Paste accepts one TSV
column, maximum 200 rows. Leading =, +, -, @ are refused. The page keeps the
plan's NO_CALCULATION initialization setting; no formulas are included in snapshots.
The native edit hook is deliberately tested with keyboard formula entry. Formula
counts are checked after each scripted operation and immediately after the native
probe, before any undo can conceal a created formula. Values in column C must
survive input edits, paste, fill and undo.

The scripted paste/fill helpers exercise setValues, not native clipboard/fill-handle
paths. Zero formula cells does not prove zero engine work. Native input locking,
clipboard, fill, undo, offscreen copy/find and visual fidelity require separate
evidence. Do not interpret a helper-only success as SPEC acceptance.

## Complete the remaining evidence

Inspect the pinned package and instrument its calculation service/events. Prove
zero recalculation for native edit, paste, fill and undo, including preservation
of authoritative server values. Verify input-only editing, native clipboard/fill,
offscreen copy/find with windows evicted and restored, and merges/freezes/styles.
Use both sheet sizes. Keep traces and raw details in results-private.

After observing these behaviors, prepare an aggregate JSON evidence file with this
shape (false below means not yet proved):

```json
{
  "univer_version": "1.0.4",
  "sheets": {
    "grid_4088x312": {
      "zero_client_recalculation": false,
      "native_edit_paste_fill_undo": false,
      "offscreen_copy_find": false,
      "merges_freezes_styles_editing": false
    },
    "rows_100000x17": {
      "zero_client_recalculation": false,
      "native_edit_paste_fill_undo": false,
      "offscreen_copy_find": false,
      "merges_freezes_styles_editing": false
    }
  }
}
```

Run `npm run gate -- /absolute/path/to/evidence.json`. Pass requires every measured
check and every evidence flag for both sheets. Without evidence the report remains
`pass: false`. Evidence flags are reviewer attestations, not automatic proof.
If APIs fail or eviction is impossible, retain failure evidence, update STATE.md,
and evaluate the SPEC's Glide fallback against the same criteria.

## VERIFY

These names/contracts could not be checked against installed 1.0.4 or live docs:

- Exports `createUniver`, `LocaleType.EN_US`, `mergeLocales` from `@univerjs/presets`;
  createUniver's return member `univerAPI` and accepted options `locale`, `locales`, `presets`.
- Export `UniverSheetsCorePreset` from `@univerjs/presets/preset-sheets-core`;
  default `enUS` export from its `/locales/en-US` path; CSS path
  `@univerjs/presets/lib/styles/preset-sheets-core.css`.
- Preset options `container`, `formula.initialFormulaComputing` and literal
  `NO_CALCULATION`. In particular, whether this only affects initialization.
- Facade methods `createWorkbook`, `getActiveWorkbook`, `getSheetBySheetId`,
  workbook `save`, worksheet `getRange(row, column, rowCount?, columnCount?)`.
- Range methods `setValues` (2D CellData objects retaining style IDs and no formulas),
  `getValue`, `clear` (actual sparse storage eviction, no retained undo copies),
  `activate` (scroll/focus semantics), and `univerAPI.undo` (asynchronous completion
  and undo grouping for each setValues call).
- `univerAPI.addEvent`, `univerAPI.Event.BeforeSheetEditEnd`, event payload
  `value` / `context.value` / CellData `v`, `f`, `si`, and `event.cancel = true`.
  Verify native typing cancellation and find the documented equivalents for paste/fill
  interception. Those paths are not claimed covered by this edit event.
- Workbook snapshot schema `id`, `name`, `appVersion`, `locale`, `sheetOrder`,
  `sheets`, `styles`; sheet `rowCount`, `columnCount`, sparse `cellData`,
  `mergeData` bounds and `freeze` fields `xSplit`, `ySplit`, `startRow`, `startColumn`;
  cell style references `s` with `bg.rgb` and `cl.rgb`.
- Calculation-service/event API names for continuous zero-recalculation observation
  are unknown and intentionally not invented. Claude must identify/instrument the
  installed version before setting `zero_client_recalculation: true`.

Record any resolved API substitutions here and append deviations to
`.planning/excel-grid/sol/deviations.md`; retain the sizes and acceptance thresholds.

## RESOLVED (Claude, 2026-10-10, verified in Ego Lite against installed 1.0.4)

- Sheets preset is the separate package `@univerjs/preset-sheets-core` (also its `/locales/en-US` and `/lib/index.css`).
- Locale constant is `LocaleType.En_US`; calculation mode is `CalculationMode.NO_CALCULATION` (startup only).
- `BeforeSheetEditEnd` passes a RichTextValue: text is `value.getData().body.dataStream`. Setting `cancel = true`
  leaves the editor open, so the handler calls `getActiveWorkbook().endEditingAsync(false)`, which re-fires the
  event; a re-entry flag stops the loop. Without these three, typed formulas were created AND calculated by Univer.
- The gate now runs in Ego Lite (`gate-ego.mjs`); Playwright was removed. Result: `.planning/excel-grid/results/univer-gate-2026-10-10.json`.
