# P6.5 Refresh data (2026-10-12)

Sir: the template's Power Query connections (2 connections feeding 2 query tables) must refresh from the app.

## Decision

Excel refreshes the data itself. On the copy's private Excel the server calls `RefreshAll`, then
`CalculateUntilAsyncQueriesDone`, under the same dialog watcher and 600 s limit as macros. What the refresh
wrote is found by the same macro diff (P6.3 decision 5), committed as one revision with history action
`refresh`, and published like an edit. No Power Query logic lives in the repo.

This supersedes the old pqtools plan on branch `feat/v2-refresh` (a Python re-implementation of the queries),
which is left untouched.

## Pieces

- Importer: `refreshable` in the snapshot meta when `xl/connections.xml` has a `<connection>` (IMPORT_VERSION 10).
- Engine: `ExcelHost.refresh_all`; `ExcelEngine.refresh_data` shares `_record` with `run_macro`.
- Worker/pool: the macro request named `RefreshAll` runs the refresh; no new message type.
- Server: `POST /api/copies/{id}/refresh`, sharing the macro flow.
- Web: "Refresh data" toolbar button, shown when `refreshable`.
- Check: `ops/excel_smoke.py --refresh`.

## Not done here

- Real refresh against the office workbook is not run on this Mac (no Excel). It needs the office server run in
  DEPLOY.md.
