# Dashboard track

## Scope and authorization

2026-10-11. User GO-TOKEN approves all six dashboard steps and local acceptance. Branch feat/v2-dashboard at f74d131. Plan: ../PLAN-dashboard.md. Synthetic fixtures only; engine/calc pipeline and fidelity features are outside this track.

## Current status

Preflight complete: branch/base verified; existing snapshot/results/history/presence interfaces identified. No application code changed yet. Exact gpt-5.6-luna profile exec capability probe returned AVAILABLE. Global disk capacity fluctuated from 136 MiB to 2 GiB with approximately 7.7 GiB swap in use; execute one worker sequentially and monitor capacity. Shared .venv and web/node_modules symlinks point to the integration worktree; dependencies must not be changed.

Routing follow-up: Luna's proposed fixture used six fields and the wrong pivot range; controller rejected that scope. Terra stopped without implementation and claimed its lane unavailable, contradicted by the exact session turn_context (`gpt-5.6-terra`, high). After these two execution failures, the OMNI planner takeover uses one gpt-5.6-sol xhigh worker. The incomplete toy fixture/test is being corrected before any commit. Step 1 remains unchecked. Capacity recovered to approximately 11 GiB; no user data was removed.

## Verification

Step 1: observed missing-fixture red, missing-metadata red, and four edge tests with five expected failing assertions. Final Python suite: 303/303, 240.606 seconds, `step1-python-full-1.log`. Focused fixture/import/default-style/display-order/count/index/range/nonfinite tests green. Ruff: `ruff check pp run_v2.py tests/v2 ops/fixtures/dashboard.py`, clean, `step1-ruff-final.log`. Content guard exit 0, `step1-content-guard.log`. Full logs live in ignored `results-private/dashboard/`.

Full generated fixture has SourceData!A1:S98701, 98,701 rows including the header, 19 cache fields and 1,864,970 populated worksheet cells. Binary size 8,697,360 bytes; temporary workbook deleted automatically. The fixture deliberately omits optional pivot cache records and the importer never reads cache records. Generation wall time 105.59 seconds is not a pivot-refresh benchmark. Evidence: `step1-full-fixture.log`.

A second intermediate full suite was cancelled at the controller's phase handoff instruction; that exit 130 is not counted as a green run. Final all-phase double suites, web checks/builds, reviews, performance and browser acceptance remain pending. Historical integration tests do not verify dashboard work.

Template IMPORT_VERSION is 7. Per-sheet `pivots` and `slicers` metadata retain source and grid bounds, typed cache shared items/formats, fields/axes/report filters/measure formats, subtotal/sort/item order/hidden flags, root row/column display items, captions and style flags. Slicer metadata retains cache links, field and items with selected/no-data state, default style, sort/cross-filter flags, columns/item height, caption/start/lock settings and two-cell anchor offsets. Related-part validation rejects external/escaping/missing/malformed references, duplicate package members, invalid counts/indices, nonfinite numeric items and coordinates outside the Excel grid.

## Next task

Step 2: current-source pivot service and coherent, authorized copy reads, filtering/layout and 100k-record cold/warm benchmark. Category totals will retain existing per-copy access unless the owner answers the pending clarification. No Microsoft Excel app or Calibri/Aptos font was found locally; local screenshots cannot independently establish Excel pixel parity.
