# Dashboard track

## Scope and authorization

2026-10-11. User GO-TOKEN approves all six dashboard steps and local acceptance. Branch feat/v2-dashboard at f74d131. Plan: ../PLAN-dashboard.md. Synthetic fixtures only; engine/calc pipeline and fidelity features are outside this track.

## Current status

Preflight complete: branch/base verified; existing snapshot/results/history/presence interfaces identified. Step 1 import is complete. Exact gpt-5.6-luna profile exec capability probe returned AVAILABLE. Global disk capacity fluctuated from 136 MiB to 2 GiB with approximately 7.7 GiB swap in use; execute one worker sequentially and monitor capacity. Shared .venv and web/node_modules symlinks point to the integration worktree; dependencies must not be changed.

Routing follow-up: Luna's proposed fixture used six fields and the wrong pivot range; controller rejected that scope. Terra stopped without implementation and claimed its lane unavailable, contradicted by the exact session turn_context (`gpt-5.6-terra`, high). After these two execution failures, the OMNI planner takeover uses one gpt-5.6-sol xhigh worker. The incomplete toy fixture/test was replaced with the approved full metadata shape before committing Step 1. Capacity recovered to approximately 11 GiB; no user data was removed.

## Verification

Step 1: observed missing-fixture red, missing-metadata red, and four edge tests with five expected failing assertions. Final Python suite: 303/303, 240.606 seconds, `step1-python-full-1.log`. Focused fixture/import/default-style/display-order/count/index/range/nonfinite tests green. Ruff: `ruff check pp run_v2.py tests/v2 ops/fixtures/dashboard.py`, clean, `step1-ruff-final.log`. Content guard exit 0, `step1-content-guard.log`. Full logs live in ignored `results-private/dashboard/`.

Full generated fixture has SourceData!A1:S98701, 98,701 rows including the header, 19 cache fields and 1,864,970 populated worksheet cells. Binary size 8,697,360 bytes; temporary workbook deleted automatically. The fixture deliberately omits optional pivot cache records and the importer never reads cache records. Generation wall time 105.59 seconds is not a pivot-refresh benchmark. Evidence: `step1-full-fixture.log`.

A second intermediate full suite was cancelled at the controller's phase handoff instruction; that exit 130 is not counted as a green run. Final all-phase double suites, web checks/builds, reviews, performance and browser acceptance remain pending. Historical integration tests do not verify dashboard work.

Template IMPORT_VERSION is 8. Per-sheet `pivots` and `slicers` metadata retain source and grid bounds, typed cache shared items/formats, fields/axes/report filters/measure formats, subtotal/sort/item order/hidden flags, root row/column display items, captions and style flags. Slicer metadata retains cache links, field and items with selected/no-data state, default style, sort/cross-filter flags, columns/item height, caption/start/lock settings and two-cell anchor offsets. Related-part validation rejects external/escaping/missing/malformed references, duplicate package members, invalid counts/indices, nonfinite numeric items and coordinates outside the Excel grid.

Focused Step 1 commit: `21b5fc7`, `feat: import pivot and slicer metadata`, normal hooks passed. No co-author trailer, binary workbook, real data or shared dependency symlink was staged.

Anchor correction: failing out-of-range regression repaired; all four frames now occupy distinct intervals within C-H, rows 17-30. Import suite 9/9 and staged gitleaks clean; only anchor hunks were amended into Step 1. Logs: `step1-anchor-red.log`, `step1-anchor-import-suite.log`, `step1-anchor-security.log`.

## Agreed refresh contract

Step 2 implemented: POST `/api/copies/{copy_id}/sheets/{sheet}/pivots/{name}/refresh` queries current data; body has page filters `{field,items}`, slicer filters `{name,field,items}`, and collapsed members `{field,item}`. Item selections accept cache indices and stable typed value selectors `{kind,value}` for new current-source axis values; returned items include stable selectors, dynamic items have `cache_index: null`. Null clears/all, repeated filters intersect and omitted entries retain imported defaults. Response has pivot identity/style/location/captions/revisions, normalized filters, display columns/rows, locked grid cell records with raw values/formats/row kinds, and slicer item selection/availability/no-data states. Pending computation returns 409 `calculation_pending` with both revisions. Invalid view selections are 422; inaccessible or absent copy/pivot is 404. Personal view persistence is Step 3, separate from copy input edits.

## Step 2 handoff

Current-source service and refresh API read authorized copy inputs plus exact-revision calculated results from one SQLite read snapshot. It rejects pending calculations, invalid typed selections and revoked access even on cached results. Filters intersect, dynamic axis values retain stable typed selectors, labels/formats and locked cell coordinates are returned in grid order, and outline/collapse/subtotal/row-column-grand-total behavior is covered. Startup prepares immutable source projections; bounded result/source caches and safe numeric measure deltas reduce repeat work. No engine/calc pipeline file changed.

Tests first: missing service, API contract, formula-driven source, dynamic axis, incremental refresh and column-grand-total failures were repaired. Focused command covers test_dashboard_import, test_template_importer, test_pivot_service, test_pivot_api, test_workbook_api, test_ops_api, test_security and test_http_transport: 64/64 twice (13.186 s and 9.545 s), logs `step2-focused-final-{1,2}.log`. Ruff clean twice, content guard exit 0, diff check clean and staged gitleaks clean. Logs are under ignored `results-private/dashboard/`.

100k synthetic records on this verified 8 GB Mac14,2: startup/source preparation 0.999605 s, first uncached result refresh 0.577600 s, new filter 0.355830 s, cache hit 0.005787 s, one calculated measure change 0.012804 s, peak process RSS 236.44 MB (`step2-benchmark-100k-final.log`). Repeat: first result 7.743297 s, filter 6.153175 s, warm 0.131671 s, preparation 3.749360 s (`step2-benchmark-100k-final-2.log`). High external CPU/load and about 6.5 GB swap were observed; no unrelated process was stopped. This variance leaves performance acceptance open. Do not present one green as a stable under-one-second guarantee. Final review/security/browser/all-suite gates remain pending.

## Step 3 live-browser correction

Tests first for personal views, output locking, dynamic overlap, strict nested selectors, multiple-selection report filters/accessibility and grid integration. Focused Python: 75/75 (`step3-python-focused-final-3.log`). Web: 112/112 twice (`step3-web-final-4a.log`, `step3-web-final-4b.log`). Ruff, pinned Prettier 3.6.2, content guard exit 0 and diff check pass. After the zoom metrics and CSS cascade correction, web suite 113/113 twice (`step3-web-final-6a.log`, `step3-web-final-6b.log`), pinned Prettier/build/guard/diff check pass. These are phase evidence, not final all-task acceptance.

Actual browser uncovered short sheet extents, an unregistered embedded scroll service, requested-versus-applied clamped scrolling and header offsets counted twice. Regression tests reproduce each boundary; the corrected coordinate adapter uses native header-inclusive cell rectangles and applied skeleton scrolling. Imported default row height is preserved (fixture 15 pt becomes 20 px), and sheet extents cover pivot/slicer bounds. Live initial Department B1 x110/y137, Region C17 x174/y457 height260; clamped scroll frame/native row17 y388; at 75% Region x131/y372 height195. Images: `step3-normalized-popup.png`, `step3-normalized-scroll.png`, `step3-normalized-zoom.png`. Zoomed inner chrome now scales with the anchored frame; regression red/green and live capture `step3-final-zoom-menu.png`: frame height195 and page button height15, matching 75% row metrics; open menu stays above sibling controls. Current synthetic server port 60128, existing Ego TaskSpace 43; neither is final acceptance yet. No engine/calc pipeline or fidelity implementation changed.

GET/PUT personal views are stable typed selectors, authorized and CSRF-protected, per user per copy, persisted separately from inputs/revisions/history. Imported output and report-filter cells stay locked even if workbook styles say unlocked; growing pivots reject unrelated input/nonblank-cell overlap as `pivot_overlap`. Four slicer frames render in-grid with filter intersection, selected/no-data states, multi-select/clear, columns and scrolling. Report filters use an in-cell dropdown with typed checklist selection and Cancel/OK/Clear.

## Step 4 refresh handoff

In-grid A4 Refresh has a busy state. Coherent calc/calc_state publications debounce refresh; calculation_pending retries cancel across sheet/dispose, and generations reject older results. Personal-view PUTs serialize on a persistent queue across navigation; reopen waits for queued writes, and latest save failure refreshes the last server-confirmed view so cells/slicers and remembered filters agree. Dynamic collapsed members now accept stable typed selectors alongside imported cache indices.

Observed lifecycle, dynamic-collapse, navigation/save-order and confirmed-result recovery reds then greens. Python 29/29 twice (`step4-python-focused-1.log`, `step4-python-focused-final.log`). Full web 120/120 twice (`step4-web-closeout-2a.log`, `step4-web-closeout-2b.log`). Ruff, pinned Prettier, build, guard exit 0 and diff check pass. Latest build source `0fc7bc428cab09350b14e8dac652fee2aeba77178984c4aeff0b12bb2f27a236`, bundle `f4a61ea265e02eb564d068619321e71ac7e17b1483cb53f02a5359221f0ebfdb`. Final current-code browser verification is still pending.

Current Step 3 service 100k benchmark now passes again: cold prepared result 0.625041 s, new filter 0.338543 s, one calculated-value change 0.010545 s, cache hit 0.005798 s, RSS 245.88 MB. Startup source preparation 1.392814 s is separate from refresh latency. Log/JSON `benchmark-current-step3.*`; generated workbook/snapshot were temporary and deleted. Retain earlier failing repeat; final same-code two benchmark greens remain required. Intermediate own server stopped, existing TaskSpace 43 retained at about:blank.

## Next task

Step 3 focused commit: `0670be1`, `feat: render in-grid pivot controls`, hooks/staged gitleaks clean. Step 4 explicit/debounced refresh is verified for its focused commit. Then refresh, category totals and team overview. Final 100k performance must be resolved along with the full acceptance gates. Original Step 2 target: current-source pivot service and coherent, authorized copy reads, filtering/layout and 100k-record cold/warm benchmark. Category totals will retain existing per-copy access unless the owner answers the pending clarification. No Microsoft Excel app or Calibri/Aptos font was found locally; local screenshots cannot independently establish Excel pixel parity.
