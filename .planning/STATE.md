# STATE

## 2026-10-12 13:45 IST (Claude)

feat/v2-integration at 595808a, pushed (Sir's y; gate clean; no office files on the remote).
- Dashboard track merged (ed352f2): cell-based pivot service, in-grid slicers, report filters,
  personal views, category totals, team overview (copies page "Live now / team"). 699 Python,
  152 web OK.
- Real template: its pivot reads the external Power Query connection, which that service cannot
  rebuild; the import failed and took 8765 down. fff35f6 degrades unsupported pivots to Excel's
  saved values (security/malformed parts still fail). 8765 back up, Coupon opens.
- Running: Excel-only pivot + slicers (PLAN-P6.4 revision) on feat/v2-dash-engine.

## 2026-10-12 12:20 IST (Claude, P6.5 refresh data)

Merged into feat/v2-integration at 7930bbe (648 Python OK, 121 web OK, guard 0). "Refresh data" runs Excel's own
RefreshAll on the copy's private Excel, waits for the queries, and commits what changed as one `refresh`
revision through the macro diff. `refreshable` (IMPORT_VERSION 10) comes from xl/connections.xml.
Route POST /api/copies/{id}/refresh; web toolbar button shown only when refreshable.
- Tests: new tests/v2/test_refresh_data.py (17) plus COM refresh tests in test_excel_macros.py; web 121 OK;
  ruff clean; content guard 0. Bundle rebuilt (static/dist).
- Not run here: a real RefreshAll (no Excel on the Mac). Office check: `ops\excel_smoke.py --refresh`, must print PASS.
- Plan: .planning/excel-grid/PLAN-P6.5-refresh.md. Old feat/v2-refresh (pqtools) stays superseded.

## 2026-10-12 11:45 IST (Claude)

Branch feat/v2-integration at c530436 (unpushed; Sir runs `git push`, answers y).
- P6.3 macros merged (M1 buttons, M2 engine, M3 flow): the template's 5 buttons are read from its
  drawings and drawn on their sheets; pressing one runs the real VBA in the copy's private Excel,
  presses OK on its message boxes, saves what it changed as one `macro` revision (values, plain
  cells and new formulas) and publishes results. Mac and values-only mode answer 409.
- Speed fix 1ade04f: rows on big protected sheets took 3-9 s per window (a 16,384-column scan;
  BU RECAP sent 1.9M empty editable cells, the browser grew to 4 GB). Now 5-130 ms warm, about
  1.3 s cold on BU RECAP; sheet switch 1.7-3.4 s in the browser, heap under 100 MB. Editable empty
  cells stop at the sheet's used width (documented shortcut).
- c530436: zero-size leftover VML button skipped (5 buttons, not 6). IMPORT_VERSION 9: servers
  reimport the template once at start.
- Tests: 629 Python OK, 118 web OK, ruff clean (pp, tests/v2), content guard 0.
Office server next (Sir): `"engine": "excel"` in settings.json, then
`.venv\Scripts\python ops\excel_smoke.py` and `... excel_smoke.py --macros`; both must print PASS.
Next here: P6.5 Power Query refresh on the Excel engine (RefreshAll + the macro diff path); the old
feat/v2-refresh worktree (pqtools, based on bcba313) is superseded and left untouched.

## 2026-10-12 (Claude, M3 macros flow)

Branch feat/v2-macros-flow: template buttons run end to end on the Excel engine (P6.3 M3). Press ->
pipeline lock -> Excel run on a worker -> one `macro` revision (inputs, plain values, formula overrides)
-> calc publish. Engine failures are 502 `macro_failed`; Mac/formualizer and cached templates stay 409.
Rows and export show non-input overrides and write formulas as formulas. `ops/excel_smoke.py --macros`
is the office check (prints counts only). Not run here: the office `--macros` smoke (no Excel on the Mac).
Next: Sir runs `excel_smoke.py --macros` on the office server; then the browser check of each button.

## 2026-10-12 10:00 IST (Claude)

Branch feat/v2-integration at b6e10e8 (37 commits not yet pushed; Sir runs
`git push` and answers y at the gate).
- Excel engine built (288e01d, eda0e81): settings "engine": "excel" makes each
  copy's worker drive a private hidden Excel (macros and events off, manual
  calculation, own temp copy of the file). A static dependency index
  (pp/engine/depindex.py, numpy, cached beside the prepared package) tells it
  which formula cells an edit can reach; it reads only those back in row
  blocks and publishes what changed. Import mode "saved": no import engine
  pass, edits allowed. Real file on the Mac: index build 103 s once, cone
  1-9 s per edit (largest cones about 540k formulas).
- Transfer track merged (b6e10e8): download dialog (pick sheets, .xlsx or
  .xlsm) and upload into the working copy or a new copy with a preview.
- Tests: 588 Python OK, 109 web OK, ruff clean.
- LAN 8765 still values-only on the Mac (no Excel here).
Office server next: set "engine": "excel" in settings.json, then
`.venv\Scripts\python ops\excel_smoke.py` must print PASS (it proves no
changed cell is missed and that reverting leaves no overlay). Then: cone speed,
macros as server actions (P6.3), Power Query refresh (P6.5).
MiMo v2.6 Pro: not usable (OpenCode Go timed out, ClinePass 403); Space Bunny
is the only free lane.

## 2026-10-12 09:15 IST (Claude)

Engine on the real file: load now 1.0-1.2 GB (was >25 GB) after trusting
Excel's own-cell array results (b08bb15; 664,077 t="array" cells, none spill).
Trace budgets capped at u32 (TRACE_U32_MAX). Formula map cached per prepared
template (04e9426). First native dependency trace on an edit still passed
3.5 GB: next is replacing _dependent_targets with a precomputed reverse index
(scratchpad reach.py logic), evaluating visible targets only (formualizer
evaluate_cells recomputes stale chains on demand, verified).
UI: sheet arrows, template card on home, values-only copies report
calculated (no "Calculating..."), recovery skipped in values-only mode.
Export verified on the real copy: 22 sheets, VBA kept, 7.9 s.
Not started: import (upload into copy or new copy), macros, data refresh.

09:30 update: visible-window evaluation on an edit also passed 3.5 GB, so
formualizer cannot run realistic edits on this file within the Mac or office
memory. Sir confirmed the office Windows server has Microsoft Excel. Decision:
real Excel via COM as the engine (exact results, VBA macros, Power Query).
All 12 branches pushed to the private repo (gitleaks clean, no office files).
Pre-push hook fixed to find the Max HQ gate from worktrees (fb50ed7, unpushed).

## 2026-10-12 08:10 IST (Claude)

Sir asked again to see the whole template with every tab. Values-only mode
(da0a435, d52dee1, d3c1df1): settings "template_results": "cached" imports
without the engine pass (Excel's saved results), records "results" in
meta.json, opens copies read only (status bar says so), refuses edits with
409 calculation_paused, and gives the app no engine pool so copy creation
never starts a worker. Real file: import 51 s, 134 MB peak; 22 sheets (13
visible, 9 hidden as in Excel); copy opens in 4.3 s; each tab switches in
about 30 ms; browser heap 45 MB. Tests: 446 Python OK, 101 web OK.

LAN: 0.0.0.0:8765 now runs ~/PricePlanning-private/lan-real (real template,
values only, brand). lan-demo is stopped. New database: users sign up again.
Next: cell-level reachability result -> pruned engine package -> turn
template_results back to "engine".

## 2026-10-12 06:20 IST (Claude, free lanes down)

Branch feat/v2-integration at the merge of feat/v2-brand. New since 4b9bc6b:
13a2640 structured table references resolved to absolute ranges before native
load (#All/#Data/#Headers/#Totals, column spans, bare names; [@Col] raises);
da0aa53 + style follow-up: optional "brand" settings block, /api/brand and
/brand/<asset> public, branded app bar and sign-in. Tests: 438 Python OK, 101
web OK, ruff clean, content_guard 0.

Real template (counts only): import and shared/structured expansion now pass.
The engine is the blocker: 1,894,479 formula cells; the Python template map
takes 357 s and 1.1 GB; native load 2.8 s; the FIRST evaluation of any single
cell builds the whole dependency graph and passed 2.5 GB in under 60 s (an
unguarded run reached a 25 GB footprint). EvaluationConfig knobs
(enable_parallel off, max_threads 1, range_expansion_limit 1, flat cache 0)
do not change it. Sheet-level reachability from editable cells keeps 93% of
formulas, so sheet pruning is not enough. Next: cell-level reachability from
input cells (formulas no input can reach keep Excel's cached value and are
dropped from the engine package); measure the live fraction first.

LAN app: 0.0.0.0:8765 runs from ~/PricePlanning-private/lan-demo on a
synthetic demo workbook with the private brand; real template not served yet.


## Current phase

Branch: feat/v2-integration. Native Formualizer load is implemented and measured.
The startup blocker is closed: no formula is registered per cell at load, and
native load reports zero parity mismatches on the synthetic fixture.

Owner decision honoured: native formualizer==0.11.1, one resident engine per
active copy, loaded once from the template, input overlays only, visible-only
affected evaluation, stale blocks calculated on request, atomic per-revision
publication, engine_memory_budget_mb with idle LRU and editor protection.

Root cause of the 47 of 150 mismatch is inside the pinned engine, not our code.
Formualizer 0.11.1 expands a shared-formula follower by translating the master
text with the follower offset, and that translation also shifts a cross-sheet
qualifier whose sheet name parses as a cell address, rewriting S01!$A$1 as
T1!$A$1. Excel never translates the sheet part. Every follower of such a family
read a sheet that does not exist and evaluated to #REF!. Minimal reproduction:
.planning/excel-grid/results/formualizer-shared-formula-repro.py. The library is
not patched. The package is rewritten once per immutable template so every
shared follower carries its own explicit formula, which keeps native load a
single read and needs no per-cell registration.

Scale results, aggregate JSON in
.planning/excel-grid/results/engine-scale-native-2026-10-11.json, synthetic only:

  scale  prepare  load    resident after load  first  warm   broad visible  broad full
  5%     1.91 s   10.03 s 71.8 MiB             0.003 s 0.0002 s 4.13 s       3.26 s
  20%   10.10 s   32.89 s 152.0 MiB            0.014 s 0.0002 s 18.48 s     26.69 s
  50%   15.09 s   92.12 s 263.7 MiB            not reached on this host

The previous per-cell adapter load was 258.73 s at 5% and never finished 20%
inside a 600 s timeout. Load is no longer the blocker. Resident engine size
scales about 1.03 MiB per 1000 formulas, which puts the 1.89M formula template
near 2.0 GiB and inside both the 3000 MB Mac and 16000 MB office budgets.

At 50% the load completed in 92.12 s at 263.7 MiB resident, then the child was
killed by the host during the full-workbook dependency trace at 955 MB on an
8 GiB Mac with swap already above 5 GiB. That is a host limit of this machine,
not a load failure, and it is recorded as failed rather than reported as passed.

Known limits, all measured rather than assumed: visible-only evaluation reduces
evaluation and publication but not the dependency trace, so at 5% it was no
faster than the full edit; the trace is now the dominant edit cost. Editor
protection exists as a pool API but nothing calls protect() yet, because the
connection registry has not been wired. Defined-name expansion still registers
formulas after load; the synthetic fixtures exercise none.

## Last handoff

2026-10-11 native Option A session. Uncommitted work from the interrupted run
did not exist; the tree was clean at a25989c and nothing was inherited.

Shipped and committed, each item verified before commit:
  da9720d expand shared formulas before native template load, 0 of 150 mismatches
  787953d counts-only native parity audit for the controller real-workbook runs
  1764d01 visible-only evaluation with stale deferral and on-demand resolution
  10db8af shared formulas split across a read boundary, plus scale measurements
  cf028e3 engine memory budget, idle LRU eviction and editor protection

Two real defects were found by testing rather than review. Cells straddling a
read boundary were copied through unexamined, so their shared followers leaked
into the prepared package and the 20% load failed closed; a test now forces the
boundary with 64 byte reads. recalculate_stale resolved each cell on demand
during the publish capture, which zeroed the delta; the marks are now cleared
before evaluation and restored on failure.

Verification for this session: Python 431 of 431 twice (48.05 s and 48.84 s),
web 101 of 101 twice, content guard exit 0, Ruff clean on every changed file.
Synthetic fixtures only. No real workbook or private directory read, no port
bound, no push, no dependency change, no external model review. The supervisor
runs reviews after this runner.

Aggregate evidence only is committed. Raw logs remain under results-private.
Free disk 14174 MiB at the last check, far above the 2500 MiB floor.

## Next task

Controller owns the real-workbook evidence. Run the counts-only audit against the
real template and report the aggregates:

    .venv/bin/python -m ops.native_parity --template /path/to/template.xlsm \
        --sample 20000 --output .planning/excel-grid/results/native-parity.json

Pass requires zero mismatches across the lazy sample and zero after revert. The
audit prints counts only. If it fails, mismatches_by_function names the family to
investigate first. The same audit runs against a synthetic fixture as a control.

Then, in priority order:

1. Make the dependency trace cheaper. It now dominates every broad edit and the
   visible-only path does not yet reduce it. This is the largest remaining
   latency lever and it is what limits 50% and above.
2. Wire the connection registry to pool.protect and release_protection, so
   editor presence is real rather than only an available API.
3. Decide defined-name expansion. If the real template uses names inside
   formulas, those cells are still registered per cell after load. Measure the
   count on the real workbook before choosing whether to extend the pre-load
   rewrite to names.
4. Re-run the 5/20/50 scale set on the office server. The 50% load succeeded
   here; its full broad edit was killed by this 8 GiB Mac, so the trace cost at
   scale is still unmeasured off-host.

Unverified and owner-controlled: office Windows and Excel acceptance, LAN access,
any push or PR, and any dependency change. engine formualizer stays 0.11.1.


## Dashboard track (merged from feat/v2-dashboard)

v2 dashboard, approved 2026-10-11. Branch `feat/v2-dashboard` from `f74d131`. Steps 1-5 are implemented and committed: import `21b5fc7`, current-source pivot `1798ce9`, in-grid controls `0670be1`, refresh `a131b5b`, category totals `395ca30`. Step 6 team overview is in progress. Final acceptance remains open. Integration owns the engine/calc pipeline; fidelity owns conditional formatting, pictures/shapes, notes, tables and autofilter. Synthetic fixtures only; no private workbook access or Docker.

## Last handoff

2026-10-11 resumed from timeout. Current user resource rule supersedes the earlier checkpoint: stop at `df -m /` Available below 2500 MB, or swap more than 1536 MB above this resume baseline 5278.62 MB (limit 6814.62 MB). Initial free disk 6071 MB. Focused Python exposed three category filter boundary failures; fixed null slicer lists and non-string Copy selectors returning 500. Python 30/30 twice (9.819 s, 25.474 s), web 125/125 twice, ruff clean, pinned Prettier 3.6.2 clean, content guard exit 0, staged gitleaks clean. Category totals committed as `395ca30`. Exact `claude-opus-5-5` independent review is running via read-only CLI; identity must be checked from JSON output before counting it. Shared `.venv` and `web/node_modules` symlinks are untouched and untracked.

Existing Ego TaskSpace 43/p1 was resumed successfully and remains at about:blank; no current server has been started in this resume. Finish it with keep: [] after final browser check. Prior screenshots and checks are historical and do not verify final code. Full logs are in ignored `results-private/dashboard/`; track report contains prior phase evidence.

Routing: prior Luna and Terra execution attempts both failed, documented in the track. One exact `gpt-5.6-sol` xhigh takeover worker owns Step 6, while root handles reviews and acceptance. No push, PR, merge, deployment or Office acceptance authorized by this track plan. Excel pixel comparison and Calibri/Aptos fonts remain unavailable; do not claim measured pixel parity.

## Next task

Finish and commit Step 6 team overview. Read the pending Opus review, verify exact returned model identity, resolve concrete findings test-first and re-review critical/high repairs. Complete independent exact planner review, scoped security/dependency checks, 100k benchmark twice, full Python twice and web twice on final code, content guard 0, ruff, pinned Prettier and two identical builds. Then run a synthetic real server on a free port excluding 8000/8765/8766; capture pivot, page filters, slicer selection/multi-select/clear, category totals and team overview in existing TaskSpace 43; finish keep: [] and stop server immediately. Update this state and tracks/dashboard.md and commit handoff.
