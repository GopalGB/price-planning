# STATE

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

