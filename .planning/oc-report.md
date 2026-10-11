# Runner report: native Formualizer load (2026-10-11)

Branch `feat/v2-integration`. Synthetic fixtures only. No real workbook or
private directory was read, no port bound, no push, no dependency change.

## STATUS

The first item is closed and the whole approved native design is implemented and
measured. The startup blocker is gone: template load no longer registers a single
formula, and native load reports 0 of 150 parity mismatches on the synthetic
fixture instead of 47.

Load time at 5% fell from 258.73 s to 10.03 s. 20%, which previously never
finished inside a 600 s timeout, now loads in 32.89 s. 50% loads in 92.12 s.

## CHANGED

da9720d expand shared formulas before native template load.
  Root cause is inside formualizer 0.11.1, not our code. Expanding a shared
  follower translates the master text with the follower offset, and that also
  shifts a cross-sheet qualifier whose sheet name parses as a cell address, so
  S01!$A$1 becomes T1!$A$1 and the follower reads a sheet that does not exist.
  Excel never translates the sheet part. Minimal reproduction committed under
  .planning/excel-grid/results/. The library is not patched: pp/engine/native_load.py
  rewrites the package once per immutable template so every shared follower
  carries its own explicit formula, which keeps native load a single read and
  registers nothing per cell.

787953d counts-only native parity audit for the controller.
  Lazy-evaluates a large random sample against the values Excel cached, runs
  edit-then-revert cycles on input cells, and reports mismatches by function
  name. Counts only: no values, addresses, formula text or defined names.

1764d01 visible-only evaluation with stale deferral.
  After an edit, affected cells inside the blocks clients hold are evaluated and
  published; the rest are marked stale, never answered from the Excel cache, and
  calculated when their block opens or the cell is read.

10db8af shared formulas split across a read boundary, plus scale measurements.
  Cells straddling a read boundary were copied through unexamined, so their
  followers leaked into the prepared package and the 20% load failed closed.

cf028e3 engine memory budget, idle LRU eviction and editor protection.
  engine_memory_budget_mb, default 3000 MB on the Mac and 16000 MB on the
  office server, replaces the unused engine_ram_budget_mb.

Also updated: STATE.md, PLAN-P6-real-template.md, LESSONS.md, and the aggregate
results JSON.

## NUMBERS

Parity on the synthetic fixture, seed 20261011, scale 0.0001:
  native load, before the fix   47 of 150 mismatches, all shared followers
  native load, after the fix     0 of 150 mismatches
  per-cell formula registrations at load   0

Scale measurements, synthetic, aggregate JSON
.planning/excel-grid/results/engine-scale-native-2026-10-11.json:

  scale  prepare  load     resident    first    warm     broad visible        broad full
  5%     1.91 s   10.03 s  71.8 MiB    0.0029 s 0.0002 s 4.13 s / 25,300     3.26 s / 75,000
  20%   10.10 s   32.89 s 152.0 MiB    0.0138 s 0.0002 s 18.48 s / 43,750    26.69 s / 300,000
  50%   15.09 s   92.12 s 263.7 MiB    not reached on this host

  Previous per-cell adapter load at 5% was 258.73 s batched, 455.15 s unbatched,
  and 20% timed out at 600 s.
  5% and 20% completed with every published value matching its expected value.
  At 20%, 256,250 of 300,000 cells were correctly deferred as stale by the
  visible-only edit and resolved with zero wrong values.
  Resident engine size scales about 1.03 MiB per 1000 formulas, putting the
  1.89M formula template near 2.0 GiB, inside both budgets.

  The 50% load completed. Its full broad edit was then killed by this 8 GiB Mac
  at 955 MB resident with swap already above 5 GiB. Recorded as failed. It is a
  host limit of this machine, not a load failure.

Verification:
  Python      431 of 431 twice, 48.05 s and 48.84 s
  Web         101 of 101 twice
  Content guard  exit 0
  Ruff        clean on all 14 changed Python files

## TESTS

New: tests/v2/test_native_shared_expansion.py (9), test_native_parity.py (6),
test_visible_evaluation.py (8), test_engine_memory_budget.py (21).

The shared-formula tests pin the library defect, prove the prepared template
carries explicit follower formulas, and force a 64 byte read boundary so the
straddling-cell bug cannot return. The parity audit has three tests that
deliberately break the engine or the stored cache, because an audit that cannot
go red is worse than no audit. The visible-block tests prove a stale downstream
cell is never answered from the Excel cache and that an untouched cell still is.

## NEXT

Controller runs the real-workbook audit; it must report zero mismatches:

    .venv/bin/python -m ops.native_parity --template /path/to/template.xlsm \
        --sample 20000 --output .planning/excel-grid/results/native-parity.json

Then, in priority order: make the dependency trace cheaper, since it now
dominates every broad edit and the visible-only path does not yet reduce it;
wire the connection registry to the pool protection API, which exists but has
no caller; decide defined-name expansion after counting its cells on the real
workbook; and re-run 5/20/50 on the office server.

Not done and not claimed: the service does not report open blocks to the engine,
so visible-only behaviour is engine-level and tested, not yet end to end through
the API. Editor protection is an available API with no caller. Defined-name
expansion still registers formulas after load. The 50% trace cost is unmeasured
off-host.

The exact model review and the planner review were not run: this runner is
instructed not to call any external model CLI, so those gates remain open for
the supervisor.