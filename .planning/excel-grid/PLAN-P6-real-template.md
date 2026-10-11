# P6: Cached startup and synthetic template scale

Date: 2026-10-11. Approved by the owner's GO-TOKEN in this session.
Branch: `feat/v2-integration`. Synthetic fixtures only.

## Goal and approved scope

Accept single-cell OOXML array formulas without changing their export XML, retain the
template-map ISFORMULA shim, avoid full calculation on template/copy load, and measure
the pinned engine against the controller's aggregate workbook shape.

## Constraints

Keep `formualizer==0.11.1` and all dependency locks/wheels unchanged. Never inspect a
real workbook or `~/PricePlanning-private`. Never use ports 8000, 8765, or 8766.
Generated workbook artifacts and raw logs stay in `results-private/`; only aggregate
metrics enter `.planning/excel-grid/results/`. No push or deployment in this round.

## Execution phases

- [x] Array compatibility: tests first for single-cell array load, import, locked
  input mask, edit recalculation, and export preservation. Keep multi-cell arrays
  and dataTable rejected. Patch `pp/engine/formualizer.py` and matching synthetic
  tests; patch importer/export only if their behavior requires it. Success: new
  tests fail before the patch and pass after it, with existing extent tests green.
- [ ] Cached startup: read pinned installed APIs and probe only synthetic XML.
  Seed adapter values from typed XML caches, preserve ISFORMULA reference-map
  semantics, and target edited dependents rather than `evaluate_all`. Check spills,
  missing caches, shared followers, typed values, and warm copy clears. A missing
  native clean-cache API is a documented limitation, not an assumed capability.
  Files: engine adapter, importer/worker where required, regression tests.
  Success: cached template/copy open performs no full evaluation and edits produce
  correct deltas without changing unrelated cached results.
- [ ] Scale evidence: deterministic seeded streaming generator and bounded runner
  under `ops/`, with generator/runner regression tests. Match 22 sheets, 9 hidden,
  approximate 5.2M cells/1.5M formulas, array/shared/normal proportions, and the
  largest-sheet concentrations. Write cached values and exercise ordinary and
  broad edits at 5%, 20%, 50%. Monitor `sysctl vm.swapusage`; terminate the owned
  benchmark and stop increasing scale if growth exceeds 1536 MiB from the resume baseline or free disk drops below 2500 MiB. Success: recorded
  load seconds, peak RSS, edit latency, correctness counts, and explicit stop or
  completion evidence in `results/engine-scale-2026-10-11.json`.
- [ ] Verification and handoff: Python suite twice, web suite twice, content guard
  0, ruff, secret/SAST/dependency scans, exact `claude-opus-5-5` independent review
  and planner review. Resolve defects, update STATE headings and this results plan,
  and make focused conventional commits. Success: all required checks accounted
  for, review model identity verified, host/RAM recommendation distinguishes
  measured evidence from extrapolation.

## Review focus

Malformed or foreign array extents must fail closed. Formula caches preserve bool,
number, empty text, error, and date-system distinctions. Shared follower formulas
are expanded before ISFORMULA rewriting. Spill followers cannot become editable
inputs. Dependency inspection must not silently truncate large fan-out or skip
dynamic references. Clearing and restoring warm overrides must invalidate results.
Root review found a spill-growth case requiring a regression: cached `C1=B2` remains
0 after `SEQUENCE(A1)` grows from one row to two, although B2 becomes 2. Re-trace
changed old/new spill member coordinates and recalculate their downstream formulas;
do not publish a cached downstream result from the pre-growth closure.

## API investigation

Installed `formualizer/formualizer_py.pyi` exposes `evaluate_cell`, `evaluate_cells`,
`get_eval_plan`, `dependents`, and inspection staleness. Its load `strategy` argument
is ignored; import is eager. A synthetic cached-XML probe shows formula `<v>` values
are ignored by the native loader and start `NeverEvaluated`. No public clean-cache
setter is exposed. Targeted evaluation leaves unrelated formulas untouched, but a
cold affected closure may evaluate native precedents absent from its own cache.

A tiny native fixture exposed an additional 0.11.1 graph limitation: all 150 formulas
were present, but dependent tracing reached only 103, omitting 47 shared followers.
Disabling span evaluation did not fix the gap; explicitly registering expanded
formulas restored 150 targets. The generator also needed its CONCAT cache corrected.

Ruling: after Luna and Terra acceptance failures, route the remaining engine,
sparse-publication, and fixture repairs to `gpt-5.6-sol` under the two-failure router
takeover rule. Full worker-suite socket errors are sandbox restrictions; repeat
the decisive full suite in the primary authorized environment without weakening tests.

## Results and recommendation

Execution is in progress. Current resume baseline swap used: 8875.31 MiB; initial free disk: 5054 MiB. Historical prior-run observations follow.
Array compatibility: commit `2ec0493`, six new synthetic tests, 53 targeted tests
green after the observed rejection failures; staged gitleaks clean. Review remains
pending until the complete round diff exists. No scale or release gate is claimed.
Pre-benchmark resource peak observed: 9247.38 MiB swap used, 2275.76 MiB above the
session baseline. Large benchmarks were held while that limit was exceeded. The
runner must retain a common baseline across scales and a minimum free-disk guard.

## Independent review checkpoint

Exact `claude-opus-5-5` identity was verified from returned `modelUsage` for both
review requests. Scale review returned FIX-FIRST for detached-child interruption
safety and partial-metric/prefix/cache-validation gaps. Those repairs now pass
30 focused tests twice and the prefixed native tiny fixture (150 cached/native,
1 ordinary delta, 150 broad deltas); independent re-review is pending.

Engine review returned FIX-FIRST for trace-budget, same-batch spill and clear/restore
cases, plus cumulative-overlay publication cost after broad edits. These are being
reproduced against the pinned native engine before resolution. No release claim.
The earlier spill repair passed 39 focused tests twice. Its restricted full run
reached 330 tests with three socket EPERM errors and one transient streamed-export
ENOSPC error. After capacity recovered, the primary exact streamed-export test
passed. The final primary full runs still remain.

Host pressure has varied substantially during this session before any large benchmark:
free disk dipped to 527 MiB, then recovered to 15 GiB; observed swap used reached
9875.81 MiB against the original 6971.62 MiB baseline. The common baseline and
resource guards remain mandatory. No scale timings are currently accepted.

## 19:45 IST resume checkpoint

Focused tests: 112/112 PASS. Primary Python run 1: 368/368 PASS; run 2:
370/370 PASS while runner tests were being added, so two final unchanged-suite
runs remain. Web: 101/101 PASS twice. Scoped ruff clean; broader old ops/tests
scan reports 141 existing lint findings (raw log results-private/p6-resume-ruff.log).
Content guard exit 0. Staged gitleaks clean on all three commits. pip-audit on
requirements.lock found no known vulnerabilities. Semgrep 151 Python rules, 0
findings; its single server.py timeout was rerun with 60 seconds and no timeout.

Commits: fixture 9dce2ba; cached startup, dirty closure and sparse publication
0e2e386; bounded benchmark runner 63e4391. Exact claude-opus-5-5 review is running.
Planner gpt-5.6-sol review is running. The scale runner is measuring 5%, then 20%,
then 50% in separate owned child processes with a common 8875.31 MiB baseline.
It deletes generated artifacts after each scale while retaining small progress/log
files, stops at >1536 MiB swap growth or <2500 MiB free disk, and records skipped
scales. No actual large-scale timing is claimed until the aggregate report exists.
