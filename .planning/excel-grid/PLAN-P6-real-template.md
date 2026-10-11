# P6: Cached startup and synthetic template scale

Date: 2026-10-11. Approved by the owner's GO-TOKEN.
Branch: `feat/v2-integration`. Synthetic fixtures only. No push or deployment.

## Approved native-load execution (2026-10-11)

Owner GO-TOKEN supersedes the template-shared worker attempt below. One native
Formualizer engine per active copy, native load once, input overlays only.
Excel snapshot caches serve untouched cells. Edited downstream cells cannot fall
back to those caches. Evaluate affected visible blocks and defer other affected
cells until requested. Preserve atomic input-revision publication.

Fresh session resource baseline: swap 6061.94 MiB; stop above 7597.94 MiB or
free disk below 2500 MiB. Guard: results-private/native-resource-live.json.
No private workbook reads, dependency changes, forbidden ports, push or deployment.

- [x] Review and drop conflicting inherited shared-engine/reload changes; retain
  compatible tests. Commit native loader and independent resident copy behavior
  after targeted synthetic load/edit/revert correctness checks. Dropped; da9720d.
- [x] Connect bounded transitive dependents, visible row blocks, deferred stale
  requests and atomic revision publication. Prove invisible downstream cells never
  return old caches, including concurrent revisions and reopened copies. Engine
  side complete, 1764d01. The service does not yet report open blocks.
- [x] Add engine_memory_budget_mb (3000 Mac, 16000 office), idle LRU, connected
  editor protection and memory telemetry with synthetic capacity tests.
  cf028e3. Protection is an available pool API with no caller yet.
- [x] Build counts-only native parity script for controller use, prove sampling,
  edit/revert and function-count reporting on synthetic fixtures. 787953d, with
  three tests that break the engine or the cache to prove the audit goes red.
- [x] Measure 5%/20%/50% sequential synthetic native load, first/warm/broad visible
  and full edits, peak RSS; commit aggregate JSON after each finished item. 10db8af.
  5% and 20% completed with every value matching; the 50% load completed at
  92.12 s and its full broad edit was killed by this 8 GiB Mac at 955 MB.
- [x] Verify Python twice, web twice, content guard 0, scoped Ruff. 431 of 431
  twice, 101 of 101 twice, guard exit 0, Ruff clean. The exact model review and
  the planner review belong to the supervisor and were not run by this runner.

Ruling: drop template-keyed single mutable engine and spill export/reload fallback
because both conflict with the owner's latest explicit per-copy native decision.
Original interrupted diff preserved in results-private/native-redirect-start.diff.

Checkpoint: inherited conflicting code dropped; native implementation not started.
Guard fired at swap 7747.25 MiB (growth 1685.31 MiB). Session stopped without
resetting baseline. Native synthetic preflight produced 47/150 mismatches, all
shared followers; alternate umya loader raised OSError. Raw logs are under
results-private/native-*. Aggregate: results/native-load-preflight-2026-10-11.json.
Planner gpt-5.6-sol/xhigh started but did not complete before stop; no review verdict.
Final Python/web twice, scoped Ruff, content guard, security and Claude gates pending.

## Historical resident-engine owner decision and stopped attempt

Owner decision 2026-10-11: load once and keep it loaded. The approved follow-up is
tracked in WARM-ENGINE-HANDOFF.md and WARM-CORE-BRIEF.md. It is incomplete and
BLOCKED_RESOURCE_SWAP. No new scale benchmark was run, so measured results below
remain historical and do not establish the resident implementation's performance.

Live resource guard fired at 20:58:19 IST: baseline swap 4589.25 MiB, observed
6277.06 MiB, growth 1687.81 MiB above the 1536 MiB allowance. Free disk 4747 MiB,
above the 2500 MiB floor. The parent stopped the implementation worker. No engine
scale load was running; the host-wide growth is not attributed to Formualizer.
Aggregate: results/warm-resource-live.json. Private stop log: warm-resource-stop.json.

API evidence: pinned 0.11.1 exposes set_value(None), set_values_batch,
set_formulas_batch, evaluate_cells, changelog actions and undo/redo. It has no public
native clone/copy, deepcopy/pickle or native engine snapshot/restore. to_xlsx_bytes
and from_bytes provide XLSX export/reparse, not resident-state restoration. No native
load-progress callback exists; honest adapter milestone progress is the proposed UI
path. Native scalar clearing works, but blocked-spill clearing and follower values
remain unsafe in isolated probes. A production fallback reload violates the owner
decision and is not an accepted solution.

Retained uncommitted patch passes 18 focused pool/warm tests and shares a child by
template path. It still has unresolved spill reload and publication/error-state
issues. Background preparation, UI progress, memory budget/LRU/editor protection,
N=1/3/5 resident benchmarks and final full verification/reviews remain unfinished.
No implementation item is accepted as complete or committed from this attempt.

Decision: keep the supported candidate (one resident base plus serialized copy
overlays and logical-state LRU), repair correctness before further measurement, and
resume on a quiet host or office server with adequate headroom. A 100% fitted curve
is unavailable because only the prior 5% scale completed. The prior single-point
linear estimate of 13.2 GiB peak process RSS is unverified: it exceeds 3 GB and does
not prove fit within 16 GB for the whole server. Do not build partitioning or
reachable-only evaluation in this round. If completed measurements exceed a one-base
budget, the next smallest experiment is reachable-formula evaluation or partitioning.

## Goal and approved scope

Accept single-cell OOXML array formulas without changing export XML, answer
ISFORMULA from the template formula map, avoid full calculation on template/copy
load, and measure the pinned engine against the supplied aggregate workbook shape.
Full-scale viability is not established: the 20% synthetic load timed out.

## Constraints

Keep `formualizer==0.11.1`, requirements.lock and vendor wheels unchanged. Never
inspect a real workbook or the private workbook directory. Never bind ports 8000,
8765 or 8766. Raw logs and fixtures stay in results-private; only aggregates enter
results. Resume swap baseline: 8875.31 MiB. Stop at swap growth above 1536 MiB or
free disk below 2500 MiB. Run one scale at a time, timeout each attempt at 600 seconds,
reap only owned process groups, and delete generated fixtures after every attempt.

## Execution phases

- [x] Array compatibility: accept single-cell t="array" in load/import/input-mask,
  preserve t="array" on export, reject multi-cell arrays and dataTable with clear
  errors. Commit 2ec0493 plus synthetic round-trip and locked-input regressions.
- [x] Cached startup and dirty calculation: typed XML cache baseline, targeted native
  bootstrap for missing caches/spills, bounded dependency tracing, shared follower
  registration, template-map ISFORMULA, sparse revision-checked worker publication.
  Coverage includes typed caches, chains, names/LET, row/column references, spill
  growth/shrink, clears/restores, stale worker publication and snapshot versioning.
- [x] Scale tools and evidence: deterministic seeded streaming generator under ops,
  cached formulas, 22 sheets/9 hidden, matching counts/concentrations/encodings.
  5% completed; 20% timed out; 50% explicitly skipped. Evidence recorded below.
- [x] Verification and handoff: final Python twice, web twice, guard, scoped Ruff, security
  scans and exact claude-opus-5-5 repair closure. STATE updated; evidence committed.

## API investigation and chosen adapter

Installed formualizer/formualizer_py.pyi exposes evaluate_cells, trace,
get_eval_plan and set_formulas_batch. load_workbook documents strategy as ignored
and uses eager Calamine import (lines 2096-2105). Native synthetic probes show XML
formula caches start NeverEvaluated. There is no public setter that installs cached
formula values while marking them clean, nor a working lazy loader flag in 0.11.1.
Dependencies were not upgraded.

The adapter retains typed XML caches outside the native engine. Template and copy
open do not call evaluate_all. Only missing caches/potential spills bootstrap via
targeted evaluation. Edits invalidate a bounded transitive closure; dynamic roots
are conservative, shared followers are explicitly registered, and downstream spill
members are retraced. Native cold targets may still calculate their precedents.
The native eager import and formula registration remain startup work.

A 150-formula native probe initially reached only 103 dependents because shared
followers were omitted from native tracing. Expanded registration fixes all 150.
Contiguous same-row registration now uses set_formulas_batch. Static defined names
are expanded only in native formulas, preserving original export formulas; LET
shadowing is protected. Complex names/dynamic references recalculate on every edit.
Whole-row, whole-column and quoted-sheet references have post-cache edit proofs.
Bare short names such as Tax and out-of-grid cell-like names are distinguished
from actual Excel addresses; regression tests prove downstream and LET behavior.
Structured/external references and non-static named-range colon operators fail
closed because native evaluation/dirty tracking was reproduced as unsafe.

IMPORT_VERSION is 7. Active templates reimport under the existing startup path;
older different-hash snapshots require explicit reimport and return a named 409
error. Existing overrides/copy data are preserved. Volatile XML caches are retained
on open under the requested cached-startup contract and refreshed on an edit.
The supplied inventory does not include TODAY/NOW/RAND.

## Measured scale results

Source benchmark commit: c1b06fc; seed 20261011. See
results/engine-scale-2026-10-11.json. Timings measure native adapter load, not full
application import, database initialization or office-server startup.

| Scale | Cells | Formulas | Load | First edit | Broad edit | Peak child RSS |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| 5% | 260000 | 75000 | 258.73 s | 0.423 s | 4.751 s | 675.6 MiB |
| 20% | 1040000 | 300000 | incomplete at 600 s attempt timeout | unmeasured | unmeasured | sampled at least 202.2 MiB |
| 50% | not generated | not generated | skipped after 20% timeout | unmeasured | unmeasured | unmeasured |

At 5%, ordinary edit returned exactly 1 correct delta; broad edit returned all
75000 correct deltas. There were 33200 single-cell arrays, 34000 shared and 7800
normal formulas. ru_maxrss is the completed child peak; sampled RSS is a lower bound
for failed attempts. 20% attempt duration was 600.511 s including 4.775 s generation.
No completed load/edit timing is inferred for that attempt.

The prior unbatched run also timed out at 20%. Its 5% load was 455.15 s, first edit
0.558 s, broad edit 5.717 s, peak child RSS 577.9 MiB. Live host pressure varied,
so the faster batched run is not an isolated comparison. Latest run minimum free
disk was 6922.94 MiB; swap never exceeded the retained 8875.31 MiB baseline. Neither
resource stop fired. Both attempts' generated fixtures were deleted; only small
progress/log files remain. Prior metrics are retained in the aggregate report.

The fixture mirrors the supplied cell/formula counts, largest-sheet concentrations
and OOXML encodings. It uses a bounded common-function palette, three-row lookup
ranges and one broad hub. It does not reproduce the real dependency graph or all
58 inventoried functions. Small 150-formula probes separately compare native values
with synthetic caches, including relative shared followers.

## Recommendation and remaining blocker

P6_LARGE_SCALE_STARTUP: cached calculation fixes correctness and edit cost, but
0.11.1 eager load is still too slow to establish full-size viability on this host.
Do not ship this as a proven real-template engine path. The next bounded experiment
should profile the 20% startup stages on the office server with the same synthetic
seed; isolate native import, expanded registration and trace construction before
choosing a smaller adapter repair or requesting a dependency change.

Use the 8 GiB Mac for small synthetic development. Start office testing with at
least 32 GiB RAM and one engine worker; 64 GiB is a planning allowance if multiple
copies must stay warm. These are capacity estimates, not measured office acceptance.
Linear extrapolation of the completed 5% child peak alone is about 13.2 GiB per
full-scale engine; import, database, server and OS need headroom. Nonlinear memory
and startup costs remain unmeasured. RAM alone does not fix a CPU-bound loader.
The configured default is one worker (pp/config.py:230; docs/DEPLOY.md:68).

Real-workbook parity, Windows/Excel/LAN behavior, full-size startup, and deployment
remain unverified. P1/P5 release gates are not advanced by these synthetic results.

## Verification and independent review

Exact claude-opus-5-5 review identified runner limits, cache versioning, dependency
name and short-name issues. All repairs were reviewed; final narrow closure is SHIP
with returned modelUsage exactly claude-opus-5-5 (p6-tax-review.json). This was
code-only independent review; native tests were executed by the primary session
and planner worker. Planner gpt-5.6-sol verdict: NO BLOCKING DEFECTS. Focused final
engine/cache/runner modules pass 83/83 twice. Final code commit afc320a. The ISFORMULA validation-order repair also received
SHIP with exact returned claude-opus-5-5 identity (p6-isformula-review.json).

Final unchanged Python runs: 387/387 PASS twice (85.657 s and 71.249 s),
using python -m unittest discover -s tests -v.
The earlier 386/387 run exposed ISFORMULA(TargetCell) validation after name expansion;
commit afc320a moves original-syntax validation before that expansion. Web passes 101/101 twice, scoped Ruff
passes, content guard exits 0, staged gitleaks is clean, Semgrep runs 151 rules on
41 Python files with 0 findings/no timeout; the final adapter patch was rescanned
with 151 rules/1 file and 0 findings, and pip-audit requirements.lock reports
no known vulnerabilities. Full raw logs are retained in results-private with
p6-final and p6-resume prefixes. Broader Ruff reported 141 existing findings in
legacy ops/tests (p6-resume-ruff.log); that unrelated baseline was not repaired.

The second final Python run emitted an ignored ZipFile destructor ValueError
while an export-route test passed. The command still exited 0 with all 387 tests
passing. This cleanup warning remains a minor observation, not a failed test.
Final resource observation at 20:27 IST: disk 6917 MiB available, swap 4669.25 MiB,
below the retained baseline. No resource stop was needed during final verification.
