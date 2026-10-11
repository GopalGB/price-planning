# P6: Cached startup and synthetic template scale

Date: 2026-10-11. Approved by the owner's GO-TOKEN.
Branch: `feat/v2-integration`. Synthetic fixtures only. No push or deployment.

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
