# STATE

## Current phase

P6 implementation is committed on `feat/v2-integration`; large-scale engine viability
remains BLOCKED (`P6_LARGE_SCALE_STARTUP`). Latest code commit: `afc320a`. Final
handoff commit follows. Approved GO-TOKEN: 2026-10-11. No push or deployment.

Single-cell array compatibility/export preservation, template-map ISFORMULA, typed
XML cache startup, dirty dependency recalculation and sparse revision-checked
publication are implemented. Engine remains `formualizer==0.11.1`. IMPORT_VERSION 7
rejects stale cached snapshot formats with a named reimport error. Multi-cell arrays
and dataTable remain rejected; unsupported structured/external/named-colon references
fail closed rather than silently returning stale native values.

Prior v2 integration remains feature-complete on synthetic Mac tests. Office
Windows/Excel/LAN acceptance, real-workbook parity and paid Strix acceptance remain
unverified. Synthetic P6 results do not advance P1/P5 release gates.

## Last handoff

2026-10-11 20:27 IST final resume handoff: previous run was interrupted by a runner timeout. Read STATE/status,
ran 112 focused tests successfully and committed retained work in focused items.
Array compatibility `2ec0493`; generator `9dce2ba`; cache/delta engine `0e2e386`;
bounded runner `63e4391`; named dependency/cache-format repair `65c190a`; runner
cleanup repair `8b5a076`; batch registration `c1b06fc`; unsafe reference rejection
`181e267`; whole-range regressions `802e0f5`; short-name fix `fccb177`; ISFORMULA
validation-order fix `afc320a`.

Scale aggregate: `.planning/excel-grid/results/engine-scale-2026-10-11.json`.
Plan/results/recommended path: `.planning/excel-grid/PLAN-P6-real-template.md`.
Batched 5%: 260000 cells/75000 formulas, load 258.7317 s, first edit 0.4232 s
(1 correct delta), broad edit 4.7513 s (75000 correct deltas), child peak RSS
708395008 bytes (675.6 MiB). At 20%: 1040000 cells/300000 formulas, load incomplete
at 600 s attempt timeout; edits unmeasured. At 50%: skipped after 20% timeout.
Prior unbatched 5% load: 455.15 s; 20% also timed out. Fixtures were deleted after
each attempt; only small owned progress/log files remain. Benchmark source:
`c1b06fc`; later correctness guards were not included in those timings.

Resource baseline: 8875.31 MiB swap; stop at growth above 1536 MiB or free disk
below 2500 MiB. Latest scale run minimum disk: 6922.94 MiB; swap never exceeded
baseline. No resource guard fired. Initial resume disk: 5054 MiB. Host pressure
varied, so timing comparisons are not isolated. No real workbook/private directory
was read. No forbidden port was bound.

Exact returned `modelUsage` identity `claude-opus-5-5` verified throughout reviews.
Repairs closed with SHIP in `p6-tax-review.json` and `p6-isformula-review.json`.
Planner `gpt-5.6-sol` verdict: NO BLOCKING DEFECTS after final repair. Initial final
Python run passed 386/387: ISFORMULA named-argument rejection was bypassed by name
expansion. Existing test reproduced it; validation now precedes expansion. Focused
83/83 twice and adapter/cache 48/48 twice pass. Final unchanged full Python runs:
387/387 PASS twice (85.657 s and 71.249 s). Second run emitted an ignored ZipFile
destructor cleanup warning during a passing export-route test; recorded in the plan.

Web: 101/101 twice PASS. Scoped Ruff clean; content guard exit 0; staged gitleaks
clean. Semgrep: 151 rules/41 Python files, 0 findings/no timeout; final adapter patch
rescanned with 151 rules/1 file, 0 findings. pip-audit requirements.lock: no known
vulnerabilities. Broad Ruff reports 141 existing legacy ops/tests findings, not
repaired. Raw logs: `results-private/p6-final-*`, `p6-resume-*`, `p6-isformula-*`.

Prior controller checkpoint: `0079f66` prevents startup on an occupied port;
`0ff02dd` corrects dropdown value types. Prior Round 5 review defects were closed.
Known minor limits include font-dependent General precision, stale-template backup
fail-closed, legacy restore relocation, unmeasured retry churn, and refresh returning
to the copy list.

## Next task

Profile 20% synthetic startup on the office server using the same seed and pinned
engine. Separate eager native import, expanded formula registration and dependency
trace construction before choosing a narrow adapter repair or requesting a
dependency change. Do not call the full-size engine path viable from this evidence.

Recommended initial office host: 32 GiB RAM, one engine worker; 64 GiB planning
allowance for multiple warm copies. Unverified linear extrapolation of the 5% child
peak alone is 13.2 GiB per full-scale engine; native load CPU cost remains a blocker.
The 8 GiB Mac is for small-fixture development. Estimate limits are in the P6 plan.

Owner: push branch/open PR only after reviewing this blocker and green CI. Office
acceptance and paid Strix actions remain owner-controlled. No push was performed. Final disk: 6917 MiB available; swap: 4669.25 MiB.
