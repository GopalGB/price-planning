# STATE

## Current phase

Branch: feat/v2-integration. Source HEAD: a05d30b. Documentation checkpoint follows.
2026-10-11 native Option A task is BLOCKED_RESOURCE_SWAP. Fresh baseline swap
6061.94 MiB; guard observed 7747.25 MiB, growth 1685.31 MiB exceeding 1536 MiB.
Disk at stop 5670 MiB, minimum this session 3409 MiB. Stop flag is durable in
results-private/native-resource-stop.json. Do not restart during this stopped
session or reset its baseline. Host-wide growth is not attributed to the engine.

Owner decision remains native formualizer==0.11.1 loading once per active copy,
visible-only affected evaluation, stale block calculation, independent resident
copy engines, engine_memory_budget_mb and idle LRU with editor protection.
The previous single mutable template-keyed engine is superseded. This session
dropped inherited pool/worker/pool-test changes and parked the warm-clear test
in results-private/native-redirect-start.diff. No new production code remains.
No engine implementation item, full verification or review gate completed.

Native default-loader synthetic preflight: 150 cached formulas, 47 mismatches,
all shared-formula followers. Native Workbook.load_path backend=calamine
reproduced 47/150; backend=umya raised OSError. This is synthetic evidence only,
not a real-template verdict. Aggregate report:
.planning/excel-grid/results/native-load-preflight-2026-10-11.json.

Prior implementation and verification below describe starting HEAD a05d30b and
historical adapter work. They do not verify the new native design.

## Last handoff

2026-10-11 native resume stopped by live swap guard before implementation. Exact
planner invocation header confirmed gpt-5.6-sol/xhigh; its review did not finish
and was terminated with only its owned subprocesses. No execution-lane worker
was started, no real workbook/private directory read, no ports bound, no push.
Both native synthetic probes completed; no Python full suite, web suite, Ruff,
content guard or new exact Claude review ran in this stopped session.

Inherited incompatible engine changes were dropped with git restore to HEAD.
Original full diff preserved in results-private/native-redirect-start.diff.
Superseded untracked WARM briefs and the previous guard JSON were archived as
results-private/superseded-* and dropped from the worktree.
Compatible regression test is parked there, not committed as failing code.
Updated plan and lesson NATIVE-LOAD-KEEP-RESIDENT record the owner correction.
Checkpoint commit is documentation and aggregate evidence only.

Historical handoffs follow, including superseded candidate design and baselines.

2026-10-11 20:58 IST resident attempt stopped. Baseline swap 4589.25 MiB; observed
6277.06 MiB, growth 1687.81 MiB, exceeding the 1536 MiB limit. Free disk 4747 MiB,
above the 2500 MiB floor. Recorded in results/warm-resource-live.json and private
warm-resource-stop.json. Parent interrupted /root/resident_core immediately. The
guard is host-wide; engine attribution is not established, and no scale load was
running. Do not reset the baseline or resume during this stopped session.

Retained partial code: template-keyed shared worker, ordinary no-reload clear,
dead-child fail-closed behavior and revised pool/warm tests. Focused suite 18/18
passed before the stop; full Python twice, web twice, Ruff, content guard and exact
Claude review were not completed for this attempt. No core implementation item is
accepted as finished. Known open issues include a forbidden spill export/reload
fallback, copy-switch delta publication and resetting the engine after errors.
Background startup/progress and memory/LRU/presence policy are not implemented.

Pinned API investigation is in private warm-planner-preflight.md. No clone,
deepcopy/pickle or native snapshot/restore exists. XLSX byte round-trip reparses;
batch value/formula writes, targeted evaluation and undo/redo exist. Scalar None
clear is supported, but native blocked-spill clearing/restoration is not proven
correct without reload. The candidate remains one resident base plus serialized
copy overlays, with logical state LRU and an explicit base memory limit.

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

  2026-10-11 resident execution checkpoint: locally changed `pp/engine/pool.py` and
  `worker.py` so one child is keyed by template path and serially overlays logical
  copies. Focused warm/pool tests pass 18/18. Ordinary cleared inputs no longer reload
  a renamed source; a blocked dynamic spill still uses a documented Formualizer 0.11.1
  export/reload fallback because native targeted evaluation cannot rebuild spill
  metadata. Commit was sandbox-blocked because Git cannot create the linked-worktree
  index lock outside the writable boundary. No benchmark or external action ran.
  Dead resident children now fail closed with a named server-restart error, rather
  than silently recreating a worker.

## Next task

Start a new safe session with lower host pressure. Keep the approved native
Option A decision. First reproduce and diagnose missing native shared-formula
followers using existing synthetic fixture (seed 20261011, scale 0.0001) and pinned
0.11.1. Close native-load correctness without per-cell formula registration or
silent cached downstream results. Then implement the unchecked native execution
items at the top of PLAN-P6-real-template.md, restore the parked warm-clear test,
commit each finished item and run all required final verification/review gates.
Controller owns >=20k real-workbook parity and edit/revert runs. No dependency
upgrade, per-cell loader, template-shared mutable worker or spill reload is approved.

Historical next-task notes below are superseded by this checkpoint and owner
native-load decision; old startup timings are not native-engine measurements.

Resume the approved resident task only in a new safe session after host pressure
falls. Read WARM-CORE-BRIEF.md and WARM-ENGINE-HANDOFF.md, retain synthetic-only
scope, take a fresh baseline and supervise the guard. Repair the partial core's
known correctness/reload issues before committing it. Then finish background startup,
UI progress, memory/presence policy, one-scale-at-a-time benchmarks and all review
gates. No fitted curve or full-size fit decision exists for the resident design.

The earlier Luna sandbox incorrectly treated historical swap as a current guard;
the parent corrected it. The real stop above is a separate verified live event.
Terra's commit was sandbox-denied; root Git access is available, but no unfinished
implementation was committed. Final required verification remains pending.

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
