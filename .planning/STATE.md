# STATE

## Current phase
P6 engine viability work is in progress on `feat/v2-integration`, approved 2026-10-11.
Array compatibility is committed as `2ec0493`; deterministic scale fixtures are
committed as `9dce2ba`, cached startup and sparse publication as `0e2e386`.
Benchmark runner guard repair, scale evidence, and independent review remain pending.

v2 integration is feature-complete and verified on the Mac with synthetic fixtures, branch `feat/v2-integration`, HEAD after this handoff. Round 5 (eight review defects) plus two controller fixes are committed. Not pushed. Office Windows/Excel/LAN/load, real-workbook parity and paid Strix acceptance remain unverified; this is not a P5 release claim.

## Last handoff

2026-10-11 19:49 IST: exact review returned verified modelUsage
claude-opus-5-5 and FIX-FIRST. HIGH outdated runner limits and LOW fixture
cleanup already resolved by 63e4391. Named-range trace budget, import baseline
versioning, bootstrap error wrapping and repeated process-group cleanup are under
review/repair. Cached volatile-on-open finding is being checked against explicit
cached-startup scope and the inventoried function list. Review log:
results-private/p6-resume-review.json. Scale 5% remains in load; no timing accepted.


2026-10-11 19:44 IST resume checkpoint: focused tests 112/112 PASS; primary full
Python run 1 368/368 PASS; web 101/101 PASS twice. Python run 2 is running.
Scoped ruff (pp, tests/v2, new scale tools) clean; broad legacy scan has 141
pre-existing findings in older ops/tests. Content guard exit 0. Staged gitleaks
clean for both new commits. pip-audit requirements.lock found no known vulnerabilities.
Semgrep found 0 findings with one rule timeout; rerun in progress. Exact
claude-opus-5-5 review is running and remains a gate, not a completion claim.
Resume swap baseline 8875.31 MiB and minimum free disk 2500 MiB; stop at growth
over 1536 MiB. No real workbook/private path was read, no push or server launch.

2026-10-11 19:33 IST resume: stopped at the owner's disk guard before focused tests,
review, benchmarks, or commits. `df -g /` reported 3 GB available (below the 4 GB
floor); `sysctl vm.swapusage` reported 9896.44 MB used, 11264.00 MB total. This is
the observed resume swap baseline, not evidence of growth during this run. Branch
is still `feat/v2-integration`; all prior tracked and untracked work remains in
place. No real workbook or private workbook directory was read. Free disk capacity
first, then resume focused verification and exact `claude-opus-5-5` review.

2026-10-11 P6 cache-repair checkpoint: uncommitted engine changes replace direct/regex
dependency targeting with bounded native transitive tracing, invalidate affected XML
caches, avoid empty-batch spill scans, and recalculate clear-rebuild closures. New
cache tests cover a cached chain, cross-sheet range, dynamic downstream, typed cache,
potential spill input masking, clear rebuild, and truncation. Cache plus adjacent
engine suites are green (52 tests), ruff and content guard are green. Full discovery
completed 310 tests with three environment errors only: this sandbox denies the
loopback binds in test_runtime, v2.test_http_transport, and v2.test_run_v2. Raw logs:
results-private/p6-cache-terra-*.log. Exact `claude-opus-5-5` review remains pending;
no commit or push was made.

2026-10-11 06:40 IST, controller (claude-opus-5-5). Commits after Round 5: `0079f66` refuse to start when another program holds the port (found live: a first start bound 0.0.0.0:8000 beside a loopback listener on 8000); `0ff02dd` literal dropdown members take the type Excel enters on pick, IMPORT_VERSION 6 (independent review finding N1).

Checks run by the controller: Python 294/294 twice, web 101/101 twice, content guard 0, ruff clean. Browser acceptance in Ego TaskSpace 38 over the LAN URL on port 8765: signup, first category, copy, Inputs!B3=24, Calc!B3 shows 0.583333 on a default-width column, download read with openpyxl (Inputs!B3=24, formulas intact, no calcChain, fullCalcOnLoad). After the IMPORT_VERSION bump the template re-imported at restart and the existing copy reopened at rev 1 with the same values.

Independent Claude review of fbd67a1..4903d79: all seven original defects CLOSED, General width CLOSED; each new regression test fails on the old code. Open, not blocking: N2 General precision depends on the browser's fallback font (display only); N3 one stale template row makes backup fail closed; backups made before Round 5 restore without template relocation; copy retry limit (3) under heavy edit churn untested; refreshing the page returns to the copies list instead of the open copy.

Codex usage limit reached; Astra/Sol unavailable until 2026-10-14 09:03 IST.

## Next task
Finish exact Opus review, commit bounded runner with current resource guards and
per-scale fixture cleanup, finish second Python run, then run scales sequentially
with one retained 8875.31 MiB swap baseline. Write aggregate report and host
recommendation. Never claim 20%/50% complete if a guard or timeout stops 5%.

Owner actions: push `feat/v2-integration` from Terminal.app (pre-push gate needs `y`) and open the PR; office test (pull, put the .xlsm in `template/`, run `run_v2.bat`, open from a second PC; if port 8000 is taken the launcher now says so, set "port" in settings.json). After Codex resets: one Astra pass over `4903d79..HEAD`, then the low items above. Strix paid run stays gated on the owner's OK and a per-scan budget key.
