# P0 execution deviations

- Task 1 (resume): the optional Max HQ pre-commit hook uses Bash syntax (`set -o pipefail`, arrays and `[[ ]]`). Invoke it with bash rather than sh so POSIX sh implementations do not reject it. A regression test models a strictly POSIX-only sh.

- Task 1: the proposed local denylist starter `catalyst brands` matches existing tracked launcher content, so populating it makes `ops/content_guard.py --all` fail. The guard and ignored denylist are prepared, but no denylist is populated until Sir confirms tokens that do not match approved tracked files.
- Task 1: the plan invokes tests as `python -m unittest tests.test_*`, but the repository had no `tests/__init__.py`. Add the empty package marker so the exact plan commands target the named modules.
- Task 3: `run.bat` has three explicit interpreter branches and no `PYTHON` variable. Installer integration preserves that detection structure and invokes the installer in each branch.
- Task 3: `install_env.py` sends pip output to stderr so Windows launcher command substitution receives only the virtual-environment interpreter path on stdout.
- Task 3: Windows offline wheel vendoring includes `requirements-dev.lock`, since the Task 5 office script needs openpyxl.
- Task 4: write-only openpyxl sheets omit reliable dimension metadata for read-only reopening; the scale test opens normally to validate dimensions.
- Task 5: synthetic openpyxl formula cells have no Excel cached values. The smoke test checks benchmark completion and the public/private boundary, not P1 formula-parity quality.
- Task 5: direct dependents are queried with an explicit limit and marked truncated when the engine reports truncation; capped counts are never represented as complete fanout.

- Task 3 (resume blocker): PyPI requests are denied by the sandbox network allowlist. Hash-checked lock generation and Windows wheel downloads remain pending. Keep requirements.txt until verified replacement locks exist. The installer created .venv, but no dependency install completed. Temporary uv/pip caches avoid writes outside the sandbox.
- Task 3 (resume): preserve --demo and unbuffered output in run.sh, matching the existing launcher behavior rather than changing the app mode during dependency integration. Use the supplied /opt/homebrew/bin/uv on this Mac.
- Task 5 (resume): openpyxl 3.1.5 read-only rows contain EmptyCell instances without row/column attributes. Enumerate coordinates in the cached-value pass and close both workbooks. This keeps the parity sample intent.
- Task 5 (resume): ps is forbidden in this macOS sandbox. Use resource.getrusage on macOS as a conservative peak-RSS measurement, retaining current-RSS sources on Linux and Windows. Mac after-load/after-evaluation metrics therefore report high-water RSS rather than instantaneous RSS.
- Task 5 (resume): Formualizer 0.11.1 exposes truncation.incomplete, not truncation.truncated. Request at most 10,000 direct dependents and report fanout_truncated. Synthetic coverage includes more than 10,000 dependents and both watchdog limits.
- Task 5 (resume): public parity keys use stable sheet-name hashes instead of raw sheet names to satisfy the aggregate-only/no-label rule. Raw sheet names remain in private mismatch records.
- Task 5 (resume): create results-private before redirecting input references and stop the office script on dependency, input-discovery, manifest or benchmark errors instead of continuing after failure.
- Execution boundary: keep feat/p0-rails for all tasks as explicitly requested. Git config/staging/commits are blocked by sandbox permissions on .git/config and .git/index.lock. No hooks were bypassed and no commits were created. Task 5 Step 6 is deliberately skipped by user instruction. The exact gpt-5.6-luna routing probe was blocked at ab.chatgpt.com; local work continued in the executor session explicitly assigned by the user.
