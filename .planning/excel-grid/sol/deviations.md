# P0 execution deviations

- Task 1: the proposed local denylist starter `catalyst brands` matches existing tracked launcher content, so populating it makes `ops/content_guard.py --all` fail. The guard and ignored denylist are prepared, but no denylist is populated until Sir confirms tokens that do not match approved tracked files.
- Task 1: the plan invokes tests as `python -m unittest tests.test_*`, but the repository had no `tests/__init__.py`. Add the empty package marker so the exact plan commands target the named modules.
- Task 3: `run.bat` has three explicit interpreter branches and no `PYTHON` variable. Installer integration preserves that detection structure and invokes the installer in each branch.
- Task 3: `install_env.py` sends pip output to stderr so Windows launcher command substitution receives only the virtual-environment interpreter path on stdout.
- Task 3: Windows offline wheel vendoring includes `requirements-dev.lock`, since the Task 5 office script needs openpyxl.
- Task 4: write-only openpyxl sheets omit reliable dimension metadata for read-only reopening; the scale test opens normally to validate dimensions.
- Task 5: synthetic openpyxl formula cells have no Excel cached values. The smoke test checks benchmark completion and the public/private boundary, not P1 formula-parity quality.
- Task 5: direct dependents are queried with an explicit limit and marked truncated when the engine reports truncation; capped counts are never represented as complete fanout.
