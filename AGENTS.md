# Price Planning - agent rules (read first, every harness: Codex, Claude Code, Claude cloud)

## What this is
Excel-like multi-user web app for the Spring 2027 price-plan template. Design: `.planning/excel-grid/SPEC.md`.
Current phase and next task: `.planning/STATE.md`. Phase plans: `.planning/excel-grid/PLAN-P*.md`.

## Hard rules
1. Never commit the .xlsm, any spreadsheet/office file, VBA exports, databases, logs, or cell values from the
   real workbook. CI and the pre-commit hook enforce this; never bypass them (`--no-verify` is forbidden).
2. Real-workbook scripts (`ops/bench_real.py`, `ops/find_inputs.py`) write committable aggregates to
   `.planning/excel-grid/results/` and everything else to `results-private/` (gitignored).
3. Work on a branch, open a PR, merge only when CI is green. Never push to `main`.
4. Before ending a session, update `.planning/STATE.md` (branch, commit, tests run, next task).
5. No em-dashes. No co-author trailers in commits.
6. Template formulas are locked in v1; users edit input cells only (SPEC D3).
7. Engine is `formualizer==0.11.1` until P1 says otherwise. Do not upgrade dependencies without updating
   `requirements.lock` and `vendor/wheels/` together (see docs/DEPLOY.md).

## Setup (once per clone)
    python ops/setup_dev.py

## Tests
    python -m unittest discover -s tests -v
    python ops/content_guard.py --all

## Office server
See `docs/DEPLOY.md`: `git pull`, then `run.bat`.
