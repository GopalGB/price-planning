# STATE

## Current phase
Track A v2 backend exact-revision repair and final confirmed regression follow-up are implemented in the
`feat/v2-backend` worktree. Local safe tests are green. Controller acceptance remains pending the two
root-owned localhost transport checks, independent review, security gates, canonical repeats, and Git integration.

## Last handoff
2026-10-10, Codex gpt-6.1-sol high, planner takeover after two executor failures.

- Repaired honest revision-zero calculation and replay, full formula snapshot persistence, shared-formula
  ISFORMULA handling, typed bool-versus-number comparisons, and immutable inherited fork baselines.
- Repaired live access revocation, session-specific logout, expiry rechecks, monotonic per-copy broadcasts, and
  same-revision calculation replay. Added importer progress and cleanup, import-only isolation, and `/template/`
  ignore coverage.
- Preserved immutable template source, mask, metadata, engine behavior, and export across replacement of the
  configured workbook path. Exact RED/GREEN evidence is indexed in `.planning/excel-grid/backend-units.md`.
- Focused controller contracts passed 39/39 after the helper-signature refactor. Two consecutive sandbox-safe
  repository runs passed 132/132. Ruff, compileall, content guard, and the touched-code structural audit passed.
- `tests/test_runtime.py` and `tests/v2/test_http_transport.py` were intentionally excluded because root owns
  actual TCP verification. Independent model review, security scans, and office acceptance were not run here.
- Controller-owned dependency locking, installs, Git index changes, commits, and pushes were not touched.
- Repaired actual legacy scrypt login, 1-based migration coordinates, durable migrated calculation jobs,
  UTC ISO migration timestamps, native error inputs, exact typed/event deltas, per-copy publication pipelines,
  bounded streaming bodies, `_xlfn.` and quoted ISFORMULA handling, fail-closed validation coverage, and
  malformed control-payload type checks.
- Final focused contracts passed 75/75. The final broad safe suite passed 148/148. Ruff, content guard, and the
  complete `pp/` function-size/signature audit passed. Evidence is in `backend-final-pending-result.md`.

## Next task
1. Root runs `tests/test_runtime.py` and `tests/v2/test_http_transport.py` with localhost binding available.
2. Root runs the canonical safe suite twice, independent review, and security gates, then sends concrete findings back.
3. Root reviews and commits the units in `.planning/excel-grid/backend-units.md`.

## Results waiting on Sir
- Office Excel parity against one approved template copy, LAN second-machine verification, measured load and
  performance evidence, review, and security acceptance remain outside this repair session.
