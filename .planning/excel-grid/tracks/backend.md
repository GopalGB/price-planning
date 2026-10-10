# Track A backend handoff

## Outcome

The authorized Track A backend is implemented in the `feat/v2-backend` worktree. It is locally verifiable with
synthetic fixtures and is ready for controller review and atomic commits. Release completion is not claimed
until the environment-blocked loopback checks and controller-owned independent gates pass.

## Implemented contracts

- Atomic template import with source hashing, 256-row blocks, disk-backed shared strings, shared-formula
  expansion, defined names, validation metadata, full render styles, protection inheritance, and publish-on-success.
- Formualizer `0.11.1` adapter with formula discovery, required compatibility behavior, changed-formula deltas,
  exact-revision replay, one warm spawned process per active copy, removed-override restoration, busy protection,
  and LRU eviction.
- Additive SQLite schema and idempotent legacy migration with WAL, foreign keys, busy timeout, copy state,
  operation idempotency, cell history, durable events, calculation jobs, and crash recovery.
- Transactional batch edits with role checks, input-mask enforcement, coordinate and validation bounds, formula
  injection rejection, optimistic versions, all-or-nothing conflicts, and actor/payload-bound operation IDs.
- Scrypt authentication, hashed sessions, expiry, strict cookies, trusted Origin and CSRF checks, login throttling,
  category membership, effective-role precedence, private-copy invisibility, owner controls, and frozen forks.
- Workbook metadata, rows, cell history, XLSM package-preserving export, ordered operation/calculation WebSocket
  events, replay, presence, safe stale-socket removal, and effective-role-aware revocation.
- Single-process Starlette/Uvicorn runtime in `run_v2.py`, explicit import-only mode, startup migration/recovery,
  static fallback, actual request-body limits, security headers, JSON API errors, and sanitized internal failures.

## TDD and verification evidence

- Per-unit RED/GREEN logs and ownership: `.planning/excel-grid/backend-units.md`.
- Final hardening RED: `.planning/excel-grid/final-hardening-red.log`.
- Final hardening GREEN: `.planning/excel-grid/final-hardening-green.log`, 17/17 tests passed.
- Final v2 pass 1: `.planning/excel-grid/backend-v2-final-pass1.log`, 55/55 tests passed.
- Final v2 pass 2: `.planning/excel-grid/backend-v2-final-pass2.log`, 55/55 tests passed.
- Ruff: `.planning/excel-grid/backend-ruff-final.log`, all checks passed.
- Compile: `.planning/excel-grid/backend-compile-final.log`, exit 0.
- Repository suite: `.planning/excel-grid/backend-full-final.log`, 119 run, 117 passed, 2 environment errors,
  zero assertion failures. Both errors are `PermissionError: [Errno 1] Operation not permitted` at localhost bind.
- Content guard: `.planning/excel-grid/backend-content-guard-final.log`.

## Boundaries and remaining gates

- Only generated synthetic workbooks and worktree-local temporary data were used. No real/private workbook,
  database, cell values, credentials, or production logs were accessed.
- The managed sandbox prevents loopback bind, so `tests/test_runtime.py` and `tests/v2/test_http_transport.py`
  must run in the controller environment without changing or skipping them.
- Scoped Gitleaks passed for `pp/`, `tests/v2/`, and `run_v2.py`; final logs are
  `.planning/excel-grid/backend-gitleaks-pp-final.log`, `backend-gitleaks-tests-final.log`, and
  `backend-gitleaks-runner-final.log`. Semgrep crashes before scanning because its CA trust store reports empty
  trust anchors; Bandit and pip-audit are unavailable. Tool-failure evidence is in
  `.planning/excel-grid/backend-security-preflight.log` and `.planning/excel-grid/backend-semgrep-local.log`.
- Controller must run the full security audit plus independent `gpt-6-astra` and exact `claude-opus-5-5` reviews,
  resolve all CRITICAL/HIGH findings, and verify returned model identity before accepting Track A.
- Controller retains ownership of dependency files, lock regeneration, installs, Git/index operations, commits,
  pushes, CI, Office/Windows/LAN validation, real-workbook parity, and performance acceptance.
