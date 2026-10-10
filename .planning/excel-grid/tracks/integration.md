# v2 integration handoff

Date: 2026-10-11. Branch: `feat/v2-integration`. Scope: the approved local integration checklist, synthetic fixtures only. No push, PR, deployment, company workbook or production database access.

## Changes

The colliding `pp/export.py` module is deleted. The HTTP route calls Track C's `pp.export.xlsm.export_xlsm` on a worker thread with sorted zero-based `(sheet, row, col, value)` tuples, unique output paths and normal-response/failure cleanup. `WorkbookService.source_and_overrides` validates every stored override against the copy's immutable template input mask before creating an export file. Invalid migrated or forged locked overrides return 422. The Track C escaping, streaming, atomic rename, formula refusal and calcChain behavior are retained.

The spill regression uses a genuinely blank styled child with no `<f>` or `<v>` element. The native engine classifies it as calculated, the snapshot excludes it from the input mask, a mixed edit is rejected atomically even after deleting its result cache, and HTTP export preserves its original XML. A separate occupied spill target remains editable. A temporary SEQUENCE-specific shortcut was caught, removed and never committed; importer and engine remain unchanged from the merge base. HTTP tests also assert the media type, edited input, original member order, package preservation and calculation mode.

Backup and server share `pp.config.resolve_database_path`. Precedence is explicit `--db` for backup, then `PP_DATABASE_PATH`, then `DATA_DIR`, then settings `database_path`, then settings `data_dir`, then repository `data/planning.db`. The legacy `app/runtime.py` resolver remains because legacy callers still use it. Launcher tests copy and execute the real `run_v2.py --import-template` with a synthetic template and require the snapshot's `meta.json`; only the environment installer is stubbed. Both existing launchers already dispatch to that real runner and are preserved.

Live operations carry the exact persisted history timestamp in `at`; old replay payloads obtain their original event timestamp instead of a new wall-clock value. API.md documents the field. A regression and an isolated failing mutation prove no-category listings include private named shares from another category, without including unrelated copies; revocation and category filtering are covered. Existing production listing logic already satisfied this contract.

Prettier is pinned exactly to `3.6.2`. All six `web/src/*.mjs` files were formatted; each minified esbuild transform is unchanged, and Astra independently verified exact formatting against the base. The source/build metadata changed, while the application bundle hash remains `9a06b98eeca33881b888da8c4100642686413ce17e64912e0c0d0ba633505239`. Ignore rules cover track/backend logs, frontend screenshots and runtime evidence; committed track reports remain intact.

## Tests and commands

Behavior changes were exercised red then green. Preserved full logs are under the ignored `results-private/integration/`, including export/config/time red logs, named-share mutation, old-event timestamp regression, CSP regression, and `11-spill-blocker-red.log` (two intended failures) followed by the exact pair green. Export focused suite passed 11/11 twice; importer/native spill coverage passed 14/14; the added HTTP route regression passed 1/1. Formatting did not change logic.

The pre-handoff acceptance sequence measured the following results. After the handoff commit, the controller repeats the complete sequence in the same order, preserves `final-*` logs/manifests and makes no later commit. These files are local evidence, not committed runtime data.

| Command | Measured result | Evidence |
| --- | --- | --- |
| `.venv/bin/python -m unittest discover -s tests` | 242 tests, zero failures/errors, twice | `readiness-after-repair.log`, `handoff-python-2.log`; final `final-python-1.log`, `final-python-2.log` |
| `npm test --prefix web` | 33 tests, 33 pass, twice | `handoff-web-1.log`, `handoff-web-2.log`; final corresponding `final-web-*` |
| `python3 ops/content_guard.py --all` | exit 0 | `handoff-content-guard.log`; final `final-content-guard.log` |
| `npm run build --prefix web` | twice; all 170 output files have identical SHA-256 manifests, including BUILD.json | `handoff-build-{1,2}.log` and `.sha256`; final corresponding `final-build-*` |
| `ruff check pp run_v2.py tests/v2` | all checks passed | `ruff-final.log` |
| `npm run format:check --prefix web` | all six formatted sources accepted | `prettier-controller.log` |

## Browser and CSP evidence

The actual runner is executed with its original contents from an ignored synthetic runtime directory, with repository `PYTHONPATH` and settings bound to `127.0.0.1:8765`. The template is produced solely by `ops.fixtures.make_semantic`. Browser script: `ops/integration_browser.mjs`, run through `ego-browser nodejs` with `integrationRoot` set to this worktree. It reuses one TaskSpace, 37, throughout probes and acceptance. The final successful invocation awaits `finish({keep: []})` exactly once.

The initial real-browser probe recorded eight CSP violations: four inline style violations and four data-image violations. Evidence: `headers-before.txt`, `csp-before.json`, `csp-before-console.json`. Only `style-src 'self' 'unsafe-inline'` and `img-src 'self' data:` were added. Default/script sources remain self; object-src none, base-uri none and frame-ancestors none remain. No blob worker or eval permission was added because none was observed as necessary.

The complete synthetic browser flow covers signup, category/copy creation, native Univer input editing, a dependent formula update received through WebSocket, op/history timestamp equality, locked formula refusal in both UI and HTTP, undo, a second edit, browser download and openpyxl validation of the edited input, preserved formula and fullCalcOnLoad. Post-change evidence records zero CSP violations. The expected server-side locked write produces one HTTP 422 console entry, not a CSP failure. Evidence: `browser-final.json`, `browser-final-console.json`, `browser-shots/`, `server-final.log`, and ignored `runtime/exported.xlsm`. Browser screenshots, downloads, synthetic accounts and auth/session state remain uncommitted.

## Review and security

Exact `claude-opus-5-5` was used and its returned `modelUsage` and canonical model identity were verified. First review found a HIGH unmasked stored-export override; the boundary guard and regressions close it. Final review: no open CRITICAL/HIGH, SHIP conditional on the final suite/content guard. Evidence: `opus-review.json`, `opus-final-review.json`. Required independent `gpt-6-astra` read-only review: SHIP for the scoped integration diff, no new CRITICAL/HIGH; `astra-review.txt`. Planner lane `gpt-5.6-sol` at xhigh performed the corrective implementation and focused verification after executor escalation; `planner-repair.txt`.

`security-auditor` ran, including staged secret scans before each code commit. Gitleaks found no secrets; required Semgrep configs (`p/python`, `p/javascript`, `p/owasp-top-ten`) returned zero findings and errors; pip-audit against `requirements.lock`, npm audit and explicit OSV npm lockfile audit reported no known vulnerabilities. Evidence: `security-auditor.log`, `*-staged-security.log`, `semgrep-final.json`, `pip-audit-python311.log`, `npm-audit.json`, `osv-web.json`. Broad auto Semgrep and Bandit retain fixed-identifier SQL/operator-template XML warnings: Bandit has 11 MEDIUM and no HIGH. Both independent reviews assessed these paths as non-exploitable; scanner warnings were not hidden or described as a clean Bandit run.

## Unverified items and retained concerns

This is local synthetic integration acceptance, not office release approval. Windows batch execution, Microsoft Excel/VBA recalculation, office LAN/user load, real-workbook parity and real legacy-data acceptance were not run. No company workbook or production database was opened. The paid Strix gate was not invoked or approved; P5 remains pending.

Both reviewers retain the pre-existing MEDIUM risk that migrated baselines/history can contain non-input coordinates. Before real deployment, an owner-approved count-only audit must establish whether such legacy rows exist; affected rows need a reviewed migration. Export now refuses bad stored overrides rather than silently writing them. Opus also notes pre-existing inability to re-edit/undo a cleared spill blocker once it becomes a calculated result. These are outside the requested integration changes and do not invalidate the ordinary-input undo acceptance.

Known LOW limits: cancellation/disconnection during threaded export can leave an orphan file despite normal cleanup; exports have no per-user concurrency limit; unsupported legacy value types surface as generic errors; the synthetic browser's pixel clicks depend on the inspected fixture layout. Future engine/importer changes must revalidate spill classification and stored overrides. No claim of sweeping orphan files or production migration safety is made.

## Commits and next action

Logical code commits: `1aa6f34` frontend pin/format, `c72eaf7` exporter/mask/CSP, `4b5e2ce` resolver/real-launcher smoke, `46c59d0` timestamp/share contract, `aeaf4d7` synthetic browser harness/hygiene. The handoff commit contains this report, PLAN.md and STATE.md. No commit has co-author trailers; hooks were not bypassed.

Next owner action: review these local commits, push this branch and open a PR when ready, then require green CI and separately authorize office/legacy/Excel/Strix acceptance. Nothing was pushed or deployed in this session.
