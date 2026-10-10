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

## Review round 2

Date: 2026-10-11. User GO-TOKEN approved STEP 0 and A-E on `feat/v2-integration`. This section supersedes round 1's test counts and next-task status for the reviewed integration branch. Work remains local and synthetic; no push, real workbook, production database or office acceptance was performed.

STEP 0 merged `feat/v2-frontend` through `3e10cc0` with `git merge --no-ff feat/v2-frontend`. Merge commit `88188de` has parents `7269bd7` and `3e10cc0`. Each conflicted `web/src` file takes the frontend version. The preserved `git diff -w a186003..HEAD~ -- web/src` was inspected against the pre-merge parent: round 1's source changes were formatting only, and `git diff 1aa6f34..7269bd7 -- web/src` is empty. Pinned Prettier was reapplied and `static/dist` rebuilt. API conflict resolution keeps the persisted operation timestamp and frontend legacy metadata-refetch guidance. STATE retains `Current phase`, `Last handoff` and `Next task`. Before A-E, Python 242/242 and web 47/47 passed. Evidence: `merge.log`, `round1-web-semantic.diff`, `merge-resolution-note.txt`, `merge-format.log`, `merge-build.log`, `step0-python.log`, `step0-web.log`.

### Changes and regression evidence

All logs below are preserved under ignored `results-private/review-round-2/`. No logs, cell values, synthetic Office files or databases are committed.

| Phase | Change and independent expectation | Red then green evidence |
| --- | --- | --- |
| A | Shared `pp.values.validate_text` rejects unsupported XML C0 controls except tab/LF/CR, U+FFFE/U+FFFF, lone surrogates and text above 32767 characters with 422 `invalid` before committing. Track C `pp/export/xlsm.py` reuses the validator. An API edit to an implicit unlocked B2 with no template `<c>` exports and reads back with valid controls and literal OOXML escape text intact. | `A-red-expected.log`: five intended unsupported-text failures; the existing length rejection and Track C implicit-cell insertion already passed. `A-green.log`: 4/4. Commit `7930a6b`. |
| B | Each authenticated mutation passes the persisted session check into its write transaction after reading the body and, for ops, after acquiring the pipeline lock. Real paused-body tests cover category create/member update, copy create/fork/patch/share, edit and undo. Revoked/expired requests return literal 401 and leave mutation tables unchanged. | `B-mutation-red-2.log`: 3 tests, 10 intended failures with checks disabled. `B-green-2.log`: auth module 9/9. The original failed call-count test was replaced with behavioral tests; its failed logs remain available. Commit `33b396a`. |
| C | `Database.read_transaction` issues BEGIN and commits on the same connection. Replay and rows retain one snapshot from revision metadata through events/results. A real second-connection calculation publication fires immediately after the metadata query. The old rows/result pair stays consistent, and replay sends pending state before the newer calc, which remains deduplicated correctly. | `C-red.log`: both review interleavings fail. `C-green.log`: 2/2. Commit `62f0607`. |
| D | Native date, time and datetime values become numeric Excel serials using the source workbook's date1904 flag. Literal assertions require DATE(2026,1,1)=46023, TIME(12,30,0)=0.5208333333333333, datetime=46023.25, 1904 DATE=44561 and the 1900 leap-day boundary 59/61. | `D-red.log`: both date-system tests fail on string results. `D-green.log`: 2/2. Commit `ad2d547`. |
| E | Forks copy original input editor/time. History recursively reuses immutable parent entries through each fork's revision cutoff and marks inherited entries with `origin`. A later parent edit is excluded. Fork-local edit and undo clear current origin markers while retaining earlier history. | `E-red.log`: Alice attribution is replaced by Bob/fork time. `E-green.log`: the full attribution, frozen history and local edit/undo regression passes. Commit `79ea428`. |

The API document now describes inherited history `origin`, text rejection and transaction-level session checks. No engine or dependency version was upgraded.

### Acceptance and review

Controller acceptance on code commit `79ea428`: full Python 251/251 and web 47/47 pass (`controller-python-1.log`, `controller-web-1.log`, both exit 0). Consolidated B-E backend coverage passes 78/78 (`final-BE-focused.log`). Ruff and pinned Prettier checks pass (`controller-ruff.log`, `controller-format.log`). Content guard passes (`final-content-guard.log`, exit 0).

The worker's full-suite attempts are preserved and were not called green: Python 244 tests with two loopback-bind permission errors, web 45/47 with two listener EPERM failures. The controller ran the same complete suites successfully in the authorized environment. Evidence: `sandbox-acceptance-errors.txt`, `final-python.log`, `final-web.log` and controller logs above.

The first exact Opus 5.5 review found one additional blocking MEDIUM in Track C's existing escaping: overlapping literal `_xHHHH_` strings and an escape prefix followed by CR were corrupted on export. Both were reproduced through exporter and implicit-blank API round trips (`export-escape-red.log`: 15 tests, 3 intended failures). Commit `3340943` uses one lookahead substitution for literal prefixes and CR; `export-escape-green.log` passes 26/26, and the scoped Ruff gate passes. A broader lint attempt included a legacy test module outside the established gate and exposed seven pre-existing style findings; that failed log is retained.

Final exact `claude-opus-5-5` review: SHIP, original escaping blocker closed, no CRITICAL/HIGH or blocking MEDIUM remains. Both initial and final JSON results verify `modelUsage` and `canonicalModel` as exactly `claude-opus-5-5`; evidence `opus-review.json`, `opus-final-review.json`. Fresh `gpt-5.6-sol` xhigh review of A-E is SHIP (`planner-review.txt`); the incremental escaping review is also SHIP (`planner-final-review.txt`).

Final acceptance on corrected code `3340943` ran in the controller environment: Python 251/251 twice, web 47/47 twice, content guard exit 0. Evidence: `final-controller-python-{1,2}.log` and corresponding exit files, `final-controller-web-{1,2}.log` and exit files, `final-controller-content-guard.log` and exit file. These are full suites with zero failures, errors or skipped tests. The final evidence commit changes only planning/API documents; no code changes follow these runs. Content guard is repeated after sealing the documents.

Execution first used exact `gpt-5.6-luna` low; linked-worktree Git writes were blocked by its workspace sandbox. The controller completed authorized Git operations. Exact `gpt-5.6-terra` then implemented A and the initial B change; after its inadequate call-count verification failed, exact `gpt-5.6-sol` xhigh completed B-E. The original outputs and corrective handoff are retained in `executor.jsonl`, `repair.jsonl`, `planner-repair.jsonl`, `planner-repair-final.txt`. No model was silently substituted and no hook was bypassed.

### Security and retained limits

`security-auditor` ran. Every merge/code commit had a clean staged gitleaks scan. Active `requirements.lock` pip-audit and `web/package-lock.json` npm audit report no known vulnerabilities. Scoped Semgrep (`p/python`, `p/javascript`, `p/owasp-top-ten`) is clean. Final code secret/SAST refreshes are retained in `gitleaks-final.log` and `semgrep-final.json`.

The broad audit is not clean: auto Semgrep reports 32 pre-existing SQL/XML/CI findings, and OSV finds seven advisories in unchanged `proto/univer-gate/package-lock.json` for development-only Vite 6.3.5, including two HIGH. The active web lockfile uses Vite 6.4.4. Prototype dependency remediation is outside A-E and no dependencies were changed. Evidence: `security-auditor.log` (exit 1), `npm-audit.json`, `pip-audit.log`. Opus and planner distinguish these from the clean active dependency audits and assess constant-identifier SQL as non-injectable. Do not interpret scoped acceptance as repository-wide security clearance.

Windows/Excel/VBA, real-workbook parity, legacy-data audit, LAN/load, paid Strix and P5 release acceptance remain unverified. This round does not repeat the round 1 browser flow. Nested fork history is implemented, but the new provenance regression explicitly covers one fork level. Opus retains LOW notes for generic export errors on old invalid stored text and unverified Excel UTF-16 length/date1904 numeric-input function parity; these were not silently changed. No push or deployment was performed.
