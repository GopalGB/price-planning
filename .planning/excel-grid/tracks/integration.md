# v2 integration handoff

## Round 5

Date: 2026-10-11. Branch: `feat/v2-integration`. Code anchor: `3a230c0`. All eight requested defects and concrete review follow-ups are fixed in eight focused conventional commits. User GO-TOKEN authorized every phase. Final code reviews, automated acceptance and real-server HTTP/export checks pass. Current-code visual browser acceptance is pending the named blocker `BROWSER-TASKSPACE-DELETED`; this is not DONE or a release approval.

Scope: synthetic fixtures only. No private directory or real workbook was read. Port 8000 was never bound or disturbed. No unowned process was stopped. Nothing was pushed, merged, published or deployed. Native engine remains `formualizer==0.11.1`; dependency pins and vendor wheels are unchanged.

### Fixes and test-first evidence

Each requested defect has an observed failing synthetic regression before its production fix and a passing run after. Concrete review follow-ups also followed red/green order and were folded into the responsible commit. Existing tests and behavioral assertions were retained. Complete original outputs are in ignored `results-private/round-5/`; no runtime workbook, database, screenshot or log is committed.

| Defect | Focused commit and final behavior | Reproduction and verification receipts |
| --- | --- | --- |
| 1 HIGH, spill membership | `b7ecc61`: authoritative style/kind metadata creates absent spill cells, restores contracting inputs and preserves null formulas before advancing revisions. Loaded sparse blocks and public row/column extents grow; released temporary blocks cannot be repopulated. | `1-planner-red.log`, `1-review-red.log`, `1-extents-red.log`, `1-temporary-controller-red.log`; corresponding controller green runs and final suites. |
| 2 HIGH, portable backup | `283f9fa`: verified immutable template bundle and relocated restore paths recover older H1 copies and pending H1 calculation with H2 configured and original root unavailable. | `2-controller-red.log` reproduces HTTP 500/FileNotFoundError; affected suite 26/26. `2-env-repro-red.log` then `2-env-repro-green.log`, 7/7. |
| 3 MEDIUM, consistent clipboard | `956deeb`: every row block must share one input/calculation revision pair. Retry the entire range at most three times; churn never writes a mixed clipboard. Navigation fences remain. | `3-red.log`, `3-green.log`, `3-controller-web.log`; final suites. |
| 4 MEDIUM, spill-blocker undo | `844657f`: original explicit/implicit inputs remain authorized through spills; undo/redo separates semantic input lineage from displayed rollback state. Fresh/reloaded cells, editing paths and formula bar use authoritative input metadata. Coalesced calculation intervals republish touched inputs across the covered revisions; pending display uses its optimistic semantic value. | `4-red.log`, `4-followup-red.log`, `4-review-red.log`, `4-reload-pre-fix-behavior-red.log`, `4-coalesced-red.log`, `4-formula-controller-red.log`, `4-interval-controller-red.log`, `low-controller-red.log`; corresponding green runs and final suites. |
| 5 MEDIUM, OOXML text | `5838637`: decode each shared/inline text node once before joining; preserve protected literal prefixes, carriage returns and surrogate pairs across runs. Persisted history, undo and export recover the original value. | `5-red.log`, `5-history-green.log`, `5-review-red.log`, `5-review-green.log`; XML/openpyxl round trips and final suites. |
| 6 MEDIUM, dropdown values | `308f3d4`: dropdown selections preserve the selected member's native type. Typed numeric edits retain generic parsing. | `6-red.log` reproduces zero POST for string member; `6-green.log`, `6-controller-web.log`; final suites. |
| 7 MEDIUM, cell metadata | `24c5589`: cell detail uses `_input_record` within one read snapshot for ver/by/at, untouched version zero and legacy metadata recovery. | `7-red.log`, `7-controller-python-green.log`, `7-controller-web.log`; final suites. |
| 8 MEDIUM, display width | `3a230c0`: default 64px width is shared by snapshot, live and merged paths. Canvas measures the actual imported/rendered font with consistent numeric padding. Deterministic fallback shows 0.583333 at 64px and trims at 63px. | `8-red.log`, `8-followup-red.log`, `8-browser-red.log`, `8-review-red.log`, `8-ci-red.log`, `low-controller-red.log`; corresponding green runs, final suites. Final visual check pending. |

The original full-suite failure exposed legacy `tests/test_core.py` setting DATA_DIR globally. The new backup test isolates DATA_DIR/PP_DATABASE_PATH/PP_TEMPLATE_PATH while retaining every restore assertion. The failed run remains in `full-suite-first-python-failure.log`. An initial stdin multiprocessing probe failed at its spawn entry point; it is not counted as the actual backup regression. Fixture/configuration failures and module-name errors are retained separately from genuine red reproductions.

### Lossless decisions

Backup format version 2 comprises online SQLite `.db`, a `.json` manifest and a sibling `.templates` directory. Every registered immutable template is included, even when no current copy references it. Each source SHA-256 directory contains the source and complete snapshot; source identity and snapshot hashes are verified. Move all three companions together. Restore stages and verifies assets/database, relocates paths under `<restore>/templates/<hash>`, and publishes the database last. Unsafe paths, corruption, copy failures and destination races fail without overwriting existing data. Legacy DB-only backups remain supported but cannot recreate omitted assets.

Pinned Formualizer treats explicit None as a present blank that blocks native SEQUENCE. When cleared-override membership changes, calculation uses the existing lossless exporter to omit those nodes in a transient source, reloads the engine and applies the exact revision's overrides. Immutable source and formulas remain unchanged. This adds export/reload work on clear-set changes; large-workbook latency was not measured.

Original input eligibility is separate from displayed spill membership. Rows, cell detail and calculation events carry `input`, `input_value` and authoritative version/editor/time when appropriate. Formula-only spill followers stay locked. Touched inputs are selected across `(previous_calculated_revision, published_revision]`; client interval tracking advances covered blocks without inventing an uncovered gap. Temporary hydration evicts cells, input versions and loaded-block markers together. Formula-bar reconciliation uses the current store and preserves focused drafts.

OOXML `<t>` nodes have ST_Xstring semantics. Decode escapes once per node, then join rich text and combine surrogate pairs; partial escape fragments spanning nodes remain literal. Protected `_x005F_` prefixes remain literal after one decode. Import cache version 5 also accounts for imported font names. Sources: [Open XML Text](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.text?view=openxml-3.0.1), [MS-OE376 escaping](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/bd0aa042-434a-4ca7-b25f-4e1fd25a954d).

Canvas and Univer use the same point-to-pixel size, bold/italic and font family. Default family is Calibri, Arial, sans-serif; General numeric cells reserve explicit 1px left/right padding. Imported single-family names are quoted/escaped, while comma-containing or invalid names use the shared safe fallback because the renderer parses comma-separated families. Raw imported metadata remains intact. Node fallback uses deterministic recorded default-font metrics; it does not establish actual custom-font browser accuracy. Historic width assertions are retained under an explicit legacy measurement fixture and supplemented by default-width cases. Python-only imported-font regressions also pass without Node on PATH; no backend CI dependency was added.

### Final automated and server evidence

The complete current-code sequence passes. The docs-only handoff triggers the same sequence again; final logs and manifests are the decisive post-commit receipts. Python resolves through repository `.venv/bin` PATH for the exact requested commands.

| Command | Result | Evidence under results-private/round-5 |
| --- | --- | --- |
| `python -m unittest discover -s tests -v`, twice | 290 tests each, zero failures/errors/skips | `final-python-{1,2}.log` |
| `npm test --prefix web`, twice | 101 tests each, zero failures/cancellations/skips | `final-web-{1,2}.log` |
| `python ops/content_guard.py --all` | exit 0 | `final-guard.log` |
| `ruff check pp run_v2.py tests/v2` | exit 0 | `final-ruff.log` |
| `npm run format:check --prefix web` | pinned Prettier 3.6.2, exit 0 | `final-prettier.log`; `final-new-tests-prettier.log` checks 20 newly added test/fixture files |
| `npm run build --prefix web`, twice | all 170 output files have identical SHA-256 manifests, including BUILD.json | `final-build-{1,2}.log`, `final-build-{1,2}-manifest.json` |

Coverage: Python's 284/321 new executable lines are covered (88.47%), including native child workers. JavaScript's 438/443 added nonblank lines are covered (98.87%) by union of process-local V8 whole-line ranges using UTF16 offsets and matching source URLs. These are changed-line metrics, not whole-application or branch coverage. Raw reports and explicit mapping: `python-coverage.json`, `coverage-web.log`, `coverage-proof.json`, `coverage-proof.py`.

Current actual `run_v2.py` server used a fresh temporary synthetic root on free port 55409. First-start settings creation was exercised with binding intercepted before choosing that port; port 8000 was never bound. Template is generated from `ops.fixtures.make_semantic` plus bare vertical/horizontal spill sheets. HTTP signup, category/copy creation, Inputs!B3=24, dependent Calc!B3 numeric/formula verification, default-width metadata with no override, row growth 1->2, column growth 2->3 and downloaded workbook read with openpyxl all pass. Export preserves inputs, formulas and fullCalcOnLoad. Evidence: `runtime.json`, `last-http-acceptance.log`, `real-server-http-proof.json`. This HTTP check does not verify browser rendering or WebSocket delivery.

The earlier actual browser flow in TaskSpace 41 passed signup/category/copy/edit, dependent WebSocket update, visible 0.583333, download/openpyxl and zero CSP violations on port 64097. It predates final review corrections and used an explicit 64px width, so it is historical evidence only. Receipts: `pre-followup-browser-proof.json`, `pre-followup-download-proof.json`, `pre-followup-calc-b3.png`. No final default-width visual proof is claimed.

### Review and security

Exact `claude-opus-5-5` reviews returned SHIP after all concrete findings were repaired. Final tiny follow-up verifies sole modelUsage key and canonicalModel exactly `claude-opus-5-5`, with no error. Independent exact `gpt-5.6-sol` xhigh reviews likewise returned SHIP; final incremental review used the supplied diff. Both are source reviews, not test or browser execution. Receipts: `opus-sealed-review.json`, `opus-last-review.json`, `planner-sealed-review.txt`, `planner-last-review.txt`, `review-proof.json`. The implementation route transparently escalated from the two failed executor lanes to Sol; no hidden model substitution.

`security-auditor` staged secret scans passed before focused code commits and folded corrections. Current explicit-source gitleaks scans 36 tracked application source files and reports zero leaks. Scoped Semgrep p/python, p/javascript and p/owasp-top-ten reports zero findings/errors. Locked active Python pip-audit and web npm audit report zero known vulnerabilities. Evidence: `*-staged-security.log`, `final-secrets.json`, `final-secret-scope.json`, `final-sast.json`, `pip-audit.log`, `npm-audit.json`. This establishes the stated current-source and active-dependency scope, not optional deep-history or legacy scanner clearance.

### Remaining gate and handoff

`BROWSER-TASKSPACE-DELETED`: the controller finished TaskSpace 41 before final review corrections, and reclaim returned `task space not found: 41`. [Ego browser skill](/Users/gopalmacbook/.claude/skills/ego-browser/SKILL.md:70) requires exactly one space for the goal, says "Never use a new TaskSpace to recover", and requires asking when the existing space cannot continue. No replacement or alternate browser was opened. Only explicit permission for one replacement TaskSpace can close the current-code default-width visual/WebSocket/download check. Lesson: `BROWSER-FINAL-GATE`; spill lesson: `SPILL-UNDO-INPUT-LINEAGE`.

Next action: approve one replacement Ego TaskSpace for the prepared final synthetic browser check. Retain it through every gate and finish once after success. Only owned server processes may be stopped. No push, PR, merge or deployment is authorized. Physical Windows batch, Microsoft Excel/VBA recalculation, real-workbook parity, office load and P5 acceptance remain separate.

## Earlier integration history

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


## Round 3

Date: 2026-10-11. User GO-TOKEN approved all three fixes, test-first local commits and synthetic acceptance on `feat/v2-integration`. Status: BLOCKED at the required fresh-server browser check. Code, reviews and automated acceptance are complete; real browser acceptance is not claimed. No push, dependency upgrade or company workbook access occurred.

### Changes and test-first evidence

| Change | Result | Commit |
| --- | --- | --- |
| Zero-config office startup | Server startup exclusively creates missing settings with `0.0.0.0`, port 8000 and demo mode true. Existing settings are preserved; import-only mode and the pure config loader retain conservative behavior. Empty origins derive canonical loopback/hostname, usable local IPv4 and specific-bind origins; explicit lists replace them and matching stays exact. Printed URLs must be assigned, reachable through the effective bind and allowed, preferring 10.x. Docs now require no manual JSON editing. | `80791ca`, `318515a` |
| Visible editable inputs | Only unfilled `k == "i"` cells receive #EAF2FB. Template fills retain their style. The default-on Highlight inputs toolbar toggle repaints loaded cells and persists per origin; both storage access and writes are guarded. Dist is rebuilt. | `067e965` |
| Actionable first run | The empty main panel contains an enabled primary Create your first category button, dispatched to the existing category modal. The sidebar action remains available with unique identifiers. | `067e965` |

Full logs are under ignored `results-private/round-3/`; none are committed. `planner-python-red-genuine.log` records 18 tests with seven intended failures for address provenance, first-start settings and share ordering. `planner-web-red.log` has six behavioral failures for snapshot/paint, toolbar, category action and unavailable storage plus two listener EPERM failures, which are environmental and not regression evidence. Focused corrections passed. Initial mock-target and wrong-interpreter errors are preserved but are not counted as valid reds.

Review-discovered regressions were also reproduced before correction. `review-fix-python-red-genuine.log` has eight expected assertions for stale DNS, bind reachability, browser port 80 canonicalization and IPv6 brackets. `review-fix-final-{safe-loader,decode,runner}-red.log` covers conservative pure defaults, undecodable interface output, route-fallback share printing and IPv6 wildcard suppression. `provenance-fix-runner-red.log` rejects stale numeric DNS when interfaces are absent and requires the UDP socket's actual local source; `provenance-fix-alias-red.log` covers mixed-case hostname binds. DNS answers now serve only as route hints; numeric fallback origins come from confirmed local sockets. Final focused Python 25/25 and web 22/22 pass in `provenance-fix-{python,web}-green.log`.

The routing lane used exact `gpt-5.6-luna` low, then `gpt-5.6-terra`, then controller-directed `gpt-5.6-sol` xhigh takeover after two inadequate implementation/verification attempts. Logs: `executor.jsonl`, `repair.jsonl`, `planner-repair.jsonl` and subsequent repair handoffs. No model was silently substituted. Lesson `INTERFACE-ADDRESS-BOUNDARY` records the trusted-address boundary. Worker socket-permission failures are preserved and are not called green.

### Final verification and security

Controller full Python command: `.venv/bin/python -m unittest discover -s tests -v`, 268/268 twice. Web command: `npm test --prefix web`, 53/53 twice. Both have zero failures, errors or skipped tests. Recovery files: `final-controller-python-{1,2}.log`, `final-controller-web-{1,2}.log`, with matching exit files all 0. Code was frozen during these runs and no source change followed them.

`ruff check pp/config.py run_v2.py tests/v2/test_config.py tests/v2/test_run_v2.py` and pinned `npm run format:check --prefix web` pass. `python ops/content_guard.py --all` exits 0. Two `npm run build --prefix web` runs yield identical hashes for every one of 170 dist files, including BUILD.json: `final-build-{1,2}.log`, `final-build-{1,2}.sha256.json`, `final-build-proof.json`. The final planning-only handoff repeats content guard. A stdlib trace measurement of newly executable Python lines covers 106/112 (94.64%) in pp/config.py and 53/57 (92.98%) in run_v2.py; `new-python-coverage.json` records the uncovered lines. This is focused new-line Python coverage, not a whole-app or browser coverage claim.

Final exact Opus review returns SHIP with no blocking finding. `opus-final-4-review.json` verifies both the sole modelUsage key and its canonicalModel as `claude-opus-5-5`, with is_error false and CLI exit 0. The installed review wrapper hardcodes an older model, so the exact-model direct Claude CLI was used with no tools or MCPs and a read-only diff payload. Independent final `gpt-5.6-sol` xhigh review returns SHIP, with fresh contract probes for numeric provenance, aliases, origin/bind filtering and IPv6 behavior: `planner-final-4-review.txt`. Opus's two unverified UI call sites were checked in current source: snapshot and live painting supply template style and toggle; the panel delegates to the same categoryModal used by the sidebar.

`security-auditor` ran. Every code commit had a clean staged secret scan and normal hooks. Locked active Python dependencies and active web npm audit are clean (`pip-audit.log`, `npm-audit.json`, both exit 0). Final scoped Semgrep `p/python`, `p/javascript`, `p/owasp-top-ten` reports no findings (`semgrep-final-4.json`, exit 0). The broad audit exits 1 with 34 retained pre-existing SQL/XML/CI/generated-vendor findings and seven unused prototype Vite advisories, including two HIGH. No repository-wide clearance is claimed; dependencies were unchanged. No commit contains a co-author trailer.

### Browser acceptance blocker and checkpoint

Named blocker: PORT-8000-OCCUPIED. The required `http://127.0.0.1:8000` listener belongs to unrelated CitiPoints PID 99858, launched by `uv run uvicorn citipoints_api.main:app --host 127.0.0.1 --port 8000` from the separate career-change CitiPoints API workspace. Its `/docs` returned HTTP 200. Owner evidence is in `port-8000-owner.log`. The service has not been stopped. Permission to temporarily stop and restore it was requested because that action affects an unrelated service outside the authorized worktree; no response has been assumed.

A fresh temporary runtime has been prepared with only a copied current `run_v2.py` and synthetic `template/`; settings.json is absent. Its location and initial-entry assertion are in `runtime.json`. Startup has not run, so created live settings and printed live share URLs remain unverified. One Ego TaskSpace 39 exists; it has only been used to verify actual browser URL canonicalization (`canonical-origin-browser.json`). No application signup or category/copy/edit/download flow has occurred. It remains reserved for resumption, with `finish({keep: []})` pending after successful acceptance.

The prepared ignored harness `browser.mjs` reuses the integration browser flow and checks both loopback and the printed LAN URL in that one TaskSpace: unique signup, primary first-category action, copy creation, input edit with dependent WebSocket formula update, preference persistence after reload, download and openpyxl verification. `prepare-runtime.py` and `start-runtime.py` are saved beside it. Resume after G frees port 8000 or authorizes stopping/restoring CitiPoints, run the requested actual browser flow, then update this evidence and STATE. The fresh-runtime browser gate and P5/office acceptance are not complete.

## Round 4: display fidelity, 2026-10-11

### Outcome and scope

Completed locally on `feat/v2-integration`, starting and ending at HEAD `fbd67a1`, with an uncommitted reviewable diff. Synthetic fixtures only. No dependency changes, company workbook access, commit, push, PR, merge or deployment. Template formulas remain server-owned and locked.

PATCH `web/src/model.mjs`, `web/src/grid-data.mjs`, `web/src/grid.mjs`, `web/test/grid.test.mjs`, `web/test/operations.test.mjs`; rebuild `static/dist`. General uses at most 11 characters without a known width and a documented conservative 12px-per-character floor budget when the column width is known. It counts signs, fits integers and rounding carry, preserves nonzero values, and falls back to uppercase signed scientific notation at >=1e11, <1e-9, or when fixed decimals cannot fit. Hashes represent an unfit narrow value. Merged anchors sum their column widths in both initial snapshots and live repainting. Raw numeric values stay numeric. Non-General preview formatting delegates to the already pinned Univer formatter.

The synthetic Calc fixture asserts literal displays for General, 0, 0.00, #,##0, #,##0.00, 0%, 0.0%, $#,##0.00, [$$-409]#,##0.00, m/d/yy, mm/dd/yyyy, d-mmm, negative red parentheses and @. It checks actual emitted workbook styles through snapshot and PlanningGrid.paint, red color, exact raw-value equality, width/sign/magnitude/carry boundaries and merged spans. The existing input-fill and toggle assertions remain protected. Literal formatting references are Microsoft's [General guidance](https://support.microsoft.com/en-us/excel/reset-a-number-to-the-general-format), [custom number-format guidance](https://support.microsoft.com/en-us/excel/review-guidelines-for-customizing-a-number-format) and [date-format guidance](https://support.microsoft.com/en-us/excel/format-a-date-the-way-you-want-in-excel).

### Red then green and reviews

All complete logs are in ignored `results-private/round-4/`. `controller-display-red.log` reproduces the negative integer trailing dot, wrong m/d/yy layout and untrimmed General value. `planner-red.log` protects the 11-character ceiling, narrow nonzero values, sign and rounding overflow. `width-floor-red.log` reproduces the 40px budget defect before the one-line floor fix; `width-floor-green.log` passes after it. `merge-red.log` reproduces ### for a two-column 80px merged span; `merge-green.log` shows literal 123456 in snapshot and paint. Focused final formatter/grid tests pass 25/25.

Behavioral mutations of the actual style mapping, disabled-highlight guard and toggle repaint fail the strengthened tests: `grid-style-mutation-red.log`, `highlight-guard-mutation-red.log`, `toggle-mutation-red.log`. Each source file was restored byte-for-byte and focused checks passed. Caller evidence in `raw-edit-copy-caller-proof.log` confirms that formula/edit seeds and clipboard paths use raw values, while formatNumber is display-only.

Final `opus-last-review.json` returns SHIP and verifies the sole modelUsage key and canonicalModel as exactly `claude-opus-5-5`, with is_error false and CLI exit 0. Its concrete medium findings on emitted-style coverage, highlight coverage and merge width were resolved and verified. Its low raw-value assertion suggestion was applied afterward with exact equality, followed by fresh web checks and deterministic builds. Independent `planner-last-review.txt` from exact `gpt-5.6-sol` xhigh returns SHIP. `review-proof.json` records model and verdict validation. Width estimation and display-only history rounding are disclosed limits; no blocking finding remains.

### Final controller verification

`final-python-{1,2}.log`: `.venv/bin/python -m unittest discover -s tests -v`, 269/269 twice, zero failures or skips. `final-web-{1,2}.log`: `npm test --prefix web`, 58/58 twice, zero failures or skips. The controller's permitted loopback binding closes the two EPERM failures seen only inside worker sandboxes. Pinned Prettier and `python ops/content_guard.py --all` exit 0. Newly generated chunks are included in the content guard through intent-to-add Git entries.

`final-build-{1,2}.log` and SHA-256 manifests show two successful builds with identical bytes for all 170 dist files, including BUILD.json: `final-build-proof.json`. Authored source/tests/planning pass git diff --check. The broad diff's seven vendor help-string whitespace diagnostics are unchanged baseline lines, verified in `vendor-whitespace-proof.json`; generated strings were not manually rewritten.

The `security-auditor` skill ran. Current source secret and scoped JavaScript/OWASP SAST checks pass: `final-secrets.log`, `final-sast.json`, both exit 0. Active web npm/OSV audits and pinned Python lock audit pass: `npm-audit.json`, `osv-web.log`, `pip-lock-audit-retry.log`. The initial scope-only OSV invocation found no package source and was corrected to the lockfile; Python 3.14 ensurepip failed, then the fully pinned no-pip/no-deps audit succeeded. No repository-wide or unrelated prototype security clearance is claimed.

### Real browser acceptance and cleanup

A fresh temporary root contains only the synthetic semantic template and copied current run_v2.py before initialization. `first-start.log` records main's first-start settings creation and runtime initialization with the default bind intercepted before any port-8000 binding. After that first start, runtime settings.json was updated to free port 61184. The actual unmodified server then started on 0.0.0.0:61184 and printed `Share with your team: http://192.168.1.5:61184` (`server.log`, `runtime.json`).

One Ego TaskSpace 40 opened that printed LAN URL through `ego-browser nodejs`, signed up a synthetic planner, created a category/copy, and used the native formula input to set Inputs!B3 to 24. The server returned Calc!B3 raw value 0.5833333333333333 and original formula =IFERROR(1-Inputs!B2/Inputs!B3,""). Actual canvas fillText captured 0.583333, and the controller inspected the screenshot showing selected locked cell B3 with exactly 0.583333. `browser-proof.json`, `browser-actions.log`, `calc-b3.cleaned.png` are the evidence. CSP violations: zero.

`browser-finish.log` confirms awaited `finish({keep: []})` once for TaskSpace 40. The task's own temporary server PID 27788 was verified against its temporary cwd and stopped. Port 8000 remained owned by unrelated PID 99858 before and after (`port-8000-{before,after}.log`); it was never stopped or rebound. Screenshot metadata cleaning reports no provenance chunks present and an exiftool strip pass.

This verifies the requested Mac browser display and synthetic format strings. Conservative width estimates are not exact desktop Excel glyph metrics; hidden columns inside merged spans still count toward the estimate. Real workbook values, physical office devices, Windows/Excel/LAN/load, the broader Round 3 export flow and P5 acceptance remain outside this round.

Final handoff checks: STATE headings match the original headings exactly; repository document tests pass 4/4; handoff content guard and authored-source whitespace check exit 0. `acceptance.json` consolidates final evidence. Installed getPatternPreview uses `throws: false` and catches formatter errors, verified in `formatter-throw-mode.log`.
