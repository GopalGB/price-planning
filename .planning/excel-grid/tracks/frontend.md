# Track B frontend handoff

Date: 2026-10-10. Branch: `feat/v2-frontend`. Status: DONE_WITH_CONCERNS for local synthetic frontend work; office acceptance remains pending. Implementation commit: `1061a45`; committed bundle and browser evidence: `6baab71`. No push performed.

[HIGH] Built the Vite app under `web/` and committed `static/dist/`. Reused the API contract, proto visual language and resolved Univer gate imports; backend paths `pp/`, `app/`, and `ops/` were not changed. No private workbook paths or external Office files were accessed. All fixtures and screenshots are synthetic.

The UI includes login/signup, category membership, Copies, create/fork/share dialogs, named roles, Excel green app bar, toolbar, external formula bar, sheet tabs, status summary and history. The active sheet applies metadata and loads three 256-row blocks. Input-only edit, Delete, paste and fill send atomic versioned server operations with optimistic rollback. Undo/redo use accepted server cell versions, preserve remote conflicts and reset history on missing live confirmation. Formula text stays outside Univer; calculation startup and later engine commands are blocked. Live ops/calcs, calculation state, reconnect/resync and validated presence are wired.

The loopback mock implements auth, categories/members, copies/fork/shares, workbook metadata, streamed rows, cell detail, history, ops, WS and synthetic download. It computes dependent values server-side for 4088x312 and 100000x17 sheets. Demo users automatically join Category A. Its generated `.xlsm` contains demo columns A-D and recalculation flags, with no original template or VBA. Production template-preserving export remains the backend's responsibility.

## Verification

| Check | Result |
|---|---|
| `npm test --prefix web` | 33 passed, 0 failed, two consecutive greens |
| Node coverage over state, grid helpers and views | 97.12% lines; 65.84% branches; DOM/Univer adapter excluded |
| `.venv/bin/python -m unittest discover -s tests` | 64 passed, 0 failed, consecutive greens |
| `python3 ops/content_guard.py --all` | Exit 0 after staging committed bundle/evidence |
| `npm audit --prefix web --json` | Zero vulnerabilities |
| Semgrep JavaScript + OWASP rules on web source/mock/scripts | Zero findings, zero errors |
| `security-auditor --staged` | Gitleaks clean over 11.68 MB of staged source and bundle |
| Exact `claude-opus-5-5` review | Returned identity verified; SHIP, no remaining MEDIUM/HIGH/CRITICAL |
| `ego-browser nodejs < web/test/ego.mjs` | Exit 0; one TaskSpace (35), `finish({keep: []})` completed |
| `npm run build --prefix web` twice | Identical source/bundle hashes; largest asset 1,460,991 bytes |

The browser passed signup, copy creation, native input edit and dependent WS update, locked-cell refusal, typed formula rejection, server undo, native paste, native fill, history, sharing and viewer refusal. Client calculation count and client formula count were zero. At row 99,000 the grid held 3 cached blocks and 768 populated rows; sampled heap was 93.753361 MB. This is one sampled heap measurement, not a peak bound.

Five screenshots in `frontend-shots/` were visually inspected: Copies home, initial workspace, history, long sheet and viewer. They remain local because the existing content guard rejects PNG binaries. No guard was bypassed. `frontend-validation.json` commits their SHA-256 checksums and sizes. PNG inspection found no C2PA or AI metadata.

`BUILD.json` pins Univer 1.0.4 and records dependency licenses. Source SHA-256: `4ef0833ecc6afa47880ecc112997b27492f20bd2ae4bbdecedfff8b14e6b6119`. Bundle SHA-256: `9a06b98eeca33881b888da8c4100642686413ce17e64912e0c0d0ba633505239`. Vite 6.4.4 and ws 8.22.0 are exact devDependency pins; earlier vulnerable dev versions were replaced. No Python dependencies were added.

Full original test/build/review/scan output remains in ignored `frontend-*.log` files in this worktree. Committed aggregates: `frontend-ego.json`, `frontend-review.json`, `frontend-validation.json`, `frontend-build-first.json`. Developer commands and mock limitations are in `web/README.md`.

## Not done or not verified

- Real backend integration, real workbook formula/VBA/export fidelity, Windows/Edge/Chrome office LAN acceptance, desktop Excel clipboard behavior, scrolling FPS, latency targets and peak heap were not verified.
- The full section 5 interaction matrix was not browser-tested. Offscreen copy/find, multi-range shortcuts, data validation controls, named presence with a second client, forced WS reconnect/resync and live role downgrade still need integrated acceptance. Their implementation is present; the core synthetic browser checks above are the verified subset.
- Reviewer LOW notes remain: history refresh briefly replaces the formula display while details reload; selecting a sparse blank cell can retain the previous history panel content. No MEDIUM or higher finding remains.
- Screenshots are local only. No push, PR, merge or deployment was performed, as instructed.

Next: backend owner integrates the committed `static/dist/` and runs the remaining office acceptance checks. The Mac local Track B report does not certify full P3 office release readiness.

## Review fixes

Requested 2026-10-10; verification crossed midnight into 2026-10-11. Branch: `feat/v2-frontend`. Source commit: `5d095431c07530b3b6f5e4f2171c144fcae319b7`. Deterministic bundle commit: `20f3c64b8565f8367871f00b06692964b9670cf6`. This section supersedes the earlier frontend defect/readiness claims for the reviewed behavior.

[HIGH] Fixed the five defects in `~/PricePlanning-work/logs/astra-frontend.md`. Undo rebases only an uninterrupted local version chain; the second undo in `0 -> local 1 -> remote 99 -> local 2 -> undo -> undo` makes a historical CAS request, gets 409, leaves 99 and reports the named editor. Overlapping cells roll back atomically, and old/duplicate live events cannot poison undo history. Partial snapshots certify their own cached block; formula patches use cell revisions and the save indicator waits for every cached block. Copy, go-to, block loading and resync reject stale sheet generations and operation tokens. Shared with me loads all accessible copies through unfiltered `/api/copies`, including a private Category B copy shared by email to an A-only recipient.

[HIGH] Live operations use authoritative server `at`. Missing legacy timestamps become null and trigger version/generation-guarded detail hydration. Bulk events above eight eligible cells reload cached blocks sequentially, avoiding thousands of concurrent detail requests. The mock sends the same ISO time as cell detail/history. API.md documents the timestamp and input detail metadata contract. **The real backend must add `at` at integration; the backend owner will do this.** No backend implementation was changed.

Each row below was tested RED before its implementation and GREEN afterward. Commands ran from `web/`: `node --test --test-name-pattern='<pattern>' <files>`. Counts are passed/total; RED failures were deterministic behavioral mismatches. The already-supported unfiltered mock endpoint passed in the share RED run while the missing UI loading path failed.

| Defect | Pattern and test files | RED | GREEN | Full original output |
|---|---|---|---|---|
| Undo / remote edit | `older undo\|multi-cell undo\|duplicate and older`, `test/client.test.mjs` | 0/3 | 3/3 | `/tmp/priceplanning-fix1-{red,green}.log` |
| Partial snapshot formula / Saved | `older formula calculation`, `test/grid.test.mjs` | 0/1 | 1/1 | `/tmp/priceplanning-fix2-{red,green}.log` |
| Sheet switch during copy / go-to | `deferred multi-block copy\|deferred go-to`, `test/grid.test.mjs` | 0/2 | 2/2 | `/tmp/priceplanning-fix3-{red,green}.log` |
| Cross-category named share | `shared navigation\|named private share`, `test/views.test.mjs test/mock.test.mjs` | 1/2 | 2/2 | `/tmp/priceplanning-fix4-{red,green}.log` |
| Live timestamp / metadata hydration | `operation timestamp\|editor metadata\|missing operation timestamp\|mock API enforces`, `test/operations.test.mjs test/views.test.mjs test/grid.test.mjs test/mock.test.mjs` | 0/4 | 4/4 | `/tmp/priceplanning-fix5-{red,green}.log` |
| Opus bulk-refetch finding | `bulk missing timestamps`, `test/grid.test.mjs` | 0/1 | 1/1 | `/tmp/priceplanning-fix5-bulk-{red,green}.log` |

The approved GSD execution completed one verified fix at a time. Luna execution was escalated to Terra after repeated patch failures; the planner lane then took over when the 409 contract remained unmet. The final independent planner review returned SHIP. Exact `claude-opus-5-5` returned SHIP with identity verified in `modelUsage`, after its MEDIUM bulk-request finding was fixed. The installed review wrapper rejects this required model, so the supported CLI was invoked directly with the exact identifier. Review evidence: `/tmp/priceplanning-final-sol-review.md`, `/tmp/priceplanning-fixes-opus-review.json`, `/tmp/priceplanning-fixes-opus-final.json`.

| Acceptance check, executed in this order | Result |
|---|---|
| `npm test --prefix web`, twice | 47 passed, 0 failed on each run |
| `.venv/bin/python -m unittest discover -s tests`, twice | 64 tests OK on each run |
| `python3 ops/content_guard.py --all` | Exit 0 |
| `npm run build --prefix web`, twice | Exit 0; identical BUILD.json and bundle SHA-256 |

Bundle SHA-256: `182e5cd9bff400940dfedfb1e61460d998f2e92823393711b9b67b380bccd4be`. Source SHA-256: `91712bacc58dee6b5547b47fb74f2eb1ecef75fac51e29b08e69ff087e680712`. Full ordered acceptance output: `/tmp/priceplanning-accept-{npm,python,build}-{1,2}.log` and `/tmp/priceplanning-accept-guard.log`. After the documentation commit, the same ordered checks are repeated against the final committed state into `/tmp/priceplanning-final-*`; no further commit is made after that pass.

Security: npm audit reported zero vulnerabilities; Semgrep JavaScript/OWASP scan returned zero findings and zero errors; `security-auditor --staged` reported no secrets for source and bundle commits. Full output: `/tmp/priceplanning-fixes-audit.json`, `/tmp/priceplanning-fixes-sast-final.json`, `/tmp/priceplanning-fixes-secrets-{code,bundle}.log`. No dependencies were changed. Vite's existing large-chunk and ignored third-party module-directive warnings remain.

One preliminary Python repeat timed out at the existing launcher test's five-second login request (`test_first_admin_restart_and_real_waitress_launcher`). Full trace: `/tmp/priceplanning-pre-python-2.log`. No out-of-scope backend/test changes were made; two consecutive reruns passed 64/64, followed by the ordered acceptance pair above. The final post-commit pair is also required before handoff.

Nonblocking review notes: an advancing 409 can leave time unknown until a row refresh; a remote change seen only by resync can make undo refuse locally without a named 409; legacy metadata hydration can leave time unknown when a snapshot remains behind calculation or a cell is only temporarily hydrated. These paths preserve data and never display a previous or invented timestamp. The real backend's required `at` removes the legacy live-message path. Office/backend integration and fresh browser/visual acceptance were not performed during these fixes.

Only `web/`, `static/dist/`, API.md, STATE.md and this report changed. No `pp/`, `app/`, or `ops/` edits, real company workbook access, push, PR, merge or deployment. Existing untracked screenshots were preserved. Raw logs remain outside commits. Required STATE headings remain intact.
