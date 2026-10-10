# Approved v2 integration plan

Date: 2026-10-11. Branch: feat/v2-integration. User supplied approval and all-phase execution authorization.
Spec: excel-grid/SPEC.md, excel-grid/API.md and user integration checklist.
Goal: one coherent synthetic-tested v2 app using the merged tracks.

- [x] Export integration: retain Track C package, port Track A assertions, reject array children through input mask. Files: pp/server.py, pp/export.py, tests/v2/test_export.py, template importer if regression requires it. Success: focused red then green export and mask tests.
- [x] Configuration and launchers: pp.config owns v2 database resolution; backup shares flag > env > settings > default precedence; replace fake entry point with real import smoke. Files: pp/config.py, pp/backup.py, tests/test_backup.py, tests/test_launchers_v2.py, tests/v2/test_config.py. Preserve legacy app/runtime.py users. Success: focused red then green, real synthetic import succeeds.
- [x] Live/access contracts: op at equals persisted history time, replay consistent; test all-category named shares. Files: pp/ops.py, pp/live.py, tests/v2/test_ws.py, tests/v2/test_copies_api.py, excel-grid/API.md. Success: focused contract tests green.
- [x] Hygiene/readability: ignore runtime evidence, preserve reports, exact pinned prettier formats web/src/*.mjs without logic changes; rebuild static/dist. Success: web tests and content guard green.
- [x] Browser CSP: run real run_v2.py with ops.fixtures.make_semantic only, inspect console in Ego TaskSpace 37, permit only evidenced directives. Files: pp/server.py, security header tests, synthetic browser evidence script. Success: app edits and WS work without CSP violations.
- [x] Security and independent review: exact claude-opus-5-5 identity and critical/high resolution; planner review; security-auditor with scanner results reported. Success: review and audit evidence or named blockers.
- [x] Final acceptance and handoff: logical commits first, Python twice, web twice, content guard, deterministic build twice, real-server browser signup/create/edit/WS/locked refusal/undo/download/openpyxl; write tracks/integration.md and update STATE.md with required headings. Any later commit triggers final sequence rerun.

Out of scope: company workbook access, real data, push, PR, merge, deployment, Windows/Excel/office LAN acceptance, paid Strix run.
Evidence: ignored results-private/integration/. Preserve full outputs there.

## Review round 2, approved 2026-10-11

User GO-TOKEN approves STEP 0 and all A-E phases in this worktree on `feat/v2-integration`. Scope remains local, synthetic and without dependency changes or push.

- [x] STEP 0: merge `feat/v2-frontend` with `--no-ff`, choose frontend conflicted sources, preserve round 1 semantics, format/rebuild, retain STATE headings. Success: merge `88188de`, Python 242 tests and web 47 tests green before A-E.
- [x] A: XML-compatible input text and length validation, API edit/export implicit blank round trip. Files: pp/values.py, pp/export/xlsm.py, tests/v2. Validation failed as intended for five unsupported-text cases; focused four-test green. Track C implicit blank behavior already passed. Commit `7930a6b`.
- [x] B: persisted session revalidation within every mutation transaction after body and pipeline waits. Files: pp/auth.py, pp/server.py, mutation services and tests/v2. Behavioral mutation red: 10 expected failures across revoked routes, expiry and pipeline wait; auth module green 9/9. Full logs B-mutation-red-2.log and B-green-2.log.
- [x] C: coherent replay/rows read snapshots and monotonic WebSocket calculation state. Files: pp/db.py, pp/live.py, pp/workbook_service.py, tests/v2. Both query-boundary interleavings failed as intended, then passed. Evidence: C-red.log and C-green.log.
- [x] D: temporal engine results in Excel serial units, honoring template date1904. Files: pp/engine/formualizer.py, tests/v2. Both 1900/1904 tests failed as intended, then passed; literal DATE/TIME/datetime and leap-day boundary expectations. Evidence: D-red.log and D-green.log.
- [x] E: inherited fork input editor/time and history provenance; local edits clear origin markers. Files: pp/copies.py, pp/ops.py, pp/workbook_service.py, tests/v2. Alice/Bob attribution regression failed as intended; provenance pinned against later parent edits and local edit/undo markers pass. Evidence: E-red.log and E-green.log. Commit `79ea428`.
- [x] Verify and handoff: exact `claude-opus-5-5` identity and final SHIP, fresh planner A-E and incremental SHIP, security-auditor with retained broad findings disclosed, hooked merge/code commits, Python 251/251 twice, web 47/47 twice, content guard exit 0. Review-discovered escaping blocker reproduced red and fixed in `3340943` (26 export tests green). Review round 2 appended; STATE headings intact. Full evidence in results-private/review-round-2; no push.

Execution uses one model worker; controller handles authorized Git operations because linked-worktree metadata is outside the worker sandbox. Office/Excel/real-workbook/P5 acceptance remains outside this scope.

## Round 3, approved 2026-10-11

User GO-TOKEN authorizes all phases, test-first local commits on feat/v2-integration.
- [x] Office startup: create missing settings with LAN/demo defaults, derive exact local origins when empty, print share URLs with 10.x preference. Files: pp/config.py, run_v2.py, LAN helper if needed, tests/v2, docs/DEPLOY.md. Success: missing/empty/explicit/foreign regressions red then green.
- [x] Input visibility: tint unfilled inputs #EAF2FB, preserve explicit fills, default-on storage toggle guarded by try/catch. Files: web/src, web/test, static/dist. Success: style and preference regressions red then green.
- [x] First category: primary button in empty panel opens existing category modal. Files: web/src/views.mjs, web/test. Success: regression red then green.
- [x] Review/security: exact claude-opus-5-5 identity, planner review, resolve blocking findings, scoped audit, hooked local commits.
- [ ] Acceptance: Python twice, web twice, content guard, two identical builds; fresh temp runtime with synthetic template only and no settings; real browser signup on loopback and printed LAN URLs, category/copy/input/formula/export verified with openpyxl; one Ego TaskSpace finished keep: []; append Round 3 evidence and preserve STATE headings.

Out of scope: push, dependency changes, real workbook/data, office Windows/Excel/load acceptance.
Evidence: results-private/round-3/. Controller owns PLAN/STATE/report/browser/commits; executor owns scoped code/tests/docs/dist.

Code commits: `067e965`, `80791ca`, `318515a`. Final exact Opus 5.5 and planner reviews SHIP. Controller Python 268/268 twice, web 53/53 twice, content guard exit 0, 170 dist files byte-identical including BUILD.json. Acceptance remains unchecked: required fresh-server browser flow is blocked by unrelated CitiPoints occupying port 8000, pending permission to stop and restore it or owner freeing the port. One Ego TaskSpace 39 is reserved; no app browser acceptance is claimed.

## Round 4, approved 2026-10-11

GO-TOKEN authorizes all phases on feat/v2-integration at fbd67a1. Synthetic fixtures only; no push, dependency upgrades, real workbook or port-8000 intervention.

- [x] Reproduce display failures first with literal Excel expectations in a synthetic format-sweep sheet and grid display tests. Cover General signs, integers, width trimming, >=1e11 and <1e-9, all requested formats.
- [x] Patch existing formatting and grid render paths with the smallest shared display solution. Preserve raw numeric values, editing and server-owned formulas. Success: focused red then green evidence, full web green.
- [x] Review exact claude-opus-5-5 identity and independent gpt-5.6-sol xhigh; security-auditor and formatting checks. Resolve blocking findings.
- [x] Final verification: Python twice, web twice, content guard 0, two byte-identical deterministic builds. Fresh temp runtime first-start settings then configured free port; one Ego TaskSpace, printed LAN URL, Inputs!B3=24 and visible Calc!B3=0.583333; finish keep: [].
- [x] Append Round 4 evidence and preserve STATE headings with truthful test/browser/review status and next action.

Files: web/src/model.mjs, web/src/grid-data.mjs, web/src/grid.mjs as needed, web/test synthetic fixtures/tests, static/dist, .planning/PLAN.md, .planning/STATE.md, .planning/excel-grid/tracks/integration.md. Controller owns planning, runtime/browser/review/evidence; one exec worker owns web code/tests/dist. Logs: results-private/round-4/.

Acceptance evidence: Python 269/269 twice, web 58/58 twice, content guard 0, 170 dist files byte-identical including BUILD.json, exact Opus 5.5 and independent planner SHIP. TaskSpace 40 opened printed LAN URL on free port 61184; Inputs!B3=24 produced visible Calc!B3=0.583333. Finished keep: []; temporary server stopped; port 8000 untouched. See tracks/integration.md, Round 4.
