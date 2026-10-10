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
