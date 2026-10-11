# Approved fidelity track

Date: 2026-10-11. Branch: feat/v2-fidelity, base f74d131.
Approval: user GO-TOKEN approves the five ordered features and local acceptance.
Goal: import and render Excel template presentation using synthetic fixtures only.

- [x] 1. Conditional formatting: tests first, six rule types and differential styles; copy-current values through public engine boundary; row-block overrides, client rendering and edit refresh; isolated future calculation hook. Files: pp/template, pp/fidelity, pp/workbook_service.py, web/src, tests/v2, web/test, ops/fixtures. Success: literal expected styles and relative/cross-sheet formulas, no full-sheet fetch. Implementation and wide312-column row-block regression corrected; final Opus acceptance blocked by weekly quota.
- [x] 2. Drawings: synthetic anchored pictures and text shapes; macro button POST stub with access checks and clear 501. Files: importer/fidelity, server action route only, web, tests, fixtures. Success: anchored render and endpoint tests.
- [x] 3. Notes: row-block notes, red markers and hover text. Files: importer/fidelity, workbook service, web, tests, fixtures. Success: literal text and hover/marker tests.
- [ ] 4. Tables and AutoFilter: header styling, banding, filter controls and per-user local visibility; no sort. Files: importer/fidelity, workbook service, web, tests, fixtures. Success: filter isolation and bounded streaming tests.
- [ ] 5. Sheet views: tab colors, gridlines, zoom, hidden-sheet control, frozen panes; retain sizes, merges and protection. Files: importer/fidelity, workbook service, web, tests, fixtures. Success: metadata and renderer tests.
- [ ] 6. Review/security: exact claude-opus-5-5 review identity, planner review, resolve blocking findings, scoped security audit.
- [ ] 7. Acceptance: Python twice, web twice, content guard 0, ruff, pinned Prettier, two identical builds; real synthetic server on a free port excluding 8000/8765/8766; one Ego TaskSpace, screenshots per feature, finish keep: []; required STATE headings and tracks/fidelity.md updated.

Out of scope: pp/engine changes, calculation pipeline changes, real/private workbook access, dependency upgrades, sort, slicer interaction, action execution, push/PR/merge/deployment.
Controller owns plan, commits, independent reviews, browser and final evidence. One routed sequential executor owns scoped code/tests. Executor is not alone in the repository and must preserve others' edits.
Evidence directory: ignored results-private/fidelity. Each feature receives a focused conventional commit without co-author trailers or em-dashes.
