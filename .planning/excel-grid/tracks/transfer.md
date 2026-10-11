# Transfer track

Date: 2026-10-11. Branch: feat/v2-transfer. Base: f74d131.
Approval: owner GO-TOKEN approved, proceed, with the full download/upload specification and final checks.

## Goal and acceptance

Non-technical Excel users download the whole workbook or a selected sheet subset as .xlsx (default) or .xlsm, preserving the original template package. Every output reopens with openpyxl and passes structural OOXML relationship/content-type validation. Subset removal cleans sheet-dependent parts and freezes cross-sheet formulas to computed values while preserving number formats. Exports carry template hash, copy id and input_revision in a small custom document property.

Users upload Excel files by content, preview changed inputs, conflicts with author/time and ignored edits, then update as one undoable operation or create a new copy in an accessible category. Missing provenance falls back to sheet names and formula-map fingerprint. Different templates are refused. Uploaded macros are neither executed nor retained. Stream size-limited uploads (100 MB default), reject archive bombs and unsafe XML. Enforce editor/owner update authorization and category-access creation authorization.

## Approved phases

- [x] A1: Synthetic red/green package export tests, .xlsx conversion, provenance, subset cleanup and output validation. Reuse pp/export/xlsm.py; create focused transfer helpers only where no existing helper fits. Include per-removal-case tests and a synthetic 55 MB performance measurement against 30 seconds/500 MB.
- [x] A2: Download dialog, default all sheets including hidden, format default .xlsx, select all/none. Preserve Download .xlsm. Server route and web tests, one focused conventional download commit.
- [x] B1: Synthetic red/green upload parsing, content/size/archive/XML safety, provenance/structure identification, input-only preview, conflict author/time and ignored-edit reasons.
- [x] B2: Preview-confirmed update/new-copy endpoints using the existing operation pipeline, one undoable operation, upload filename source/uploader attribution, selected conflicts and authorization. Upload UI on home and inside copy, regression tests and one focused upload commit.
- [ ] C: Exact claude-opus-5-5 review with returned identity, planner review, security-auditor; resolve blocking findings test-first.
- [ ] D: Refresh after review corrections: Python twice, web twice, content guard 0, ruff, pinned Prettier, two identical builds. One real-server Ego TaskSpace on a free port excluding 8000/8765/8766: all-sheet and two-sheet downloads/openpyxl, edit/download/upload preview/update/undo/new-copy. Finish keep: []. Preserve STATE headings and hand off evidence.

## Ownership and boundaries

Transfer owns download/upload package helpers, routes/config strictly needed for transfer, transfer UI/tests and build artifacts. Engine/calculation pipeline is read-only. Formatting/drawing rendering and pivot/slicer dashboards belong to other tracks. Preserve their APIs and edits. Synthetic fixtures only, no private workbook paths, no Docker, no dependency upgrades without lock/wheels together, no push/main/merge/deploy. Controller owns Git, state, review, runtime/browser and evidence. One routed executor owns transfer implementation and regression tests.

## Evidence and current checkpoint

Latest corrections: `e6c8a89` fixes H1 self-closing cells (red, regression green twice, controller export18/18). `baceaa6` fixes H3 stale upload merging (red, final regressions4/4 twice, controller upload33/33) and preserves the existing change/revert conflict contract. `3c16953` bounds upload metadata and detaches XML nodes (controller upload39/39). `348f0a8` bounds worker concurrency/XML nesting/tokens/prepared bytes and removes cancelled/startup artifacts (controller upload/routes/resource55/55). All new boundaries have red evidence and focused greens twice. Changed-file Ruff, staged secret scans and commit content guards passed. Medium/low disposition and exact clean re-review remain in progress. All earlier full-suite, browser and build evidence below is historical until refreshed against corrected source.

Local feature commits: `d4dddd4` download, `2ead2d7` upload. Both have matching built assets and passed staged secret/content guards without bypassing hooks. Full source remains identical to the final acceptance hash receipt. No dependency changes or outbound publication.

Final Python365/365 twice, web108/108 twice, guard0, Prettier3.6.2 pass, changed Python Ruff pass, full-tree141 Ruff findings equal the baseline with zero introduced findings. Focused threaded coverage132/132, transfer92% and upload91%. Full170-file build trees match byte-for-byte. Detailed commands, logs and receipts are in ignored `results-private/transfer/`; final Python/web/build/guard/Ruff evidence uses the `final-*` prefix.

Synthetic57.97MiB XML transfer:7.786s/257.73MiB and7.835s/189.27MiB in final full-suite runs. Earlier timing failure43.006s and provisional runs are retained honestly. No real-template timing claim.

Real-server browser port63630, single TaskSpace44: all-sheet and two-sheet downloads, openpyxl/OOXML validation, hidden-sheet preservation, formula freeze/number format, edited-file preview, one-operation update/undo, isolated new copy, copies-page update/undo and unchanged-result messaging passed. Read-only database receipts confirm2 copies/5 operations. Browser finished keep:[]; runtime stopped and port closed. See `browser-acceptance.json`, `browser-finish.json`, download/openpyxl receipts and rendered snapshots.

Independent planner findings were corrected test-first, including shared/array/named/table dependencies, scoped vendor namespaces, immutable prepared downloads, trusted large-upload inverse lineage and bounded formula-aware parsing/preview storage. Final residual closure has no blocking defects. Final SAST5 findings/0 scanner errors were independently dispositioned as protected parsing or fixed-column parameterized SQL. Dependency audits report0 vulnerabilities; staged secret scans report0 leaks. No scanner suppression.

## Remaining gate

2026-10-11 19:51 IST: exact `claude-opus-5-5` review succeeded with verified returned modelUsage/canonicalModel and session `8ba75123-b41f-46db-8877-57a615c43987`. The review is not clean:3 HIGH hypotheses and all medium/low findings are assigned to one exact planner takeover worker for reproduction, focused fixes and disposition. Full detailed review is preserved in ignored `opus-review-resume-complete.md`. Owner's uploaded-value default for true conflicts remains binding. Commit each corrected item before continuing. Latest resume verification: transfer71/71, full Python365/365 once; other full runs had only XML timing failures48.374s/40.263s under host load139.54. Focused benchmark passed11.331s and a full-run benchmark passed9.753s. No test relaxation; preserve failure logs. Changed Ruff/Prettier3.6.2/guard passed; full Ruff141 unchanged. Source and browser-tested bundle hashes still match until corrections start.

Resume resource blocker (2026-10-11): RESOURCE_DISK_LOW. Initial `df -g /` reported 3 GB available, below the owner's 4 GB stop threshold. Swap used=9912.44M is the resume baseline. No tests, builds, browser, worker or exact Claude retry ran this session. Initial Git status had no tracked changes; only untracked `.venv` and `web/node_modules`, with HEAD `ae79f33`. No cleanup was attempted after the stop threshold triggered. Previous acceptance remains historical evidence.

Historical quota failure: exact `claude-opus-5-5` previously returned is_error=true and empty modelUsage. The subsequent successful exact review above supersedes that blocker. Phase C remains unchecked for unresolved findings and clean re-review. Existing Ruff141 baseline findings remain a disclosed concern.

## Resumption

Read STATE and this track. Latest owner limits supersede older thresholds: stop below2500 MB free disk or swap above10411.31 MiB (resume baseline8875.31 +1536). Finish current exact-review corrections and disposition, then re-review and rerun affected acceptance. Preserve the approved scope, private-workbook prohibition, separate feature commits and source receipts. No push/merge/deploy authorization.
