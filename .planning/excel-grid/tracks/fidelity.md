# Fidelity track

2026-10-11. feat/v2-fidelity starts at f74d131. User approved all five features and local acceptance.

## Scope
Synthetic fixtures only. No private workbook access. pp/engine and calculation pipeline remain read-only.

## Progress
Tables/AutoFilter resumed and controller-verified: Python36/36, web31/31, content guard0, Ruff and Prettier green. New resource guard: disk2500 MB, swap6814.62 MB. Phase4 local commit follows; review and final acceptance pending.

2026-10-11 19:33 IST resume stopped at the resource gate: df -g / available 3 GB, df -k / available 3323888 KiB, swap used 9912.44 MB baseline. Tables/AutoFilter changes and tests are present but uncommitted and unverified in this resume. No source edits, tests, commits, review calls or browser checks ran; checkpoint documentation only. Pending exact claude-opus-5-5 review is deferred by low disk, not a newly verified quota failure. TaskSpace 42 remains retained.

Wide CF block correction is green: scoped Python30/30, CF11/11; 50x312 and256x312 emit15600 and79872 expected styles without warnings, using one public engine load for TRUE(). Per-target volatile/named evaluation is preserved. Table work resumes next.

Notes Phase3 implementation passes scoped Python27/27 and final popup web23/23. Drawings are committed082fea3; table styles/filtering are next. All features remain unaccepted pending Opus and final whole-track checks.

CF locally committed f764aaa; drawing feature implementation passes relevant Python 32/32 and web 26/26. Final acceptance remains unaccepted pending exact Opus review and whole-track gates. Notes are next. Earlier checkpoints below are historical.

CF F1-F7 correction implementation is tested (full Python 302/302, web 105/105) but unaccepted. Fresh exact Opus 5.5 re-review is blocked by weekly quota, HTTP 429, with provider-reported reset 7:30pm (Asia/Calcutta). Controller continues approved local work; no substitute or paid lane. Final acceptance remains blocked.

Resume: remaining CF importer/evaluator findings have regression coverage. Relevant executor checks passed (Python 25/25, web 23/23); controller full Python 300/300 and full web 103/103 passed. Content guard 0, project Ruff and pinned Prettier green. Fresh Opus 5.5 returned FIX-FIRST (verified identity); F1-F7 are in a second regression-fix pass before any commit. This is the current checkpoint; low-disk details below are historical.

CF draft and regression fixes are present but not accepted or committed. Other four features remain pending. Exact Opus 5.5 returned FIX-FIRST; remaining findings and large-reference limits need resolution.

Grounding completed; approved execution plan: ../PLAN-fidelity.md. First step: conditional formatting tests and row-block evaluator. Luna and Terra attempts did not provide valid acceptance tests; Sol xhigh took over under the two-failure routing rule. No feature commit yet. Session: 01a12961-bb90-71b3-af7a-07bc42b2cd29.

## Verification
Latest focused checks: Python 22/22 and web 23/23. Earlier full Python 296/296 passed; earlier full web 102/103 failed a cache regression, since fixed in focused tests only. Final duplicate suite runs, content guard, ruff, pinned Prettier, builds, security closure and browser acceptance are pending. Execution stopped on repeated sub-1-GiB capacity readings; controller last measured 180297728 bytes free. Owned executor exited cleanly.

No passing feature tests or final acceptance claimed yet. Runtime evidence belongs in results-private/fidelity. Active Python lock audit: no known vulnerabilities. Active web lock audit: zero findings. Exact Opus probe: canonical claude-opus-5-5. These are preliminary gates, not a code review or feature acceptance. TaskSpace 42 reserved. Low-disk incident preserved in STATE; checks remain sequential.

## Next
Recover to at least 4 GB free, confirm swap growth stays within 2 GB of the 9912.44 MB resume baseline, then test and commit the uncommitted tables/AutoFilter feature. Run pending exact claude-opus-5-5 review and implement sheet views before final acceptance gates. TaskSpace 42 stays open until successful acceptance.

Implement conditional formatting first, then drawings, notes, tables/filters and sheet view in order.

## CF contract and bounds

`conditional_style_overrides(snapshot, values, query)` returns `ConditionalFormatResult(styles, warnings)`. `/rows` streams `style_override` per cell and generic `conditional_format_warnings` per block; a failing rule rolls back its own styles without hiding ordinary values or healthy styles. Workbook sheet metadata includes `has_conditional_formats`. Calc refresh touches loaded CF blocks only and coalesces concurrent events.

Bounds: 100000 formula targets/request, 256 formulas and 10000 referenced cells/public-engine load, 100000 duplicate-scan cells/rule, 100000 override records/request, 31 warnings plus a truncation marker. Two cached value/template blocks. Skipped formatting is visibly warned; extended x14 CF and external/structured formula references remain unsupported. No real template fidelity is claimed.

## Drawings contract

Internal PNG/JPEG/GIF assets only, with authenticated copy access. Macro identifiers remain verbatim; authorized POST action returns 501 and a clear message, viewer writes fail 403 and inaccessible copies fail 404. Anchors use installed Univer public skeleton, scroll and zoom APIs with DOM presentation, because the pinned preset has no sheets-drawing plugin. Limits: 500 anchored drawings, 100 asset references, 5 MiB per asset, 25 MiB cumulative. Slicer presentation is view only.

## Notes contract

Legacy notes are imported into row-block snapshots and returned as optional cell.note with author/text. Red markers and safe DOM hover/focus tooltips use live merged-cell/freeze/scroll/zoom geometry. Eviction and sheet switches remove markers; focused overlay controls do not execute destructive grid shortcuts. Limits: 10000 notes and 64 KiB UTF-8 per note/author. Notes are read only; threaded comments are not supported.

## Wide CF regression

Controller reproduction initially showed0 of15600 styles for a bounded50x312 TRUE() range. The target guard is now100000, with existing reference/batch/override limits intact. Identical proven-pure translated formulas share an evaluation and preserve aliases; named references, unknown functions and known volatile/context-dependent functions remain per-target. API tests assert literal first/last styles, complete counts and no warnings for79872 targets, plus atomic alias-failure rollback.
