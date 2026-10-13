# Fidelity track

2026-10-11. feat/v2-fidelity starts at f74d131. User approved all five features and local acceptance.

## Scope
Synthetic fixtures only. No private workbook access. pp/engine and calculation pipeline remain read-only.

## Progress
2026-10-11 22:50 final acceptance pass. HIGH3 committed 425ca95: a conditional formula that returns an error is a no-match for that cell only, while a public engine failure still warns and rolls back its own rule. Sheet views committed a9a809c: tab colours, gridline visibility, per-sheet zoom, scroll position, frozen split, hidden rows and columns, merges, the active sheet and hidden-sheet states are imported and projected. The browser acceptance found two real defects, both fixed with regression tests in 301943d: overlay layers were detached when Univer remounted the grid container, and the template scroll cell was written into the Univer freeze split. Evidence is results-private/fidelity/final-ego.json and results-private/fidelity/shots/. Historical notes follow.

HIGH1 committedf958f06: local hide/show now passes through the real grid command guard. Behavioral RED5/7; focused GREEN7/7; relevant grid37/37 and controller37/37. Content guard0, pinned Prettier/diff check green. HIGH2 next; all final gates remain pending.

2026-10-11 21:04 resume approved. Latest resource rule supersedes every historical swap ceiling: stop and record only if `df -m /` Available is below2500 MB; swap growth alone is not a stop reason. Initial controller disk5108 MB. Same exact Sol takeover session resumed, sequential HIGH1 then HIGH2 then HIGH3, each RED/GREEN and a focused commit, before Phase5. Ego TaskSpace42 verified agent-owned and retained with p1. No final acceptance claim.

2026-10-11 resume STOPPED by current user resource gate. Baseline swap5278.62 MB; ceiling6814.62 MB. Controller measured swap6882.50 MB (growth1603.88 MB), disk3436 MB, so execution stopped immediately. No server or browser check started in this resume, no owned executors remain. Tables/AutoFilter committed7b4ae03 after Python36/36, web31/31, content guard0, global Ruff, pinned Prettier and diff check green. Phase5 added only unrun tests/v2/test_fidelity_views.py and web/test/fidelity-views.test.mjs; production unchanged. Never stage .venv or web/node_modules.

Pending exact Claude review now RAN successfully: response modelUsage contains only claude-opus-5-5, is_error false; verdict FIX-FIRST. Evidence results-private/fidelity/resume-opus-review.json and resume-opus-result.md. Three HIGH findings require regression tests and fixes: (1) local hide/show rows are cancelled by PlanningGrid.command because calls lack the internal mutation guard; (2) pp/server.py async rows handler blocks the event loop during synchronous CF work, use existing asyncio.to_thread pattern; (3) one formula-result error rolls back healthy CF targets, treat error result as false while retaining EngineError handling. MED findings to investigate in same pass: deterministic function batching/template-block reuse, explicit cell styling precedence over table styles, overlay scroll redraw/asset caching. Header clipping needs coverage; table-order memory claim is unverified and must not be accepted without source evidence. Full review is preserved verbatim. No shipping/acceptance claim.

Security: branch gitleaks f74d131..7b4ae03 no leaks; pinned Python audit no known vulnerabilities; production web audit0. Semgrep p/python flagged only fixed committed table-catalog ET.parse sink; trusted-path triage in resume-security-triage.md. Final all-source scan/full web dependency audit, coverage, planner review, duplicate full suites/builds and browser remain pending.

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
Final acceptance pass on feat/v2-fidelity at 301943d: Python 321/321 twice, web 124/124 twice, content guard exit 0, `ruff check pp run_v2.py tests/v2` clean, pinned Prettier clean, two identical production builds (bundle 3efadf59), and a real-server Ego check on port 8791 against the synthetic eight-sheet fixture covering all six CF rule types, drawings, notes, tables, filters and sheet views. Earlier historical gates below are superseded.

No passing feature tests or final acceptance claimed yet. Runtime evidence belongs in results-private/fidelity. Active Python lock audit: no known vulnerabilities. Active web lock audit: zero findings. Exact Opus probe: canonical claude-opus-5-5. These are preliminary gates, not a code review or feature acceptance. TaskSpace 42 reserved. Low-disk incident preserved in STATE; checks remain sequential.

## Next
Current stop rule: `df -m /` Available below2500 MB. Swap growth alone does not stop work. Fix the three HIGH exact Opus findings preserved in results-private/fidelity/resume-opus-result.md, one focused commit each. Tables committed7b4ae03; two view tests remain unrun/uncommitted. Then finish Phase5 and all acceptance gates.

## CF contract and bounds

`conditional_style_overrides(snapshot, values, query)` returns `ConditionalFormatResult(styles, warnings)`. `/rows` streams `style_override` per cell and generic `conditional_format_warnings` per block; a failing rule rolls back its own styles without hiding ordinary values or healthy styles. Workbook sheet metadata includes `has_conditional_formats`. Calc refresh touches loaded CF blocks only and coalesces concurrent events.

Bounds: 100000 formula targets/request, 256 formulas and 10000 referenced cells/public-engine load, 100000 duplicate-scan cells/rule, 100000 override records/request, 31 warnings plus a truncation marker. Two cached value/template blocks. Skipped formatting is visibly warned; extended x14 CF and external/structured formula references remain unsupported. No real template fidelity is claimed.

## Drawings contract

Internal PNG/JPEG/GIF assets only, with authenticated copy access. Macro identifiers remain verbatim; authorized POST action returns 501 and a clear message, viewer writes fail 403 and inaccessible copies fail 404. Anchors use installed Univer public skeleton, scroll and zoom APIs with DOM presentation, because the pinned preset has no sheets-drawing plugin. Limits: 500 anchored drawings, 100 asset references, 5 MiB per asset, 25 MiB cumulative. Slicer presentation is view only.

## Notes contract

Legacy notes are imported into row-block snapshots and returned as optional cell.note with author/text. Red markers and safe DOM hover/focus tooltips use live merged-cell/freeze/scroll/zoom geometry. Eviction and sheet switches remove markers; focused overlay controls do not execute destructive grid shortcuts. Limits: 10000 notes and 64 KiB UTF-8 per note/author. Notes are read only; threaded comments are not supported.

## Wide CF regression

Controller reproduction initially showed0 of15600 styles for a bounded50x312 TRUE() range. The target guard is now100000, with existing reference/batch/override limits intact. Identical proven-pure translated formulas share an evaluation and preserve aliases; named references, unknown functions and known volatile/context-dependent functions remain per-target. API tests assert literal first/last styles, complete counts and no warnings for79872 targets, plus atomic alias-failure rollback.
