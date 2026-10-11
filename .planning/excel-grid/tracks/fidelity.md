# Fidelity track

2026-10-11. feat/v2-fidelity starts at f74d131. User approved all five features and local acceptance.

## Scope
Synthetic fixtures only. No private workbook access. pp/engine and calculation pipeline remain read-only.

## Progress
CF F1-F7 correction implementation is tested (full Python 302/302, web 105/105) but unaccepted. Fresh exact Opus 5.5 re-review is blocked by weekly quota, HTTP 429, with provider-reported reset 7:30pm (Asia/Calcutta). Controller continues approved local work; no substitute or paid lane. Final acceptance remains blocked.

Resume: remaining CF importer/evaluator findings have regression coverage. Relevant executor checks passed (Python 25/25, web 23/23); controller full Python 300/300 and full web 103/103 passed. Content guard 0, project Ruff and pinned Prettier green. Fresh Opus 5.5 returned FIX-FIRST (verified identity); F1-F7 are in a second regression-fix pass before any commit. This is the current checkpoint; low-disk details below are historical.

CF draft and regression fixes are present but not accepted or committed. Other four features remain pending. Exact Opus 5.5 returned FIX-FIRST; remaining findings and large-reference limits need resolution.

Grounding completed; approved execution plan: ../PLAN-fidelity.md. First step: conditional formatting tests and row-block evaluator. Luna and Terra attempts did not provide valid acceptance tests; Sol xhigh took over under the two-failure routing rule. No feature commit yet. Session: 01a12961-bb90-71b3-af7a-07bc42b2cd29.

## Verification
Latest focused checks: Python 22/22 and web 23/23. Earlier full Python 296/296 passed; earlier full web 102/103 failed a cache regression, since fixed in focused tests only. Final duplicate suite runs, content guard, ruff, pinned Prettier, builds, security closure and browser acceptance are pending. Execution stopped on repeated sub-1-GiB capacity readings; controller last measured 180297728 bytes free. Owned executor exited cleanly.

No passing feature tests or final acceptance claimed yet. Runtime evidence belongs in results-private/fidelity. Active Python lock audit: no known vulnerabilities. Active web lock audit: zero findings. Exact Opus probe: canonical claude-opus-5-5. These are preliminary gates, not a code review or feature acceptance. TaskSpace 42 reserved. Low-disk incident preserved in STATE; checks remain sequential.

## Next
Recover disk capacity, then resume exact Sol session 01a12961-bb90-71b3-af7a-07bc42b2cd29 with results-private/fidelity/sol-next-review.md. Resolve CF findings and rerun full checks before first feature commit. TaskSpace 42 stays open until successful acceptance.

Implement conditional formatting first, then drawings, notes, tables/filters and sheet view in order.

## CF contract and bounds

`conditional_style_overrides(snapshot, values, query)` returns `ConditionalFormatResult(styles, warnings)`. `/rows` streams `style_override` per cell and generic `conditional_format_warnings` per block; a failing rule rolls back its own styles without hiding ordinary values or healthy styles. Workbook sheet metadata includes `has_conditional_formats`. Calc refresh touches loaded CF blocks only and coalesces concurrent events.

Bounds: 10000 formula targets/request, 256 formulas and 10000 referenced cells/public-engine load, 100000 duplicate-scan cells/rule, 100000 override records/request, 31 warnings plus a truncation marker. Two cached value/template blocks. Skipped formatting is visibly warned; extended x14 CF and external/structured formula references remain unsupported. No real template fidelity is claimed.
