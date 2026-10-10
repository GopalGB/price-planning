# Validation status

## Confirmed in this build

The original file was ingested with 22 sheets, 1,894,479 formulas, shared formula masters, saved values, style metadata, and cell validation rules. The original XLSM was not changed. One `#REF!` and one `#VALUE!` existed in the supplied workbook.

Twenty-three automated tests passed on Linux. Demo checks confirm startup without administrator or settings prompts, LAN address selection, self-signup with ordinary permissions, private-copy access, sharing without an administrator, disabled signup outside demo mode, and wrong-origin rejection. The operating checks cover first administrator setup, restart without prompts, the real bundled Waitress HTTP server, duplicate launch rejection, and an online SQLite backup containing committed WAL changes that restored into a new directory. An HTTP integration test against the actual imported workbook passed demo self-signup without administrator-created accounts, private-copy access, editor sharing, stale-write rejection, independent duplication, CSRF rejection, shared event delivery, and viewer protection. A cost of 10 and retail of 20 recalculated BU RECAP J23 to an IMU of 50%.

Representative formula checks: **101 matched out of 103**. Two differences remain in Dept OTD LY E2 and F2: the saved workbook values are empty strings and the native interpreter returns zero. This is a confirmed difference against the saved cache; the cause has not been established. A fresh Excel recalculation is needed to distinguish stale saved caches from a semantic difference in the native interpreter.

## Unverified

- All remaining formulas and nonblank pricing boundary cases.
- The effect of edits on every downstream report, coupon calculation, sales curve, and departmental total.
- Dynamic-array edge cases, wildcard/search semantics, and all optional lookup modes.
- Original VBA workflows and workbook event handlers.
- Pivot refresh, slicers, charts, merged layouts, conditional formatting, and full XLSM export.
- Real source files, scheduled source refresh, and MicroStrategy metric switching.
- Windows batch execution and Windows file-lock behavior on the target server; optional externally provided PostgreSQL; HTTPS deployment; browser visual behavior; load capacity; and a target-server restore drill. Linux backup/restore with a small fixture passed.

## Inputs required to complete acceptance

1. A populated price-plan copy freshly recalculated and saved in the intended Excel version.
2. The Calendar source file and the previous-season rotational source file.
3. Representative LY and forecast exports, including the business meaning of the selected MFC/SBP/WBP metric.
4. Business-approved expected outputs for representative groups: percentage price, dollar price, CEX/LTS, landing windows, cross-year weeks, X/XB freezing, coupon exclusions, and rounding.

## Required acceptance cases

| Area | Evidence required |
| --- | --- |
| Inputs | Blank vs zero, text identifiers, duplicates, valid/invalid department and class, override restrictions |
| Pricing | Default and LY-main selection, discount/dollar prices, RD floors, CEX/LTS flags, frozen X/XB, rotation duration |
| Sales | Planned sales and curve changes update daily allocation and marketing block totals |
| Projections | Group contribution changes flow into each fiscal week and reconcile to the plan |
| Coupons | Secondary discounts, exclusions and coupon mixes reproduce out-the-door AUR |
| Sources | New files refresh the correct fields; removed source rows cannot retain old values |
| Reports | BU, fiscal week, department and store/flagship totals match a freshly recalculated workbook |
| Collaboration | Concurrent writes conflict predictably, sharing changes take effect, forks stay independent |
| Operations | Target Windows startup, HTTPS, workload benchmark, interrupted import recovery, real database restore, optional PostgreSQL if selected |

The production replacement remains incomplete until these cases pass. Formula text preservation and a passing baseline sample are not a substitute for changed-input parity.
