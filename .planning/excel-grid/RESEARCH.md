# Research: formula engine (A) and browser grid (B) for the Price Planning web app

Date: 2026-10-10. Read-only research. Facts file: `FACTS.md` (same folder).
Method: registry JSON APIs (PyPI, npm, GitHub) via curl; shallow clones / npm tarballs / wheels read locally (function tables, source, docs, changelogs); WebFetch/WebSearch for vendor docs. Context7 was not used (library docs came straight from the repos, which are newer than any index).
Marking: V = verified this session from the named source. S = taken from a web-search summary only (page itself not fetched). UNVERIFIED = not checked. No version below is from memory.

Target function list checked: LET, XLOOKUP, FILTER, UNIQUE, SEQUENCE, SWITCH, IFS, NUMBERVALUE, MAXIFS, CONCAT (the ten asked for), plus a wider set from FACTS.md (SUBSTITUTE, IFERROR, MATCH, VLOOKUP, INDEX, COLUMN, TRIM, ISERROR, SUMIFS, COUNTIFS, SUMPRODUCT, MINIFS, HLOOKUP, ISFORMULA, OFFSET, INDIRECT, CELL, TEXT, FLOOR.MATH, DAYS, COUNTA, etc.).

---

## 0. Bottom line

**(A) Engine ranking**
1. **Formualizer 0.11.1** (PyPI wheel, MIT OR Apache-2.0). Only candidate with ALL of: pip wheel for Windows, every function on our list, dynamic-array spills, a true dependency graph with dirty-subgraph recalc, undo/redo, and a published 1M-formula benchmark. Gate it on a spike with the real .xlsm (our file is about 1.9x the size of its biggest published benchmark and has a very different formula mix).
2. **IronCalc 0.8.3** (PyPI wheel, MIT/Apache-2.0). Every function on our list, spill supported in the released 0.8.3. But its evaluator recomputes the whole workbook on `evaluate()` (no dirty tracking found in source) and no performance numbers are published. Good second opinion and fallback; weak for live recalc on 1.9M formulas.
3. **Univer formula engine, headless Node** (Apache-2.0). Function names all present, but needs Node (portable node.exe is feasible), JS only, no published formula-engine numbers outside a vendor "Pro" report, and server-side calc is listed as a Pro feature.
4. **LibreOffice headless/UNO**: use as an offline parity oracle, not the live engine.
Rejected: HyperFormula (no LET, NUMBERVALUE, CONCAT, INDIRECT; GPLv3 or paid; Node), `formulas` (no LET, SEQUENCE, OFFSET, INDIRECT; EUPL), pycel (last release 2021; no LET, SEQUENCE, SWITCH; no spilling).
Keep the current Python interpreter (101/103 parity) as the incumbent and as a third cross-check.

**(B) Grid ranking**
1. **Univer OSS (Apache-2.0)**, front-end only, values fed from the server engine. It is the only free grid that ships canvas rendering, number formats, merged cells, freeze panes, autofill, Excel-HTML clipboard and undo in the open-source tier. Pro-only: xlsx import/export, collaboration, edit history, server-side calc, so we write the xlsx-to-snapshot converter and the history ourselves (we already have history).
2. **Glide Data Grid 6.0.3** (MIT). Canvas, very fast, has fill handle, paste/copy hooks, frozen columns, custom canvas cell renderers. Gaps: no undo, merges are horizontal spans only, you hand-draw every format/border/fill, last npm release 2024-02-03, React peer 16-18.
3. **FortuneSheet 1.0.4** (MIT). Complete UI but couples to its own formula parser and xlsx import is an external plugin; performance on 1.2M cells unknown.
Not suitable: AG Grid Community (range selection and clipboard are Enterprise), Handsontable (commercial use needs a paid license), Jspreadsheet CE (DOM table; xlsx import and formulas are "Enterprise"), RevoGrid OSS (undo, merged cells, smart autofill are Pro), Luckysheet (archived), x-spreadsheet (last release 2021).

---

## 1. Comparison table (A) engines

| | Formualizer | IronCalc | HyperFormula | Univer engine (headless) | `formulas` | pycel | LibreOffice UNO |
|---|---|---|---|---|---|---|---|
| Latest version / date | 0.11.1 / 2026-10-07 (PyPI) V | 0.8.3 / 2026-07-31 PyPI; GH release 2026-08-03; npm wasm 0.8.4 / 2026-08-01 V | 3.5.0 / 2026-10-09 (npm; CHANGELOG) V | 1.0.4 / 2026-10-10 (npm) V | 1.3.4 / 2026-03-11 (PyPI) V | 1.0b30 / 2021-10-13 (PyPI) V | not checked (24.8 added the new functions, S) |
| License | MIT OR Apache-2.0 V | MIT/Apache-2.0 (npm), Apache-2.0 (GH) V | GPLv3 or proprietary key; LICENSE.txt V | Apache-2.0 core V | EUPL 1.1+ V | no license field on PyPI (UNVERIFIED) | MPL 2.0 (UNVERIFIED) |
| Windows wheel (py3.10+) | cp310-abi3 win_amd64 V | cp310..cp313 win_amd64 (and win32) V | n/a (JS) | n/a (JS) | pure Python wheel V | pure Python wheel V | needs LO install (portable build UNVERIFIED) |
| LET | yes V | yes V | **no** V | yes V | **no** V | **no** V | yes (24.8, S) |
| XLOOKUP | yes V | yes V | yes V | yes V | yes V | yes V | yes (24.8, S) |
| FILTER | yes V | yes V | yes V | yes V | yes V | yes V | yes (24.8, S) |
| UNIQUE | yes V | yes V | yes V | yes V | yes V | yes V | yes (24.8, S) |
| SEQUENCE | yes V | yes V | yes V | yes V | **no** V | **no** V | yes (24.8, S) |
| SWITCH | yes V | yes V | yes V | yes V | yes V | **no** V | UNVERIFIED |
| IFS | yes V | yes V | yes V | yes V | yes V | yes V | UNVERIFIED |
| NUMBERVALUE | yes V | yes V | **no** V | yes V | yes V | yes V | UNVERIFIED |
| MAXIFS | yes V | yes V | yes V | yes V | yes V | yes V | UNVERIFIED |
| CONCAT | yes V | yes V | **no** (only CONCATENATE and operator HF.CONCAT) V | yes V | yes V | yes V | UNVERIFIED |
| Other FACTS.md functions missing | none of the 41 checked V | none of the 41 checked (doc page exists for each) V | LET, NUMBERVALUE, INDIRECT, CONCAT, CELL, ISFORMULA | none of the 33 checked V (name table) | LET, SEQUENCE, OFFSET, INDIRECT, CELL | LET, SEQUENCE, SWITCH | n/a |
| Dynamic-array spill | yes: spill, `A1#`, `#SPILL!`, xlsx write-back (CHANGELOG) V | yes in code (SpillCell, spill functions) and in tag v0.8.3; docs page still says "formulas don't spill yet" (stale) V | array formulas; FILTER/UNIQUE/SEQUENCE exist; spill semantics UNVERIFIED | "array formulas" (S); UNVERIFIED depth | no spill (UNVERIFIED) | explicitly "not going to do any spilling" (source comment) V | yes |
| xlsx/xlsm import | `load_workbook`; xlsm package-backed in the CellRune sibling only; xlsm in Formualizer UNVERIFIED | `load_from_xlsx`; xlsm UNVERIFIED | none (data in via arrays) | Pro only (xlsx), browser API rejects .xlsm | `ExcelModel`, xlsx | openpyxl-based | native |
| Incremental recalc / dep graph | yes: dependency graph, dirty subgraph only V | no: `evaluate()` = "every formula of every sheet" in a pass (source header comment) V | yes (graph, `dependencies.md`) V | yes, in-engine (UNVERIFIED how) | whole-model compile; UNVERIFIED | limited | recalcs dirty cells, UNVERIFIED |
| Published perf on big workbooks | 1M-formula financial model: load 4.2 s, first calc 0.38 s, change 1 input + recalc 52 ms, 45 MB after calc; 52 workbooks, ~1/4 the time of LibreOffice 24.2 (README, vendor-run, shared machine) V | none found in README/docs V | none found for engine at this scale | "Performance Report" (Pro): 20,000 VLOOKUP formulas 4.73 s; 10M-cell load ~1 s (S, page fetch returned 404) | none | none | Formualizer README table: e.g. product lookups 53.3 s vs 0.35 s V |
| Undo/redo, dependents API | `undo`, `redo`, `begin_action`, `dependents`, `precedents` (binding source) V | user model has history (UNVERIFIED depth) | undo/redo built in | Pro history | no | no | no |
| Runtime needs on server | Python only | Python only | **Node** | **Node** | Python | Python | LO install + UNO |
| Maintenance signal | 192 stars, pushed 2026-10-10, single maintainer org (PSU3D0), repo created 2025-03-10 V | 4,203 stars, pushed 2026-10-09 V | 2,806 stars, pushed 2026-10-09 V | 22,526 stars, pushed 2026-10-10 V | 500 stars, pushed 2026-08-07 | stale | n/a |

Sources (all fetched 2026-10-10):
- PyPI JSON: https://pypi.org/pypi/formualizer/json , /ironcalc/json , /formulas/json , /pycel/json , /cellrune/json
- npm registry: https://registry.npmjs.org/hyperformula , /@univerjs%2fcore , /@ironcalc%2fwasm
- GitHub API repos (license, pushed_at, stars): api.github.com/repos/PSU3D0/formualizer, ironcalc/IronCalc, handsontable/hyperformula, dream-num/univer, vinci1it2000/formulas
- Formualizer README + `docs-site/src/generated/functions-meta.json` (409 functions) + `CHANGELOG.md` + `docs-site/content/docs/core-concepts/dependency-graph-and-recalc.mdx`: https://github.com/PSU3D0/formualizer (clone at commit of 2026-10-10)
- IronCalc: https://github.com/ironcalc/IronCalc (clone HEAD 2026-10-09; `base/src/functions/mod.rs`, `base/src/evaluation.rs` header, `docs/docs.ironcalc.com/src/features/error-types.md`, tag v0.8.3 grep), docs https://docs.ironcalc.com/functions/lookup-and-reference.html
- HyperFormula: https://github.com/handsontable/hyperformula (LICENSE.txt, `src/license/functionCapabilities.ts`, `src/i18n/languages/enGB.ts`, CHANGELOG 3.5.0 entry dated 2026-10-09)
- Univer: npm tarballs `@univerjs/engine-formula@1.0.4` (function-names tables and per-function folders), README https://raw.githubusercontent.com/dream-num/univer/dev/README.md
- formulas 1.3.4 wheel and pycel 1.0b30 wheel: source grepped for function registrations.

### Per-engine notes

**Formualizer** (https://github.com/PSU3D0/formualizer, Python API `formualizer`).
- Function coverage: all 41 names tested are in the generated function table (409 entries), including LET, SEQUENCE, NUMBERVALUE, SWITCH, XLOOKUP, FILTER, UNIQUE, MAXIFS, CONCAT, ISFORMULA, OFFSET, INDIRECT, CELL. Presence in the table is not proof of Excel-exact behaviour; the README says edge cases are checked against Excel.
- Spill: CHANGELOG lists spill writeback, `A1#`/ANCHORARRAY, `#SPILL!` handling. An older doc, `docs/compatibility.md`, still lists "dynamic-arrays (currently skipped)" as a corpus test category; treat as stale.
- Recalc: edits mark downstream dependents dirty and recalc visits only dirty subgraphs; copied formulas are evaluated as "families" over Arrow columns. This is exactly the shape of our workbook (hundreds of thousands of copied formulas per sheet).
- Python API methods found in binding source: `load_workbook`, `load_path`, `set_value`, `set_formula`, `evaluate_cell(s)`, `evaluate_all`, `undo`, `redo`, `begin_action`, `dependents`, `precedents`.
- Caveats: young project (created 2025-03), one org; our workbook has 1.894M formulas and 650 MB of sheet XML versus a 1M-formula published benchmark; benchmark numbers are vendor-run on a "shared machine". The mix (3.6M SUBSTITUTE, 487k NUMBERVALUE, 881k VLOOKUP, 259k XLOOKUP, 93k LET) may not hit the fast "family" path. VBA: xlsm macros are not executed anyway; whether load of .xlsm works is UNVERIFIED (test).
- Windows: `formualizer-0.11.1-cp310-abi3-win_amd64.whl` exists (V). abi3 means one wheel for Python 3.10+.

**IronCalc** (https://github.com/ironcalc/IronCalc).
- All 41 names have a doc page under `docs/docs.ironcalc.com/src/functions`; the ten required ones are in the `Function` enum and dispatch table in `base/src/functions/mod.rs`; LET has `fn_let` in `logical/let.rs`.
- Spill: `Cell::SpillCell` in `cell.rs` and `spill_functions.rs` at HEAD; also present in tag v0.8.3 (4 hits for SpillCell, Filter dispatch present). The error-types doc still says "formulas don't spill yet": stale vs code.
- Recalc: `evaluation.rs` header: "Evaluation of every formula in the workbook ... then every formula of every sheet." No dirty-cell tracking found. Edit-one-cell = full pass over 1.9M formulas. This is the decisive weakness for "live recalculation".
- Python bindings (`bindings/python/src`): `load_from_xlsx`, `load_from_bytes`, `create_user_model_from_xlsx`, `set_user_input`, `evaluate`. Also has WASM and Node bindings.
- Unsupported (docs): real-time collaboration, charts, pivot tables.
- No published performance numbers found (searched README and docs).

**HyperFormula 3.5.0** (npm). Best-in-class dependency graph and batch ops, but: LET, NUMBERVALUE, CONCAT, INDIRECT, CELL, ISFORMULA absent from the function tables (V: `src/license/functionCapabilities.ts`, `enGB.ts`). License: "dual-licensed ... proprietary or GPLv3, determined by the license key" (LICENSE.txt, V). 3.5.0 adds a new license-key format in which a proprietary key can grant only a subset of functions and others return `#LIC!` (CHANGELOG, V). Internal-use implication (not legal advice): GPLv3 obligations attach on distribution/conveying; running it only on a server you control and never shipping it to others is generally not "distribution", but the browser would also receive the code if run client-side, and users are employees of a company using a product built on it. A commercial key is the clean route; price not found. Since it lacks four functions we use heavily (LET 93k, NUMBERVALUE 487k, CONCAT 362k), it fails on coverage regardless of license.

**Univer formula engine.** `@univerjs/engine-formula@1.0.4` has per-function folders including `logical/let`, `math/sequence`, `logical/switch`, `logical/ifs`, `text/numbervalue`, `statistical/maxifs`, `text/concat`, `lookup/xlookup|filter|unique`. README: Headless Mode "supports Node.js >=18.17.0", for "server-side workbook/document processing, formula calculation". Pro features listed in README: "Real-time collaboration, edit history, import/export, printing, charts, pivot tables, sparklines, outlines, shapes ... data connectors and server-side calculation". No independent perf numbers. Vendor "Performance Report" (S): 1M-cell load ~0.25 s, 10M ~1-1.5 s, 20,000 VLOOKUP 4.73 s, scrolling 50-60 FPS on an i9-13900H (page URL https://docs.univer.ai/guides/pro/performance-report, search summary only; WebFetch gave 404). Scale of 20k formulas in 4.7 s is poor against 1.9M formulas, so do not count on it as the recalc engine.

**`formulas` 1.3.4**: 597 functions registered, but LET, SEQUENCE, OFFSET, INDIRECT, CELL missing (V). EUPL 1.1+ (copyleft, internal use fine). Pure-Python compile of an Excel model; at 1.9M formulas, expect very high memory/time (UNVERIFIED, no numbers).

**pycel 1.0b30**: last release 2021-10-13; no LET, SEQUENCE, SWITCH; source comment "not going to do any spilling" (V). Out.

**LibreOffice headless/UNO**: 24.8 added FILTER, LET, RANDARRAY, SEQUENCE, SORT, SORTBY, UNIQUE, XLOOKUP, XMATCH (S: https://linuxiac.com/libreoffice-24-8-office-suite , https://www.debugpoint.com/libreoffice-24-8-features/ ; secondary sources). Formualizer README compares itself to LibreOffice Calc 24.2 and reports LO slower on all 9 shown workloads (e.g. 53.3 s vs 0.35 s). Needs an install and a UNO bridge; "portable LO without admin" UNVERIFIED. Use only as a one-off oracle to diff cell values against the xlsm's cached values.

**Newer 2025-2026 engines found**: CellRune 0.1.23 (PyPI, 2026-10-01, MIT OR Apache-2.0, Rust, cp310-cp312 win_amd64 wheels, headless XLSX/XLSM read and deterministic calc, repo github.com/emulette/cellrune). Its PyPI description mentions XLOOKUP, spill references, merged ranges, tables, xlsm round-trip; LET, SEQUENCE, FILTER, UNIQUE, SWITCH, NUMBERVALUE, MAXIFS are not mentioned in the description (UNVERIFIED either way). Too new (0.1.x) to bet on; worth a one-hour probe only if Formualizer fails the spike.

### Node.js feasibility
Only matters for Univer-headless or HyperFormula. Node LTS v24.21.0 (2026-09-07) and `node-v24.21.0-win-x64.zip` exist at nodejs.org/dist (HTTP 200 on the zip, V). A zip needs no admin; vendor `node.exe` in the repo and start it from Python with `subprocess`. Feasible but adds a second runtime and an IPC layer. Not needed for the recommended path (Formualizer/IronCalc are pip wheels; the grid is static JS built on the dev Mac and served as files).

---

## 2. Comparison table (B) grids

| | Univer OSS | Glide Data Grid | FortuneSheet | RevoGrid | Jspreadsheet CE | AG Grid Community | Handsontable | Luckysheet | x-spreadsheet |
|---|---|---|---|---|---|---|---|---|---|
| Latest / date (npm) | 1.0.4 / 2026-10-10 | 6.0.3 / 2024-02-03 | @fortune-sheet/react 1.0.4 / 2025-11-06 | 4.28.3 / 2026-10-05 | 5.0.4 / 2025-08-25 (CE); Pro `jspreadsheet` 13.0.1 / 2026-10-03 | 36.2.0 / 2026-09-16 | 18.1.2 / 2026-10-07 | 2.1.13 / 2021-01-19; repo ARCHIVED | 1.1.9 / 2021-05-20 |
| License | Apache-2.0 | MIT | MIT | MIT (core) | MIT (CE) | MIT (community) | proprietary; commercial use needs paid license | MIT | MIT |
| Virtualized / canvas | canvas | canvas; "scales to millions of rows", lazy | canvas (UNVERIFIED perf) | virtual scroll, DOM | DOM table, `lazyLoading`/pagination options | row virtualization (UNVERIFIED here) | virtualized DOM (UNVERIFIED) | canvas | canvas |
| 4000x300 (1.2M cells) | designed for it; vendor claims 10M-cell loads (S) | yes by design (data pulled on demand by callback) | UNVERIFIED | UNVERIFIED | unlikely (DOM) | UNVERIFIED | UNVERIFIED | UNVERIFIED | UNVERIFIED |
| Number formats, fill, borders | yes (`sheets-numfmt`, styles) | you draw them | yes | partial/custom | partial | custom | yes | yes | yes |
| Merged cells | yes (`add-worksheet-merge` command) | horizontal spans only | yes (UNVERIFIED) | Pro only | `mergeCells` option | n/a (not a sheet) | yes | yes | yes |
| Frozen panes | yes (`set-frozen` command) | `freezeColumns`; frozen rows only trailing | yes | pinned rows/cols (OSS) | `freezeColumns` | pinned cols/rows | yes | yes | yes |
| Double-click edit | yes | yes | yes | yes | yes | yes | yes | yes | yes |
| Fill handle drag | yes (`auto-fill` command, UI controller) | `fillHandle` + `onFillPattern` | yes (UNVERIFIED) | Pro ("smart autofill") | UNVERIFIED | **Enterprise** | yes | yes | no (UNVERIFIED) |
| Range copy/paste, Excel clipboard | yes: clipboard service with `htmlIsFromExcel`, html-to-usm and usm-to-html | `getCellsForSelection` copy; `onPaste` hook (TSV) | yes (UNVERIFIED HTML) | OSS "copy/paste from Excel" | yes | **Enterprise** | yes | partial | partial |
| Undo/redo | yes | no | yes | Pro | yes (`onundo`, `redo`) | cell-edit undo only (UNVERIFIED) | yes | yes | yes |
| Custom cell render (locked vs input) | yes (permissions, styles) | full canvas renderer API | limited | cell templates | `cells` callback | yes | yes | limited | limited |
| Own formula engine you must bypass | yes: formula plugins are optional packages; but `@univerjs/sheets` lists `@univerjs/engine-formula` as a dependency, so the engine is bundled; running with it disabled is UNVERIFIED | none (pure grid) | yes, forked handsontable formula-parser | none | basic built-in | none | optional HyperFormula | yes | yes |
| xlsx import/export in free tier | **Pro** | none | external plugin | none | **Enterprise** | none | none (paid) | external plugin | none |
| Collaboration / history | **Pro** | none | op callbacks `onOp`, demo backend | none | Pro | none | none | none | none |
| Fit | best | good if merges/undo built by us | fallback | no | no | no | no (license) | no | no |

Sources: npm registry (versions/dates/licenses); GitHub API (Luckysheet `archived: true`, pushed 2025-08-19; x-spreadsheet pushed 2024-08-07; glide pushed 2026-01-21); RevoGrid feature table and Pro split from https://github.com/revolist/revogrid (WebFetch); Glide https://github.com/glideapps/glide-data-grid README + `API.md` in npm tarball 6.0.3 (V: `fillHandle`, `onFillPattern`, `onPaste`, `getCellsForSelection`, `freezeColumns`, `span` horizontal only); FortuneSheet README https://github.com/ruilisi/fortune-sheet (WebFetch); AG Grid https://www.ag-grid.com/javascript-data-grid/range-selection/ ("enterprise: true") and /clipboard/ ("enabled by default in the enterprise version"); Handsontable `LICENSE.txt` ("dual-licensed - depending on whether your use [is] for commercial purposes ...", V); Jspreadsheet CE d.ts in npm tarball (V: `freezeColumns`, `lazyLoading`, `mergeCells`, undo/redo) and https://jspreadsheet.com/docs/pro/features (Enterprise: Formulas, Export/Import XLSX; WebFetch); Univer: tarballs `@univerjs/sheets-ui@1.0.4` (auto-fill controller, clipboard with `htmlIsFromExcel`), `@univerjs/sheets@1.0.4` (commands: `add-worksheet-merge`, `set-frozen`, `auto-fill`), https://docs.univer.ai/guides/sheets/features/import-export (Pro exchange-client packages need a conversion backend; browser API does not accept .xlsm; CSV/TSV export one sheet).

### Univer details that matter
- Pro (paid or restricted): xlsx import/export (needs a conversion backend), collaboration, edit history, server-side calculation, printing, pivots, charts (README + import-export doc). Unlicensed Pro use has watermarks and caps (S: https://docs.univer.ai/guides/sheets/getting-started/pro). Price: not found.
- OSS: workbook model, ranges, selection, number format, filter/sort, data validation, conditional formatting, hyperlinks, comments, find/replace (README); merge, frozen panes, autofill and clipboard commands/services exist in the 1.0.4 packages (V).
- Feeding server values: cells are `{v, f, s}` in the snapshot. Setting `v` from our Python engine and handling `onCommandExecuted` edits to POST changes is the natural pattern; disabling local calc is UNVERIFIED and should be tested in the spike (render a snapshot with `f` and `v` but no formula plugin registered).
- We must write: xlsx-to-Univer-snapshot converter (styles, number formats, merges, freeze, column widths) in Python with a streaming reader (650 MB sheet XML rules out naive openpyxl); change-history and who-edited-what; live co-editing (our own broadcast of cell edits).
- Risk: whole sheet snapshot lives in browser memory. Largest sheet A1:KZ4088 is about 1.27M cells (4088 x 312). Dept OTD LY is A1:AP10001 (~420k). LY Data is A1:Q100000 (1.7M cells, 99,998 formulas). Verify memory and scroll FPS on an office PC; if too heavy, split or window the sheet data.

### Glide details
Glide is a pure grid: you own the model. Upsides: lowest-risk performance, `getCellContent` pulls only visible cells, so server-fed values are natural. Downsides: you implement formats/borders/fill/merges drawing, undo/redo stack, Excel-HTML paste (paste hook receives plain strings), frozen top rows; React 16-18 peer; no release since 2024-02-03.

---

## 3. Recommended path

1. **Spike (1 day, decides A):** load the real .xlsm in Formualizer and IronCalc on the Mac (8 GB) and a Windows PC. Record: load time, peak RSS, first full calc, single-input edit recalc time, and cell-by-cell diff against the xlsm's cached values, with attention to LET (93k), NUMBERVALUE (487k), CONCAT (362k), XLOOKUP (259k), SEQUENCE/FILTER/UNIQUE spills. Compare against the existing Python interpreter's 101/103 parity.
2. **Engine decision rule:** if Formualizer loads under available RAM and edit-recalc is sub-second, adopt it behind a thin Python adapter (`set_value`, `evaluate_cells(visible range)`, `dependents`). If memory fails, keep the visible-range on-demand interpreter and use Formualizer/IronCalc as the regression oracle.
3. **Grid:** prototype Univer OSS with a server-fed 4000x300 snapshot (values only, formulas shown as text, locked vs input styling, fill handle, paste from Excel). Pass criteria: scroll 60 FPS, paste of 100x20 from Excel, fill-drag, undo, memory under a stated budget. Fallback Glide.
4. **Packaging:** build the JS bundle on the dev Mac (node there), commit the built static files (vendor), serve from Python/waitress. Server needs no Node. Wheels (`formualizer`, optional `ironcalc`) vendored as files in a `wheels/` folder and installed with `pip install --no-index --find-links`.
5. **Licensing:** keep to MIT/Apache components. Avoid Handsontable, AG Grid Enterprise, Univer Pro, HyperFormula commercial unless Sir approves spend (any spend is his click).

---

## 4. Not verified / gaps
- Real-file behaviour of every engine at 1.9M formulas (no vendor number matches our scale).
- Whether Formualizer or IronCalc can open `.xlsm` with 22 sheets and 68 defined names without refusal.
- IronCalc `evaluate()` time on a large workbook; no dirty-tracking found does not prove none exists in a different code path (I read the evaluation header, not every module).
- LibreOffice function behaviour for SWITCH, IFS, NUMBERVALUE, MAXIFS, CONCAT; portable LibreOffice without admin.
- Univer: running without its formula plugin; actual browser memory for 1.27M-cell sheet; unlicensed Pro limits (only search-summary).
- Univer performance report figures (search summary; page fetch 404).
- FortuneSheet, Handsontable, AG Grid, RevoGrid performance at 1.2M cells.
- Handsontable and Univer Pro prices; HyperFormula commercial price.
- CellRune function coverage.
- Legal: GPLv3 internal-use reading above is a general summary, not legal advice.

Temporary clones/tarballs were kept in the session scratchpad only; no repo code was changed.
