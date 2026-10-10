**My recommendation [MEDIUM]: server-authoritative incremental calculation, a bounded pool of warm calculation workers, and Univer OSS as the first grid candidate. Benchmark Formualizer first. Keep commercial Aspose.Cells as the fallback if Formualizer cannot achieve Excel parity.**

Retain the prototype’s accounts, permissions, sparse input overrides, cell versions, and history. Replace its grid and strengthen its calculation scheduling and cache invalidation. Those existing pieces already fit the required collaboration model. [FACTS.md](/Users/gopalmacbook/Desktop/Max%20HQ/price-planning/.planning/excel-grid/FACTS.md:17)

**Performance caveat:** exact milliseconds per edit and MB per copy are **UNSURE**. `FACTS.md` gives formula counts but not dependency fan-out, scanned range sizes, unique formula families, or office-server specifications. The expectations below describe workload classes; the numerical thresholds at the end are proposed acceptance criteria, not measured performance.

**1) Recalculation architecture and engine choice**

For this workbook, edit latency is approximately:

`queue delay + dependency invalidation + affected formula/range evaluation + result publication`

A single edit might affect a few formulas or most of the workbook. Repeated lookups and large range scans can dominate even when the affected formula count is modest. Excel itself uses dependencies and a calculation chain to avoid recalculating everything. [Microsoft calculation guidance](https://learn.microsoft.com/en-us/office/vba/excel/concepts/excel-performance/excel-improving-calculation-performance)

| Server engine | Calculation behavior and expected edit-latency class | Memory per warm copy | Windows, no admin, offline deployment | Verdict |
|---|---|---|---|---|
| **Formualizer through Python** | Documented incremental graph, copied-formula families, and lookup indexing. Subsecond narrow edits are a plausible hypothesis; broad lookup/aggregate edits can take seconds or longer. **Actual latency UNSURE.** | Native workbook, graph, columns, indexes, and mutable results. No verified shared-model/COW API; initially assume each instance duplicates its model. **MB/copy UNSURE.** | Published Python 3.10+ `win_amd64` wheel; vendor the wheel and dependencies. MIT/Apache-2.0. | **First benchmark candidate**, with mandatory compatibility work. |
| **IronCalc Python bindings** | Current public `Model.evaluate()` clears evaluation state and evaluates the whole model. A small edit therefore invokes full-model work. **Interactive latency UNSURE and structurally risky.** | Native workbook and evaluation state per instance. **MB/copy UNSURE.** | Windows CPython wheels exist, including Python 3.10 x64. MIT/Apache-2.0. | Do not choose it assuming dirty-subgraph recalculation. |
| **HyperFormula on Node** | Genuine incremental dependency graph. Narrow-edit responsiveness is plausible, but required-function and array-semantic gaps disqualify it before performance becomes decisive. | Full JavaScript workbook graph per instance, plus strings and caches. **MB/copy UNSURE.** | Bundle a portable Node runtime and packages. GPLv3/commercial, not permissive. Needs a separate workbook importer. | **Reject as primary engine for this workbook.** |
| **LibreOffice headless** | Keep Calc resident and manipulate it through UNO if testing this option. Starting a process and converting a file for each edit adds substantial avoidable work. **Warm-edit latency UNSURE.** | Full Calc document state per loaded workbook, plus process overhead. **MB/copy UNSURE.** | Portable Windows distribution exists; exact headless/UNO packaging must be tested offline. | Useful secondary comparator or batch fallback. Excel remains the correctness reference. |
| **Extend the custom Python interpreter** | Could become incremental with persistent parsed formulas, reverse dependencies, range indexes, and revision-aware caches. Its current visible-range evaluation does not establish global invalidation correctness. | Python objects and caches per copy unless explicitly redesigned around shared immutable structures. **MB/copy UNSURE.** | Best existing deployment fit. | Preserve useful parts, but avoid turning it into a general Excel implementation without proving that a small set of formula families makes this tractable. |
| **Univer headless/Node** | Formula architecture describes dirty dependencies and execution queues. Full workbook compatibility and production latency are **UNSURE**. | Workbook model plus formula graph per instance. **MB/copy UNSURE.** | Bundled Node/packages feasible; verify exact OSS/Pro package boundary. | Exploratory candidate, not my first server engine. |
| **Aspose.Cells .NET** | Supports calculation-chain reuse and documents broad modern-function coverage. **Actual incremental behavior, latency, and memory on this workbook remain UNSURE.** | Workbook, calculation chain, and caches per instance unless measured otherwise. | Commercial library; a prebuilt application with bundled runtime is a plausible no-admin package, subject to clean-machine verification. | **Commercial fallback to benchmark.** |

Primary evidence: [Formualizer architecture](https://www.formualizer.dev/docs/core-concepts/dependency-graph-and-recalc), [large-workbook behavior](https://www.formualizer.dev/docs/guides/large-workbook-performance), [Windows wheel](https://pypi.org/project/formualizer/); [IronCalc evaluation source](https://docs.rs/ironcalc_base/latest/src/ironcalc_base/model.rs.html), [Windows wheels](https://pypi.org/project/ironcalc/); [HyperFormula performance](https://hyperformula.handsontable.com/docs/guide/performance.html), [licensing](https://hyperformula.handsontable.com/docs/guide/licensing.html); [LibreOffice Portable](https://portableapps.com/apps/office/libreoffice_portable); [Univer formula architecture](https://docs.univer.ai/zh-CN/blog/formula); [Aspose calculation-chain documentation](https://docs.aspose.com/cells/net/calculate-formulas/).

The critical function comparison is:

**Listed** means documented presence, not verified Excel parity. **Partial** means a documented mismatch. **UNSURE** means current evidence was insufficient.

| Engine | LET | XLOOKUP | FILTER | UNIQUE | SEQUENCE | SWITCH | NUMBERVALUE |
|---|---|---|---|---|---|---|---|
| Formualizer | Listed | Listed | Listed | Listed | Listed | **UNSURE** | Listed |
| IronCalc | Listed | **Partial: reported computed-array defect** | Listed | Listed | Listed | Listed | Listed |
| HyperFormula 3.5 | **Absent from current list** | Listed | **Partial** | **Partial** | **Partial** | Listed | **Absent from current list** |
| LibreOffice | Listed | Listed | **Array/spill differences** | **Array/spill differences** | **Array/spill differences** | Listed | Listed |
| Univer | **UNSURE** | **UNSURE** | **UNSURE** | **UNSURE** | **UNSURE** | **UNSURE** | **UNSURE** |
| Existing interpreter | **UNSURE** | **UNSURE** | **UNSURE** | **UNSURE** | **UNSURE** | **UNSURE** | **UNSURE** |
| Commercial Aspose.Cells | Listed | Listed | Listed | Listed | Listed | Listed | Listed |

Sources: [Formualizer reference/repository](https://github.com/PSU3D0/formualizer/), [IronCalc functions](https://docs.ironcalc.com/functions/), [reported XLOOKUP defect](https://github.com/ironcalc/IronCalc/issues/1338), [HyperFormula built-ins](https://hyperformula.handsontable.com/docs/guide/built-in-functions.html), [LibreOffice functions](https://help.libreoffice.org/latest/en-US/text/scalc/01/04060107.html?DbPAR=CALC), [Aspose supported functions](https://docs.aspose.com/cells/net/supported-formula-functions/).

The important compatibility findings are:

- **Formualizer needs a compatibility gate immediately.** Its documented `ISFORMULA` currently always returns `FALSE`; your workbook uses it 234 times. Resolve it against original formula metadata, or fix the engine, then test. `CELL` is absent from the published information-function registry, so treat those four uses as missing until proven otherwise. `INDIRECT(...,FALSE)` is documented as unimplemented. `SWITCH` semantics remain **UNSURE**, which matters with 87,000 uses. [ISFORMULA](https://www.formualizer.dev/docs/reference/functions/info/isformula), [information registry](https://www.formualizer.dev/docs/reference/functions/info), [INDIRECT](https://www.formualizer.dev/docs/reference/functions/reference-fns/indirect)
- **HyperFormula’s array limitations are substantive.** Its current documentation says dynamic arrays are unsupported, limits `FILTER`, and requires array dimensions to be known at parse time. It also restricts `INDEX` and `OFFSET` behavior relevant to your broader list. [Known limitations](https://hyperformula.handsontable.com/docs/guide/known-limitations.html)
- **LibreOffice function names do not establish Excel spill compatibility.** Its documentation describes array entry and non-resizing `SEQUENCE` results. Test imported spill anchors, changing dimensions, and downstream references explicitly. [SEQUENCE documentation](https://help.libreoffice.org/latest/ug/text/scalc/01/func_sequence.html?DbPAR=CALC)
- **The remaining common functions still need semantic tests.** Check lookup duplicates and wildcard escaping; blanks versus empty strings versus zero; `IFERROR` error propagation; `NUMBERVALUE` separators; `TEXT` formatting; scoped defined names; and `SUMIFS` coercion. The prototype’s 101/103 sample parity does not certify the complete function list.

For **in-browser calculation**:

| Option | Expected behavior | Memory/deployment implication | Recommendation |
|---|---|---|---|
| Univer engine | Incremental design, but exact coverage **UNSURE**. Use a Worker if tested. | Browser holds UI model plus calculation state; worker messaging may add copies. Vendored assets are feasible. | Evaluate only if it passes the same parity gate. |
| HyperFormula | Same function and array gaps as Node. A Worker prevents main-thread calculation stalls, not incorrect results. | Full engine state per tab/session. | Reject for this workbook. |
| IronCalc WASM | Moving to WASM does not remove the current full-model evaluation behavior. | WASM memory plus browser UI state and transfer buffers. Actual tab limits **UNSURE**. | Not the full-workbook v1 choice. |

**Hybrid is the recommended deployment architecture:** browser handles editing, selection, formatting, scrolling, and optimistic input display; server produces authoritative calculated values.

I would defer a *dual-calculation hybrid*, where both browser and server evaluate formulas. It adds reconciliation and two compatibility surfaces. If later justified, restrict client calculation to a proven closed dependency subset, pin the same engine/version, and label its results provisional.

For memory perspective, **1,894,479 hypothetical 16-byte result slots occupy about 28.9 MiB**. That arithmetic excludes formulas, strings, graphs, lookup indexes, styles, spill buffers, and allocator overhead. It is **not an engine RAM estimate**. The benchmark must measure both steady-state memory and peak import/recalculation memory.

**2) Making N copies cheap**

The existing immutable baseline plus input overrides is the right persistence model. Extend that principle to calculation carefully:

1. **One immutable template version.** Store its hash, formulas, static values, styles, names, dimensions, and input permissions once. New copies reference it. A fork captures its parent’s overrides at a specific revision; subsequent parent edits must not change the fork.
2. **Separate shared computation from scenario computation.** Precompute pure formulas that cannot depend on editable inputs. Share interned formula templates and static lookup indexes where the engine permits it.
3. **Per-copy mutable state:** input overrides, affected calculated values, dirty state, spill geometry, copy-specific indexes, and calculation revision.
4. **Use one warm engine instance per actively calculating copy, shared by all its editors.** Five users editing one copy should not create five engines.
5. **Bound the warm pool by measured memory.** Evict least-recently-used idle instances; retain durable inputs/history. Pin in-flight jobs, apply backpressure, and measure reload latency. A persisted copy can be cheap even when keeping it warm is expensive.

The desirable memory model is:

`shared template/graph/indexes + sum(per-copy overrides + divergent results + execution state)`

But **none of the reviewed evidence establishes that the chosen engine exposes this full shared-model/COW design**. Start with bounded independent instances; do not claim sparse database copies imply sparse engine memory.

Also, **Windows does not give Python workers Unix-style fork COW**: its multiprocessing default starts a fresh interpreter. Shared immutable memory requires deliberate support, such as memory-mapped buffers and an engine capable of using them. [Python process-start documentation](https://docs.python.org/3/library/multiprocessing.html#contexts-and-start-methods)

**One mutable engine switched between copies** saves resident instances but creates a correctness trap. A safe switch must reset every differing input, clear or correctly invalidate scenario-specific caches and spills, and recalculate before serving results. It also serializes unrelated copies. Use this only as a measured memory-saving fallback with exclusive leases and explicit reset semantics.

A baseline cached result is reusable only if its dependency state remains valid for that copy. Comparing the cell’s own input value is insufficient.

**3) Grid UI choice**

**Rank: 1. Univer OSS; 2. Glide Data Grid, conditional on filling its frozen-row gap; 3. FortuneSheet for a bounded prototype.**

| Library | License and required features | Main issue for this app | Decision |
|---|---|---|---|
| **Univer OSS** | Apache-2.0 core; spreadsheet editing, formats, fill, clipboard, frozen rows/columns, undo/redo. Pro covers import/export, collaboration/history, and server features. | Viewport rendering does not prove lazy loading of its workbook model. Exact style fidelity and discontiguous paste **UNSURE**. | **First choice.** Disable client calculation and feed server values. |
| **Glide Data Grid** | MIT; canvas grid with on-demand cell callbacks, editing, fill, rectangular paste, multiple rectangle selections. | Frozen columns supported; arbitrary frozen top rows are a gap. Undo/redo requires application work. | Best fallback for controlling browser data loading, but incomplete spreadsheet behavior. |
| **FortuneSheet** | MIT; spreadsheet editing, fill, formats, selection, paste, freeze, undo/redo. | Sparse input is converted into a dense row/column model. Large-sheet memory is a concern. | Third; require a memory/scroll test before commitment. |
| **x-spreadsheet** | MIT; canvas spreadsheet with basic editing, autofill, freeze, undo/redo. | Core selection model is limited; no proven server paging or workbook-fidelity path. | Lower fit. |
| **Jspreadsheet CE** | MIT CE; basic editing, drag-copy, rectangular paste, undo/redo, frozen columns. | Do not attribute commercial-version features to CE. Arbitrary frozen rows were not established in CE. | Does not meet the complete requirement as verified. |
| **AG Grid Community** | MIT; strong virtualization and ordinary data editing. | Required cell-range selection, fill handle, and clipboard modules are Enterprise. | **Reject Community for this feature set.** |
| **RevoGrid** | MIT core with virtualization and editing. | Multi-range selection, smart autofill, undo history, and other relevant features cross into Pro. Pinning differs from Excel freeze panes. | Reject core-only proposal for the full requirement. |
| **Handsontable** | Strong spreadsheet interaction, but current production licensing is commercial; free key is non-commercial/evaluation. | Company-internal production use does not satisfy the permissive-license requirement. | Reject under that requirement. |

Sources: [Univer OSS/Pro boundary](https://github.com/dream-num/univer#-open-source-and-pro), [disable calculation](https://docs.univer.ai/guides/sheets/features/core/formula); [Glide API](https://github.com/glideapps/glide-data-grid/blob/main/packages/core/API.md), [freezing props](https://docs.grid.glideapps.com/api/dataeditor/important-props); [FortuneSheet](https://github.com/ruilisi/fortune-sheet), [data model](https://github.com/ruilisi/fortune-sheet/blob/master/docs/zh/guide/sheet.md); [x-spreadsheet](https://github.com/myliang/x-spreadsheet); [Jspreadsheet CE](https://github.com/jspreadsheet/ce); [AG Grid cell selection](https://www.ag-grid.com/javascript-data-grid/cell-selection/), [clipboard](https://www.ag-grid.com/javascript-data-grid/clipboard/); [RevoGrid feature boundary](https://github.com/revolist/revogrid); [Handsontable license](https://handsontable.com/docs/12.2/javascript-data-grid/software-license/).

Three requirements need explicit acceptance tests:

- **Actual dimensions:** test `4088 × 312`, not just `4000 × 300`, and separately the `100000 × 17` sheet. Measure total browser heap as well as rendered cells. [Workbook extents](/Users/gopalmacbook/Desktop/Max%20HQ/price-planning/.planning/excel-grid/FACTS.md:11)
- **“Multi-range paste” needs a precise contract.** Multiple selections, rectangular TSV paste, and preserving several discontiguous source/target ranges are different capabilities. None should be inferred from a generic “clipboard supported” claim.
- **Formatting requires an importer/adapter.** Preserve number formats, dates, row heights, column widths, borders, fills, alignment, merges, and relevant conditional formats. Keep the XLSM server-owned; do not route it through a grid exporter and assume preservation.

Your HTTP LAN address also matters: asynchronous Clipboard API access requires a secure context. Keyboard clipboard events can still support conventional paste, but toolbar clipboard operations and rich content need testing at the actual deployment URL. [MDN Clipboard API](https://developer.mozilla.org/en-US/docs/Web/API/Clipboard_API)

**4) Collaboration: cell operations and versions, not CRDT for v1**

Use a **server-ordered operation stream per copy**. Fixed formulas, fixed structure, online LAN use, and no merge-back make this a good fit.

An edit operation should contain a unique operation ID, copy/template identity, target cells, typed values, and expected versions for the cells it changes. Derive the actor from the authenticated session.

The processing contract should be:

1. **Validate the entire batch:** permissions, editable-cell mask, types, limits, and expected versions. Disjoint edits can both succeed; a stale edit to the same cell returns a conflict.
2. **Commit atomically:** input changes, new cell versions, copy revision, history, and a durable recalculation job/outbox record.
3. **Calculate outside the database transaction** against a consistent revision.
4. **Publish revision-tagged results.** Track `input_revision` separately from `calculated_revision`; show pending/stale state while they differ. Never label an old calculated value current.
5. **Recover idempotently after failure:** replay durable edits, reconstruct the target revision, publish a complete result batch, and prevent an older job from replacing newer results.

WebSocket transport is suitable. The current Waitress/WSGI server is not itself the WebSocket implementation; use an appropriate bundled transport layer or sidecar. SQLite WAL remains reasonable for this scale if transactions stay short and the database is on the server’s local disk. WAL permits concurrent readers but only one writer at a time. [Waitress scope](https://docs.pylonsproject.org/projects/waitress/en/stable/), [SQLite WAL](https://www.sqlite.org/wal.html)

For user-facing behavior:

- **Attribution:** input cells show who changed them, when, and the operation. Formula cells show calculated status or triggering revision, not a misleading “edited by” label.
- **History:** retain before/after typed values, actor, revision, and batch ID. Group a paste or fill into one history action.
- **Undo:** submit a compensating operation with version checks. Do not silently undo another person’s later edit.
- **Paste/fill across locked cells:** preflight and reject the whole invalid operation with a useful explanation. Avoid silent partial application.
- **Presence:** cursors and selections are ephemeral messages, separate from durable edits.

A CRDT becomes more attractive with offline editing, structural row/column changes, or independently writable replicas. It does not eliminate business conflicts over which price should win, authorization, or calculation ordering.

**5) Top five risks and the deciding benchmark**

| Rank | Risk | Required response |
|---|---|---|
| **1** | Plausible but incorrect results from semantic differences, especially errors hidden by `IFERROR` | Differential testing against fully recalculated Excel, including types, errors, spills, and displayed business values. |
| **2** | A common input invalidates enormous lookup/aggregate chains | Measure affected formulas and scanned range sizes for real editable inputs. Test the highest-fan-out permitted edit. |
| **3** | Copies duplicate graphs, strings, indexes, and caches until the server starts paging | Measure peak and steady-state memory at several active-copy counts; enforce a bounded warm pool. |
| **4** | A virtualized grid still loads too much data or fails Excel interaction/fidelity | Test actual sheet dimensions, style density, freezing, paste, fill, undo, and server patches. |
| **5** | Saved inputs and published calculations diverge after races or crashes | Durable jobs, idempotent replay, revision-tagged results, conflict handling, and crash-recovery tests. |

**The one benchmark should be an end-to-end “edit → correct revision” replay of the actual workbook on the actual office Windows machine.**

Use the same replay against Formualizer, the existing interpreter, and the commercial fallback if available:

- Start from an Excel-recalculated baseline, then obtain Excel reference results after the same edits. Old cached XLSM values alone are insufficient.
- Include narrow, typical, and highest-fan-out input edits; a large paste; blank/text-number/error cases; and changes that resize arrays or change lookup results.
- Exercise five editors on one copy and a proposed stress case of ten active copies. These are benchmark scenarios, not assumed production requirements.
- Measure cold import time, first calculation, warm p50/p95/p99 edit latency, time until authoritative results are publishable, peak memory, incremental memory per added copy, eviction/reload time, and crash recovery.
- Compare baseline and post-edit outputs across the workbook, not just visible cells. Require exact type/error/spill agreement and business-visible numeric agreement; document any permitted floating-point tolerance explicitly.

**Proposed acceptance targets:** p95 ordinary edits within **500 ms**, p95 broad edits within **2 seconds**, no silent correctness mismatches, no cross-copy contamination, and adequate RAM headroom without paging at the agreed concurrency. These are targets, **not forecasts**. Run the replay twice consecutively before accepting a candidate.

The highest-impact missing measurement is **the fraction of the 1.9 million formulas reachable from editable inputs**. If most formulas depend only on immutable historical data, they can be computed once and shared. That could change the capacity requirement more than choosing between two engines.

```text
STATUS:   DONE_WITH_CONCERNS
CHANGED:  none - design critique only
TESTS:    FACTS.md and primary documentation reviewed; no workbook benchmark run
GATE:     approved design critique
LESSON:   none
NEXT:     Run the actual-workbook edit-to-correct-revision benchmark

DONE:     Ranked architecture, engines, UI libraries, copy model, and collaboration.
VERIFIED: Documented capabilities, license boundaries, and named compatibility gaps.
LEFT:     Runtime parity, latency, memory, and exact UI behavior remain unverified.
```