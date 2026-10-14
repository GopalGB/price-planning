# STATE

## 2026-10-14 (Claude, own tabs and columns, validation, Excel mode)

feat/v2-integration. Sir: "Validation formula is not supported" when typing; let users upload their
version with extra tabs or calculation columns and show what exists and what is new; macros say they
run on the office server; BU recap values do not flow to other tabs. Grill: keep extras in the copy,
show them in the app, editable like inputs.
- abd0366: rules the app cannot evaluate (custom formulas, lists from a range, formula bounds) accept
  the value instead of blocking the cell. The real template's custom rules on 2 sheets blocked entry.
- 6131814 (plan .planning/excel-grid/PLAN-extras.md): extra tabs and columns at or right of a template
  sheet's first empty column are imported once (template importer, Excel's saved values) and grafted
  into the template package (pp/export/extras_graft.py, written by a Haiku executor, reviewed).
  pp/extras.py CopySnapshot serves them via TemplateCatalog.for_copy: tabs at their file position with
  indexes after the template's, formulas show saved values, other cells open for typing, no validation.
  Downloads keep them in the user's tab order with fullCalcOnLoad. Preview lists "Kept in your copy" and
  "Removed from your copy"; a second upload replaces extras and their values. Columns inserted in the
  middle are refused (a template column emptied, or a heading found in another column).
- Real template (Mac, lan-check): own tab + 59-formula column: preview 87 s (import + graft), confirm
  instant, rows render, export 13 s with VBA and formulas, peak 1.46 GB.
- ac5129d: office_run.bat forces PP_ENGINE=excel when Excel is registered (an older settings.json could
  keep values mode) and prints EXCEL MODE or VALUES MODE. Macro and slicer messages now say the server
  runs without Excel. Cross-tab recalculation and macros need the Excel engine; values mode never
  recalculates. Still unverified on the office server.
- Open: extras in Excel-engine calculation need the engine's caches per grafted package (untested);
  forks do not inherit extras; extras folders are not in backups.

## 2026-10-14 (Claude, upload adapts to moved rows, .xlsb)

feat/v2-integration. Sir: edited downloads with more rows were rejected as moved, and "File is not a
safe Excel workbook".
- Cause 1: any extra row grew the sheet extent, so the whole sheet counted as "moved" and every change on
  it was ignored (reproduced on the real template: 3 appended rows lost all edits on 2 sheets).
- Now (pp/upload_align.py): sheets whose formula layout or locked labels moved are matched row by row
  against the copy (template plus saved edits) with an order-keeping alignment. A moved mapping is only
  used when every run of rows is proven by an unedited value its neighbours do not share, every pair
  shares a value or is boxed in by proven pairs, and breaking ties the other way gives the same answer;
  otherwise the sheet is refused with the reason. Rows added at the end or into blank lines apply;
  a row inserted inside the table is named and left out; deleted rows are named and left unchanged;
  an emptied row is a clear. Columns that moved (occupancy and unchanged values) and values in locked
  cells past the last filled column refuse the sheet. Template identity is version 5.
- Adversarial review (Sonnet) found 2 HIGH and 4 MEDIUM before push (no-label tables shifting
  silently, column insert clearing inputs, trailing delete clearing, identical rows, false column
  shift, crossed pairs); all fixed and turned into tests/v2/test_upload_align.py (random test: 0 wrong
  rows whenever a move is claimed).
- Cause 2: .xlsb/.xls/password files hit the generic message. Now named, and on the Excel engine
  converted with Excel Save As (pp/engine/excel.py convert_to_xlsm, untested on real Excel). Other
  unreadable files name the part and the error type; the server log has the detail.
- Files without download information are matched by sheet names; differing formulas are noted.
- Upload preview now has "Your file compared with the template": per sheet, rows with data (file vs
  template), columns empty in the file and columns only in the file (letter and top text), missing and
  extra sheets; matching sheets collapse. Real template: 20 same, the 2 edited sheets +3 rows.
- Real template (Mac): append preview 23 s; insert on 2 sheets 47 s. Tests: 847 Python OK, 198 web OK,
  content guard 0.

## 2026-10-13 night (Claude, E2E readiness, logo, 10-user load)

feat/v2-integration at bdbd9c2. CI green through 8a4771e; bdbd9c2 was running at handoff.
- E2E run on the real template (lan-check, port 8793) found and fixed two bugs (47932b8): the Upload
  button was hidden in values mode, and 17 of 22 sheets counted as "moved" because the extent check
  compared against the snapshot instead of the template's own extents (identity file version 3).
- Logo (8a4771e): logo.png or logo.svg (max 5 MB) in the app folder or its parent shows top-left
  with no brand config. Office: put it in `D:\Gopal\PP 2` or `D:\Gopal\PP 2\app`, restart.
- Speed (bdbd9c2): unedited copies share row blocks per template; colour-rule formula prep is cached
  per process; ROW/COLUMN and defined-name rewrites skip the token walk when absent. Differential on
  the real template: 473,003 rewrites, 0 differences.
- 10 users, 90 s, random sheet and block every 0.5 to 2 s, one upload plus one download mid-run (Mac
  M2, 8 cores, 8 GB): sheet load p50 1.1 s (was 2.9 s), p95 9.0 s, max 21 s; edits p50 284 ms,
  p95 1.8 s; 0 errors. Under that load: download 95 s, upload 105 s. Server CPU avg 2.2 cores, max
  4.1; memory idle 284 MB, peak 4.5 GB (upload workers).
- Known cost left: a fully cold colour-rule build of the largest sheet's first block is 34 s; the
  1.5 s budget skips the rest and the cache fills over later renders. Warm renders still tokenize
  each formula three times (CF ReferenceScan.collect, engine _expand_defined_names, _trace_budget),
  about 0.6 of 1.8 s. Next speed step: memoize tokens per formula text, or pre-warm block 0 per
  sheet at startup.
- Tests: 831 Python OK, content guard 0, ruff no new findings.
- Still waiting on IT: Excel Quick Repair, then `office_run.bat check` (P6.1, 6.3, 6.4, 6.5).
- Excel without the repair: when `Excel.Application` is unregistered ("Invalid class string") the app
  now starts Excel by its class id, and office_run.bat accepts either key; if neither exists it prints
  the Click-to-Run platform and registration lines for Claude and stays in values mode.

## 2026-10-13 evening (Claude, upload and download speed)

feat/v2-integration at 2392afe. CI green through 8c8124a; later pushes not yet checked here.
- Uploads and downloads run as background jobs (`?async=1`, `GET /api/jobs/{token}`, pp/jobs.py) with
  a progress bar; the upload dialog pre-fills the copy name from the file name.
- Upload preview on the real template (Mac, 8 GB): 350 s originally, 147 s, 39 s (1ec2989), 31 s now
  (2392afe: 4.4 s safety checks, 26 s reading sheets). Big workbooks read one sheet per spawn worker
  (MAX_SCAN_WORKERS 4, PARALLEL_MIN_BYTES 32 MB), one parse per sheet; a dead worker falls back to an
  in-process read. LinkScan skips the link parse for screened UTF-8 parts with no prefixed
  id/embed/link attribute (calcChain, pivot cache records).
- Download (export) on the real template: about 26 s to 12.6 s (b738e49): the final package check
  skips parts copied byte for byte whose .rels were also copied.
- Overlay fix (b58594c): note and filter markers use dashboard.mjs worksheet geometry and are placed
  again after a zoom change. Checked in Ego Lite on the synthetic fixture at 75/100/125/150%; the
  fixture is too small to scroll, so frozen plus scrolled placement is covered by unit tests only.
- Tests: 822 Python OK, 198 web OK, content guard 0, ruff F clean, bandit -ll exit 0, prettier clean.
- Adversarial review (Sonnet) of the upload safety changes found one real hole: a UTF-16LE part with
  no BOM passed LinkScan as UTF-8, so a dangling r:id was accepted on the in-process path (a
  regression from 2392afe; 4543416 rejected it). Fixed: any NUL byte forces the full parse; tests
  for both encodings, mutation-checked. Everything else it tried held (chunk boundaries, declared
  encodings, duplicate entries, deferred-but-unread parts, fallback path, export skip).
- Next: Sir pulls on the office server (port 8211) and retries upload.
- 18:24: lan-real (8765) had stopped (graceful shutdown, cause unknown); restarted 19:12.

## 2026-10-13 (OpenCode runner, perf/tab-cache)

perf/tab-cache at 91b0367, four commits from 7b0104e. Opening a sheet tab no longer renders
and downloads the same window twice.
- pp/row_cache.py: byte and entry bounded LRU (256 MB, 512 entries) of rendered row bodies. The key
  is copy id, template id, sheet, start, count, input_revision, calculated_revision and the copy's
  stored slicer selection hashed by content. Access is checked before the cache is read, so a user
  without a role still gets 404. A block is served again after any edit, calculation or stored
  slicer change; see `_row_cache_key`.
- pp/server.py: JsonGZipMiddleware compresses application/json over 1 KB (level 6). Workbook
  downloads are already zip containers and are passed through untouched, as are the WebSocket and
  streaming responses. One test cleared the row cache because the CF evaluation budgets are process
  tuning, not copy state (test_fidelity_cf).
- web: SheetRowCache/SheetCacheMap keep the last six visited sheets (least recently visited first
  out) and CellStore gained evictSheet. A sheet's blocks are reusable only when every one was
  rendered at the revisions the client is on, so a remote op or calc invalidates the other sheets,
  and a stale window is asked for again rather than answered from the old cache. Per-sheet zoom
  unchanged.
- Bench (`ops/bench_rows.py`, synthetic only): 800 x 312 protected sheet with a per-row conditional
  format rule, 256-row block = 4,435,955 JSON bytes (243,550 gzipped); first visit 0.331 s, second
  visit 0.001 s. Default fixture 2 x 1200 x 24: 339,603 bytes (20,607 gzipped), 0.057 s then 0.001 s.
- Tests: 804 Python OK (134.6 s), 186 web OK, prettier clean, build committed, content guard 0,
  bandit `uvx bandit -r app ops pp -ll` exit 0, ruff clean on the touched Python.
- Not run here: a real browser check, the office server, Excel, semgrep, pip-audit, gitleaks. No
  external model review; the supervisor owns that.
- Report: .planning/oc-report.md.

## 2026-10-13 (Claude, production readiness)

feat/v2-integration. CI fully green for the first time (tests 3.10/3.12, security, windows-offline x3).
- Security: defusedxml for every workbook parse (pp, app, ops); bandit now scans pp; CI actions pinned;
  nginx example forwards the live WebSocket only, 100 MB uploads. pip-audit clean, gitleaks clean.
- Daily automatic backup while run_v2 serves (newest 7, "auto_backup": false turns it off).
- Values mode (no Excel automation): typed values save, formulas stay locked. Ego Lite check found the save
  chip stuck on "Calculating..."; fixed by an empty calc message in values mode (stored revisions stay honest
  so the engine recalculates queued jobs once Excel works). Per-sheet zoom verified in Ego Lite.
- Left for Sir: Office Quick Repair on the server, office_run.bat check 4/4, real edit+download test,
  firewall port, merge PR #2 to main, repo back to private. Branch oc/editor-protection is unreviewed
  worker output, not merged.

## 2026-10-12 (OpenCode runner, pp under bandit)

fix/bandit-hardening at 9f330df. CI now scans the v2 server too: `bandit -r app ops pp -ll`, exit 0.
- XML hardening (34 B314 in pp/): every workbook parse in pp/ now goes through defusedxml -
  importer, upload, transfer, xlsm, dashboard, drawings, colors, native_load, formualizer.
  No B314 nosec was needed: defusedxml covered all of them, including the custom-target parse
  in dashboard.py (SafeET.parse with SafeET.XMLParser(forbid_dtd=True); its hand-rolled
  _DashboardXML.doctype guard was dropped because DefusedXMLParser does not dispatch to
  target.doctype, so the guard was silently dead once the parser was swapped).
- Five reviewed false positives annotated, no logic touched: B104 pp/config.py, B608 pp/db.py
  and pp/workbook_service.py, B301 in both engine caches.
- Two tests patched onto pp.upload.SafeET / pp.export.transfer.SafeET. Both had been patching
  ET.parse and ET.iterparse, names the modules no longer use, so the guard they assert was
  already inert; they now guard the parser actually in the path.
- Measured cost: defusedxml iterparse is ~1.6-2.9x slower than the stdlib parser on a 20 MB
  synthetic sheet (min-of-4, A/B, same process). The export budget test
  (test_55mb_xml_subset_path_stays_within_budget, 30 s) measured 21.9/23.1 s at HEAD and
  22.1/28.1 s after, on a host running another suite. Headroom is now thin; watch it.
- Tests: 132 template/upload/transfer OK, 104 upload/db/config/p0 OK, 57 export/db/api OK,
  28 security/ops OK, 26 native, 68 engine, 17 workbook/ops/end-to-end. bandit 0, semgrep 0,
  content guard 0, Ruff whole tree 698 now and 698 at baseline, 0 introduced.
- Not run: the whole suite (another agent holds it), web suite, Office/Excel.
  No external model review; the supervisor owns that.
- Report: .planning/oc-report.md.

## 2026-10-12 (OpenCode runner, bandit hardening)

fix/bandit-hardening at 8394df1 (3 commits after d2a45ca). CI's `bandit -r app ops -ll` was
exiting 1 on 18 MEDIUM findings; it now exits 0.
- XML hardening (51e301e): 11 B314 findings closed. This was a real defect, not a lint nit.
  Entity expansion was live: a 200 kB literal through six levels expanded 64x (12.8 MB)
  before expat's ceiling refused it. app/workbook.py, ops/find_inputs.py and the ops
  dashboard fixture now parse through defusedxml, already pinned and vendored since d2a45ca.
  Two tests in tests/test_p0_regressions.py (XmlEntityRefusal) pin refusal at both entry
  points; both failed before the change.
- B104 (76690a1): bandit's rule fires on the literal 0.0.0.0 wherever it appears, including
  the two comparisons that reject it as a public origin, so one deliberate LAN deployment
  choice produced six findings. The literal now lives in one ALL_INTERFACES constant with
  the justification and a scoped nosec. Behaviour unchanged.
- B608 (8394df1): both findings concatenated fixed literals, so nothing was injectable, but
  bandit reads a BinOp and cannot see that. Complete statements now; no suppression added.
- Checked pp/ (39 B314 findings, outside the CI scan) by driving real entity bombs through
  _validate_archive: already defended. The DOCTYPE screen runs before any parser, returning
  422 unsafe_xml or 422 invalid_workbook. Not widened, and reported rather than assumed.
- Tests: 754 Python OK twice (187.967s, 185.953s), bandit exit 0, content guard 0, Ruff whole
  tree 698 findings now and 698 at baseline, 0 introduced and 0 removed.
- Not run: web suite (no node_modules, no web file touched), semgrep, pip-audit, gitleaks,
  browser, Office/Excel. No external model review; the supervisor owns that.
- Incident: a `git stash pop` with nothing to stash popped another branch's stash@{0}
  (feat/v2-dashboard). Nothing lost; stash preserved and verified. Baselines now use a
  scratch worktree.
- Report: .planning/oc-report.md.
## 2026-10-12 (OpenCode runner, feat/auto-backup)

Daily automatic backup while `run_v2.py` serves. New `pp/auto_backup.py`: `prune(dest_dir, keep=7)`
deletes all but the newest 7 whole backup sets (`planning-<time>.db` plus its `.json` and
`.templates` siblings) and touches nothing else, never a `.partial`; `AutoBackup` runs one daemon
thread (first backup after `first_delay` 300 s, then every `interval` 86400 s), each run is
`backup()` then `prune()`, a failure is logged with `logging.exception` and the loop continues,
`stop()` sets an Event and joins. `run_v2.py` starts it after the runtime exists (live database from
`resolve_database_path`, folder `ROOT/backups`, same as `run_v2.bat --backup`) and stops it in the
existing `finally`; settings `"auto_backup": false` (new validated `Config.auto_backup`, default
true) switches it off. Docs: two sentences in the DEPLOY.md Backup section.
Tests: new `tests/v2/test_auto_backup.py` (6: prune keeps newest 7 of 9 whole sets, `.partial` and
an unrelated file untouched; first_delay=0/interval=0.2 against a real `pp.db` schema database makes
2+ verified backups and `stop()` returns in under 2 s; a raising `backup` is logged and the next run
still happens; the runner starts/stops it and honours `"auto_backup": false`). The shared fake runtime
in `tests/v2/test_run_v2.py` sets `auto_backup=False` so URL tests stay thread-free.
Next: office check is `run_v2.bat`, then confirm `backups\planning-<time>.db` appears about five
minutes after start and that the oldest set disappears after the eighth day.

## 2026-10-12 14:35 IST (Claude, P6.2 look: colors, pictures, shapes)

feat/v2-integration at 6dd69b4 (5 commits after 78bbe4c). IMPORT_VERSION 15 (8765 reimported).
- Colors (0270882): every template color reached Univer as ARGB or theme:N, which Univer reads as
  transparent. Now resolved at import to #RRGGBB (theme + tint, indexed, ARGB; Excel's tint grid
  differs by at most 2/255 per channel). Borders carry line style and color.
- Pictures and shapes (fe496e8, 7e534b9): 3 pictures, 5 linked rounded buttons (links go to cells on
  their sheet), 1 arrow; hidden shapes, slicer fallbacks, empty text boxes left out, as in Excel.
  Macro buttons drawn as filled shapes keep their look. Placed with the slicers' worksheet geometry
  (headers, frozen panes, zoom). No charts in the template.
- Grid fixes (6dd69b4): overlay layer was under the canvas (z-index), zoom dropdown never zoomed
  (operation without the command), and a drawing on the first visible sheet aborted grid.start()
  (scroll state throws on first paint): no live socket, unbound toolbar. Bisected and fixed.
- Live (5282cfb): a socket counts as occupant before calc_state reaches its client (flaky test).
- Verified in Ego Lite on 8765 (geometry/counts only): open + live socket, toolbar bound, picture
  loads, shapes clickable, link lands on its cell, 150% zoom scales overlays 1.5x.
- Tests: 743 Python OK, 161 web OK, ruff clean, guard 0.
- Fixed 2026-10-12 (2c7f158): each sheet keeps its own zoom; the dropdown follows the sheet.
- Fixed 2026-10-12 (e607dfd): engine workers spool into <temp>/pp-engine-spool/<pid>-*, removed on pool
  close, dead-server folders swept on start (not on Windows). Tests use a throwaway PP_ENGINE_CACHE_DIR.
  Tests: 751 Python OK, 162 web OK, content guard 0, prettier clean.
  Hatch pattern fills and group shapes are not drawn (none in this template).

## 2026-10-12 12:34 IST (Claude)

Excel-only pivot and slicers merged (feat/v2-dash-engine) plus b9227f2: the real pivot's cache
reads a sheet of another workbook (worksheetSource with an external-link r:id; earlier notes said
Power Query, which was wrong). It is now run by Excel; its 4 slicers import (10/78/769/52 items,
template selections kept) and show read-only on 8765 (verified in the browser, counts only).
Office checks for Sir: excel_smoke.py, --macros, --refresh, --slicers must each print PASS.
Tests: 736 Python OK, 157 web OK, ruff clean, guard 0.
Earlier entries below had guessed clock times; corrected to undated headings.

## 2026-10-12 (Claude, dashboard merged)

feat/v2-integration at 595808a, pushed (Sir's y; gate clean; no office files on the remote).
- Dashboard track merged (ed352f2): cell-based pivot service, in-grid slicers, report filters,
  personal views, category totals, team overview (copies page "Live now / team"). 699 Python,
  152 web OK.
- Real template: its pivot reads the external Power Query connection, which that service cannot
  rebuild; the import failed and took 8765 down. fff35f6 degrades unsupported pivots to Excel's
  saved values (security/malformed parts still fail). 8765 back up, Coupon opens.
- Running: Excel-only pivot + slicers (PLAN-P6.4 revision) on feat/v2-dash-engine.

## 2026-10-12 (Claude, P6.4 pivots and slicers on the Excel engine, feat/v2-dash-engine)

Excel-only pivots (Power Query source) now keep their slicers: imported as `engine: "excel"` slicers with items and
selection (IMPORT_VERSION 12). Pivot output cells are Excel's, published and never overrides. A slicer click posts
`POST /api/copies/{id}/slicers`, runs in the copy's Excel, commits one `slicer` revision and stores the selection per
copy (`pp_pivot_state`); fresh engines re-apply it. Office check still to run: `ops\excel_smoke.py --slicers` (PASS).
Plan: .planning/excel-grid/PLAN-P6.4-dashboard.md (status section lists the deviations).

## 2026-10-12 (Claude, P6.5 refresh data)

Merged into feat/v2-integration at 7930bbe (648 Python OK, 121 web OK, guard 0). "Refresh data" runs Excel's own
RefreshAll on the copy's private Excel, waits for the queries, and commits what changed as one `refresh`
revision through the macro diff. `refreshable` (IMPORT_VERSION 10) comes from xl/connections.xml.
Route POST /api/copies/{id}/refresh; web toolbar button shown only when refreshable.
- Tests: new tests/v2/test_refresh_data.py (17) plus COM refresh tests in test_excel_macros.py; web 121 OK;
  ruff clean; content guard 0. Bundle rebuilt (static/dist).
- Not run here: a real RefreshAll (no Excel on the Mac). Office check: `ops\excel_smoke.py --refresh`, must print PASS.
- Plan: .planning/excel-grid/PLAN-P6.5-refresh.md. Old feat/v2-refresh (pqtools) stays superseded.

## 2026-10-12 (Claude, macros merged)

Branch feat/v2-integration at c530436 (unpushed; Sir runs `git push`, answers y).
- P6.3 macros merged (M1 buttons, M2 engine, M3 flow): the template's 5 buttons are read from its
  drawings and drawn on their sheets; pressing one runs the real VBA in the copy's private Excel,
  presses OK on its message boxes, saves what it changed as one `macro` revision (values, plain
  cells and new formulas) and publishes results. Mac and values-only mode answer 409.
- Speed fix 1ade04f: rows on big protected sheets took 3-9 s per window (a 16,384-column scan;
  BU RECAP sent 1.9M empty editable cells, the browser grew to 4 GB). Now 5-130 ms warm, about
  1.3 s cold on BU RECAP; sheet switch 1.7-3.4 s in the browser, heap under 100 MB. Editable empty
  cells stop at the sheet's used width (documented shortcut).
- c530436: zero-size leftover VML button skipped (5 buttons, not 6). IMPORT_VERSION 9: servers
  reimport the template once at start.
- Tests: 629 Python OK, 118 web OK, ruff clean (pp, tests/v2), content guard 0.
Office server next (Sir): `"engine": "excel"` in settings.json, then
`.venv\Scripts\python ops\excel_smoke.py` and `... excel_smoke.py --macros`; both must print PASS.
Next here: P6.5 Power Query refresh on the Excel engine (RefreshAll + the macro diff path); the old
feat/v2-refresh worktree (pqtools, based on bcba313) is superseded and left untouched.

## 2026-10-12 (Claude, M3 macros flow)

Branch feat/v2-macros-flow: template buttons run end to end on the Excel engine (P6.3 M3). Press ->
pipeline lock -> Excel run on a worker -> one `macro` revision (inputs, plain values, formula overrides)
-> calc publish. Engine failures are 502 `macro_failed`; Mac/formualizer and cached templates stay 409.
Rows and export show non-input overrides and write formulas as formulas. `ops/excel_smoke.py --macros`
is the office check (prints counts only). Not run here: the office `--macros` smoke (no Excel on the Mac).
Next: Sir runs `excel_smoke.py --macros` on the office server; then the browser check of each button.

## 2026-10-12 10:00 IST (Claude)

Branch feat/v2-integration at b6e10e8 (37 commits not yet pushed; Sir runs
`git push` and answers y at the gate).
- Excel engine built (288e01d, eda0e81): settings "engine": "excel" makes each
  copy's worker drive a private hidden Excel (macros and events off, manual
  calculation, own temp copy of the file). A static dependency index
  (pp/engine/depindex.py, numpy, cached beside the prepared package) tells it
  which formula cells an edit can reach; it reads only those back in row
  blocks and publishes what changed. Import mode "saved": no import engine
  pass, edits allowed. Real file on the Mac: index build 103 s once, cone
  1-9 s per edit (largest cones about 540k formulas).
- Transfer track merged (b6e10e8): download dialog (pick sheets, .xlsx or
  .xlsm) and upload into the working copy or a new copy with a preview.
- Tests: 588 Python OK, 109 web OK, ruff clean.
- LAN 8765 still values-only on the Mac (no Excel here).
Office server next: set "engine": "excel" in settings.json, then
`.venv\Scripts\python ops\excel_smoke.py` must print PASS (it proves no
changed cell is missed and that reverting leaves no overlay). Then: cone speed,
macros as server actions (P6.3), Power Query refresh (P6.5).
MiMo v2.6 Pro: not usable (OpenCode Go timed out, ClinePass 403); Space Bunny
is the only free lane.

## 2026-10-12 09:15 IST (Claude)

Engine on the real file: load now 1.0-1.2 GB (was >25 GB) after trusting
Excel's own-cell array results (b08bb15; 664,077 t="array" cells, none spill).
Trace budgets capped at u32 (TRACE_U32_MAX). Formula map cached per prepared
template (04e9426). First native dependency trace on an edit still passed
3.5 GB: next is replacing _dependent_targets with a precomputed reverse index
(scratchpad reach.py logic), evaluating visible targets only (formualizer
evaluate_cells recomputes stale chains on demand, verified).
UI: sheet arrows, template card on home, values-only copies report
calculated (no "Calculating..."), recovery skipped in values-only mode.
Export verified on the real copy: 22 sheets, VBA kept, 7.9 s.
Not started: import (upload into copy or new copy), macros, data refresh.

09:30 update: visible-window evaluation on an edit also passed 3.5 GB, so
formualizer cannot run realistic edits on this file within the Mac or office
memory. Sir confirmed the office Windows server has Microsoft Excel. Decision:
real Excel via COM as the engine (exact results, VBA macros, Power Query).
All 12 branches pushed to the private repo (gitleaks clean, no office files).
Pre-push hook fixed to find the Max HQ gate from worktrees (fb50ed7, unpushed).

## 2026-10-12 08:10 IST (Claude)

Sir asked again to see the whole template with every tab. Values-only mode
(da0a435, d52dee1, d3c1df1): settings "template_results": "cached" imports
without the engine pass (Excel's saved results), records "results" in
meta.json, opens copies read only (status bar says so), refuses edits with
409 calculation_paused, and gives the app no engine pool so copy creation
never starts a worker. Real file: import 51 s, 134 MB peak; 22 sheets (13
visible, 9 hidden as in Excel); copy opens in 4.3 s; each tab switches in
about 30 ms; browser heap 45 MB. Tests: 446 Python OK, 101 web OK.

LAN: 0.0.0.0:8765 now runs ~/PricePlanning-private/lan-real (real template,
values only, brand). lan-demo is stopped. New database: users sign up again.
Next: cell-level reachability result -> pruned engine package -> turn
template_results back to "engine".

## 2026-10-12 06:20 IST (Claude, free lanes down)

Branch feat/v2-integration at the merge of feat/v2-brand. New since 4b9bc6b:
13a2640 structured table references resolved to absolute ranges before native
load (#All/#Data/#Headers/#Totals, column spans, bare names; [@Col] raises);
da0aa53 + style follow-up: optional "brand" settings block, /api/brand and
/brand/<asset> public, branded app bar and sign-in. Tests: 438 Python OK, 101
web OK, ruff clean, content_guard 0.

Real template (counts only): import and shared/structured expansion now pass.
The engine is the blocker: 1,894,479 formula cells; the Python template map
takes 357 s and 1.1 GB; native load 2.8 s; the FIRST evaluation of any single
cell builds the whole dependency graph and passed 2.5 GB in under 60 s (an
unguarded run reached a 25 GB footprint). EvaluationConfig knobs
(enable_parallel off, max_threads 1, range_expansion_limit 1, flat cache 0)
do not change it. Sheet-level reachability from editable cells keeps 93% of
formulas, so sheet pruning is not enough. Next: cell-level reachability from
input cells (formulas no input can reach keep Excel's cached value and are
dropped from the engine package); measure the live fraction first.

LAN app: 0.0.0.0:8765 runs from ~/PricePlanning-private/lan-demo on a
synthetic demo workbook with the private brand; real template not served yet.


## Current phase
<<<<<<< HEAD

Branch: feat/v2-integration. Native Formualizer load is implemented and measured.
The startup blocker is closed: no formula is registered per cell at load, and
native load reports zero parity mismatches on the synthetic fixture.

Owner decision honoured: native formualizer==0.11.1, one resident engine per
active copy, loaded once from the template, input overlays only, visible-only
affected evaluation, stale blocks calculated on request, atomic per-revision
publication, engine_memory_budget_mb with idle LRU and editor protection.

Root cause of the 47 of 150 mismatch is inside the pinned engine, not our code.
Formualizer 0.11.1 expands a shared-formula follower by translating the master
text with the follower offset, and that translation also shifts a cross-sheet
qualifier whose sheet name parses as a cell address, rewriting S01!$A$1 as
T1!$A$1. Excel never translates the sheet part. Every follower of such a family
read a sheet that does not exist and evaluated to #REF!. Minimal reproduction:
.planning/excel-grid/results/formualizer-shared-formula-repro.py. The library is
not patched. The package is rewritten once per immutable template so every
shared follower carries its own explicit formula, which keeps native load a
single read and needs no per-cell registration.

Scale results, aggregate JSON in
.planning/excel-grid/results/engine-scale-native-2026-10-11.json, synthetic only:

  scale  prepare  load    resident after load  first  warm   broad visible  broad full
  5%     1.91 s   10.03 s 71.8 MiB             0.003 s 0.0002 s 4.13 s       3.26 s
  20%   10.10 s   32.89 s 152.0 MiB            0.014 s 0.0002 s 18.48 s     26.69 s
  50%   15.09 s   92.12 s 263.7 MiB            not reached on this host

The previous per-cell adapter load was 258.73 s at 5% and never finished 20%
inside a 600 s timeout. Load is no longer the blocker. Resident engine size
scales about 1.03 MiB per 1000 formulas, which puts the 1.89M formula template
near 2.0 GiB and inside both the 3000 MB Mac and 16000 MB office budgets.

At 50% the load completed in 92.12 s at 263.7 MiB resident, then the child was
killed by the host during the full-workbook dependency trace at 955 MB on an
8 GiB Mac with swap already above 5 GiB. That is a host limit of this machine,
not a load failure, and it is recorded as failed rather than reported as passed.

Known limits, all measured rather than assumed: visible-only evaluation reduces
evaluation and publication but not the dependency trace, so at 5% it was no
faster than the full edit; the trace is now the dominant edit cost. Editor
protection exists as a pool API but nothing calls protect() yet, because the
connection registry has not been wired. Defined-name expansion still registers
formulas after load; the synthetic fixtures exercise none.

## Last handoff
=======
2026-10-11: fidelity track complete and locally accepted on `feat/v2-fidelity` at 301943d. All five presentation features ship: conditional formatting end to end, drawings with macro actions, notes, table styles with local filters, and sheet views. User GO-TOKEN approved all five and local acceptance. Synthetic fixtures only. Engine and calculation pipeline remained read-only throughout. Plan: `.planning/excel-grid/PLAN-fidelity.md`. Branch is NOT pushed or merged; it merges into feat/v2-integration after the other agent finishes the engine.

## Last handoff
2026-10-11 22:50 final acceptance. HIGH1 f958f06, HIGH2 f64e28d and HIGH3 425ca95 each closed with a regression test; sheet views a9a809c; two browser-found defects fixed in 301943d (overlay layers detached by Univer remount, template scroll cell leaking into the Univer freeze split) plus the rebuilt bundle 17ba91a. Acceptance evidence: Python 321/321 twice, web 124/124 twice, content guard 0, `ruff check pp run_v2.py tests/v2` clean, pinned Prettier clean, two identical production builds (bundle 3efadf59), real-server Ego check on port 8791 with the synthetic eight-sheet fixture, TaskSpace 42 finished with keep: []. Screenshots in results-private/fidelity/shots/, checks in results-private/fidelity/final-ego.json. Disk stayed above 13 GB free.

2026-10-11 resume: the previous agent stopped on a resource guard, not a failure. No uncommitted production work existed. The three HIGH Opus findings were finished first (HIGH1 and HIGH2 already committed), then sheet views, then acceptance. No external model was invoked: this run is headless, so the pending exact Opus re-review is still outstanding and must run in the supervisor lane.

HIGH1 fixed and committedf958f06: local filter apply/clear/eviction visibility mutations use PlanningGrid internal guard. Behavioral RED5 pass/2 fail; focused GREEN7/7; executor relevant grid37/37 and controller37/37. Content guard0, pinned Prettier/diff check green. HIGH2 resumed next in same Sol session. Latest disk10545 MB. Full acceptance remains pending.

2026-10-11 21:04 user-approved resume: current resource rule is `df -m /` Available at least 2500 MB. Swap growth alone is not a stop reason; all older swap ceilings below are superseded. Controller initial disk5108 MB, branch feat/v2-fidelity HEAD4ca148b. Same Sol takeover session resumed for HIGH1, then HIGH2 and HIGH3, each test-first and separately committed, before sheet views and full acceptance. Existing unrun view tests and shared environment symlinks preserved. No push or merge authorized.

2026-10-11 resume STOPPED by current user resource gate. Baseline swap5278.62 MB; ceiling6814.62 MB. Controller measured swap6882.50 MB (growth1603.88 MB), disk3436 MB, so execution stopped immediately. No server or browser check started in this resume, no owned executors remain. Tables/AutoFilter committed7b4ae03 after Python36/36, web31/31, content guard0, global Ruff, pinned Prettier and diff check green. Phase5 added only unrun tests/v2/test_fidelity_views.py and web/test/fidelity-views.test.mjs; production unchanged. Never stage .venv or web/node_modules.

Pending exact Claude review now RAN successfully: response modelUsage contains only claude-opus-5-5, is_error false; verdict FIX-FIRST. Evidence results-private/fidelity/resume-opus-review.json and resume-opus-result.md. Three HIGH findings require regression tests and fixes: (1) local hide/show rows are cancelled by PlanningGrid.command because calls lack the internal mutation guard; (2) pp/server.py async rows handler blocks the event loop during synchronous CF work, use existing asyncio.to_thread pattern; (3) one formula-result error rolls back healthy CF targets, treat error result as false while retaining EngineError handling. MED findings to investigate in same pass: deterministic function batching/template-block reuse, explicit cell styling precedence over table styles, overlay scroll redraw/asset caching. Header clipping needs coverage; table-order memory claim is unverified and must not be accepted without source evidence. Full review is preserved verbatim. No shipping/acceptance claim.

Security: branch gitleaks f74d131..7b4ae03 no leaks; pinned Python audit no known vulnerabilities; production web audit0. Semgrep p/python flagged only fixed committed table-catalog ET.parse sink; trusted-path triage in resume-security-triage.md. Final all-source scan/full web dependency audit, coverage, planner review, duplicate full suites/builds and browser remain pending.

2026-10-11 resumed with current user limits: minimum 2500 MB free; swap baseline 5278.62 MB, stop above 6814.62 MB. Initial disk 6075 MB. Tables/AutoFilter controller checks pass: focused Python 36/36 and web 31/31, content guard 0, global Ruff, pinned Prettier and diff check green. Committing Phase 4 now. Exact Opus review and Phase 5 next. Historical resource limits below are superseded by this resume.

2026-10-11 19:33 IST resume: STATUS BLOCKED by the user's 4 GB disk stop threshold. First checks confirmed branch feat/v2-fidelity and preserved uncommitted tables/AutoFilter work in ops/fixtures.py, pp/template/importer.py, pp/workbook_service.py, web/src/grid.mjs, web/src/main.mjs, plus new table-style source/licenses, pp/fidelity/tables.py, tests/v2/test_fidelity_tables.py, web/src/tables.{mjs,css} and web/test/fidelity-tables.test.mjs. Untracked .venv and web/node_modules are present; never stage these. df -g / reported 3 GB available; confirmation df -k / reported 3323888 KiB available. Swap baseline: used 9912.44 MB, total 11264 MB. No tests, source edits, commits, model invocation or browser acceptance ran in this resume. Only checkpoint documentation was updated. The historical Claude quota blocker has not been rechecked; pending exact claude-opus-5-5 review is deferred by the resource stop. No files were deleted and no private workbook was read. Resume only with at least 4 GB free and swap growth no more than 2 GB above this baseline; run focused tables tests first, fix if needed, commit the feature, then pending exact review and sheet views. Keep TaskSpace 42 retained.

2026-10-11 wide CF correction checkpoint: 50x312 produces15600 literal overrides; 256x312 produces79872, with no warnings and one public engine load/request for TRUE(). Target bound now100000, aligned with override cap. Only proven-pure formulas alias; volatile calls, named references and unknown functions stay per-target, with scoped-name and rollback regressions green. Scoped Python30/30 and CF11/11 pass, webCF5/5 unchanged; Ruff/content guard/diff check green. Notes commit1ae50ee retained. Resume Phase4 from next-tables.md; no table source was changed before the interruption. Final Opus review remains quota-blocked.

2026-10-11 controller scale probe: notes committed1ae50ee. A valid synthetic50x312 TRUE() CF range produces0 overrides instead of15600 because of formula_target_limit, while10x312 produces3120 correctly. Reproduction results-private/fidelity/probe-cf-broad.json and script. Phase4 was interrupted during preparation, no source changes yet; close this CF budget regression with tests before resuming table work. Original final Opus quota blocker remains.

2026-10-11 Phase 3 implementation checkpoint: drawings committed 082fea3. Notes import two synthetic comment parts, stream with requested row blocks, and render red merged-cell-aware markers with safe DOM hover/focus popups. Eviction/sheet-switch cleanup and overlay shortcut guards are tested. Python 27/27 scoped tests passed; final JS-only popup correction passes 4/4 focused and 23/23 notes/drawings/grid regressions, with pinned Prettier, content guard and diff checks green. Final Opus quota blocker and whole-track acceptance gates remain pending. Next: tables/AutoFilter, then sheet views.

2026-10-11 Phase 2 implementation checkpoint: CF committed f764aaa. Drawings/assets/actions implemented and locally tested: relevant Python 32/32, web 26/26; content guard 0, Ruff, pinned Prettier and diff check green. Synthetic fixtures only; no engine/calc pipeline changes. Drawing import streams and caps 500 anchors, 100 image references, 5 MiB per asset and 25 MiB total. Public Univer geometry drives DOM overlays because sheets-drawing is absent from the pinned preset. Macro actions require copy write access/CSRF and return 501. Final exact Opus review still quota-blocked; acceptance/build/browser gates pending. Next: notes, then tables/filters and sheet views.

2026-10-11 CF implementation checkpoint: F1-F7 regression fixes pass full Python 302/302 and web 105/105. Ruff including the fixture module, pinned Prettier, content guard, staged secrets and scoped SAST are green. Mandatory fresh claude-opus-5-5 re-review is BLOCKED by provider HTTP 429 usage_limit_reached: weekly limit, reported reset 7:30pm (Asia/Calcutta). No substitute, credential changes or paid lane used. Code may be committed locally as unaccepted implementation; final acceptance remains blocked. CF helper now returns ConditionalFormatResult(styles,warnings); rows stay 200 with generic conditional_format_warnings, workbook sheet has_conditional_formats flags, bounded shared public-engine loads, and coalesced refresh only for CF sheets. Next implementation: drawings/actions. Final review/browser/suites-twice/builds still pending; TaskSpace 42 retained.

Resume checkpoint, 2026-10-11: disk verified above 12 GiB; exact gpt-5.6-sol session 01a12961-bb90-71b3-af7a-07bc42b2cd29 resumed and exited cleanly. CF remaining regression fixes are present. Controller full Python 300/300 and web 103/103 passed; content guard exit 0, project Ruff (`pp run_v2.py tests/v2`) green, pinned Prettier green after formatting. Staged secret scan and scoped Python SAST passed. Fresh exact Opus 5.5 returned FIX-FIRST with verified canonical model identity. Its F1-F7 findings are being fixed in the same Sol session, evidence cf-review-round2-*. No feature commit yet. Evidence: results-private/fidelity/resume-cf-*. TaskSpace 42 verified available, retain it through final acceptance. Unused fixture compatibility shim removed; canonical builder is ops/fixtures.py. Engine/calc pipeline remain unchanged.

Fidelity CF checkpoint, 2026-10-11: no feature commit yet; HEAD f74d131. Latest bounded regression checks: Python 22/22, web 23/23. Earlier controller full Python 296/296 passed; earlier full web 102/103 exposed a cache regression that is now fixed in focused tests. Full acceptance has not been rerun. Exact claude-opus-5-5 static review returned FIX-FIRST; remaining findings still need closure and re-review (multi-area formula origin, DXF background fallback, border enums, target-range clipping and escaped sheet names). CF evaluator currently rejects more than 10000 referenced cells per batch or 100000 overrides with an explicit service error; these limits need review against the large-sheet requirement. Drawings, notes, tables/filters and sheet views remain unimplemented. No private workbook or engine/pipeline mutation occurred.

Execution stopped after the capacity guard measured 398950400 bytes free, followed by a controller reading of 180297728 bytes. The owned executor exited cleanly. No further test/build was launched after the guard. Recover capacity before resuming. Evidence: results-private/fidelity/sol-fixes-last.md, cf-fix-checkpoint-low-disk.md, cf-fix-green-python.log, cf-fix-web-relevant.log, opus-cf-review-2.json. Next routed prompt is results-private/fidelity/sol-next-review.md; resume session 01a12961-bb90-71b3-af7a-07bc42b2cd29 using exact gpt-5.6-sol xhigh. TaskSpace 42 remains open, unused for application acceptance; finish keep: [] only after final successful checks.

Fidelity checkpoint: initial Luna tests rejected for insufficient real-boundary coverage and incompatible fake engine signature. Terra's test harness still had invalid snapshot paths, operation fields and a nonexistent grid class. After the two execution failures the exact `gpt-5.6-sol` xhigh lane took over, in session `01a12961-bb90-71b3-af7a-07bc42b2cd29`. Evidence: ignored `results-private/fidelity/`, current executor log `sol.jsonl`. Ego TaskSpace `42` is reserved until final acceptance; do not finish it early. Branch HEAD remains `f74d131`; tests have not yet established fidelity acceptance. Worker sandbox socket tests fail EPERM, so the controller owns full suite runs. Active Python/web lock audits report zero known vulnerabilities; Opus probe confirms canonical `claude-opus-5-5`.

Resource incident: a hook could not write state when disk reached 190 MiB free. Only owned executors were interrupted; no other process or user file was removed. Disk then recovered above 4 GiB but continues fluctuating. Keep execution sequential and stop before heavyweight work if capacity falls below 1 GiB.

## Next task
Stop and record only if `df -m /` Available drops below 2500 MB. Still open, none blocking a merge: the MED review findings (deterministic function batching and template-block reuse, explicit cell styling precedence over table styles, overlay scroll redraw and asset caching), header-clipping coverage, the unverified table-order memory claim, the pending exact claude-opus-5-5 re-review, and a final all-source security scan plus web dependency audit on this branch. Do not push or merge without user approval.


Historical integration handoff follows; its checks do not verify this fidelity work.

2026-10-11 06:40 IST, controller (claude-opus-5-5). Commits after Round 5: `0079f66` refuse to start when another program holds the port (found live: a first start bound 0.0.0.0:8000 beside a loopback listener on 8000); `0ff02dd` literal dropdown members take the type Excel enters on pick, IMPORT_VERSION 6 (independent review finding N1).
>>>>>>> feat/v2-fidelity

2026-10-11 native Option A session. Uncommitted work from the interrupted run
did not exist; the tree was clean at a25989c and nothing was inherited.

Shipped and committed, each item verified before commit:
  da9720d expand shared formulas before native template load, 0 of 150 mismatches
  787953d counts-only native parity audit for the controller real-workbook runs
  1764d01 visible-only evaluation with stale deferral and on-demand resolution
  10db8af shared formulas split across a read boundary, plus scale measurements
  cf028e3 engine memory budget, idle LRU eviction and editor protection

<<<<<<< HEAD
Two real defects were found by testing rather than review. Cells straddling a
read boundary were copied through unexamined, so their shared followers leaked
into the prepared package and the 20% load failed closed; a test now forces the
boundary with 64 byte reads. recalculate_stale resolved each cell on demand
during the publish capture, which zeroed the delta; the marks are now cleared
before evaluation and restored on failure.

Verification for this session: Python 431 of 431 twice (48.05 s and 48.84 s),
web 101 of 101 twice, content guard exit 0, Ruff clean on every changed file.
Synthetic fixtures only. No real workbook or private directory read, no port
bound, no push, no dependency change, no external model review. The supervisor
runs reviews after this runner.

Aggregate evidence only is committed. Raw logs remain under results-private.
Free disk 14174 MiB at the last check, far above the 2500 MiB floor.

## Next task

Controller owns the real-workbook evidence. Run the counts-only audit against the
real template and report the aggregates:

    .venv/bin/python -m ops.native_parity --template /path/to/template.xlsm \
        --sample 20000 --output .planning/excel-grid/results/native-parity.json

Pass requires zero mismatches across the lazy sample and zero after revert. The
audit prints counts only. If it fails, mismatches_by_function names the family to
investigate first. The same audit runs against a synthetic fixture as a control.

Then, in priority order:

1. Make the dependency trace cheaper. It now dominates every broad edit and the
   visible-only path does not yet reduce it. This is the largest remaining
   latency lever and it is what limits 50% and above.
2. Wire the connection registry to pool.protect and release_protection, so
   editor presence is real rather than only an available API.
3. Decide defined-name expansion. If the real template uses names inside
   formulas, those cells are still registered per cell after load. Measure the
   count on the real workbook before choosing whether to extend the pre-load
   rewrite to names.
4. Re-run the 5/20/50 scale set on the office server. The 50% load succeeded
   here; its full broad edit was killed by this 8 GiB Mac, so the trace cost at
   scale is still unmeasured off-host.

Unverified and owner-controlled: office Windows and Excel acceptance, LAN access,
any push or PR, and any dependency change. engine formualizer stays 0.11.1.


## Dashboard track (merged from feat/v2-dashboard)

v2 dashboard, approved 2026-10-11. Branch `feat/v2-dashboard` from `f74d131`. Steps 1-5 are implemented and committed: import `21b5fc7`, current-source pivot `1798ce9`, in-grid controls `0670be1`, refresh `a131b5b`, category totals `395ca30`. Step 6 team overview is in progress. Final acceptance remains open. Integration owns the engine/calc pipeline; fidelity owns conditional formatting, pictures/shapes, notes, tables and autofilter. Synthetic fixtures only; no private workbook access or Docker.

## Last handoff

2026-10-11 resumed from timeout. Current user resource rule supersedes the earlier checkpoint: stop at `df -m /` Available below 2500 MB, or swap more than 1536 MB above this resume baseline 5278.62 MB (limit 6814.62 MB). Initial free disk 6071 MB. Focused Python exposed three category filter boundary failures; fixed null slicer lists and non-string Copy selectors returning 500. Python 30/30 twice (9.819 s, 25.474 s), web 125/125 twice, ruff clean, pinned Prettier 3.6.2 clean, content guard exit 0, staged gitleaks clean. Category totals committed as `395ca30`. Exact `claude-opus-5-5` independent review is running via read-only CLI; identity must be checked from JSON output before counting it. Shared `.venv` and `web/node_modules` symlinks are untouched and untracked.

Existing Ego TaskSpace 43/p1 was resumed successfully and remains at about:blank; no current server has been started in this resume. Finish it with keep: [] after final browser check. Prior screenshots and checks are historical and do not verify final code. Full logs are in ignored `results-private/dashboard/`; track report contains prior phase evidence.

Routing: prior Luna and Terra execution attempts both failed, documented in the track. One exact `gpt-5.6-sol` xhigh takeover worker owns Step 6, while root handles reviews and acceptance. No push, PR, merge, deployment or Office acceptance authorized by this track plan. Excel pixel comparison and Calibri/Aptos fonts remain unavailable; do not claim measured pixel parity.

## Next task

Finish and commit Step 6 team overview. Read the pending Opus review, verify exact returned model identity, resolve concrete findings test-first and re-review critical/high repairs. Complete independent exact planner review, scoped security/dependency checks, 100k benchmark twice, full Python twice and web twice on final code, content guard 0, ruff, pinned Prettier and two identical builds. Then run a synthetic real server on a free port excluding 8000/8765/8766; capture pivot, page filters, slicer selection/multi-select/clear, category totals and team overview in existing TaskSpace 43; finish keep: [] and stop server immediately. Update this state and tracks/dashboard.md and commit handoff.
=======
Codex usage limit reached; Astra/Sol unavailable until 2026-10-14 09:03 IST.
>>>>>>> feat/v2-fidelity
