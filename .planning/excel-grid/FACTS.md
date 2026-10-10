# Price Planning web app - facts for design (2026-10-10)
Goal: Excel-like multi-user web app for a retail price-planning workbook (Catalyst Brands, office LAN).
Users are Excel power users. Need: double-click edit, drag fill handle, copy/paste ranges, undo/redo, keyboard nav,
freeze panes, formats as in the xlsm, fast scrolling, live recalculation, who-edited-what per cell, change history.
Copies: anyone in a category creates a copy; others co-edit it live or fork their own. No merge-back (v1).
Formulas: template formulas are fixed (locked); users edit input cells only (v1). Template xlsm is always the base.
Hosting: one office Windows PC/server on LAN, Python 3.10+ available, no admin rights, no Docker, no internet installs at runtime
 (deps must be vendored/bundled). Users open http://10.x.x.x:8000 in browser. Team size ~5 per category, maybe tens total.
Repo: private GitHub, code only; the 53 MB .xlsm is NOT in git (copied onto server once).
Workbook: 22 sheets, 1,894,479 formulas, ~650 MB sheet XML, 68 defined names, VBA modules (not executed).
Biggest sheets: Sales by Mkting Block 649,797 formulas A1:KZ4088; BU RECAP 364,246 A1:EI4355; Dept OTD LY 230,000 A1:AP10001;
TY vs LY 212,263; Store vs Flagship 110,089; Fiscal Week Recap 108,571; LY Data 99,998 (A1:Q100000).
Function counts: SUBSTITUTE 3.59M, IFERROR 1.56M, MATCH 917k, VLOOKUP 881k, IF 765k, COLUMN 545k, TRIM 529k, NUMBERVALUE 487k,
INDEX 467k, CONCAT 362k, ISERROR 351k, RIGHT 338k, AND 295k, FIND 290k, XLOOKUP 259k, INT 149k, SEARCH 121k, SUMIF 105k,
LET 93k, SWITCH 87k, SUMIFS 22k, HLOOKUP 21k, IFS 14k, COUNTIFS 10k, SUMPRODUCT 7.7k, SEQUENCE 4.6k, MAXIFS 2.6k, DAYS 2.4k,
UNIQUE 756, FILTER 383, ISFORMULA 234, OFFSET 2, INDIRECT 2, CELL 4, plus TEXT, FLOOR.MATH, ROUNDUP, MOD, MINIFS, COUNTA.
Existing prototype (to replace/keep parts): Python stdlib WSGI + waitress, SQLite WAL, accounts/sessions/CSRF, copies with
per-cell input overrides on an immutable baseline, owner/editor/viewer, append-only change history, cell version conflict checks,
custom Python formula interpreter (101/103 parity on sample), recalculation only on demand for visible range; UI is a paged
grid with an inspector panel (no in-cell edit, no fill handle).
Dev Mac: 8 GB RAM. Office server: unknown specs, Windows.
