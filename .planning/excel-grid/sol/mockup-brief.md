GO-TOKEN: approved, proceed. Today is 2026-10-10.

Build ONE self-contained file: proto/mockup/index.html (inline CSS + vanilla JS; no external requests, no libraries, no web fonts). Touch nothing else in the repo. Do not commit. No em-dashes anywhere. All data is FAKE: app name "Price Planning"; users Priya Shah, Marcus Lee, Ana Torres, Dev Patel (current user, "You"), Sam Kim. Never use any real company name.

PURPOSE: show an Excel-power-user planning team what the new web app will look and feel like. It must look and behave like Microsoft Excel, but clean and modern. Light theme, Excel green accent #217346, font Segoe UI / system-ui, 13px grid text, crisp 1px gridlines #e1e1e1. Must look right at 1440x900 and not break at 1280 wide.

LAYOUT, top to bottom:
1. App bar (44px): left "Price Planning" wordmark + breadcrumb "Category A / Spring 27 - Women's Tops (shared)". Center: status chip "Saved · calculated rev 42" (green dot); after any edit it shows amber "Calculating..." for ~600 ms then green with rev+1. Right: presence avatars (colored circles PS, ML, AT with name tooltips), buttons "Share", "Fork copy", "Download .xlsm", "History", user avatar DP.
2. Toolbar (36px, ribbon-lite): Undo, Redo, Paste values, number format dropdown (General/Number/Currency/Percent), Freeze panes toggle, Find box (Ctrl+F focuses it; Enter jumps to next match), Zoom 100%. Simple inline SVG icons.
3. Formula bar: Name box (active cell, e.g. "F12") + "fx" + content. Locked formula cell: shows the formula text greyed, a small lock icon, and "read only (template formula)". Input cell: shows the value, editable (Enter commits).
4. Grid: column headers A..N, row numbers 1..60. Rows 1-3 sticky (frozen), columns A-B sticky (frozen). Row 1 title "Spring 27 Price Plan - Women's Tops". Row 3 headers: A Group #, B Description, C Dept, D Class, E Cost $, F Retail $, G IMU %, H Disc %, I Promo $, J Units Plan, K Sales $, L Wk Start, M Wk End, N Status. Rows 4-60: 57 realistic fake groups generated in JS (e.g. 1001 Knit Tee Basic, 1002 Woven Blouse, ...).
   - INPUT cells E, F, H, J, L, M: fill #EAF2FB, dark blue text, editable.
   - FORMULA cells: G = 1 - E/F, I = F*(1-H), K = I*J, N = IF(G<0.45,"Review","OK"). White/very light grey, black text, NOT editable: double-click shows toast "Template formula cells are locked".
   - Formats: $ 2 decimals, % 1 decimal, units with thousands separators, dates MM/DD/YY. Status "Review" = amber pill, "OK" = green pill.
   - Changing an input recalculates its row's formula cells instantly and flashes them light yellow for 800 ms.
5. Sheet tabs at the bottom: BU Recap (active), Sales by Mkting Block, Projections, Fiscal Week Recap, Coupon, Calendar, then muted "+ hidden (14)". Other tabs show a centered placeholder "Sheet preview not in mockup".
6. Status bar (24px): left "Ready" / "Editing"; right "Average: x   Count: n   Sum: y" for numeric cells in the selection, and "Copy: shared · 5 members".

EXCEL BEHAVIOR (must actually work):
- Click selects a cell (2px #217346 border). Shift+click or mouse drag selects a range (light green tint). Arrows move; Shift+arrows extend; Tab/Enter move right/down; Ctrl+Home goes to E4.
- Double-click, F2, or start typing on an INPUT cell edits inline (an input sized to the cell). Enter commits and moves down; Escape cancels. Delete/Backspace clears selected input cells.
- Typing text starting with "=" into an input cell is rejected with toast "Formulas can't be typed into input cells in v1".
- Fill handle: small green square at the selection's bottom-right. Dragging it down copies the value (or continues a numeric series when 2+ cells are selected), recalculates, and skips locked cells with toast "N locked cells skipped".
- Ctrl+C copies the selection as TSV (navigator.clipboard, fallback hidden textarea + execCommand). Ctrl+V pastes TSV into input cells from the active cell; locked targets are skipped with a toast.
- Ctrl+Z / Ctrl+Y undo/redo your own edits (one paste or fill = one undo step).
- Hovering an edited input cell shows a card: "Last edit: Marcus Lee · 2 min ago · was $18.00". Pre-seed ~8 cells with fake history by different users; your edits show "You · just now".
- Presence: Priya's selection = purple 2px outline on F9 with a name tag "Priya" above it; Ana's = orange on J15. Every ~6 s simulate Priya changing a random Retail cell: value changes, purple flash, status chip goes calculating then rev+1, toast "Priya Shah changed F12 to $24.00", and the history panel updates.
- Right panel (320px, toggled by "History", open by default): "Cell history: F12" for the active cell with entries (avatar, name, time, "from $18.00 to $22.00"); below, "Copy activity" with the last 6 edits across the sheet. Live updates.
- "Share" modal: visibility radios (Private / Category A members / Named people); members list with role dropdowns (Owner/Editor/Viewer) for the 5 users; Done. "Fork copy" modal: "Create your own copy of 'Spring 27 - Women's Tops' at rev 42", name field, "Create fork" -> toast. "Download .xlsm" -> toast "Writes your inputs into the original template; Excel recalculates on open".
- A "Copies" home screen, reached by clicking the wordmark (an "Open" button returns to the sheet): left nav of categories (Category A selected, B, C); table of 5 copies: Name, Owner, Visibility (Shared/Private), Last edited (by + when), Members (avatars), Actions (Open / Fork). Default view on load = the sheet.

QUALITY: pixel-tidy alignment, no overlapping elements, no console errors, keyboard focus visible, prefers-reduced-motion disables flashes. Plain DOM table is fine for 60 rows.

VERIFY before finishing: run `node -e "const s=require('fs').readFileSync('proto/mockup/index.html','utf8');const m=s.match(/<script>([\s\S]*?)<\/script>/);new Function(m[1]);console.log('js parses')"` and fix until it prints "js parses". Report the file path, size, and a 10-line list of what works and anything not implemented.
