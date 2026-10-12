# v2 API contract (shared by backend Track A and frontend Track B)

Date 2026-10-11. Source of truth for both tracks. Changing it needs a note in `.planning/STATE.md`.
All JSON. Auth = httpOnly session cookie `pp_session` (SameSite=Strict) + header `X-CSRF-Token` on every
non-GET. Errors: `{"error": "<code>", "message": "<human text>", ...extra}` with proper HTTP status.
Cell coordinates are 0-based `r`, `c` integers; `sheet` is the sheet index (0-based, template order).

## Values
`v` is one of: number, string, boolean, null (blank), or `{"error": "#DIV/0!"}`.
Display formatting is client-side from `nf` (Excel number-format code) in the style table.

## Auth
- `POST /api/auth/signup {email, name, password}` -> 201 `{user}` (demo mode only; password >= 12 chars)
- `POST /api/auth/login {email, password}` -> 200 `{user, csrf}`; 429 after 5 failures/min
- `POST /api/auth/logout` -> 204
- `GET /api/me` -> `{user: {id, email, name}, csrf}` or 401

## Categories
- `GET /api/categories` -> `{categories: [{id, name, role: "manager"|"member", members: [{id, name, email, role}]}]}`
- `POST /api/categories {name}` -> 201 `{category}` (creator becomes manager)
- `PUT /api/categories/{id}/members {email, role: "manager"|"member"|null}` (null removes) -> 200 `{category}`

## Copies
- `GET /api/copies?category={id}` -> `{copies: [{id, name, category_id, owner: {id,name}, visibility: "private"|"category",
   default_role: "editor"|"viewer", my_role: "owner"|"editor"|"viewer", input_revision, calculated_revision,
   parent: {id, name, revision} | null, last_edit: {by: {id,name}, at} | null, members: [{id, name, role}]}]}`
  The `category` query is optional. Without it, the endpoint returns every copy accessible to the current user,
  including named shares from categories where that user is not a member.
- `POST /api/copies {category_id, name, visibility, default_role}` -> 201 `{copy}`
- `POST /api/copies/{id}/fork {name}` -> 201 `{copy}` (atomic at parent's current input_revision)
- `PATCH /api/copies/{id} {name?, visibility?, default_role?}` (owner) -> 200 `{copy}`
- `PUT /api/copies/{id}/shares {email, role: "editor"|"viewer"|null}` (owner) -> 200 `{copy}`

## Workbook data (any role on the copy)
- `GET /api/copies/{id}/workbook` -> `{template: {hash, name}, input_revision, calculated_revision,
   sheets: [{index, name, hidden, rows, cols, freeze: {rows, cols}, col_widths: {c: px}, row_heights: {r: px},
   hidden_rows: [r], hidden_cols: [c], merges: [[r1,c1,r2,c2]]}],
   buttons: [{id, sheet, from: {row, col, dx, dy}, to: {row, col, dx, dy} | size: {w, h}, label, macro}]}
   (macro buttons of the template; `from`/`to` are zero-based cells with pixel offsets, `size` for one-cell anchors;
   snapshots imported before buttons existed have `buttons: []` until the template is reimported)
  `refreshable: bool` is true when the template has data connections (Power Query) that "Refresh data" can run.
   styles: [{id, nf, bold, italic, color, fill, align, valign, wrap, font_size, font_name, border: {t,r,b,l}}]}`
- `GET /api/copies/{id}/sheets/{sheet}/rows?start={r}&count={n<=500}` ->
  `{calculated_revision, input_revision, start, rows: {"r": {"c": {v, s, k, input?, input_value?, ver?, by?, at?}}}}`
  where `s` = style id, `k` = kind: `"i"` input (editable), `"f"` formula (locked), `"c"` constant (locked);
  input cells carry `ver` (cell version, 0 if never edited), `by` {id,name}, `at` ISO time of last edit.
  When a result covers an immutable input, `k` remains `"f"` for display while `input: true`, `input_value`
  and `ver` expose its editable semantic input state; formula-only cells omit these fields.
- `GET /api/copies/{id}/sheets/{sheet}/cell?r=&c=` -> `{v, s, k, input?, input_value?, ver?, by?, at?, formula: "=..." | null, validation: {type, list?, min?, max?} | null}`
- `GET /api/copies/{id}/history?sheet=&r=&c=&limit=50` (cell when r,c given, else whole copy) ->
  `{entries: [{op_id, revision, at, by: {id,name}, sheet, r, c, before, after, action: "edit"|"paste"|"fill"|"clear"|"undo", origin?: {copy_id, name, revision}}]}`
  Fork history includes parent entries only through the recorded fork revision, retaining their original editor and time.
  Inherited entries carry `origin`; local entries omit it. Nested forks retain ancestor revision cutoffs.
  A local edit or undo clears inherited markers on the current input without erasing its earlier history.

## Edits (owner/editor)
- `POST /api/copies/{id}/ops {op_id, action, cells: [{sheet, r, c, v, expected_version}]}` (<= 10,000 cells)
  Text longer than 32767 characters or containing unsupported XML characters returns 422 `invalid` atomically.
  Unsupported characters include C0 controls except tab/LF/CR, U+FFFE/U+FFFF and lone surrogates.
  Authenticated mutations revalidate their persisted, unexpired session inside the write transaction after body and pipeline waits.
  -> 200 `{op_id, input_revision, accepted: n}` | 409 `{error: "conflict", cells: [{sheet,r,c, current_version, current_v, by}]}`
  | 422 `{error: "locked"|"invalid", cells: [...], message}`. All-or-nothing. Values starting with `= + - @` as
  text are rejected 422 ("Formulas can't be typed into input cells"). Identical retry of an op_id returns the
  stored ack; same op_id with different payload -> 409 `{error: "op_id_reused"}`.
- Undo = client sends a new op with `action: "undo"` restoring `before` values with current versions.
- `POST /api/copies/{id}/macros {button, op_id}` (owner/editor; viewers 403). The server maps `button` to the
  template's macro; clients never name a macro. Runs inside the copy's pipeline lock, so edits to the copy
  wait behind it (no 409 for a running macro).
  -> 200 `{messages: [...], changed: n, input_revision}` where `messages` are the macro's dialog texts and `changed`
  the number of cells it wrote (0: nothing committed, `input_revision` unchanged). One macro press is one
  revision with history action `macro`; its cells are overrides (inputs, plain values and `{formula}` values).
  Identical retry of `op_id` returns the stored ack without running the macro; reused `op_id` -> 409 `op_id_reused`.
  Errors: 404 `unknown_button`; 409 `macros_unavailable` ("Macros run on the office server, where the app uses
  Microsoft Excel.") when there is no Excel engine or the template is `cached`; 502 `macro_failed` with
  "The macro did not finish in Excel." or, past the 600 s limit, "The macro ran too long and was stopped."
  (nothing is committed on 502). If the server dies after the commit, the normal calculation replay publishes it.
- `POST /api/copies/{id}/refresh {op_id}` (owner/editor; viewers 403, same CSRF and pipeline rules as macros). Runs
  Excel's own `RefreshAll` on the copy's private Excel, waits for the background queries, and records what it wrote
  exactly like a macro: one revision with history action `refresh`. Response, replay and `op_id` rules are the
  same as for macros: `200 {messages, changed, input_revision}`.
  Errors: 409 `refresh_unavailable` ("This template has no data connections to refresh.") when `refreshable` is
  false; 409 `macros_unavailable` when there is no Excel engine; 502 `macro_failed` with "The refresh did not
  finish in Excel." or "The refresh ran too long and was stopped." (nothing committed).
- Rows and cells: an override on a non-input cell shows its value with the cell's own kind (`k` stays `c`/`f`) and no
  `ver`/`by`/`at`; a formula override is never sent as a value (its published result is shown). On an input,
  `input_value` is the formula text with `input_formula: true`.

## Export
- `GET /api/copies/{id}/export.xlsm` -> the original template with this copy's input values written in; Excel
  recalculates on open.

## Live channel
`WS /api/copies/{id}/ws` (cookie auth + Origin check; closed with 4403 on lost access).
- client -> server: `{type: "hello", last_revision}`; `{type: "presence", sheet, sel: [r1,c1,r2,c2]}`
- server -> client:
  - `{type: "op", op_id, input_revision, by: {id,name}, at: "<ISO-8601 server time>", action, cells: [{sheet, r, c, v, ver}]}`
    `at` is the authoritative server-clock ISO-8601 timestamp persisted for the operation's history rows. For a legacy event with no `at`, the client stores `null`
    immediately and refetches cell detail; the detail version must still match before metadata is applied.
  - `{type: "calc", from_calculated_revision, calculated_revision, changes: [{sheet, r, c, v, s?, k?, input?, input_value?, ver?, by?, at?}]}` (complete batch per published interval, monotonic)
    `from_calculated_revision` is the prior published revision covered by the aggregate delta. New publications include authoritative style and kind so spill expansion creates cells and contraction restores the underlying template cell. A null value with kind `f` is a calculated blank, not a deletion. A result covering an immutable input carries `input: true`, its semantic `input_value`, and input metadata even when the displayed result did not change. Input records restored after contraction include their current input metadata. Legacy events without the interval field cover only their immediately preceding revision, and legacy value-only events retain their prior replay behavior.
  - `{type: "calc_state", input_revision, calculated_revision}` ("calculating" while they differ)
  - `{type: "presence", users: [{id, name, color, sheet, sel}]}`
  - `{type: "resync"}` when the client is too far behind: client re-fetches visible rows.
