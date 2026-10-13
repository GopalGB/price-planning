# Tab switching: cache + compression

Branch `perf/tab-cache`, four commits from `7b0104e`. Synthetic fixtures only; no company
workbook was read and nothing under `~/PricePlanning-private` was touched.

## STATUS

DONE, with the list of every input a cached row body depends on written out below, because the
cache is only as safe as that list.

- No external model review was run: this runner does not call other AI CLIs, the supervisor owns
  review.

## CHANGED

### 1. JSON responses are gzipped (`pp/server.py`)

`JsonGZipMiddleware` compresses `application/json` over 1 KB at level 6. Starlette's own
`GZipMiddleware` compresses anything large enough, which would gzip the `.xlsm`/`.xlsx`
downloads; those are already zip containers, so the wrapper holds the response head back until
the first body chunk (where the content type is known), passes streaming bodies through
untouched, and adds `vary: accept-encoding` plus the new `content-length` only when it actually
compressed something. WebSocket traffic never reaches it (scope type is not `http`). It is added
inside `SecurityBoundaryMiddleware`, so the security headers are still applied last.

### 2. Rendered row blocks are cached (`pp/row_cache.py`, `pp/workbook_service.py`)

A byte bounded (256 MB) and entry bounded (512) LRU of rendered bodies, keyed as:

```
(copy_id, template_id, sheet, start, count, input_revision, calculated_revision, pivot_state_json)
```

`WorkbookService.rows()` settles access first (`_copy`, which raises 404 without a role), then
builds the key, then reads the cache: a body cached for one member is only ever handed to another
member who has just been authorized by the same check.

Every input to the rendered body, and how the key covers it:

| Input | Covered by |
|---|---|
| `copy_id` | key |
| `template_id` (the snapshot: styles, input mask, sheet meta, merges, col/row styles, validations, conditional formats, tables, filters, pivots, defined names, per-block cells and notes, `results` mode) | key; a `TemplateSnapshot` is immutable for the life of the process (its own `block()` cache already assumes this), and an upload never rebuilds an existing template's snapshot |
| `sheet`, `start`, `count` | key |
| `pp_overrides` rows in range (value, version, `updated_by`, `updated_at`) | `input_revision`, which every commit moves |
| `pp_results` rows in range (values from any published revision) | `calculated_revision`, which every publication moves in the same transaction as the rows |
| `pp_users.name` behind each override's `updated_by` (rendered as `by`) | immutable: signup is the only write besides `password_hash` |
| `pp_pivot_state` (stored slicer selections and `refreshed`) | key, hashed by content rather than by the revision stored beside it |
| Conditional formatting, which reads current values for **any** sheet and block a rule formula references | `input_revision` and `calculated_revision` are copy-global, so one change anywhere invalidates every sheet |
| The effective calculated revision (`results == "cached"` reads as `input_revision`) | derived from rows already in the key |

Deliberately **not** an input: the CF evaluation budgets and the 1.5 s formula deadline. A render
that runs out of time differs from one that does not, so an un-cached server could answer two
identical requests differently; caching makes the answer stable instead of flapping. One existing
test (`test_formula_target_budget_is_total_and_uses_one_public_engine_load`) patched a budget
constant and expected a re-render; it now clears the cache first, with the reason in a comment.

### 3. Each sheet keeps its own rows across tab switches (`web/src/model.mjs`, `web/src/grid.mjs`)

`SheetRowCache` + `SheetCacheMap` keep the last six visited sheets (least recently visited first
out, the same LRU shape the server cache uses). `switchSheet` no longer builds a new `RowCache`
and no longer clears the whole cell store; `openSheetCache` picks the sheet's own cache and
decides whether it is reusable.

A sheet's blocks are reusable only when **every** one of them was rendered at the revisions the
client is on (`data.input_revision`/`data.calculated_revision` against
`store.inputRevision`/`store.calculatedRevision`). A remote `op` or `calc` moves those
revisions, so blocks of other sheets stop being reusable the moment one arrives, and are dropped
with their cells the next time they are opened. `calc_state` messages move the revisions too,
which is what makes the window right after your own edit safe. `switchSheet` already refuses to
run while an edit is in flight, and the server broadcasts the op event before it answers the
write, so there is no window where an edit is unaccounted for.

`CellStore.evictSheet` frees a sheet's cells when it leaves the tab cache or when its blocks are
being reloaded, so a stale window is never painted. Per-sheet zoom (`this.zooms`, the zoom
dropdown) is untouched.

## TESTS

| Gate | Result |
|---|---|
| `.venv/bin/python -m unittest discover -s tests` | 804 tests, OK, 134.6 s (run once at the end, on the final code) |
| `cd web && npm test` | 186 tests, 186 pass, 0 fail |
| `npx prettier --check src/*.mjs test/tab-cache.test.mjs` | clean |
| `npm run build` | built; `static/dist` committed in `91b0367` |
| `.venv/bin/python ops/content_guard.py --all` | exit 0 |
| `uvx bandit -r app ops pp -ll` | exit 0 (68 low, 0 medium, 0 high) |
| `uvx ruff check pp ops/bench_rows.py tests/v2/test_row_block_cache.py tests/v2/test_gzip.py` | clean |

New tests, written before the code:

- `tests/v2/test_row_block_cache.py` (13): the LRU drops the oldest entry on either budget, a
  read refreshes an entry it keeps, replacing an entry does not double count its bytes, a body
  larger than the whole budget is not stored, `clear` releases everything. Against the API: an
  unchanged copy renders once and answers the second visit from the cache; another range is its
  own block; an edit (input_revision alone), a published calculation (calculated_revision alone)
  and a stored slicer selection (content digest alone) each make the next call re-render and then
  hold again; an outsider gets 404 on a block cached for someone else; another member of the same
  copy gets the same body.
- `tests/v2/test_gzip.py` (4): a JSON block is gzipped when the client asks and round-trips to
  the plain body; a client that does not ask is answered plain; a body under the minimum size is
  left alone; `.xlsm` and `.xlsx` downloads carry no `content-encoding` and still start with `PK`.
- `web/test/tab-cache.test.mjs` (5): switching away and back issues no rows request; an `op` or a
  `calc` makes the next visit reload; the seventh visited sheet is dropped from the tab cache; a
  window scrolled out before an edit is asked for again (this one fails if `entry.reset()` is
  removed, which was checked by removing it).

Not run here: a real browser check, the office server, Excel, semgrep, pip-audit, gitleaks.

## BENCHMARK

`ops/bench_rows.py` generates the workbook into a temporary directory, imports it, and prints
the time and the bytes of the first and second `rows()` call per sheet. No binary is committed.

```
$ .venv/bin/python ops/bench_rows.py --rows 800 --cols 312 --sheets 1
generated and imported 1 sheets of 800 x 312 in 4.3 s (249,600 cells, 7,191 kB of snapshot)
cache holds 1 blocks in 3671 kB (1 hits, 1 misses)
visit                          seconds    json bytes    gzip bytes
sheet 0 first visit              0.336     4,435,955       243,550
sheet 0 second visit             0.001     4,435,955       243,550

$ .venv/bin/python ops/bench_rows.py            # default: 2 sheets of 1200 x 24
generated and imported 2 sheets of 1200 x 24 in 1.1 s (57,600 cells, 1,644 kB of snapshot)
cache holds 2 blocks in 560 kB (2 hits, 2 misses)
visit                          seconds    json bytes    gzip bytes
sheet 0 first visit              0.059       339,603        20,607
sheet 0 second visit             0.001       339,603        20,607
sheet 1 first visit              0.059       339,603        20,607
sheet 1 second visit             0.001       339,603        20,607
```

What this says, and what it does not: a 256-row block of a protected sheet with a per-row
conditional format rule costs 0.34 s and 4.4 MB of JSON, 0.24 MB over the wire once compressed
(18x), and 1 ms on the second visit. The synthetic sheet is not the real template (no spilled
pivot output, no 16,384-column styling, no cross-sheet defined names), so the absolute times are
not the real tab-switch time; the second-visit number is the part this work removes, and the
browser no longer even asks for it when the revisions have not moved.

## NEXT

- Open a real tab switch in the browser before merging: the visible check is that coming back to
  a tab you just left paints with no network request for that sheet, and that an edit in another
  tab followed by a return shows fresh rows.
- The server cache is per process and not shared. With several workers behind the reverse proxy,
  each keeps its own copy; nothing is wrong with that, but a shared cache would be the next step
  if tab-switch latency matters more than one hit.
- The 1.5 s CF formula budget is still the reason a first visit on the real template can take
  seconds. The cache hides the second visit, not the first.