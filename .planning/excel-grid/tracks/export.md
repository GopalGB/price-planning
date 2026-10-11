# Track C: export, backup, v2 launchers (branch feat/v2-export)

Date 2026-10-10. Test-first per unit (red run seen before each implementation).

## Built
- `pp/export/xlsm.py` `export_xlsm(template_path, cells, out_path)`: writes (sheet, r, c, value) into a copy
  of the original .xlsm. Touched worksheet parts are streamed row by row (1 MB text chunks, regex on complete
  `<row>` elements only); every other member, including `xl/vbaProject.bin`, is stream-copied unchanged in the
  original order. Numbers as `<v>`, text as `t="inlineStr"` (sharedStrings untouched), bool `t="b"`, None keeps
  only `r` and `s`. Existing `s` kept; new cells inherit row style (customFormat) or `<cols>` style. Cells with
  `<f>` raise ValueError and no output file is left. Rows/cells inserted in order, `spans` dropped on rewritten
  rows, `<dimension>` widened, `<sheetData/>` expanded, implicit row/cell refs handled. calcChain part, rel and
  content type removed; `fullCalcOnLoad="1"` set on calcPr (added or inserted). Writes to a temp file, then
  atomic rename. Refuses: out == template, bad indices, NaN/inf, XML-illegal chars, text > 32767, other types.
- `pp/backup.py`: `backup()` via sqlite3 online backup API from a read-only connection, journal_mode=DELETE on
  the copy, integrity_check + foreign_key_check, sidecar JSON {created_at, source, schema_version, sha256}.
  `restore()` refuses a non-empty dir, checks sha256 and integrity/foreign keys, removes the copy on failure.
  CLI `python -m pp.backup [--db] [--dest] [--restore BACKUP --dest NEWDIR]` (defaults: `$DATA_DIR` or
  `data/planning.db`, `backups/`).
- `run_v2.bat` (CRLF) and `run_v2.sh`: same Python resolution as run.bat, `ops/install_env.py`, then
  `run_v2.py %*`; `--backup` / `--restore` route to `python -m pp.backup`.
- `docs/DEPLOY.md` v2 steps (template folder or PP_TEMPLATE_PATH, run_v2.bat, backup, restore, drill);
  `.gitignore` adds `/template/`.

## Test results (this Mac, Python 3.11.4 in .venv)
`.venv/bin/python -m unittest discover -s tests`:
    Ran 86 tests in 14.655s
    OK
`python3 ops/content_guard.py --all`: exit 0.
Streaming evidence: a 20+ MB synthetic sheet part exported with tracemalloc peak 0.88 MB (64 KB chunks).

## Not verified
- No Excel or LibreOffice on this Mac: exported files were read back with openpyxl only; "opens in Excel and
  matches" (P4 gate) needs a Windows/Excel run on a synthetic fixture, then on the real template at the office.
- Not run against the real 650 MB sheet part (never touched here). Throughput unmeasured.
- `run_v2.bat` checked by content test only (no cmd.exe here). `run_v2.py` is not in this branch; Track A owns
  it, so plain `run_v2.bat` fails until that merges. `--backup` / `--restore` work without it.
- Array-formula and spill child cells have no `<f>`, so the exporter would write them; the server input mask
  must keep them out. Export assumes UTF-8 sheet XML (Excel and openpyxl always write it).
- Backup sidecar does not yet record template hash (SPEC 9 asks for template and schema hashes); needs the
  template store from Track A.

## Addendum 2026-10-10: review fixes (Astra + coordinator)
Each fix test-first; red seen before the change.
1. HIGH backup default DB: `pp.backup.default_db()` reuses `app.runtime.read_settings` / `validate_settings`
   and a new `app.runtime.resolve_data_dir` (extracted from `apply_environment`, which now calls it). Order is the
   server's: `--db`, then `DATA_DIR`, then settings.json `data_dir` (relative = under repo), then `<repo>/data`.
   CLI gained `--settings`. Tests: configured vs default vs relative vs env dirs; CLI backs up the configured DB.
2. MEDIUM restore: copy into `.restore-*.partial` inside the target folder (inside the cleanup scope), verify
   sha256 + integrity + foreign keys there, publish with `os.link` (fails if planning.db exists, never
   overwrites), then remove the staging file. Tests: injected ENOSPC mid-copy leaves the folder empty and a retry
   succeeds; a second restore refuses; a planning.db appearing after the emptiness check is kept untouched.
   Limit: needs hard links (NTFS, APFS, ext4 yes; FAT/exFAT no, which fail cleanly with OSError).
3. MEDIUM export: text gets ST_Xstring escaping before XML escaping (`_xHHHH_` -> `_x005F_xHHHH_`, CR ->
   `_x000D_`). Note: openpyxl 3.1.5 returns inline cell text still ST_Xstring-encoded, so the round-trip test
   decodes with openpyxl's own `openpyxl.utils.escape.unescape` (single pass, as Excel does); raw XML is also
   asserted. 'a\rb\r\nc', 'pre_x0041_post', '_x000D_' come back exactly.
4. run_v2.py / DEPLOY.md: untouched, per coordinator.
Also fixed (AGENTS rule 7, seen flaking 1 in 3 full runs): `ops/bench_real.py` reported `ok` when the 0.25 s
polling watchdog missed a run that exceeded `--max-seconds`; completion now checks the limit and exits 3.
Deterministic test stubs the watchdog.

Results: `.venv/bin/python -m unittest discover -s tests` three consecutive runs, each `Ran 92 tests` / `OK`
(32.3 s, 28.5 s, 21.0 s). `python3 ops/content_guard.py --all` exit 0.
