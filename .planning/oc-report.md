# Bandit hardening: pp under the scan

## STATUS

DONE. `uvx bandit -r app ops pp -ll` exits 0 with "No issues identified", and CI runs the
same command (`.github/workflows/ci.yml:36`). Every B314 in `pp/` is now a defusedxml parse
rather than a suppression. No external model review was run: this runner does not call other
AI CLIs, the supervisor owns review.

## CHANGED

XML parses switched to defusedxml, 34 flagged call sites in 9 files (plus one unflagged
default-argument binding that was a live hole). No `ET` rename; where `xml.etree.ElementTree`
was still needed (type hints, `ET.ParseError`, `ET.TreeBuilder`) the import stayed, and where
the module ended up with no non-parsing use left it was dropped so Ruff stays clean.

| File | Calls switched |
|---|---|
| pp/engine/formualizer.py | 4 `fromstring` + 1 default arg `parser=ET.iterparse` |
| pp/engine/native_load.py | 4 `fromstring` |
| pp/export/transfer.py | 4 `fromstring` + 3 `iterparse` |
| pp/export/xlsm.py | 2 `fromstring` |
| pp/template/importer.py | 5 `fromstring` + 2 `iterparse` |
| pp/template/dashboard.py | 1 `fromstring` + 1 `XMLParser` |
| pp/template/colors.py | 2 `fromstring` |
| pp/template/drawings.py | 1 `fromstring` |
| pp/upload.py | 2 `fromstring` + 2 `iterparse` |

Reviewed false positives annotated, no logic touched:

- `pp/config.py:226` `# nosec B104` LAN app: colleagues connect over the office network; origins are checked
- `pp/db.py:162` `# nosec B608` identifiers come from this database's own schema and are quoted
- `pp/workbook_service.py:487` `# nosec B608` clause names are fixed (sheet,row,col); values are bound
- `pp/engine/excel.py:441`, `pp/engine/formualizer.py:175` `# nosec B301` per-user engine cache written only by this app

Separator note: the reason is written as `# nosec B104 # reason`. Bandit reads everything after
`nosec` up to a `#` as candidate test IDs, so the brief's `# nosec B104  reason` form emitted
eleven `Test in comment: LAN is not a test name or id, ignoring` warnings per line on every CI run.
Same text, no log noise, same suppression.

Three behaviour-preserving details worth a reviewer's eye:

1. `pp/template/dashboard.py` had a hand-rolled DTD guard: a `TreeBuilder` subclass whose
   `doctype()` raised. `SafeET.XMLParser` is a Python subclass of the C `XMLParser` and does
   **not** dispatch to `target.doctype` (measured: `ET.XMLParser(target=T())` calls it,
   `SafeET.XMLParser(target=T())` does not). A plain swap would have silently accepted every
   DOCTYPE. It now uses `SafeET.XMLParser(forbid_dtd=True)` and maps `DefusedXmlException` to
   `DashboardMetadataError`, so the exception type and the "DTD" message the pinned test asserts
   are unchanged. `_DashboardXML` was deleted: after the swap its guard was unreachable.
2. `pp/template/colors.py` `theme_colors` returns `DEFAULT_THEME` for a malformed theme part.
   defusedxml raises `DTDForbidden`, which is not an `ET.ParseError`, so `DefusedXmlException`
   joined the existing except tuple rather than letting a DTD escape as a 500.
3. `_formula_cells(stream, parser=ET.iterparse)` in formualizer.py was a default argument
   bound to the stdlib parser. Bandit did not flag it (it is not a call), but every sheet part
   reached the engine through it. Changed with the rest.

Two tests patched (`tests/v2/test_transfer_upload.py`, `tests/v2/test_transfer_adversarial.py`):
both mocked `pp.upload.ET.parse` / `pp.upload.ET.iterparse` / `pp.export.transfer.ET.parse`,
names those modules no longer have. The guards they assert ("the tree parser must not reach
this part") were already inert before this change and are now actually wired to the parser in
the path. No assertion was weakened.

## TESTS

| Check | Result |
|---|---|
| `uvx bandit -r app ops pp -ll` | exit 0, "No issues identified", 0 MEDIUM |
| `semgrep scan --error --config p/python --config p/javascript --config p/owasp-top-ten --config p/secrets --quiet` | exit 0, empty |
| `.venv/bin/python ops/content_guard.py --all` | exit 0 |
| `ruff check` on the 15 touched files | clean at HEAD and clean now |
| `ruff check .` whole tree | 698 findings now, 698 in a scratch worktree at HEAD, same rule breakdown |

Focused suites, all OK, run one at a time:

```
tests.v2.test_dashboard_import tests.v2.test_template_importer tests.v2.test_colors
tests.v2.test_drawings tests.v2.test_transfer_upload tests.v2.test_transfer_adversarial   132
tests.v2.test_transfer_upload tests.v2.test_db tests.v2.test_config tests.test_p0_regressions  104
tests.test_export_xlsm tests.v2.test_export tests.v2.test_export_route tests.test_core
tests.v2.test_colors tests.v2.test_drawings                                              47
tests.v2.test_security tests.v2.test_ops tests.v2.test_ops_api                            28
tests.v2.test_native_shared_expansion tests.v2.test_structured_refs
tests.v2.test_depindex tests.v2.test_native_parity                                         26
tests.v2.test_engine_adapter tests.v2.test_engine_contract tests.v2.test_engine_cached
tests.v2.test_engine_array                                                                 68
tests.v2.test_workbook_api tests.v2.test_ops_api tests.v2.test_end_to_end                   17
```

The whole suite was not run: another agent holds it on this host.

## B314 nosec cases

**None.** Every flagged parse moved to defusedxml, including the one that passed a custom
`parser=` target. No suppression was needed for XML anywhere in `app/`, `ops/` or `pp/`.

## NEXT

1. **Perf, owner decision.** defusedxml's `iterparse` measures 1.6x to 2.9x slower than the
   stdlib parser on a 20 MB synthetic sheet (min-of-4, A/B in one process; the cause is
   `DefusedXMLParser` being a Python subclass of the C `XMLParser`, not the expat handlers:
   a subclass with no handlers set is equally slow). The blast radius is the XML hot paths:
   `pp/template/importer.py` shared strings and sheet rows, `pp/upload.py` sheet structure,
   `pp/export/transfer.py` cell iteration. On this contended host the 30 s export budget test
   measured 21.9/23.1 s at HEAD and 22.1/28.1 s after. That headroom is now thin; the number
   should be re-measured on a quiet host before anyone trusts it, and if it bites, the fix is
   a deliberate decision to keep `ET.iterparse` behind the existing `_validate_archive` DOCTYPE
   screen in `pp/upload.py` rather than an accidental one.
2. Re-measure `test_55mb_xml_subset_path_stays_within_budget` on the office server.
3. Whole Python suite on a quiet machine, then the supervisor's review.
# OpenCode runner report

Branch: `feat/auto-backup` · commit `544928e` · base `d2a45ca`

## STATUS

Done. While `run_v2.py` serves, one daemon thread takes a verified online backup of the live database
about five minutes after start and every 24 hours after that, into the same `backups\` folder the
manual `run_v2.bat --backup` uses, and keeps only the newest 7 sets. `"auto_backup": false` in
`settings.json` switches it off. No external model review was run: the runner note forbids it and the
supervisor runs reviews afterwards.

## CHANGED

- `pp/auto_backup.py` (new)
  - `prune(dest_dir, keep=7)`: keeps the newest 7 `planning-*.db` sets by file name (the backup
    timestamp sorts lexicographically), deletes the `.db` plus its `.json` and `.templates`
    siblings, and never touches `.partial` files, other names, or symlinks. Returns the removed sets;
    a missing folder is a no-op.
  - `class AutoBackup(db_path, dest_dir, first_delay=300, interval=86400, keep=7)`: `start()` launches
    one daemon thread (`pp-auto-backup`), `stop(timeout=30)` sets an Event and joins. Each tick calls
    `backup()` then `prune()`; a failure is logged with `LOGGER.exception` and the loop continues, so a
    failed backup can never stop the server. Delays are constructor arguments.
- `pp/config.py`: new validated `Config.auto_backup: bool = True` (`"auto_backup"` must be boolean),
  read like `demo_mode` / `cookie_secure`.
- `run_v2.py`: `_start_auto_backup(runtime.config)` starts `AutoBackup(resolve_database_path(ROOT),
  ROOT / "backups")` after the runtime exists and before `uvicorn.run`, prints
  `Automatic backup: daily into <folder>, keeping the newest 7.`, and the existing `finally` calls
  `auto_backup.stop()` before `runtime.close()`.
- `tests/v2/test_auto_backup.py` (new, 6 tests): prune keeps the newest 7 of 9 whole sets and leaves a
  `.partial` file and an unrelated file alone; `first_delay=0, interval=0.2` against a real
  `pp.db.Database` schema database makes 2+ verified backups, `stop()` returns in under 2 s and the
  thread is dead; a `backup` that raises once is logged on `pp.auto_backup` and the next run still
  succeeds; the runner starts and stops `AutoBackup` with the resolved live database and `ROOT/backups`,
  and starts nothing when settings say `"auto_backup": false`.
- `tests/v2/test_run_v2.py`: the shared fake runtime config gains `auto_backup=False` so the URL tests
  stay free of a backup thread.
- `docs/DEPLOY.md`: two sentences in the Backup section (when the automatic backup runs, and how to turn
  it off).
- `.planning/STATE.md`: session entry (branch, what shipped, tests, office check).

## TESTS

- `.venv/bin/python -m unittest discover -s tests` -> `Ran 758 tests ... OK` (final run, 442 s).
  An earlier full run showed `test_transfer_adversarial.test_55mb_xml_subset_path_stays_within_budget`
  at 34.5 s against its 30 s budget; with the change stashed it ran in 13.9 s and with the change
  applied in isolation 17.4 s, and it passed in the final full run, so it is machine load, not this
  change. One genuine test bug was found and fixed: `tests/test_core.py` sets `DATA_DIR` process wide at
  import, so the runner wiring test now isolates `DATA_DIR` / `PP_DATABASE_PATH` / `PP_PORT` (house
  pattern from `tests/v2/test_backup_template_bundle.py`).
- `.venv/bin/python ops/content_guard.py --all` -> exit 0.
- `ruff check pp/auto_backup.py run_v2.py tests/v2/test_auto_backup.py` (plus `pp/config.py`,
  `tests/v2/test_run_v2.py`) -> clean.
- Tests were written first: the new module did not exist yet and `python -m unittest tests.v2.test_auto_backup`
  failed with `ImportError: cannot import name 'auto_backup' from 'pp'` before the implementation.
- No real workbook, no private data folder, no port bound, no push, no PR, no dependency change.

## NEXT

- Office check (Sir, Windows): start `run_v2.bat`, confirm the line `Automatic backup: daily into
  ...\backups, keeping the newest 7.` appears, then confirm `backups\planning-<time>.db` plus its `.json`
  appears about five minutes after start. Leave the server running past day 8 to see the oldest set
  disappear. Set `"auto_backup": false` in `settings.json` to confirm no `Automatic backup:` line prints.
- Push and PR are owner-controlled: `git push` then open a PR for `feat/auto-backup` (CI must be green).
- Optional, not in this brief: `settings.example.json` does not list `"auto_backup"` yet, and the backup
  folder has no size guard (7 full sets of a large database).
- `.planning/oc-report.md` (this file) is left uncommitted on purpose: the brief asked for one commit.