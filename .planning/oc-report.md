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