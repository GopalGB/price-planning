# Deploy on the office server (Windows, no admin)

## First time
1. Install nothing as admin. Python 3.10 to 3.12 (x64) must already exist (`py -3 --version`).
2. `git clone https://github.com/GopalGB/price-planning.git` into a local folder (not a network share).
3. Copy the template `.xlsm` by hand into `<repo>\template\` (for example
   `template\Price Planning Template Spring 2027.xlsm`). The folder is gitignored; the file is never in git.
   To keep it elsewhere, set `PP_TEMPLATE_PATH` to its full path instead (or `template_path` in `settings.json`).
4. Double-click `run_v2.bat`. It creates `.venv` from `vendor\wheels` offline and starts the app.

## Backup (online, safe while the app is running)
`run_v2.bat --backup` writes a verified copy of the resolved live database into `backups\` using the SQLite online
backup API (never a file copy of the live database). Database precedence is: `--db`, then `PP_DATABASE_PATH`, then
`DATA_DIR\planning.db`, then `settings.json` `database_path`, then `settings.json` `data_dir\planning.db`, then
`<repo>\data\planning.db`. Next to each `planning-<time>.db` it writes a `.json` record: time, source, schema
version and SHA-256. Other folders: `run_v2.bat --backup --db <db> --dest <folder>`.
Without the launcher: `.venv\Scripts\python -m pp.backup` (same options).

## Update
1. `run_v2.bat --backup`
2. Stop the app (Ctrl+C in its window).
3. `git pull`
4. `run_v2.bat` (re-installs wheels automatically if `requirements.lock` changed, runs migrations).

## Restore / rollback
1. Stop the app. 2. If rolling back code: `git checkout <previous tag or commit>`.
3. `run_v2.bat --restore backups\planning-<time>.db --dest <new empty folder>`. It refuses a folder that is
   not empty, checks the SHA-256 against the `.json` record, then runs the integrity and foreign-key checks.
4. Point `settings.json` `data_dir` (or `DATA_DIR`) at that folder, then `run_v2.bat`.
Restore drill: do step 3 into a scratch folder once per month and confirm it succeeds.

## Network
Colleagues open the printed `http://<server-ip>:<port>`. If they cannot connect, IT must allow that port;
the app cannot change firewall rules. Production should sit behind the company HTTPS proxy (`ops/nginx.conf`).

## Legacy v1
`run.bat` / `run.sh` still start the v1 prototype (`run.bat --backup` is its own backup). Do not mix data folders.

## Dependency changes (developers)
Edit `requirements.in`, run `python ops/lock_deps.py`, commit `requirements.lock`, `vendor/wheels/` and
`ops/large_file_allowlist.json` together.

## Opt-in v2 backend

The existing `run.bat` launcher continues to run the legacy app. For the v2 backend,
install the locked environment with `python ops/install_env.py`, then run
`.venv\Scripts\python run_v2.py`. On macOS or Linux use `.venv/bin/python` instead.

Template selection uses `PP_TEMPLATE_PATH`, then `settings.json` `template_path`,
then the first alphabetically sorted `.xlsm` in `template/`. Keep the template and
`data_dir` on local disk. `python run_v2.py --import-template` imports the selected
template without serving HTTP. The template, snapshots, exports, and SQLite files
stay outside Git.

Configure `bind`, `port`, `data_dir`, `template_path`, `allowed_origins`, and
`max_workers` in `settings.json`. Use the exact browser origin including its port.
Keep `max_workers` at one until memory measurements justify more warm copies.
Signup requires `demo_mode: true`; use HTTPS and `cookie_secure: true` outside an
explicit HTTP demo. V2 serves `static/dist/` when present, otherwise a placeholder.
Office Excel parity, Windows operation, LAN access, and performance acceptance
still require their separate release checks.

V2 imports legacy tables once. Stop the legacy app before migration and keep an
online SQLite backup for rollback. Once v2 has migrated a data directory, do not
run `run.bat` against that directory. V2 refuses a later startup if legacy tables
have changed since migration. Keep legacy operation in a separate data directory
or restore its backup into an empty directory before reverting to the legacy app.
