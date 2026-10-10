# Deploy on the office server (Windows, no admin)

## First time
1. Install nothing as admin. Python 3.10 to 3.12 (x64) must already exist (`py -3 --version`).
2. `git clone https://github.com/GopalGB/price-planning.git` into a local folder (not a network share).
3. Copy `Price Planning Template Spring 2027.xlsm` next to `run.bat` by hand. It is never in git.
4. Double-click `run.bat`. It creates `.venv` from `vendor\wheels` offline and starts the app.

## Update
1. Stop the app (Ctrl+C in its window).
2. `run.bat --backup` (verified online SQLite backup into `backups\`).
3. `git pull`
4. `run.bat` (re-installs wheels automatically if `requirements.lock` changed, runs migrations).

## Rollback
1. Stop the app. 2. `git checkout <previous tag or commit>`. 3. Restore the backup into a new empty folder
and point `settings.json` `data_dir` at it. 4. `run.bat`.

## Network
Colleagues open the printed `http://<server-ip>:<port>`. If they cannot connect, IT must allow that port;
the app cannot change firewall rules. Production should sit behind the company HTTPS proxy (`ops/nginx.conf`).

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
