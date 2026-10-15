# Deploy on the office server (Windows, no admin)

## First time
1. Install nothing as admin. Python 3.10 to 3.12 (x64) must already exist (`py -3 --version`).
2. `git clone https://github.com/GopalGB/price-planning.git` into a local folder (not a network share).
3. Copy the template `.xlsm` by hand into `<repo>\template\` (for example
   `template\Price Planning Template Spring 2027.xlsm`). The folder is gitignored; the file is never in git.
   To keep it elsewhere, set `PP_TEMPLATE_PATH` to its full path instead (or `template_path` in `settings.json`).
4. Double-click `run_v2.bat`. It creates `.venv` from `vendor\wheels` offline, writes first-start LAN/demo
   defaults when `settings.json` is absent, and starts the app. Do not create or edit JSON for the standard
   office start.
5. Copy each printed `Share with your team` URL to colleagues. A detected `10.x` address is listed first,
   followed by other usable local addresses.

## Office PC: one command to pull, one to run

Make a folder (for example `C:\PricePlanning`), put the template `.xlsm` in it, open CMD in that folder:

    curl.exe --ssl-no-revoke -sL -o pull.bat https://raw.githubusercontent.com/GopalGB/price-planning/feat/v2-integration/office_pull.bat && pull.bat
    app\office_run.bat

`office_pull.bat` needs no installed Git and no admin: on first use it unpacks MinGit into
`%LOCALAPPDATA%\MinGit`, then clones the app into `app\` or fast-forwards it (Windows certificate store,
revocation check off, the same as `curl.exe --ssl-no-revoke`).

`office_run.bat` uses the first `.xlsm` in that folder (`PP_TEMPLATE_PATH`), keeps engine caches in
`engine-cache\` there, writes `settings.json` with the Excel engine on first use (never rewrites it), unpacks
Python 3.12 (python.org NuGet zip, no installer or admin) into `%LOCALAPPDATA%\PricePlanning` when no 3.10 to 3.12 exists (`curl --ssl-no-revoke` for proxies that block revocation
checks), installs the wheels offline, runs the four Excel checks on the first run only (`checks-done.txt`;
`app\office_run.bat check` runs them again), then starts the app.

## Backup (online, safe while the app is running)
`run_v2.bat --backup` writes a verified copy of the resolved live database into `backups\` using the SQLite online
backup API (never a file copy of the live database). Database precedence is: `--db`, then `PP_DATABASE_PATH`, then
`DATA_DIR\planning.db`, then `settings.json` `database_path`, then `settings.json` `data_dir\planning.db`, then
`<repo>\data\planning.db`. Next to each `planning-<time>.db` it writes a `.json` record: time, source, schema
version and SHA-256. Other folders: `run_v2.bat --backup --db <db> --dest <folder>`.
Without the launcher: `.venv\Scripts\python -m pp.backup` (same options).
While `run_v2.py` serves, the app also backs up on its own: once five minutes after start, then every
24 hours, into the same `backups\` folder, keeping only the newest 7 sets. Set `"auto_backup": false` in
`settings.json` to switch that off; `run_v2.bat --backup` still works by hand.

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
Colleagues open a printed `Share with your team: http://<server-ip>:<port>` URL. No JSON edit is needed for the
standard office copy, start, and share flow. If they cannot connect, IT must allow that port; the app cannot
change firewall rules. Production should sit behind the company HTTPS proxy (`ops/nginx.conf`).

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

On first v2 server start, a missing `settings.json` uses LAN defaults: bind
`0.0.0.0`, port `8000`, and demo mode enabled. It derives exact loopback,
hostname, and detected local IPv4 origins and prints shareable URLs, preferring
a detected `10.x` address. No manual JSON editing is required. Configure
`bind`, `port`, `data_dir`, `template_path`, `allowed_origins`, and
`max_workers` in `settings.json` only when you need to override those defaults.
An explicit nonempty `allowed_origins` list replaces the derived list and must
use the exact browser origin including its port.
`max_workers` defaults to one, or six with the Excel engine: each Excel copy is its own Excel
process, and the memory budget closes idle copies when they outgrow it.
Signup requires `demo_mode: true`; use HTTPS and `cookie_secure: true` outside an
explicit HTTP demo. V2 serves `static/dist/` when present, otherwise a placeholder.
Office Excel parity, Windows operation, LAN access, and performance acceptance
still require their separate release checks.

V2 imports legacy tables once. Stop the legacy app before migration and keep an
online SQLite backup for rollback. Once v2 has migrated a data directory, do not
run `run.bat` against that directory. V2 refuses a later startup if legacy tables
have changed since migration. Keep legacy operation in a separate data directory
or restore its backup into an empty directory before reverting to the legacy app.

## Excel calculation engine (office server with Microsoft Excel)

Large templates calculate in Microsoft Excel itself. In `settings.json` add `"engine": "excel"` (keep
`template_results` unset). Each open copy gets its own hidden Excel with macros and events turned off;
idle copies are closed by the memory budget (`engine_memory_budget_mb`, counts Excel's memory too).

First check on the server, before inviting users (prints counts and timings only, no cell values):

    .venv\Scripts\python ops\excel_smoke.py

It must end with `PASS`. The first start builds the template caches in the background (a few
minutes); edits made before that finish simply wait for them. Without Excel (this Mac), use
`"template_results": "cached"` to show the template read only.

Template macros (the buttons on the template) get their own check:

    .venv\Scripts\python ops\excel_smoke.py --macros

Every button runs on a fresh Excel, then a second fresh Excel applies only the recorded contents, and the
two are compared cell by cell. It prints per button: seconds, message count, contents count, formula
contents, published changes and Excel's peak memory, plus the mismatched-cell count of the replay. It must
end with `PASS`. Macros only run on the Excel engine; on this Mac they answer 409 `macros_unavailable`.

"Refresh data" (shown only for templates with Power Query connections) runs Excel's own refresh on each
copy's private Excel. The query sources must be reachable from the server's Windows account (a network
share or database the account can log in to). Excel's messages about a source (for example, that it
could not be reached) appear in the dialog. Check it with:

    .venv\Scripts\python ops\excel_smoke.py --refresh

It refreshes on a fresh Excel, replays the recorded contents on a second fresh Excel, and prints counts,
timings and Excel's peak memory only. It must end with `PASS`.

Excel-only pivots and their slicers (the template's pivot fed by Power Query) get their own check:

    .venv\Scripts\python ops\excel_smoke.py --slicers

For each Excel slicer it selects one item with data on a fresh Excel, restores all items, then a second
fresh Excel applies the final stored state and must read the same pivot cells. It prints slicer numbers
(never item texts), counts, timings and the mismatched pivot-cell count. It must end with `PASS`. Slicer
clicks in the browser change the pivot in this Excel; on the Mac they answer 409 `slicers_unavailable`.
