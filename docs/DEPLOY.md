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
