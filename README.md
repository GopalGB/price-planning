# Catalyst Brands price planning server

**LAN demo: extract the package, add your root `logo.png`, and double-click `run.bat`. No administrator account setup or Windows elevation is needed for the app.**

This is a validation build with private working copies, shared input editing, access roles, and your original workbook and SOP. It is not yet an approved replacement for the complete Excel model. Formula acceptance, VBA workflows, pivot/slicer behavior, and actual source feeds still require completion; see `docs/validation.md`.

## Start with one file

1. Use your company's approved **Python 3.10 or newer** installation; Python 3.12 is recommended. The launcher checks an optional `python/python.exe` in this folder, the Windows `py` launcher, and Python on PATH. Python itself must already be available.
2. Extract the entire package into a writable **local folder on the server**. Keep at least 2 GB of free space initially, then size storage and memory using your actual workload. Keep the database off network shares and synced folders.
3. Put the real Catalyst Brands logo at **`logo.png` beside `run.bat`**. The app reads it directly; replacing it needs only a browser refresh.
4. Double-click **`run.bat`**. You do not need to run `setup.py` separately.
5. The launcher creates demo settings automatically, binds to the server network interfaces on port 8000, and skips administrator account creation. No first-start questions are asked.
6. Wait for the workbook import. The 22 sheets and 1,894,479 formulas expand to roughly 650 MB of worksheet XML. The first start can take several minutes and prints sheet progress. Later starts reuse the imported database.
7. When **READY** appears, the browser opens the server address. Share the printed **Share with your team** URL, for example **http://10.0.0.25:8000**. The actual address comes from your server; the app cannot assign it a new 10.x IP. A detected 10.x interface is preferred. On the server itself you can also open **http://localhost:8000**.
8. Each person chooses **Create demo account**, enters a name, email and password of at least 12 characters, then chooses **New copy**. No administrator approval is needed. Returning users sign in with the same account.

The bundled Waitress web server runs directly with Python and SQLite. Startup needs no pip command, internet download, container engine, database service installation, or elevated privileges for the app. The package retains Waitress's license in `vendor/`.

Keep the server window open while people work. Press **Ctrl+C** to stop it. Double-click the same `run.bat` to restart; existing users, copies, inputs, and history remain in `data/planning.db`. Running it a second time against the same database is rejected. Closing the window or suspending the machine stops availability; continuous hosting requires an always-on machine and your company's existing process/startup arrangement.

Linux/macOS uses the same application:

```bash
sh run.sh
# Or: python3 run.py --no-browser
```

## Team access on your server

The demo starts on `0.0.0.0:8000` and prints the server's usable LAN addresses. This listen address is internal to the server configuration; people open the printed `http://10.x.x.x:8000` (or other actual server IP) URL in their browsers. They do not install Python, copy the app, or open the database file.

The startup does not request Windows elevation, change firewall rules, install a service, or download dependencies. Python must already be available and the server folder must be writable. Company network policy must already allow colleagues to reach this server and port. If localhost works but the team URL does not, check the actual IP with `ipconfig`, confirm colleagues are on the relevant network/VPN, and ask IT to verify connectivity if the port is blocked. Changing Python code cannot bypass company network restrictions.

IP detection is repeated on each startup. `settings.json` stores `public_origin: "auto"`, not an IP that becomes stale. If the server has several interfaces, the printed alternatives may include a second valid address. If no LAN address is detected, configure the actual server URL manually.

To change the port or set a specific LAN URL, stop the server and run:

```bat
run.bat --configure
```

Choose **2 = team LAN**, the desired port, the actual server URL, and **Y** for demo accounts. Restart with `run.bat`. An explicitly configured LAN URL is retained. `settings.example.json` shows the fields; `data_dir` points to local storage. Existing databases and accounts are preserved, including those from the earlier build.

For a future controlled setup behind an existing company HTTPS proxy, run `python setup.py`, choose HTTPS and **n** for demo accounts, then start **`python run.py`** using those settings. `run.bat` always enables the requested LAN demo. Managed mode creates the first administrator interactively and disables self-signup. The optional `ops/nginx.conf` is an upstream example for IT; this package does not install a proxy or certificates.

The demo is intended for your internal network. Email addresses are account identifiers; this demo does not send verification messages or integrate with company SSO.

## What works

- Users create their own demo accounts without approval; sign-in, password changes, and expiring sessions remain available.
- Automatic LAN startup and detected server URLs, without an administrator account or elevated launch.
- Personal copies and independent duplicates of saved inputs.
- Owner, editor, and viewer permissions; unshared copies are private.
- All 22 worksheets, including supporting sheets, in a paged grid.
- Editable/protected cell rules, validation, date inputs, and atomic multi-cell paste.
- Cell version checks that reject conflicting edits and an append-only change history.
- Shared updates checked every second while a workspace is open.
- Server-side native calculation for supported arithmetic, lookups, LET, conditional logic, aggregates, text functions, and several array functions.
- Explicit saved-baseline, stale, calculated, unsupported, and pending calculation states.
- Administrator CSV source refresh with formula preservation, removed-row clearing, fingerprint checks, and events for all copies.
- A server-owned SQLite database in WAL mode and a verified online backup command.

The SQL adapter also retains optional PostgreSQL support for an externally provided database. That requires its Python driver and an IT-provided connection; it is not the default, is not installed by the launcher, and was not tested here. Benchmark actual usage before choosing a larger deployment.

## Planner workflow

1. Create your demo account (or sign in), choose **New copy**, and name the business unit or scenario.
2. Open **BU recap**. Select a blue input cell, enter its value in the inspector, and save. Date inputs use a date picker and Excel date serials.
3. Use **Sales plan** and **Projections** for their inputs; all original sheets remain available in the selector.
4. Choose **Calculate visible cells** to evaluate the selected range. Calculations run on demand, not as a complete background recalculation of every formula. Unsupported functions and work limits are shown explicitly. Narrow the range if necessary.
5. Review coupons and reports. Amber values are saved results requiring recalculation after edits or source changes.
6. Choose **Manage access** to add an existing colleague as editor or viewer. An administrator does not automatically gain access to every private plan.
7. Choose **Duplicate** for an independent copy of saved inputs. Shared source refreshes intentionally apply to all copies.
8. Use **Change history** to review edits. Synchronization is approximately one second and requires a connection; no offline merge queue is provided.

**Paste / import** accepts up to 2,000 cells in one atomic request. Every target must pass its input rule; formula injection into inputs is blocked. CSV export includes saved baseline results and edited inputs, with its unverified calculation status. It is not a full-fidelity XLSM export.

## Demo accounts and sharing

Every person creates a separate account through **Create demo account**. Accounts are ordinary planners with no administrator privileges. Each person creates private copies, duplicates saved inputs, and shares a chosen copy with an existing colleague through **Manage access**. Ask a colleague to sign up before sharing with their email. Choose editor for shared changes or viewer for read-only review. Users change their passwords through **My account**; a password change invalidates their sessions.

One account is not shared across the team: separate identities preserve private copies, permissions, and names in the change history. On a shared computer, sign out before the next person signs in. On different computers, people use the same server URL.

## Refresh shared sources

Demo startup uses the saved baseline and creates no administrator. Global source refresh stays restricted to an administrator in managed mode; self-signup does not grant that privilege. For this demo, use planning inputs and per-copy paste. The existing source command is retained for a later controlled setup.

Place a reviewed UTF-8 CSV in `sources/`. Confirm its sheet and starting-cell mapping against the original template. The optional `--header` verifies the first row before importing rows below it. For a Rotation export matching A1:G:

```bat
run.bat --source-refresh --file sources/Rotation.csv --sheet Rotation --start A1 --header --actor your-admin-email
```

On other platforms, use `python3 run.py` with the same arguments. This command uses `settings.json` and the same data directory as the server. It can run while the server is open and publishes changes transactionally. It rejects invalid actors, mismatched headers, out-of-bounds data, and formula-like source content. It preserves source formulas, clears old raw tails, skips identical files, and sends a refresh event to each copy.

Validate mappings against a disposable database first. Calendar has a wide layout; LY sources may start at B3 to preserve their key column. The actual Calendar, rotational and MicroStrategy files were not supplied. Power Query, workbook Windows paths, and scheduled feed integration are not executed by this build.

## Backups and recovery

From Command Prompt in this folder:

```bat
run.bat --backup
```

This can run while users are connected. It uses SQLite's backup API, includes committed WAL changes, verifies the resulting database, and saves a uniquely named file under `backups/`. It closes database handles before renaming for Windows compatibility. Do not replace this with a raw copy of an active database. Keep backups in your approved backup destination along with the application, original workbook, SOP, branding, and settings. Backups contain user and planning data.

For a restore, stop the server. Copy a verified backup into a **new empty local directory**, name it `planning.db`, and point `settings.json`'s `data_dir` to that directory. Preserve the previous data directory for rollback. Restart with `run.bat`. Run an isolated restore check before relying on recovery.

Errors are recorded in `data/logs/server.log` (or the configured data directory), with bounded rotating files. Workbook import progress appears in the server window. A port-in-use error means stop the existing process or configure another port; the launcher does not terminate other programs.

## Architecture and verification

Browsers call the authenticated JSON API. One immutable workbook baseline is shared; each copy stores versioned input overrides rather than another 50 MB workbook. The server stores accounts, sessions, membership, overrides, source overlays, and events in its local SQLite database. Transactions serialize writes; stale cell versions cannot overwrite newer changes. Formula evaluation executes neither VBA nor arbitrary Python/JavaScript.

The database stays on the server's local disk. SQLite WAL requires readers and writers on the same host; team members connect to the web API, not a shared database file. Reference: https://www.sqlite.org/wal.html

```bash
python -m unittest discover -s tests -v
python ops/check_parity.py
```

Twenty-three automated checks passed on Linux, including demo startup without stdin/admin/settings prompts, self-signup without elevated roles, separate copy permissions, sharing, wrong-origin rejection, LAN address preference, managed-mode initialization/restart, duplicate launch protection, and WAL-safe backup/restore. Full demo collaboration checks also passed through bundled Waitress with the actual imported workbook. Windows batch execution, Windows file-lock behavior, HTTPS, browser visual QA, and capacity testing remain unverified on the target machine.

`ops/smoke_http.py` creates QA users and copies; use only an evaluation database, with `PLANNING_TEST_MODE=1`. `ops/check_parity.py` compares a sample with the supplied saved Excel cache, not fresh Excel outputs. It matched **101 of 103** values; two differences remain. See `docs/validation.md` before operational use.

## Folder layout

```text
run.bat              Single Windows LAN demo entry point
run.py / run.sh      Direct Python / Unix entry points
settings.json        Automatically created demo/access/storage settings
settings.example.json
app/                 API, database, importer, calculations, runtime, sources
static/              Browser interface and styles
vendor/              Bundled Waitress wheel and license
ops/                 Optional proxy example and verification tools
docs/                Workbook audit, VBA inventory, parity and acceptance report
tests/               Calculation, collaboration, and operating checks
sources/             Your reviewed source exports
logo.png             Your Catalyst Brands logo, supplied by you
data/                Created locally; users, copies, inputs, events, logs
backups/             Created by the backup command
Price Planning Template Spring 2027.xlsm
Price Plan Template - Overview.docx
```

The download excludes the development database, QA accounts, credentials, local settings, logs, and backups. Your original XLSM and SOP are preserved unchanged.
