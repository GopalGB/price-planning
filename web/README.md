# Price Planning frontend

The app speaks `docs/API.md`. The production server serves committed `static/dist/`; Node is required only for development and rebuilding.

## Development

Run `npm ci --prefix web`, then start these in separate terminals from the repository root:

```
npm run mock --prefix web
npm run dev --prefix web
```

Open `http://127.0.0.1:5173`. Sign up with a synthetic name/email and a password of at least 12 characters. The mock automatically joins demo users to Category A so sharing can be tried without setup. It runs only on loopback, stores everything in memory and never reads a workbook.

Both Univer packages are pinned to 1.0.4. Only server values enter the grid. Formula text is fetched separately for the read-only formula bar. Startup calculation and later formula engine commands are blocked. The grid retains at most three 256-row windows, plus at most 10,000 temporary input records while an operation hydrates offscreen targets. Server echoes supply actual cell versions for undo; an accepted edit without live confirmation refreshes values and resets local history.

The mock computes dependent columns C and D on the server. Its export is a generated synthetic OOXML `.xlsm` package containing columns A-D, with recalculation requested on open. It has no VBA or original template formatting. Real template-preserving export belongs to the backend contract.

## Checks and build

```
npm test --prefix web
npm audit --prefix web
npm run build --prefix web
```

The build writes `static/dist/BUILD.json`: SHA-256 over sorted relative file names and bytes for `web/` and the bundle, excluding `node_modules` and `BUILD.json`, plus the pinned Univer version and dependency licenses. Rebuilding unchanged source must reproduce the same manifest and bundle hashes.

Windows/LAN clipboard behavior, desktop Excel export fidelity, real backend integration, and measured scrolling frame rate require office acceptance. The Mac synthetic run does not establish those results.
