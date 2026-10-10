# Track B frontend handoff

Date: 2026-10-10. Branch: `feat/v2-frontend`. Status: DONE_WITH_CONCERNS for local synthetic frontend work; office acceptance remains pending. Implementation commit: `1061a45`; committed bundle and browser evidence: `6baab71`. No push performed.

[HIGH] Built the Vite app under `web/` and committed `static/dist/`. Reused the API contract, proto visual language and resolved Univer gate imports; backend paths `pp/`, `app/`, and `ops/` were not changed. No private workbook paths or external Office files were accessed. All fixtures and screenshots are synthetic.

The UI includes login/signup, category membership, Copies, create/fork/share dialogs, named roles, Excel green app bar, toolbar, external formula bar, sheet tabs, status summary and history. The active sheet applies metadata and loads three 256-row blocks. Input-only edit, Delete, paste and fill send atomic versioned server operations with optimistic rollback. Undo/redo use accepted server cell versions, preserve remote conflicts and reset history on missing live confirmation. Formula text stays outside Univer; calculation startup and later engine commands are blocked. Live ops/calcs, calculation state, reconnect/resync and validated presence are wired.

The loopback mock implements auth, categories/members, copies/fork/shares, workbook metadata, streamed rows, cell detail, history, ops, WS and synthetic download. It computes dependent values server-side for 4088x312 and 100000x17 sheets. Demo users automatically join Category A. Its generated `.xlsm` contains demo columns A-D and recalculation flags, with no original template or VBA. Production template-preserving export remains the backend's responsibility.

## Verification

| Check | Result |
|---|---|
| `npm test --prefix web` | 33 passed, 0 failed, two consecutive greens |
| Node coverage over state, grid helpers and views | 97.12% lines; 65.84% branches; DOM/Univer adapter excluded |
| `.venv/bin/python -m unittest discover -s tests` | 64 passed, 0 failed, consecutive greens |
| `python3 ops/content_guard.py --all` | Exit 0 after staging committed bundle/evidence |
| `npm audit --prefix web --json` | Zero vulnerabilities |
| Semgrep JavaScript + OWASP rules on web source/mock/scripts | Zero findings, zero errors |
| `security-auditor --staged` | Gitleaks clean over 11.68 MB of staged source and bundle |
| Exact `claude-opus-5-5` review | Returned identity verified; SHIP, no remaining MEDIUM/HIGH/CRITICAL |
| `ego-browser nodejs < web/test/ego.mjs` | Exit 0; one TaskSpace (35), `finish({keep: []})` completed |
| `npm run build --prefix web` twice | Identical source/bundle hashes; largest asset 1,460,991 bytes |

The browser passed signup, copy creation, native input edit and dependent WS update, locked-cell refusal, typed formula rejection, server undo, native paste, native fill, history, sharing and viewer refusal. Client calculation count and client formula count were zero. At row 99,000 the grid held 3 cached blocks and 768 populated rows; sampled heap was 93.753361 MB. This is one sampled heap measurement, not a peak bound.

Five screenshots in `frontend-shots/` were visually inspected: Copies home, initial workspace, history, long sheet and viewer. They remain local because the existing content guard rejects PNG binaries. No guard was bypassed. `frontend-validation.json` commits their SHA-256 checksums and sizes. PNG inspection found no C2PA or AI metadata.

`BUILD.json` pins Univer 1.0.4 and records dependency licenses. Source SHA-256: `4ef0833ecc6afa47880ecc112997b27492f20bd2ae4bbdecedfff8b14e6b6119`. Bundle SHA-256: `9a06b98eeca33881b888da8c4100642686413ce17e64912e0c0d0ba633505239`. Vite 6.4.4 and ws 8.22.0 are exact devDependency pins; earlier vulnerable dev versions were replaced. No Python dependencies were added.

Full original test/build/review/scan output remains in ignored `frontend-*.log` files in this worktree. Committed aggregates: `frontend-ego.json`, `frontend-review.json`, `frontend-validation.json`, `frontend-build-first.json`. Developer commands and mock limitations are in `web/README.md`.

## Not done or not verified

- Real backend integration, real workbook formula/VBA/export fidelity, Windows/Edge/Chrome office LAN acceptance, desktop Excel clipboard behavior, scrolling FPS, latency targets and peak heap were not verified.
- The full section 5 interaction matrix was not browser-tested. Offscreen copy/find, multi-range shortcuts, data validation controls, named presence with a second client, forced WS reconnect/resync and live role downgrade still need integrated acceptance. Their implementation is present; the core synthetic browser checks above are the verified subset.
- Reviewer LOW notes remain: history refresh briefly replaces the formula display while details reload; selecting a sparse blank cell can retain the previous history panel content. No MEDIUM or higher finding remains.
- Screenshots are local only. No push, PR, merge or deployment was performed, as instructed.

Next: backend owner integrates the committed `static/dist/` and runs the remaining office acceptance checks. The Mac local Track B report does not certify full P3 office release readiness.
