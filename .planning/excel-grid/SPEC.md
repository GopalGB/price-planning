# Price Planning Web App v2 - Excel-like multi-user planning (SPEC)

Date: 2026-10-10 · Owner: Sir (G) · Status: DRAFT for Sir's review
Inputs: FACTS.md, RESEARCH.md, SPIKE-ENGINE.md, astra/architecture-critique.md

## 1. Goal

Excel users in a category (example: Category A, 5 planners) open one URL on the office LAN and work on the
Spring 2027 price-plan model the way they work in Excel: double-click a cell, type, drag the fill handle,
paste from Excel, undo. Every formula of the original template is always there and recalculates live.
Anyone can share a copy for live co-editing, or fork a private one. Every input change shows who made it and when.

### Decisions locked with Sir (2026-10-10)

| # | Decision |
|---|---|
| D1 | Code lives in a PRIVATE GitHub repo `GopalGB/price-planning`. The .xlsm, .docx, data, logs, backups are NEVER committed. |
| D2 | Runs on one office Windows PC/server (16 GB+ RAM) on the LAN. Colleagues use a browser. No cloud. |
| D3 | v1: template formulas are fixed and locked. Users edit input cells only, with full Excel-style editing. |
| D4 | Copies: shared copy with live co-editing + personal forks. No merge-back in v1. |
| D5 | Architecture A: Univer OSS grid in the browser, Formualizer engine on the server, existing accounts/copies/history kept. |
| D6 | Security gate includes Strix, run WITHOUT Docker (local backend), Sir approves each paid run. |
| D7 | Build team: Claude plans and verifies; Codex `gpt-6.1-sol` (ultra) and Haiku execute bounded tasks; Codex `gpt-6-astra` critiques every design step and reviews every diff (read-only). |

### Success criteria (acceptance, measured on the office server)

1. **Feels like Excel:** the interaction checklist in section 5 passes in Edge and Chrome at the LAN URL.
2. **Correct:** parity harness reports 0 silent mismatches against an Excel-recalculated reference copy
   (types, errors, spills, displayed values; numeric tolerance 1e-9 relative, documented).
3. **Fast:** p95 ordinary input edit to authoritative values on screen <= 500 ms; p95 broad edit <= 2 s;
   sheet switch <= 1.5 s; scrolling 60 fps on the largest sheet (Sales by Mkting Block, 4088 x 312).
4. **Multi-user:** 5 editors on one copy plus 10 active copies with no cross-copy contamination,
   no lost edits, and conflicts reported to the user, never silently overwritten.
5. **Traceable:** every input cell shows last editor + time; full per-cell and per-copy history.
6. **Secure:** security gate (section 9) green; Strix finds no HIGH/CRITICAL.
7. **Operable:** one double-click `run.bat` start on Windows without admin rights; verified backup + restore.

## 2. Out of scope for v1 (named so nobody assumes them)

Users writing new formulas or editing template formulas · inserting/deleting rows or columns · VBA macros ·
pivot tables, slicers, charts · merge of forks back into a master · offline editing · SSO/LDAP ·
internet hosting. Conditional formatting arrives in Phase 4, not on day one.

## 3. Architecture

```
Browser (Univer OSS, client calc OFF)          Office server (Python 3.10+, Windows, no admin)
 ├ grid render, selection, edit, fill,   WS    ├ ASGI app (Starlette + Uvicorn, vendored wheels)
 │ paste, undo, presence  <──────────────────> │  ├ auth, sessions, CSRF, rate limits
 ├ shows server values only                    │  ├ copy/category/permission service
 └ optimistic input echo, "calculating" badge  │  ├ op log (SQLite WAL): ops, versions, history
                                               │  ├ calc scheduler -> warm engine pool
                                               │  │    one Formualizer instance per ACTIVE copy,
                                               │  │    LRU eviction by measured RAM budget
                                               │  ├ template store: parsed once from the .xlsm
                                               │  │    (styles, formats, merges, freeze, validation,
                                               │  │     lock mask, names) -> per-sheet snapshot files
                                               │  └ export: patch inputs into the original .xlsm
                                               └ data/: planning.db, snapshots/, logs/, backups/
```

### 3.1 Template store (one per template version)

- On first start the server opens the .xlsm once and produces: per-sheet render snapshots (values, number
  formats, styles, merges, column widths, row heights, frozen panes, hidden rows/cols/sheets, data validation),
  the **input mask** (which cells users may edit: unlocked cells on protected sheets, plus the existing
  prototype's input rules), and the template hash. Snapshots are streamed, never the whole 650 MB XML in memory.
- The baseline engine state (Formualizer workbook, fully evaluated) is cached to disk if the engine supports it;
  otherwise it is rebuilt at start and the cost is reported in the startup log.

### 3.2 Engine adapter (`app/engine/`)

- One interface: `load(template)`, `apply(ops) -> changed cells`, `get(range)`, `dependents(cell)`, `evaluate_dirty()`.
- Formualizer is the only implementation in v1. An Aspose.Cells adapter is the documented fallback (paid, Sir's call).
- **Formualizer is PROVISIONAL.** The 2026-10-10 spike measured only load (about 5 s, 150 MB); evaluation memory,
  parity and edit latency are unmeasured. P1 blocks P2 until two consecutive real-workbook replays on the office
  Windows server pass baseline and post-edit parity, latency and memory targets. Ordinary, broad and worst-fan-out
  edits are fixed in `ops/bench_manifest.json`. A fallback decision alone does not pass P1.
- **Compatibility shims and semantic parity, required before P1 exits:**
  - `ISFORMULA` (234 uses): Formualizer returns FALSE. Answer from the template formula map at load time.
  - `CELL` (4 uses), `INDIRECT` (2 uses), `OFFSET` (2 uses): verify; implement via engine custom functions or
    precompute if their arguments are static.
  - `SWITCH` (87k uses): semantic test set before trusting it.
- **Static-region optimisation:** share only formulas PROVEN pure and transitively independent of every input
  cell and every source overlay. Any formula with an unresolved dependency (INDIRECT, OFFSET, CELL, dynamic
  ranges) stays dynamic. Precompute reference targets, not mutable values. Source overlays are versioned per
  copy/fork. Engine-level sharing is used only if the engine demonstrably supports it; otherwise the RAM budget
  assumes independent instances.
- **Calculation runs in worker processes, never on the ASGI event loop.** Each copy's engine is serialized to one
  worker at a time and always evaluates an exact reconstructed revision.

### 3.3 Copies, categories, permissions

- **Category:** a named group (example "Category A") with members (category roles: manager, member).
  Any member creates a copy in it. Managers add/remove members.
- **Copy:** template version + its input overrides. Visibility: `private` (owner + named shares only),
  `category` (every category member gets the copy's default role, editor or viewer, chosen by the owner),
  plus named shares that may only target existing users. Roles per copy: owner, editor, viewer; the highest of
  inherited and named role wins.
- **Authorization on EVERY read and write:** copy listing, snapshots, cell values, history, export, fork,
  WebSocket subscribe. Removing a member or share revokes live access immediately (open sockets are closed).
- **Fork:** created atomically at one revision = parent's overrides frozen at that revision, plus a provenance
  record (parent, revision, who forked). Later parent edits never leak in. History in the fork shows inherited
  values as "from parent @ revision" and new edits by their own editors.
- Storage stays sparse: a copy is a list of input overrides, not another 53 MB file.

### 3.4 Edit pipeline (server-ordered op stream per copy, no CRDT)

1. Browser sends `{op_id, copy_id, cells:[{ref, value, expected_version}]}`. One paste or fill = one op.
2. Server first deduplicates: `op_id` is unique per copy and bound to actor + payload hash; an identical retry
   returns the stored acknowledgement, a different payload under the same `op_id` is rejected.
3. Inside `BEGIN IMMEDIATE` (SQLite busy_timeout 5 s, bounded retries) the server re-checks role, input mask,
   types, data validation, size limit (10,000 cells) and expected versions, then commits overrides, new cell
   versions, copy `input_revision`, history rows and a durable recalc job row. Any failure rejects the whole op
   with a clear message (no partial apply).
4. Calc runs outside the transaction in a worker against that exact revision. Results are published as one
   complete batch tagged `calculated_revision`, monotonically: an older job can never overwrite a newer result.
   The UI shows "calculating" while `calculated_revision < input_revision` and never labels old values current.
5. All editors of the copy receive the op and the changed values over WebSocket. On reconnect a client sends its
   last revision and gets either an ordered replay or a revision-consistent snapshot. Remote patches never create
   outgoing edits or local undo entries. Presence (who is where) is a separate ephemeral channel.
6. Undo = a compensating op with version checks. It never silently undoes someone else's later edit.
7. Crash recovery: on start, replay committed ops whose recalc job is unfinished; idempotent by `op_id`.

### 3.5 Grid (Univer OSS)

- Client formula calculation disabled; cells render values from the server. UNPROVEN: Univer documents
  `NO_CALCULATION` for initialization only. **Univer gate (end of P0, before P2 starts):** with pinned OSS
  package versions, prove (a) bounded row loading and eviction, (b) zero client recalculation across edit, paste,
  fill and undo, (c) server values preserved after paste/undo, (d) offscreen copy and find work, on synthetic
  4088 x 312 and 100000 x 17 sheets, heap < 1.5 GB. If Univer fails, Glide Data Grid replaces it only after
  passing the SAME gate including merges, freezes, styles and editing.
- Only the active sheet is loaded. Large sheets stream in row blocks with styles.
- Locked cells look and behave locked (grey hint, edit refused locally and on server). Input cells keep the
  template's blue input styling.
- Formula bar shows the formula text of a locked cell (read only) and the value of an input cell.

### 3.6 Export

"Download as Excel" writes the copy's input values into a copy of the ORIGINAL .xlsm by patching only those
cells' XML, keeping VBA, formulas, pivots and formats intact. Excel recalculates on open (fullCalcOnLoad).

## 4. Data model (SQLite WAL, one file on the server's local disk)

users, sessions, categories, category_members, templates, copies (template_id, parent_copy_id,
fork_revision, visibility, input_revision, calculated_revision), copy_members (role), overrides
(copy_id, sheet, row, col, value, type, version, updated_by, updated_at), ops (op_id, copy_id, actor,
created_at, payload), history (op_id, cell, before, after), calc_jobs, source_overlays, events.
Existing prototype tables are migrated, not dropped.

## 5. Excel interaction checklist (v1 acceptance)

Double-click / F2 / start typing to edit · Enter, Tab, Shift+Enter, arrows, Ctrl+arrows, Ctrl+Home/End ·
Shift / Ctrl multi-range selection · fill handle drag (copy and series), double-click fill handle ·
Ctrl+D / Ctrl+R · copy/paste to and from desktop Excel (TSV + HTML), paste values only into inputs ·
Delete clears inputs · Ctrl+Z / Ctrl+Y · Ctrl+F find · sheet tabs (template's hidden sheets stay hidden,
"show hidden" toggle for owners) · frozen panes, widths, heights, number formats, fonts, fills, borders,
merges as in the .xlsm · data validation dropdowns and date pickers · status bar Sum/Average/Count of
selection · zoom · hover shows last editor + time on input cells · per-cell history panel · presence colours
for other editors · "calculating" indicator · conflict message on stale edit.
Note: the async Clipboard API needs HTTPS or localhost; keyboard Ctrl+C/V works over the LAN http URL.
Toolbar copy/paste buttons are tested at the real LAN URL.

## 6. Repo and multi-agent workflow (Mac <-> GitHub <-> office)

- Repo root = this folder as a standalone git repo (`origin` = github.com/GopalGB/price-planning, private).
  Max HQ ignores `/price-planning/` (done 2026-10-10). P0 records root, origin and migrated files in STATE.md.
- Runtime: CPython 3.10 to 3.12 on Windows 10/11 x64. Uvicorn runs with the asyncio loop and the `websockets`
  backend; every dependency is a pinned, hash-checked wheel in `vendor/wheels/` installed with
  `pip install --no-index --require-hashes`. P0 proves offline, no-admin install and a second LAN machine
  connecting. Firewall: IT must allow the chosen port (no app can bypass it). Production uses HTTPS/WSS with
  Secure cookies behind the company's certificate; plain http is for the LAN demo only.
- Every handoff (Mac, office, Claude cloud) ends by updating `.planning/STATE.md`: branch, commit, tests run,
  next task. Every harness is told to read AGENTS.md first. Deploys: tested `main` commit, backup, migrate,
  rollback steps in `docs/DEPLOY.md`.
- `main` is always runnable. All work on branches `feat/...`, `fix/...`, merged by PR.
- `AGENTS.md` (single source; read natively by Codex, Claude Code, Claude cloud) holds: architecture summary,
  rules (no .xlsm in git, no secrets, tests before PR, no em-dashes, no co-author trailers), how to run tests,
  the security gate, and the current phase from `.planning/STATE.md`.
- `.gitignore` blocks `*.xlsm *.xlsx *.docx data/ backups/ settings.json logs/ sources/*.csv`, plus a
  pre-commit hook that refuses any file > 5 MB or any spreadsheet extension.
- CI (GitHub Actions, free private-repo minutes): unit + integration tests on SYNTHETIC fixtures only
  (deterministic generator: semantic fixtures for blanks vs "" vs 0, text, errors, spills, plus scale fixtures
  with the real sheet dimensions; no company data), frontend build check, semgrep, gitleaks, bandit, pip-audit,
  npm audit, and a content allowlist that fails the PR on any spreadsheet/office file, any file > 5 MB not on the
  vendored-binary allowlist, or business strings from a local denylist. A PR cannot merge red.
- The frontend bundle under `static/dist/` is built and committed ON THE PR BRANCH by the developer or agent.
  CI rebuilds deterministically and fails on any difference; CI never pushes. Source hash, bundle hash and
  dependency licences are recorded in `static/dist/BUILD.json`. The office server never needs Node.
- Office flow: `git pull` then `run.bat`. Office Codex/Astra sessions push branches and open PRs; this Mac pulls them.
- Real-workbook tests (`ops/bench_real.py`, `ops/parity_real.py`) run only where the .xlsm exists. Their
  COMMITTABLE report holds aggregate metrics, versions and hashes only (timings, RAM, match counts). Mismatch
  details, formulas, cell values, business labels and raw logs stay in local `results-private/` (gitignored).
- Parity oracle: typed goldens generated by desktop Excel (baseline and after each manifest edit), distinguishing
  blank, empty string, zero, text, error type and spill size. Any engine or importer change re-runs real-workbook
  acceptance.

## 7. Phases (each ends with tests green + Astra diff review + Sir-visible demo)

| Phase | Delivers | Exit gate |
|---|---|---|
| P0 Repo + rails | Private repo, AGENTS.md, STATE.md, CI + content allowlist, size hook, synthetic fixture generator, vendored hashed wheels, `ops/bench_real.py` + manifest, Univer gate prototype | CI green; offline no-admin install + second-machine LAN test pass on Windows; Univer gate pass (or Glide passes same gate); Sir runs bench on office server and pushes the aggregate report |
| P1 Engine | Engine adapter, compat shims, static-region proof, typed-golden parity harness | Two consecutive office replays pass parity, latency and memory targets. If not: fallback brought to Sir; P2 stays blocked |
| P2 Server | ASGI app, migrated accounts/copies/history, categories, op stream, warm pool, WebSocket | Concurrency + crash-recovery tests green |
| P3 Grid | Univer UI, streamed snapshots, full section 5 checklist, presence, history UI | Checklist passes at LAN URL, heap gate met |
| P4 Fidelity | .xlsm export, conditional formats, CSV sources, backups, Windows launcher | Export opens in Excel and matches; restore drill passes |
| P5 Release | Strix local run, security fixes, office acceptance with real users | Section 1 success criteria all met |

## 8. Risks (ranked) and responses

1. **Plausible but wrong numbers** (IFERROR hides engine differences): differential parity vs Excel-recalculated
   reference, including after edits. Sir supplies one copy recalculated and saved in Excel.
2. **One input invalidates huge lookup chains:** measure fan-out of the worst real input in P1; static-region
   sharing; per-copy calc queue so one slow copy never blocks others.
3. **RAM per warm copy:** measured in P1 on the 16 GB server; LRU pool sized from the measurement.
4. **Browser memory on 1.27M-cell sheet:** streamed rows; heap gate in P3 with Glide fallback for that sheet.
5. **Inputs and results diverging after crash/race:** revision-tagged results, durable jobs, idempotent replay.
6. **Company data leaving the office:** D1 rules, size hook, CI uses synthetic data only, Strix sees code only.
7. **Formualizer is young (0.11.x):** pin the version, vendor the wheel, keep the adapter thin so Aspose or a
   patched fork can replace it.

## 9. Security gate

Every PR: semgrep (p/python, p/javascript, p/owasp-top-ten), gitleaks, bandit, pip-audit, npm audit,
Astra read-only security review of the diff. App controls: scrypt password hashes, httpOnly SameSite=Strict
session cookie, CSRF token on every write, WebSocket auth bound to the session and Origin-checked,
login rate limit 5/min, write rate limit, server-side role + input-mask check on every cell, formula and
CSV-injection blocking on input values, op size limits, no stack traces to clients, audit log.

**Strix without Docker (status: `STRIX_ISOLATION_BLOCKED` until containment tests pass).** The agents SDK
`UnixLocalSandboxClient` gives NO filesystem or network isolation by itself; budget caps and approval are not
containment. Plan: a registered `local` backend in `security/tools/strix-local.py` (Max HQ) that launches every
Strix tool command under macOS `sandbox-exec` with a deny-by-default profile: writes only inside a throwaway work
directory, reads only the app source copy and tool binaries, network only to 127.0.0.1:<test port> plus the LLM
provider endpoint, and an environment stripped of every credential except the Strix provider key. Execution host:
this Mac only. Caido CLI, sqlmap and nuclei installed locally. Containment tests that must pass first: attempts to
read `~/.ssh`, `.env`, the real .xlsm, and to reach any non-allowed host all fail. Then: target only the test
instance with a throwaway database and fake users, quick mode, `--max-budget-usd 1`, Sir's OK per run.
Backups: SQLite online backup API only (never copy the live file); backup records template and schema hashes;
restore drill restores into an empty directory and checks integrity, foreign keys, revisions and history while
writes are running.

## 10. Open items Sir provides

1. One copy of the template, populated and freshly recalculated + saved in desktop Excel (parity reference).
2. Office server: run `ops/bench_real.py` after P0 and push the results file.
3. Category names and who belongs to each (can start with one test category).
