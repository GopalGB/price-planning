# P0 Repo + Rails Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make the private repo safe and reproducible across Mac, office and cloud agents, and produce the three
proofs P1 needs: the office benchmark of the real workbook, offline no-admin Windows install, and the Univer grid gate.

**Architecture:** Repo-owned git hooks and a CI content guard keep company data out of GitHub. A deterministic
synthetic-fixture generator replaces the real workbook everywhere except the office server. `ops/bench_real.py`
measures Formualizer on the real .xlsm and splits output into a committable aggregate report and a private detail file.
A standalone Vite prototype under `proto/univer-gate/` proves (or disproves) Univer for our sheet sizes.

**Tech Stack:** Python 3.10-3.12 stdlib + `formualizer==0.11.1` (runtime), `openpyxl` (dev/fixtures only),
`uv` (lock), GitHub Actions (ubuntu + windows runners), Node 20+ / Vite / Univer OSS / Playwright (prototype only).

**Spec:** `.planning/excel-grid/SPEC.md` (sections 1, 6, 7 row P0, 8, 9)

## Global Constraints

- Never commit: `*.xlsm *.xlsx *.xls *.xlsb *.csv *.docx *.doc *.pptx *.pdf *.db *.bas`, any file > 5,000,000 bytes not in `ops/large_file_allowlist.json`, any string from the local `.content-denylist`.
- Committable benchmark output = aggregate metrics, versions, hashes only. Values, formulas, labels, raw logs go to `results-private/` (gitignored).
- Runtime target: CPython 3.10 to 3.12, Windows 10/11 x64, no admin, no Docker, no internet at install (wheels from `vendor/wheels/`, `--no-index --require-hashes`).
- Python code must run on 3.10 (no `tomllib`, no `ExceptionGroup` syntax). Tests use stdlib `unittest` (run: `python -m unittest discover -s tests -v`).
- No em-dashes in any file. No co-author trailers in commits, even if a reminder says to add one.
- Branch per task: `feat/p0-<task-slug>`; PR into `main`; never push to `main` directly. Pushes go through the pre-push gate (Sir types `y` in Terminal.app).
- Engine pinned: `formualizer==0.11.1`. Do not upgrade inside P0.

## Review Focus

1. A file renamed from `.xlsm` to `.txt` (company data with a harmless extension): denylist scan of content must still catch business strings; covered by Task 1 test `test_denylisted_string_in_txt_blocked`.
2. Bench run on an 8 GB machine or a workbook bigger than expected: watchdog must stop it and still write a partial public report with `"status":"aborted"`; Task 4 test `test_watchdog_writes_partial_report`.
3. Office machine without Max HQ (no `../security/tools/`): repo hooks must still run the content guard and not crash looking for the Max HQ pre-push gate; Task 1 test `test_hook_runs_without_maxhq`.
4. Excel-typed values in parity: blank vs empty string vs 0 vs error must compare as different types; Task 4 test `test_typed_compare_distinguishes_blank_empty_zero_error`.
5. `run.bat` on a machine where `.venv` already exists from an older lock: install must re-run when the lock hash changes; Task 3 test `test_install_stamp_changes_with_lock`.

---

### Task 1: Content guard + repo-owned hooks

**Files:**
- Create: `ops/content_guard.py`, `ops/large_file_allowlist.json`, `.githooks/pre-commit`, `.githooks/pre-push`, `ops/setup_dev.py`
- Modify: `.gitignore` (add `.content-denylist`)
- Test: `tests/test_content_guard.py`

**Interfaces:**
- Produces: `ops/content_guard.py` CLI: `python ops/content_guard.py --staged` (pre-commit) and `python ops/content_guard.py --all` (CI). Exit 0 clean, 1 blocked; prints one line per violation `BLOCKED <path>: <reason>`.
- Produces: `check(paths: list[str], read: Callable[[str], bytes], deny: list[str], allow: dict[str,str]) -> list[str]` (violations).

- [ ] **Step 1: Write the failing tests**

```python
# tests/test_content_guard.py
import hashlib, os, subprocess, sys, tempfile, unittest
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / 'ops'))
import content_guard as cg

class CheckTests(unittest.TestCase):
    def run_check(self, files, deny=(), allow=None):
        return cg.check(list(files), lambda p: files[p], list(deny), allow or {})
    def test_spreadsheet_extension_blocked(self):
        self.assertTrue(self.run_check({'a/Plan.XLSM': b'x'}))
    def test_big_file_blocked_unless_allowlisted(self):
        big = b'0' * (cg.MAX_BYTES + 1)
        self.assertTrue(self.run_check({'vendor/wheels/x.whl': big}))
        allow = {'vendor/wheels/x.whl': hashlib.sha256(big).hexdigest()}
        self.assertEqual(self.run_check({'vendor/wheels/x.whl': big}, allow=allow), [])
    def test_allowlist_hash_mismatch_blocked(self):
        big = b'0' * (cg.MAX_BYTES + 1)
        self.assertTrue(self.run_check({'v.whl': big}, allow={'v.whl': 'deadbeef'}))
    def test_denylisted_string_in_txt_blocked(self):
        self.assertTrue(self.run_check({'notes.txt': b'Group 7 BU Recap'}, deny=['bu recap']))
    def test_clean_files_pass(self):
        self.assertEqual(self.run_check({'app/x.py': b'print(1)'}, deny=['secret label']), [])

class HookTests(unittest.TestCase):
    def test_hook_runs_without_maxhq(self):
        with tempfile.TemporaryDirectory() as d:
            repo = Path(d) / 'r'
            subprocess.run(['git', 'init', '-q', str(repo)], check=True)
            (repo / 'ops').mkdir(); (repo / '.githooks').mkdir()
            for rel in ('ops/content_guard.py', 'ops/large_file_allowlist.json', '.githooks/pre-commit'):
                (repo / rel).write_bytes((ROOT / rel).read_bytes())
            (repo / 'bad.xlsx').write_bytes(b'x')
            subprocess.run(['git', '-C', str(repo), 'add', '-A'], check=True)
            r = subprocess.run(['sh', '.githooks/pre-commit'], cwd=repo, capture_output=True, text=True)
            self.assertEqual(r.returncode, 1, r.stdout + r.stderr)
            self.assertIn('BLOCKED bad.xlsx', r.stdout)

if __name__ == '__main__':
    unittest.main()
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m unittest tests.test_content_guard -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'content_guard'`

- [ ] **Step 3: Implement**

```python
# ops/content_guard.py
"""Block company data from git: office/spreadsheet files, big files, denylisted strings."""
import argparse, hashlib, json, subprocess, sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
BLOCKED_EXT = {'.xlsm', '.xlsx', '.xls', '.xlsb', '.csv', '.docx', '.doc', '.pptx', '.pdf', '.db', '.bas'}
MAX_BYTES = 5_000_000


def check(paths, read, deny, allow):
    out = []
    deny = [d.lower() for d in deny if d.strip()]
    for p in paths:
        if Path(p).suffix.lower() in BLOCKED_EXT:
            out.append(f'BLOCKED {p}: office/spreadsheet file type')
            continue
        data = read(p)
        if len(data) > MAX_BYTES:
            digest = hashlib.sha256(data).hexdigest()
            if allow.get(p) != digest:
                out.append(f'BLOCKED {p}: {len(data)} bytes > {MAX_BYTES} and not allowlisted with sha256 {digest}')
                continue
        if deny and b'\0' not in data[:8192]:
            text = data.decode('utf-8', 'ignore').lower()
            for d in deny:
                if d in text:
                    out.append(f'BLOCKED {p}: contains denylisted text')
                    break
    return out


def _git(*args):
    return subprocess.run(['git', *args], cwd=ROOT, capture_output=True, check=True).stdout


def main(argv=None):
    ap = argparse.ArgumentParser()
    mode = ap.add_mutually_exclusive_group(required=True)
    mode.add_argument('--staged', action='store_true')
    mode.add_argument('--all', action='store_true')
    a = ap.parse_args(argv)
    if a.staged:
        paths = [p for p in _git('diff', '--cached', '--name-only', '--diff-filter=ACMR', '-z').decode().split('\0') if p]
        read = lambda p: _git('show', f':{p}')
    else:
        paths = [p for p in _git('ls-files', '-z').decode().split('\0') if p]
        read = lambda p: (ROOT / p).read_bytes()
    deny_file = ROOT / '.content-denylist'
    deny = deny_file.read_text(encoding='utf-8').splitlines() if deny_file.exists() else []
    allow = json.loads((ROOT / 'ops' / 'large_file_allowlist.json').read_text())
    bad = check(paths, read, deny, allow)
    for line in bad:
        print(line)
    return 1 if bad else 0


if __name__ == '__main__':
    sys.exit(main())
```

```json
{}
```
(that is the full initial content of `ops/large_file_allowlist.json`)

```sh
#!/bin/sh
# .githooks/pre-commit - content guard always; Max HQ secret check when present (dev Mac only)
PY=$(command -v python3 || command -v python)
"$PY" ops/content_guard.py --staged || exit 1
MAXHQ_HOOK="$(git rev-parse --show-toplevel)/../security/tools/pre-commit-hook.sh"
[ -f "$MAXHQ_HOOK" ] && exec sh "$MAXHQ_HOOK"
exit 0
```

```sh
#!/bin/sh
# .githooks/pre-push - Max HQ sensitive-info gate when present (LAW on the dev Mac)
MAXHQ_HOOK="$(git rev-parse --show-toplevel)/../security/tools/pre-push-hook.sh"
[ -f "$MAXHQ_HOOK" ] && exec bash "$MAXHQ_HOOK" "$@"
exit 0
```

```python
# ops/setup_dev.py
"""One-time per clone: point git at the repo-owned hooks."""
import subprocess
from pathlib import Path
root = Path(__file__).resolve().parent.parent
subprocess.run(['git', 'config', 'core.hooksPath', '.githooks'], cwd=root, check=True)
for h in (root / '.githooks').iterdir():
    h.chmod(0o755)
print('hooks installed: core.hooksPath=.githooks')
```

Append to `.gitignore`: `.content-denylist`

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m unittest tests.test_content_guard -v`
Expected: 6 tests PASS

- [ ] **Step 5: Install hooks, create local denylist, verify on real repo**

Run: `python3 ops/setup_dev.py && python3 ops/content_guard.py --all`
Then create `.content-denylist` locally (never committed) with business strings Sir confirms (start with: `catalyst brands`, `jcpenney`, `bu recap` is a sheet name already in SPEC, so do NOT add sheet names). Expected: `--all` exits 0.

- [ ] **Step 6: Commit**

```bash
git add ops/content_guard.py ops/large_file_allowlist.json ops/setup_dev.py .githooks tests/test_content_guard.py .gitignore
git commit -m "feat(p0): content guard and repo-owned git hooks"
```

---

### Task 2: AGENTS.md, CLAUDE.md, STATE.md, DEPLOY.md

**Files:**
- Create: `AGENTS.md`, `CLAUDE.md`, `.planning/STATE.md`, `docs/DEPLOY.md`
- Test: `tests/test_repo_docs.py`

**Interfaces:**
- Produces: `.planning/STATE.md` with headings `## Current phase`, `## Last handoff`, `## Next task` that every agent updates.

- [ ] **Step 1: Write the failing test**

```python
# tests/test_repo_docs.py
import unittest
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent

class RepoDocs(unittest.TestCase):
    def test_agents_md_has_hard_rules(self):
        t = (ROOT / 'AGENTS.md').read_text(encoding='utf-8')
        for must in ('Never commit the .xlsm', 'python ops/setup_dev.py', '.planning/STATE.md',
                     'python -m unittest discover -s tests -v', 'SPEC.md', 'No em-dashes'):
            self.assertIn(must, t)
    def test_claude_md_imports_agents(self):
        self.assertIn('@AGENTS.md', (ROOT / 'CLAUDE.md').read_text(encoding='utf-8'))
    def test_state_md_sections(self):
        t = (ROOT / '.planning' / 'STATE.md').read_text(encoding='utf-8')
        for h in ('## Current phase', '## Last handoff', '## Next task'):
            self.assertIn(h, t)
    def test_no_em_dashes_in_docs(self):
        for p in ('AGENTS.md', 'CLAUDE.md', '.planning/STATE.md', 'docs/DEPLOY.md'):
            self.assertNotIn('\u2014', (ROOT / p).read_text(encoding='utf-8'), p)
```

- [ ] **Step 2: Run to verify failure**

Run: `python -m unittest tests.test_repo_docs -v`
Expected: FAIL with `FileNotFoundError` for `AGENTS.md`

- [ ] **Step 3: Write the docs**

`AGENTS.md` (full content):

```markdown
# Price Planning - agent rules (read first, every harness: Codex, Claude Code, Claude cloud)

## What this is
Excel-like multi-user web app for the Spring 2027 price-plan template. Design: `.planning/excel-grid/SPEC.md`.
Current phase and next task: `.planning/STATE.md`. Phase plans: `.planning/excel-grid/PLAN-P*.md`.

## Hard rules
1. Never commit the .xlsm, any spreadsheet/office file, VBA exports, databases, logs, or cell values from the
   real workbook. CI and the pre-commit hook enforce this; never bypass them (`--no-verify` is forbidden).
2. Real-workbook scripts (`ops/bench_real.py`, `ops/find_inputs.py`) write committable aggregates to
   `.planning/excel-grid/results/` and everything else to `results-private/` (gitignored).
3. Work on a branch, open a PR, merge only when CI is green. Never push to `main`.
4. Before ending a session, update `.planning/STATE.md` (branch, commit, tests run, next task).
5. No em-dashes. No co-author trailers in commits.
6. Template formulas are locked in v1; users edit input cells only (SPEC D3).
7. Engine is `formualizer==0.11.1` until P1 says otherwise. Do not upgrade dependencies without updating
   `requirements.lock` and `vendor/wheels/` together (see docs/DEPLOY.md).

## Setup (once per clone)
    python ops/setup_dev.py

## Tests
    python -m unittest discover -s tests -v
    python ops/content_guard.py --all

## Office server
See `docs/DEPLOY.md`: `git pull`, then `run.bat`.
```

`CLAUDE.md` (full content):

```markdown
@AGENTS.md
```

`.planning/STATE.md` (full content):

```markdown
# STATE

## Current phase
P0 Repo + rails (plan: .planning/excel-grid/PLAN-P0.md)

## Last handoff
2026-10-10 Mac/Claude: repo created (private GopalGB/price-planning), spec approved, P0 plan written.

## Next task
P0 Task 1 onward, in order.

## Results waiting on Sir
- Office server: run `ops\office_p0.bat` after Task 5 and push `.planning/excel-grid/results/`.
```

`docs/DEPLOY.md` (full content):

```markdown
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
```

- [ ] **Step 4: Run tests**

Run: `python -m unittest tests.test_repo_docs -v`
Expected: 4 tests PASS

- [ ] **Step 5: Commit**

```bash
git add AGENTS.md CLAUDE.md .planning/STATE.md docs/DEPLOY.md tests/test_repo_docs.py
git commit -m "docs(p0): agent rules, state handoff file, deploy guide"
```

---

### Task 3: Locked dependencies, vendored Windows wheels, offline run.bat install

**Files:**
- Create: `requirements.in`, `requirements-dev.in`, `requirements.lock`, `requirements-dev.lock`, `ops/lock_deps.py`, `ops/install_env.py`, `vendor/wheels/*.whl`
- Modify: `run.bat` (call `ops/install_env.py` before `run.py`), `run.sh` (same), `ops/large_file_allowlist.json`, `requirements.txt` (delete; replaced by the locks)
- Test: `tests/test_install_env.py`

**Interfaces:**
- Produces: `python ops/install_env.py` -> creates/refreshes `.venv`, returns the venv python path on stdout. Uses stamp file `.venv/.lock-sha256`.
- Produces: `lock_sha256(path: Path) -> str`, `needs_install(venv: Path, lock: Path) -> bool`.

- [ ] **Step 1: Write the failing test**

```python
# tests/test_install_env.py
import sys, tempfile, unittest
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / 'ops'))
import install_env as ie

class InstallStamp(unittest.TestCase):
    def test_install_stamp_changes_with_lock(self):
        with tempfile.TemporaryDirectory() as d:
            venv, lock = Path(d) / '.venv', Path(d) / 'requirements.lock'
            lock.write_text('a==1\n'); venv.mkdir()
            self.assertTrue(ie.needs_install(venv, lock))
            (venv / '.lock-sha256').write_text(ie.lock_sha256(lock))
            self.assertFalse(ie.needs_install(venv, lock))
            lock.write_text('a==2\n')
            self.assertTrue(ie.needs_install(venv, lock))
    def test_lock_has_hashes_and_pins_engine(self):
        t = (ROOT / 'requirements.lock').read_text()
        self.assertIn('formualizer==0.11.1', t)
        self.assertIn('--hash=sha256:', t)
```

- [ ] **Step 2: Run to verify failure**

Run: `python -m unittest tests.test_install_env -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'install_env'`

- [ ] **Step 3: Implement**

`requirements.in`:
```
formualizer==0.11.1
waitress==3.0.2
```
`requirements-dev.in`:
```
-r requirements.in
openpyxl==3.1.5
```

```python
# ops/lock_deps.py
"""Regenerate hashed locks (all platforms) and download Windows wheels for offline install."""
import json, hashlib, subprocess, sys
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
WHEELS = ROOT / 'vendor' / 'wheels'

def run(*a):
    subprocess.run(a, cwd=ROOT, check=True)

run('uv', 'pip', 'compile', 'requirements.in', '--universal', '--generate-hashes', '-o', 'requirements.lock')
run('uv', 'pip', 'compile', 'requirements-dev.in', '--universal', '--generate-hashes', '-o', 'requirements-dev.lock')
WHEELS.mkdir(parents=True, exist_ok=True)
for old in WHEELS.glob('*.whl'):
    old.unlink()
for pyver in ('3.10', '3.11', '3.12'):
    run(sys.executable, '-m', 'pip', 'download', '--no-deps', '--only-binary=:all:', '--platform', 'win_amd64',
        '--python-version', pyver, '--require-hashes', '-r', 'requirements.lock', '-d', str(WHEELS))
allow_path = ROOT / 'ops' / 'large_file_allowlist.json'
allow = {k: v for k, v in json.loads(allow_path.read_text()).items() if not k.startswith('vendor/wheels/')}
for w in sorted(WHEELS.glob('*.whl')):
    data = w.read_bytes()
    if len(data) > 5_000_000:
        allow[w.relative_to(ROOT).as_posix()] = hashlib.sha256(data).hexdigest()
allow_path.write_text(json.dumps(allow, indent=2, sort_keys=True) + '\n')
print('locked; wheels:', len(list(WHEELS.glob('*.whl'))))
```

```python
# ops/install_env.py
"""Create or refresh .venv offline from vendor/wheels; print the venv python path."""
import hashlib, os, subprocess, sys, venv
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent

def lock_sha256(lock):
    return hashlib.sha256(Path(lock).read_bytes()).hexdigest()

def needs_install(venv_dir, lock):
    stamp = Path(venv_dir) / '.lock-sha256'
    return not stamp.exists() or stamp.read_text().strip() != lock_sha256(lock)

def venv_python(venv_dir):
    return Path(venv_dir) / ('Scripts/python.exe' if os.name == 'nt' else 'bin/python')

def main():
    venv_dir, lock = ROOT / '.venv', ROOT / 'requirements.lock'
    if needs_install(venv_dir, lock):
        if not venv_python(venv_dir).exists():
            venv.EnvBuilder(with_pip=True).create(venv_dir)
        args = [str(venv_python(venv_dir)), '-m', 'pip', 'install', '--require-hashes', '-r', str(lock)]
        if os.name == 'nt':
            args[4:4] = ['--no-index', '--find-links', str(ROOT / 'vendor' / 'wheels')]
        subprocess.run(args, check=True)
        (venv_dir / '.lock-sha256').write_text(lock_sha256(lock))
    print(venv_python(venv_dir))

if __name__ == '__main__':
    main()
```

`run.bat`: before the existing line that launches `run.py`, insert:
```bat
for /f "delims=" %%P in ('"%PYTHON%" ops\install_env.py') do set "VENV_PY=%%P"
if not defined VENV_PY ( echo Dependency install failed. & exit /b 1 )
set "PYTHON=%VENV_PY%"
```
(`%PYTHON%` is the interpreter variable the existing `run.bat` already resolves; keep its detection logic unchanged.)
`run.sh`: replace its body with:
```sh
#!/bin/sh
set -eu
cd "$(dirname "$0")"
PY=$(python3 ops/install_env.py)
exec "$PY" run.py "$@"
```

- [ ] **Step 4: Generate locks and wheels, verify size rule**

Run: `python3 ops/lock_deps.py && git rm -q requirements.txt && python3 ops/content_guard.py --all; python -m unittest tests.test_install_env -v`
Expected: lock files contain `formualizer==0.11.1` and `--hash=sha256:` lines; `vendor/wheels/` holds 3 formualizer win_amd64 wheels (or 1 abi3 wheel) plus waitress; content guard exit 0; 2 tests PASS.

- [ ] **Step 5: Commit**

```bash
git add requirements.in requirements-dev.in requirements.lock requirements-dev.lock ops/lock_deps.py ops/install_env.py vendor/wheels ops/large_file_allowlist.json run.bat run.sh tests/test_install_env.py
git commit -m "build(p0): hashed locks, vendored Windows wheels, offline venv install"
```

---

### Task 4: Synthetic fixtures + typed comparison

**Files:**
- Create: `ops/fixtures.py`, `ops/typed.py`
- Test: `tests/test_fixtures.py`, `tests/test_typed.py`

**Interfaces:**
- Produces: `make_semantic(path: Path) -> dict` (writes small workbook; returns `{"inputs": ["Inputs!B2", ...], "formulas": {"Calc!B2": "<formula>", ...}}`), `make_scale(path: Path, rows: int, cols: int, lookup_rows: int) -> None`.
- Produces: `typed(v) -> tuple[str, object]` returning one of `('blank', None) ('empty', '') ('num', float) ('text', str) ('bool', bool) ('error', '#DIV/0!' etc.)`; `same(a, b, rel=1e-9) -> bool`.

- [ ] **Step 1: Write failing tests**

```python
# tests/test_typed.py
import sys, unittest
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / 'ops'))
from typed import typed, same

class Typed(unittest.TestCase):
    def test_typed_compare_distinguishes_blank_empty_zero_error(self):
        kinds = [typed(None)[0], typed('')[0], typed(0)[0], typed('#N/A')[0], typed(True)[0], typed('abc')[0]]
        self.assertEqual(kinds, ['blank', 'empty', 'num', 'error', 'bool', 'text'])
        self.assertFalse(same(None, ''))
        self.assertFalse(same('', 0))
        self.assertFalse(same(0, '#N/A'))
    def test_numeric_tolerance(self):
        self.assertTrue(same(1.0, 1.0 + 1e-12))
        self.assertFalse(same(1.0, 1.001))
    def test_bool_is_not_number(self):
        self.assertFalse(same(True, 1))
```

```python
# tests/test_fixtures.py
import sys, tempfile, unittest, zipfile
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / 'ops'))
import fixtures

class Fixtures(unittest.TestCase):
    def test_semantic_has_protected_inputs_and_modern_functions(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / 'semantic.xlsx'
            meta = fixtures.make_semantic(p)
            self.assertGreaterEqual(len(meta['inputs']), 5)
            joined = ' '.join(meta['formulas'].values())
            for fn in ('XLOOKUP', 'LET', 'FILTER', 'UNIQUE', 'SEQUENCE', 'SWITCH', 'IFS', 'NUMBERVALUE', 'IFERROR', 'SUMIFS'):
                self.assertIn(fn, joined)
            xml = zipfile.ZipFile(p).read('xl/worksheets/sheet1.xml').decode()
            self.assertIn('sheetProtection', xml)
    def test_scale_dimensions(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / 'scale.xlsx'
            fixtures.make_scale(p, rows=200, cols=30, lookup_rows=500)
            from openpyxl import load_workbook
            wb = load_workbook(p, read_only=True)
            self.assertEqual(wb['Grid'].max_row, 200)
            self.assertEqual(wb['Grid'].max_column, 30)
            self.assertEqual(wb['Lookup'].max_row, 500)
```

- [ ] **Step 2: Run to verify failure**

Run: `python -m unittest tests.test_typed tests.test_fixtures -v`
Expected: FAIL with `ModuleNotFoundError`

- [ ] **Step 3: Implement**

```python
# ops/typed.py
"""Excel-typed values: blank, empty string, number, text, bool and error are different things."""
import math
ERRORS = {'#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#SPILL!', '#CALC!', '#GETTING_DATA'}

def typed(v):
    if v is None:
        return ('blank', None)
    if isinstance(v, bool):
        return ('bool', v)
    if isinstance(v, (int, float)):
        return ('num', float(v))
    s = str(v)
    if s == '':
        return ('empty', '')
    if s in ERRORS:
        return ('error', s)
    return ('text', s)

def same(a, b, rel=1e-9):
    ta, tb = typed(a), typed(b)
    if ta[0] != tb[0]:
        return False
    if ta[0] == 'num':
        return math.isclose(ta[1], tb[1], rel_tol=rel, abs_tol=1e-12)
    return ta[1] == tb[1]
```

```python
# ops/fixtures.py
"""Deterministic synthetic workbooks. No company data. Used by CI and by the Univer gate."""
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Protection

def make_semantic(path):
    wb = Workbook()
    inp = wb.active; inp.title = 'Inputs'
    inp.protection.sheet = True
    rows = [('Cost', 10), ('Retail', 20), ('Blank', None), ('EmptyText', ''), ('Rate', '1.234,5')]
    inputs = []
    for i, (label, val) in enumerate(rows, start=2):
        inp.cell(i, 1, label)
        c = inp.cell(i, 2, val)
        c.protection = Protection(locked=False)
        inputs.append(f'Inputs!B{i}')
    lk = wb.create_sheet('Lookup')
    for i in range(1, 11):
        lk.cell(i, 1, f'K{i}'); lk.cell(i, 2, i * 1.5); lk.cell(i, 3, 'A' if i % 2 else 'B')
    calc = wb.create_sheet('Calc')
    formulas = {
        'B2': '=Inputs!B2*2',
        'B3': '=IFERROR(1-Inputs!B2/Inputs!B3,"")',
        'B4': '=IF(Inputs!B4="","blank","set")',
        'B5': '=IF(ISBLANK(Inputs!B4),1,0)',
        'B6': '=IFERROR(1/0,"err")',
        'B7': '=_xlfn.XLOOKUP("K3",Lookup!A1:A10,Lookup!B1:B10,"none")',
        'B8': '=_xlfn.LET(_xlpm.x,Inputs!B2,_xlpm.x+1)',
        'B9': '=_xlfn.SWITCH(Inputs!B2,10,"ten",20,"twenty","other")',
        'B10': '=_xlfn.IFS(Inputs!B2>15,"hi",TRUE,"lo")',
        'B11': '=_xlfn.NUMBERVALUE(Inputs!B6,",",".")',
        'B12': '=SUMIFS(Lookup!B1:B10,Lookup!C1:C10,"A")',
        'B13': '=_xlfn.SUM(_xlfn._xlws.FILTER(Lookup!B1:B10,Lookup!C1:C10="B"))',
        'B14': '=ROWS(_xlfn.UNIQUE(Lookup!C1:C10))',
        'B15': '=SUM(_xlfn.SEQUENCE(3))',
        'B16': '=VLOOKUP("K5",Lookup!A1:C10,2,FALSE)',
        'B17': '=SUBSTITUTE(TRIM("  a  b "),"a","x")',
    }
    for ref, f in formulas.items():
        calc[ref] = f
    wb.save(path)
    return {'inputs': inputs, 'formulas': {f'Calc!{k}': v for k, v in formulas.items()}}

def make_scale(path, rows=4088, cols=312, lookup_rows=100000):
    wb = Workbook(write_only=True)
    lk = wb.create_sheet('Lookup')
    for r in range(1, lookup_rows + 1):
        lk.append([f'K{r}', r * 0.5, f'G{r % 97}'] + [r % 13] * 14)
    grid = wb.create_sheet('Grid')
    for r in range(1, rows + 1):
        row = [f'K{(r * 7) % lookup_rows + 1}', r % 50 + 1]
        for c in range(3, cols + 1):
            row.append(f'=IFERROR(VLOOKUP($A{r},Lookup!$A:$C,2,FALSE)*$B{r}+{c},0)' if c % 3 else f'=SUBSTITUTE(TRIM($A{r}),"K","")')
        grid.append(row)
    wb.save(path)
```

- [ ] **Step 4: Run tests**

Run: `.venv/bin/pip install --require-hashes -r requirements-dev.lock && python -m unittest tests.test_typed tests.test_fixtures -v`
Expected: 5 tests PASS

- [ ] **Step 5: Commit**

```bash
git add ops/typed.py ops/fixtures.py tests/test_typed.py tests/test_fixtures.py
git commit -m "test(p0): synthetic semantic and scale fixtures, typed value comparison"
```

---

### Task 5: Real-workbook benchmark (office) with public/private split and watchdog

**Files:**
- Create: `ops/find_inputs.py`, `ops/bench_real.py`, `ops/bench_manifest.json`, `ops/office_p0.bat`, `.planning/excel-grid/results/.keep`
- Test: `tests/test_bench.py`

**Interfaces:**
- Consumes: `typed.same`, `fixtures.make_semantic` (tests).
- Produces: `find_inputs(xlsm: Path) -> list[str]` (unlocked cells on protected sheets, `"Sheet!A1"`, sorted, streaming XML).
- Produces: `bench_real.py --xlsm P --manifest M --public-dir D1 --private-dir D2 [--max-rss-mb N] [--max-seconds N]`; public JSON keys: `status, host, python, formualizer, xlsm_sha256, load_s, evaluate_all_s, rss_mb_after_load, rss_mb_after_eval, parity:{sheet:{checked,matched}}, edits:{narrow|median|worst:{fanout_direct, p50_ms, p95_ms, max_ms}}, aborted_reason`.

- [ ] **Step 1: Write failing tests**

```python
# tests/test_bench.py
import json, os, subprocess, sys, tempfile, unittest
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / 'ops'))
import fixtures, find_inputs

PY = sys.executable

class Bench(unittest.TestCase):
    def setUp(self):
        self.d = tempfile.TemporaryDirectory(); self.dir = Path(self.d.name)
        self.wb = self.dir / 'semantic.xlsx'; self.meta = fixtures.make_semantic(self.wb)
    def tearDown(self):
        self.d.cleanup()
    def test_find_inputs_lists_unlocked_cells_on_protected_sheet(self):
        self.assertEqual(find_inputs.find_inputs(self.wb), sorted(self.meta['inputs']))
    def run_bench(self, *extra):
        manifest = self.dir / 'm.json'
        manifest.write_text(json.dumps({'parity_sheets': ['Calc'], 'parity_sample': 50,
                                        'candidate_inputs': self.meta['inputs'], 'edit_values': [12, 30], 'repeats': 3}))
        r = subprocess.run([PY, str(ROOT / 'ops' / 'bench_real.py'), '--xlsm', str(self.wb), '--manifest', str(manifest),
                            '--public-dir', str(self.dir / 'pub'), '--private-dir', str(self.dir / 'priv'), *extra],
                           capture_output=True, text=True, timeout=300)
        pub = json.loads(next((self.dir / 'pub').glob('*.json')).read_text())
        return r, pub
    def test_public_report_has_aggregates_only(self):
        r, pub = self.run_bench()
        self.assertEqual(pub['status'], 'ok', r.stderr)
        self.assertIn('narrow', pub['edits'])
        text = json.dumps(pub)
        for leak in ('XLOOKUP', 'twenty', 'K3', 'Cost'):
            self.assertNotIn(leak, text)
        self.assertTrue(list((self.dir / 'priv').glob('*.json')))
    def test_watchdog_writes_partial_report(self):
        r, pub = self.run_bench('--max-rss-mb', '1')
        self.assertEqual(pub['status'], 'aborted')
        self.assertIn('rss', pub['aborted_reason'])
```

- [ ] **Step 2: Run to verify failure**

Run: `python -m unittest tests.test_bench -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'find_inputs'`

- [ ] **Step 3: Implement `ops/find_inputs.py`**

```python
# ops/find_inputs.py
"""List input cells: unlocked cells on protected sheets. Streams the XML; never loads the whole sheet."""
import re, sys, zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
RNS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'

def _unlocked_style_ids(z):
    root = ET.fromstring(z.read('xl/styles.xml'))
    ids = set()
    for i, xf in enumerate(root.find(f'{NS}cellXfs')):
        p = xf.find(f'{NS}protection')
        if p is not None and p.get('locked') in ('0', 'false'):
            ids.add(str(i))
    return ids

def _sheets(z):
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    target = {r.get('Id'): r.get('Target') for r in rels}
    for s in wb.find(f'{NS}sheets'):
        t = target[s.get(f'{RNS}id')].lstrip('/')
        yield s.get('name'), t if t.startswith('xl/') else 'xl/' + t

def find_inputs(xlsm):
    out = []
    with zipfile.ZipFile(xlsm) as z:
        unlocked = _unlocked_style_ids(z)
        for name, part in _sheets(z):
            protected, cells = False, []
            with z.open(part) as f:
                for _, el in ET.iterparse(f):
                    if el.tag == f'{NS}c':
                        if el.get('s') in unlocked and el.find(f'{NS}f') is None:
                            cells.append(el.get('r'))
                        el.clear()
                    elif el.tag == f'{NS}sheetProtection' and el.get('sheet') in ('1', 'true'):
                        protected = True
                    elif el.tag == f'{NS}row':
                        el.clear()
            if protected:
                out += [f'{name}!{c}' for c in cells]
    return sorted(out)

if __name__ == '__main__':
    for ref in find_inputs(Path(sys.argv[1])):
        print(ref)
```

Note for the implementer: openpyxl writes `sheetProtection` with `sheet="1"`. Excel may write it without `sheet` when all defaults; if `find_inputs` returns nothing on the real file, print the raw `sheetProtection` attributes into `results-private/` and adjust the test and the condition together.

- [ ] **Step 4: Implement `ops/bench_real.py`**

```python
# ops/bench_real.py
"""Benchmark Formualizer on a workbook. Public report = aggregates only; details stay private."""
import argparse, hashlib, json, os, platform, socket, statistics, subprocess, sys, threading, time
from datetime import datetime, timezone
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent))
from typed import same

def rss_mb():
    if os.name == 'nt':
        import ctypes, ctypes.wintypes as w
        class PMC(ctypes.Structure):
            _fields_ = [('cb', w.DWORD), ('PageFaultCount', w.DWORD), ('PeakWorkingSetSize', ctypes.c_size_t),
                        ('WorkingSetSize', ctypes.c_size_t), ('QuotaPeakPagedPoolUsage', ctypes.c_size_t),
                        ('QuotaPagedPoolUsage', ctypes.c_size_t), ('QuotaPeakNonPagedPoolUsage', ctypes.c_size_t),
                        ('QuotaNonPagedPoolUsage', ctypes.c_size_t), ('PagefileUsage', ctypes.c_size_t),
                        ('PeakPagefileUsage', ctypes.c_size_t)]
        c = PMC(); c.cb = ctypes.sizeof(PMC)
        ctypes.windll.psapi.GetProcessMemoryInfo(ctypes.windll.kernel32.GetCurrentProcess(), ctypes.byref(c), c.cb)
        return c.WorkingSetSize / 1e6
    if Path('/proc/self/statm').exists():
        return int(Path('/proc/self/statm').read_text().split()[1]) * os.sysconf('SC_PAGE_SIZE') / 1e6
    return int(subprocess.run(['ps', '-o', 'rss=', '-p', str(os.getpid())], capture_output=True, text=True).stdout) / 1e3

class Report:
    def __init__(self, a):
        self.a, self.pub, self.priv = a, {'status': 'running'}, {'mismatches': []}
        stamp = datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%SZ')
        host = hashlib.sha256(socket.gethostname().encode()).hexdigest()[:8]
        self.name = f'bench-{host}-{stamp}.json'
        self.pub.update(host=host, python=platform.python_version(), os=platform.platform(terse=True))
    def write(self):
        for d, data in ((self.a.public_dir, self.pub), (self.a.private_dir, self.priv)):
            Path(d).mkdir(parents=True, exist_ok=True)
            (Path(d) / self.name).write_text(json.dumps(data, indent=2, default=str))

def watchdog(rep, limit_mb, limit_s, t0):
    while rep.pub['status'] == 'running':
        m = rss_mb()
        rep.pub['rss_mb_peak'] = max(rep.pub.get('rss_mb_peak', 0), round(m, 1))
        reason = f'rss {m:.0f} MB > {limit_mb} MB' if m > limit_mb else (
            f'time > {limit_s} s' if time.time() - t0 > limit_s else None)
        if reason:
            rep.pub.update(status='aborted', aborted_reason=reason); rep.write(); os._exit(3)
        time.sleep(0.25)

def timed(fn):
    t = time.perf_counter(); out = fn(); return out, time.perf_counter() - t

def split(ref):
    sheet, cell = ref.rsplit('!', 1)
    col = ''.join(ch for ch in cell if ch.isalpha()); row = int(cell[len(col):])
    n = 0
    for ch in col.upper():
        n = n * 26 + ord(ch) - 64
    return sheet.strip("'"), row, n

def cached_values(xlsm, sheets, sample):
    from openpyxl import load_workbook
    f_wb = load_workbook(xlsm, read_only=True, data_only=False)
    v_wb = load_workbook(xlsm, read_only=True, data_only=True)
    out = {}
    for s in sheets:
        refs = []
        for row in f_wb[s].iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith('='):
                    refs.append((c.row, c.column))
            if len(refs) >= sample:
                break
        refs = refs[:sample]; want = set(refs)
        for row in v_wb[s].iter_rows(max_row=max(r for r, _ in refs) if refs else 0):
            for c in row:
                if (c.row, c.column) in want:
                    out[(s, c.row, c.column)] = c.value
    return out

def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument('--xlsm', required=True); ap.add_argument('--manifest', required=True)
    ap.add_argument('--public-dir', default='.planning/excel-grid/results')
    ap.add_argument('--private-dir', default='results-private')
    ap.add_argument('--max-rss-mb', type=float, default=12000); ap.add_argument('--max-seconds', type=float, default=3600)
    a = ap.parse_args(argv)
    import formualizer
    from importlib.metadata import version
    m = json.loads(Path(a.manifest).read_text())
    rep = Report(a); t0 = time.time()
    rep.pub.update(formualizer=version('formualizer'), xlsm_sha256=hashlib.sha256(Path(a.xlsm).read_bytes()).hexdigest())
    threading.Thread(target=watchdog, args=(rep, a.max_rss_mb, a.max_seconds, t0), daemon=True).start()
    expected = cached_values(a.xlsm, m['parity_sheets'], m['parity_sample'])
    wb, rep.pub['load_s'] = timed(lambda: formualizer.load_workbook(a.xlsm))
    rep.pub['rss_mb_after_load'] = round(rss_mb(), 1)
    _, rep.pub['evaluate_all_s'] = timed(wb.evaluate_all)
    rep.pub['rss_mb_after_eval'] = round(rss_mb(), 1)
    parity = {}
    for (s, r, c), exp in expected.items():
        got = wb.get_value(s, r, c); p = parity.setdefault(s, {'checked': 0, 'matched': 0})
        p['checked'] += 1
        if same(exp, got):
            p['matched'] += 1
        else:
            rep.priv['mismatches'].append({'sheet': s, 'row': r, 'col': c, 'formula': wb.get_formula(s, r, c),
                                           'expected': exp, 'got': got})
    rep.pub['parity'] = parity
    fan = sorted(((len(wb.dependents(ref).dependents), ref) for ref in m['candidate_inputs']))
    picks = {'narrow': fan[0], 'median': fan[len(fan) // 2], 'worst': fan[-1]} if fan else {}
    rep.priv['picked_inputs'] = {k: v[1] for k, v in picks.items()}
    edits = {}
    for label, (fanout, ref) in picks.items():
        s, r, c = split(ref); times = []
        for i in range(m['repeats']):
            v = m['edit_values'][i % len(m['edit_values'])]
            wb.set_value(s, r, c, v)
            _, dt = timed(wb.evaluate_all); times.append(dt * 1000)
        times.sort()
        edits[label] = {'fanout_direct': fanout, 'p50_ms': round(statistics.median(times), 1),
                        'p95_ms': round(times[min(len(times) - 1, int(0.95 * len(times)))], 1), 'max_ms': round(times[-1], 1)}
    rep.pub['edits'] = edits
    rep.pub['total_s'] = round(time.time() - t0, 1)
    rep.pub['status'] = 'ok'; rep.write()
    print(json.dumps(rep.pub, indent=2))

if __name__ == '__main__':
    main()
```

`ops/bench_manifest.json` (real workbook; `candidate_inputs` is filled on the office machine by `office_p0.bat`):
```json
{
  "parity_sheets": ["BU RECAP", "Sales by Mkting Block", "Projections", "Fiscal Week Recap", "Dept OTD LY", "Coupon"],
  "parity_sample": 2000,
  "candidate_inputs": [],
  "edit_values": [10, 20, 0, 15.5, 7],
  "repeats": 20
}
```

`ops/office_p0.bat`:
```bat
@echo off
setlocal
cd /d "%~dp0\.."
for /f "delims=" %%P in ('py -3 ops\install_env.py') do set "VENV_PY=%%P"
if not defined VENV_PY ( echo install failed & exit /b 1 )
"%VENV_PY%" -m pip install --no-index --find-links vendor\wheels --require-hashes -r requirements-dev.lock || echo openpyxl not vendored: see note
set "XLSM=Price Planning Template Spring 2027.xlsm"
"%VENV_PY%" ops\find_inputs.py "%XLSM%" > results-private\inputs.txt
"%VENV_PY%" -c "import json;m=json.load(open('ops/bench_manifest.json'));l=[x.strip() for x in open('results-private/inputs.txt',encoding='utf-8') if x.strip()];m['candidate_inputs']=l[::max(1,len(l)//50)][:50];json.dump(m,open('results-private/manifest.json','w'),indent=2)"
"%VENV_PY%" ops\bench_real.py --xlsm "%XLSM%" --manifest results-private\manifest.json
"%VENV_PY%" ops\bench_real.py --xlsm "%XLSM%" --manifest results-private\manifest.json
echo Done. Commit .planning\excel-grid\results\*.json on a branch and push.
```
Note: `openpyxl` is a dev dependency; `lock_deps.py` must also download `requirements-dev.lock` wheels for win_amd64 (add a second `pip download` loop over `requirements-dev.lock`) so this runs offline. `mkdir results-private` before first use (add `if not exist results-private mkdir results-private`).

- [ ] **Step 5: Run tests**

Run: `.venv/bin/python -m unittest tests.test_bench -v`
Expected: 3 tests PASS. If `wb.dependents(ref)` raises for a constant input cell, record the exact error and fall back to `len(wb.dependents(ref).dependents)` = 0 handling via try/except `formualizer.InspectionError` that sets fanout to -1; update the test to accept -1 only in that case.

- [ ] **Step 6: Small-scale run on this Mac (guarded)**

Run: `.venv/bin/python ops/bench_real.py --xlsm "Price Planning Template Spring 2027.xlsm" --manifest ops/bench_manifest.json --max-rss-mb 2500 --max-seconds 900` only if `df -k ~` shows > 5,000,000 KB free and swap used < 2 GB. Expected on this 8 GB Mac: `status: aborted` with an rss reason is an acceptable, informative result. Do not retry more than once.

- [ ] **Step 7: Commit**

```bash
git add ops/find_inputs.py ops/bench_real.py ops/bench_manifest.json ops/office_p0.bat .planning/excel-grid/results/.keep tests/test_bench.py
git commit -m "feat(p0): real-workbook benchmark with private detail split and watchdog"
```

---

### Task 6: CI (ubuntu tests + security, windows offline install)

**Files:**
- Create: `.github/workflows/ci.yml`, `.semgrepignore`
- Test: the workflow itself (green run on the PR)

- [ ] **Step 1: Write the workflow**

```yaml
# .github/workflows/ci.yml
name: ci
on: [pull_request, push]
permissions: { contents: read }
jobs:
  test:
    runs-on: ubuntu-latest
    strategy: { matrix: { python: ['3.10', '3.12'] } }
    steps:
      - uses: actions/checkout@v4
        with: { fetch-depth: 0 }
      - uses: actions/setup-python@v5
        with: { python-version: '${{ matrix.python }}' }
      - run: python -m pip install --require-hashes -r requirements-dev.lock
      - run: python ops/content_guard.py --all
      - run: python -m unittest discover -s tests -v
  security:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
        with: { fetch-depth: 0 }
      - uses: actions/setup-python@v5
        with: { python-version: '3.12' }
      - run: python -m pip install semgrep bandit pip-audit
      - run: semgrep scan --error --config p/python --config p/owasp-top-ten --config p/secrets
      - run: bandit -r app ops -ll
      - run: pip-audit --require-hashes -r requirements.lock
      - uses: gitleaks/gitleaks-action@v2
        env: { GITHUB_TOKEN: '${{ secrets.GITHUB_TOKEN }}' }
  windows-offline:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with: { python-version: '3.10' }
      - name: offline install from vendored wheels (no index)
        run: python ops/install_env.py
      - run: .venv\Scripts\python -c "import formualizer, waitress; print('ok')"
```

`.semgrepignore`:
```
vendor/
proto/univer-gate/node_modules/
```

- [ ] **Step 2: Push the branch and open the PR (Sir runs the push in Terminal.app; the gate asks `y`)**

Run (Sir): `git push -u origin feat/p0-rails` then `gh pr create --fill`
Expected: all three jobs green. If gitleaks-action needs a licence for private org repos, replace it with `run: curl -sSfL https://github.com/gitleaks/gitleaks/releases/download/v8.30.1/gitleaks_8.30.1_linux_x64.tar.gz | tar xz && ./gitleaks detect --no-banner --redact`.

- [ ] **Step 3: Commit**

```bash
git add .github/workflows/ci.yml .semgrepignore
git commit -m "ci(p0): tests, content guard, security scans, windows offline install"
```

---

### Task 7: Univer gate prototype

**Files:**
- Create: `proto/univer-gate/package.json`, `proto/univer-gate/index.html`, `proto/univer-gate/src/main.js`, `proto/univer-gate/src/data.js`, `proto/univer-gate/gate.mjs`, `proto/univer-gate/README.md`
- Output: `.planning/excel-grid/results/univer-gate-<date>.json`

**Interfaces:**
- Produces: `window.gate` in the page: `{ loadWindow(sheetId, firstRow, rowCount), heapMB(), formulaCellCount(), pasteTsv(ref, tsv), undo() }` used by `gate.mjs`.
- Gate JSON: `{univer_version, sheets:{grid_4088x312:{heap_mb_full, heap_mb_windowed, scroll_fps, formula_cells_after_ops}, rows_100000x17:{...}}, pass}`.

- [ ] **Step 1: Scaffold with pinned versions**

Before writing code, query context7 (`resolve-library-id` "univer", then `query-docs` "presets createUniver sheets core preset disable formula calculation setRangeValues") and pin the exact versions it documents in `package.json`. Starting point:

```json
{
  "name": "univer-gate",
  "private": true,
  "type": "module",
  "scripts": { "dev": "vite", "build": "vite build", "gate": "node gate.mjs" },
  "dependencies": { "@univerjs/presets": "1.0.4" },
  "devDependencies": { "vite": "6.3.5", "playwright": "1.55.0" }
}
```

- [ ] **Step 2: Data generator (values + styles only, NO formulas sent to the client)**

```js
// src/data.js - synthetic snapshot blocks; the server will send exactly this shape later
export function block(sheet, firstRow, rowCount, cols) {
  const cellData = {};
  for (let r = firstRow; r < firstRow + rowCount; r++) {
    const row = {};
    for (let c = 0; c < cols; c++) {
      const input = c === 1;
      row[c] = { v: c === 0 ? `K${r + 1}` : (r * 31 + c * 7) % 1000 / 10, s: input ? 'in' : 'calc' };
    }
    cellData[r] = row;
  }
  return cellData;
}
```

- [ ] **Step 3: Page wiring**

```js
// src/main.js
import { createUniver, LocaleType, mergeLocales } from '@univerjs/presets';
import { UniverSheetsCorePreset } from '@univerjs/presets/preset-sheets-core';
import enUS from '@univerjs/presets/preset-sheets-core/locales/en-US';
import '@univerjs/presets/lib/styles/preset-sheets-core.css';
import { block } from './data.js';

const SHEETS = { grid: { rows: 4088, cols: 312 }, tall: { rows: 100000, cols: 17 } };
const params = new URLSearchParams(location.search);
const mode = params.get('mode') || 'windowed';  // 'full' | 'windowed'
const { univerAPI } = createUniver({
  locale: LocaleType.EN_US, locales: { [LocaleType.EN_US]: mergeLocales(enUS) },
  presets: [UniverSheetsCorePreset({ container: 'app', formula: { initialFormulaComputing: 'NO_CALCULATION' } })],
});
const sheets = {};
for (const [id, d] of Object.entries(SHEETS)) {
  sheets[id] = { id, name: id, rowCount: d.rows, columnCount: d.cols,
    cellData: mode === 'full' ? block(id, 0, d.rows, d.cols) : block(id, 0, 200, d.cols), freeze: { xSplit: 1, ySplit: 1, startRow: 1, startColumn: 1 } };
}
univerAPI.createWorkbook({ id: 'wb', sheetOrder: Object.keys(SHEETS), sheets,
  styles: { in: { bg: { rgb: '#DDEBF7' }, cl: { rgb: '#1F4E79' } }, calc: { cl: { rgb: '#000000' } } } });

window.gate = {
  async heapMB() { return performance.memory ? performance.memory.usedJSHeapSize / 1e6 : -1; },
  loadWindow(sheetId, firstRow, rowCount) {
    const ws = univerAPI.getActiveWorkbook().getSheetBySheetId(sheetId);
    const data = block(sheetId, firstRow, rowCount, SHEETS[sheetId].cols);
    const values = Object.values(data).map(row => Object.values(row).map(c => c.v));
    ws.getRange(firstRow, 0, rowCount, SHEETS[sheetId].cols).setValues(values);
  },
  formulaCellCount() {
    const snap = univerAPI.getActiveWorkbook().save();
    let n = 0;
    for (const s of Object.values(snap.sheets)) for (const row of Object.values(s.cellData || {})) for (const c of Object.values(row)) if (c && (c.f || c.si)) n++;
    return n;
  },
};
```

If any import path or method name differs in the pinned version, use the name context7 documents for that version and note the change in `README.md`. Do not change the gate measurements.

- [ ] **Step 4: Gate runner**

```js
// gate.mjs - launches Vite preview + Chromium, measures, writes results JSON
import { chromium } from 'playwright';
import { spawn } from 'node:child_process';
import { writeFileSync, mkdirSync } from 'node:fs';
const srv = spawn('npx', ['vite', 'preview', '--port', '4310', '--strictPort'], { stdio: 'inherit' });
await new Promise(r => setTimeout(r, 3000));
const browser = await chromium.launch({ args: ['--enable-precise-memory-info', '--js-flags=--expose-gc'] });
import { readFileSync } from 'node:fs';
const pkg = JSON.parse(readFileSync('package.json', 'utf8'));
const out = { univer_version: pkg.dependencies['@univerjs/presets'], sheets: {} };
for (const mode of ['windowed', 'full']) {
  const page = await browser.newPage();
  await page.goto(`http://127.0.0.1:4310/?mode=${mode}`);
  await page.waitForFunction(() => window.gate, null, { timeout: 120000 });
  await page.evaluate(() => window.gc && window.gc());
  const heap = await page.evaluate(() => window.gate.heapMB());
  await page.keyboard.press('Control+Home');
  for (let i = 0; i < 50; i++) await page.mouse.wheel(0, 2000);
  await page.keyboard.type('=SUM(1,2)'); await page.keyboard.press('Enter');
  await page.keyboard.press('Control+Z');
  const formulas = await page.evaluate(() => window.gate.formulaCellCount());
  out.sheets[mode] = { heap_mb: Math.round(heap), formula_cells_after_ops: formulas };
  await page.close();
}
out.pass = out.sheets.windowed.heap_mb < 1500 && out.sheets.windowed.formula_cells_after_ops === 0;
mkdirSync('../../.planning/excel-grid/results', { recursive: true });
writeFileSync(`../../.planning/excel-grid/results/univer-gate-${new Date().toISOString().slice(0, 10)}.json`, JSON.stringify(out, null, 2));
console.log(JSON.stringify(out, null, 2));
await browser.close(); srv.kill();
```

Note: `formula_cells_after_ops` must be 0 because the app will reject formula text in input cells; the prototype proves Univer does not create formula cells on its own once we intercept the edit. If typing `=SUM(1,2)` creates a formula cell, add a `beforeCommand` hook (`univerAPI.addEvent(univerAPI.Event.BeforeSheetEditEnd, ...)` or the version's documented equivalent) that cancels edits starting with `=`, `+`, `-`, `@`, then re-run. That hook is the exact behavior P3 will use.

- [ ] **Step 5: Run the gate on this Mac (memory-guarded) and on CI is NOT required**

Run: `cd proto/univer-gate && npm install && npx playwright install chromium && npm run build && npm run gate`
Only when free disk > 5 GB. Expected: JSON with `pass: true`. If `pass: false`, keep the JSON (it is the evidence), and write the failing measurement into `.planning/STATE.md` as the P0 blocker for Sir; do not tune thresholds.

- [ ] **Step 6: Commit (no node_modules, no dist)**

```bash
git add proto/univer-gate/package.json proto/univer-gate/package-lock.json proto/univer-gate/index.html proto/univer-gate/src proto/univer-gate/gate.mjs proto/univer-gate/README.md .planning/excel-grid/results/univer-gate-*.json
git commit -m "test(p0): Univer gate prototype with heap and no-client-formula checks"
```

---

### Task 8: P0 handoff

**Files:**
- Modify: `.planning/STATE.md`, `README.md` (top section: point to AGENTS.md and docs/DEPLOY.md; mark the old LAN-demo text as "validation build, being replaced by v2")

- [ ] **Step 0 (Sir, office):** after `ops\office_p0.bat`, start `run.bat`, then open the printed `http://<server-ip>:<port>` from a SECOND office PC and sign in. Record "LAN second machine: pass/fail + port" in `.planning/excel-grid/results/office-lan-<date>.md` (no IPs or names, just pass/fail and port). This is the spec's P0 LAN exit check.
- [ ] **Step 1:** Update `.planning/STATE.md` `## Last handoff` with branch, last commit hash, tests run and their result lines, gate JSON paths. `## Next task`: "Sir runs ops\office_p0.bat on the office server, pushes results; then PLAN-P1".
- [ ] **Step 2:** Run `python -m unittest discover -s tests -v && python ops/content_guard.py --all`. Expected: all PASS, guard exit 0.
- [ ] **Step 3:** Commit `docs(p0): handoff state` and open the PR.
