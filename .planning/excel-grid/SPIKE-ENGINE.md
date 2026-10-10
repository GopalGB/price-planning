# SPIKE-ENGINE - Formualizer 0.11.1 vs IronCalc on the Spring 2027 workbook

Status: **PARTIAL. Load measured. Full evaluation, parity compare, edit latency and IronCalc were NOT completed.**
The host ran out of safe headroom twice during full-workbook evaluation and the watchdog killed the benchmark each time.
Nothing in this report is extrapolated. Anything not measured is marked NOT MEASURED with the reason.

Workspace: `/private/tmp/claude-501/-Users-gopalmacbook-Desktop-Max-HQ-price-planning/43ab80ca-d82c-4da5-9f50-8c24b8619935/scratchpad/spike`
Source (never modified): `price-planning/Price Planning Template Spring 2027.xlsm` (53,146,147 bytes, 22 sheets, 1,894,479 formulas per brief).
Copies used: `wb.xlsm` and `wb.xlsx` in the workspace, byte-identical (`cmp` passed).

## 1. Versions

| Package | Version | Note |
|---|---|---|
| formualizer | 0.11.1 | pinned, installed as requested |
| ironcalc | 0.8.3 | unpinned in the brief, latest installed |
| openpyxl | 3.1.5 | used only for cached-value extraction |
| Python | 3.11.4 | `/Library/Frameworks/Python.framework/Versions/3.11/bin/python3`, venv in workspace |
| Host | macOS Darwin 27.0.0, 8 GB RAM, 10 GB swap pool (dynamic, 7-10 GB reported) | |

`formualizer.__version__` does not exist; the version above is from `pip list`.

## 2. Exact commands

```
python3 -m venv WORKDIR/venv
WORKDIR/venv/bin/pip install formualizer==0.11.1 ironcalc openpyxl
cp "<SRC>" WORKDIR/wb.xlsm && cp "<SRC>" WORKDIR/wb.xlsx && cmp wb.xlsm wb.xlsx
WORKDIR/watchdog.sh WORKDIR            # background; kills bench_(formualizer|ironcalc).py at free disk < 2,500,000 KB or RSS > 3,500,000 KB; polls every 2 s
timeout 1200 WORKDIR/venv/bin/python bench_formualizer.py extract
timeout 1200 WORKDIR/venv/bin/python bench_formualizer.py run ondemand   # killed by watchdog, see section 7
timeout 1200 WORKDIR/venv/bin/python bench_formualizer.py run           # first attempt, the pre-"mode" version, killed by watchdog
```

Scripts: `bench_formualizer.py` (current), `bench_formualizer.v1.py` (version before the ondemand/full split), `probe_one_cell.py` (diagnostic, killed during load).

## 3. Formualizer API used (only what was seen in `help()` / `dir()` / a toy run)

- `formualizer.load_workbook(path)` - eager load, calamine backend.
- `Workbook.sheet_names` (property), `Workbook.get_formula(sheet,row,col)` (returns `'=...'` for formulas, `None` for constants), `Workbook.get_value(sheet,row,col)`, `Workbook.evaluate_cell(sheet,row,col)`, `Workbook.set_value(sheet,row,col,value)`, `Workbook.evaluate_all()` (returns `None`; no formula count exposed), `Workbook.dependents("Sheet!A1")` returns `DependentsReport` with `.dependents` list of `Dependent(cell=..., via=...)`. A tuple address raises `TypeError: 'tuple' object is not an instance of 'str'`.
- Rows/columns 1-based.

## 4. Raw JSON outputs

### 4a. Load - `formualizer_bench_ondemand.json` (the step that completed)

```json
{
 "mode": "ondemand",
 "formualizer_version": "unknown (no __version__ attr; pip 0.11.1)",
 "python": "3.11.4",
 "steps": {
  "load": {"loaded_file": "wb.xlsm", "load_seconds": 4.01, "errors": [],
           "peak_rss_mb_after_load": 152.2, "rss_now_mb": 138.9,
           "sheet_names": ["Config and Lookup", "TESTING NOTES", "Acronyms Legend", "Sales by Mkting Block", "BU RECAP", "Store vs Flagship", "TY vs LY (Not Shifted)", "Top 10 Program", "LY Data", "Dept OTD LY", "Fiscal Week Recap", "Projections", "MFP Plan", "Error Report", "Rotation", "Calendar", "Dept CMU% LY", "Daily Sales - Open Stores", "P2 P5 SM History", "Coupon", "Pricing Format", "Dept Rollups"]},
  "first_evaluate_all": "NOT RUN in ondemand mode (see run-full attempt); cells evaluated on demand"
 }
}
```

Earlier attempt (`run.stdout.log`, first version of the script, same load): `load_seconds` 5.8, `peak_rss_mb_after_load` 103.3, `loaded_file` "wb.xlsm", no load errors. `.xlsm` loads directly; the `.xlsx` fallback was never needed.

Peak RSS note: `ru_maxrss` is only recorded after load (103 to 152 MB across the two runs; the two numbers differ because the runs differ, not because of a measurement error). **Peak RSS during evaluation: NOT MEASURED** because the process was killed before the step reported.

### 4b. openpyxl cached-value extraction - `expected.json` (completed, exit 0)

```
counts: BU RECAP 2000 | Sales by Mkting Block 2000 | Projections 2000 | Fiscal Week Recap 2000 | Dept OTD LY 2000 | Coupon 54
skipped: {}   extract_total_s: 0.8   peak_rss_mb (openpyxl process): 44.0
```

Coupon has only 54 formula cells in the whole sheet, so it is not capped at 2,000. The first 2,000 formula cells per sheet are in row-major order, so they cluster in the top rows.

## 5. Parity table

**NOT MEASURED.** The 12,054 expected cached values were extracted (table above), but no Formualizer value was computed for any of them, so there is no match count. Reason: every attempt to compute values (full `evaluate_all()` and on-demand `evaluate_cell()`) was killed by the host-protection watchdog (section 7).

| Sheet | Expected cells extracted | Formualizer match | Mismatches |
|---|---|---|---|
| BU RECAP | 2000 | NOT MEASURED | NOT MEASURED |
| Sales by Mkting Block | 2000 | NOT MEASURED | NOT MEASURED |
| Projections | 2000 | NOT MEASURED | NOT MEASURED |
| Fiscal Week Recap | 2000 | NOT MEASURED | NOT MEASURED |
| Dept OTD LY | 2000 | NOT MEASURED | NOT MEASURED |
| Coupon | 54 | NOT MEASURED | NOT MEASURED |

Cached-value type mix (from the extraction, for later comparison): BU RECAP int 714, None 607, str 505, datetime 104, float 52, time 18; Sales by Mkting Block int 1371, str 428, datetime 182, None 19; Projections int 1917, None 54, str 29; Fiscal Week Recap int 1672, str 247, None 81; Dept OTD LY int 1392, None 521, time 87; Coupon float 36, int 9, str 9. Datetime and time values are stored as Excel serials. The comparator is written (rel 1e-9, empty == '' == None, error strings compared as text), but it has not produced a result.

## 6. Edit latency (BU RECAP, rows 22-40, cols C..Z)

**NOT MEASURED.** The candidate scan and the three-cell selection were not reached in a completed run. The candidate logic counts constant numeric or blank cells in that block that have at least one direct dependent (via `dependents("'BU RECAP'!C22")`-style addresses), and takes the first three. No candidate list was printed.

## 7. Failures and watchdog events (exact text / log lines)

### 7a. Build-time failures (fixed, not engine problems)

1. `extract` run 1: `AttributeError: 'EmptyCell' object has no attribute 'row'` (openpyxl read-only cells). Fixed by using `enumerate()` indices instead of `cell.row` / `cell.column`.
2. `extract` run 2: `TypeError: Object of type time is not JSON serializable`. Fixed by converting `datetime.time` / `timedelta` to Excel day fractions.

### 7b. Full evaluation and on-demand evaluation runs (the blocking problem)

Watchdog log, verbatim (`WORKDIR/watchdog.log`):

```
started 2026-10-10 17:37:24
2026-10-10 17:40:04 KILL pid=11203 reason=free_disk_kb=1798440<2500000
2026-10-10 17:40:04 KILL pid=11206 reason=free_disk_kb=1798440<2500000
2026-10-10 17:40:04 KILL pid=11208 reason=free_disk_kb=1798440<2500000
2026-10-10 17:40:04 KILL pid=11209 reason=free_disk_kb=1798440<2500000
2026-10-10 17:42:01 KILL pid=12656 reason=free_disk_kb=2455360<2500000
2026-10-10 17:42:01 KILL pid=12659 reason=free_disk_kb=2455360<2500000
2026-10-10 17:42:01 KILL pid=12660 reason=free_disk_kb=2455360<2500000
2026-10-10 17:42:01 KILL pid=12661 reason=free_disk_kb=2455360<2500000
2026-10-10 17:45 watchdog stopped by spike at end of run; no bench process alive
```

Probe kill (`probe.log`): `17:42:23 PROBE KILL free_kb=1568996` (stricter 3.2 GB guard, `probe_one_cell.py`; the process was still in load, printed no output, exit 137).

What happened, from measurements:

- Attempt 1 (`run`, version 1: load, then `evaluate_all()`): load succeeded in 5.8 s with 103 MB peak RSS. The process was killed at 17:40:04 when free disk fell to 1,798,440 KB. The watchdog killed the `bash -c` wrapper, `timeout`, and the python process (four PIDs, one per matching process). No Python traceback was written; stderr was empty.
- Attempt 2 (`run ondemand`): load succeeded in 4.01 s with 152 MB peak RSS. The parity step (first `evaluate_cell` on BU RECAP) was in progress when the watchdog killed it at 17:42:01 with free disk at 2,455,360 KB. stderr was empty.
- Attempt 3 (`probe_one_cell.py`, one cell): killed at 17:42:23 at 1,568,996 KB free, before the load print. Not attributable on its own (see section 8).

Host state during the runs (`df -k`, `vm.swapusage`), free disk and swap used:

| Time | Free disk KB | Swap used MB | Context |
|---|---|---|---|
| 17:37 | ~6,345,560 (start) | n/a | before runs |
| 17:40:04 | 1,798,440 | n/a | attempt 1 killed |
| 17:40:27 | 1,698,848 | 6,926 | after attempt 1 exit, swap draining |
| 17:40:47 | 4,773,760 | 6,377 | recovered |
| 17:42:01 | 2,455,360 | n/a | attempt 2 killed |
| 17:42:28 | 1,464,052 | 6,302 | probe killed; no bench process alive |
| 17:43:38 | 3,602,232 | 6,038 | recovering |
| 17:45:16 | 3,138,728 | 5,934 | still falling about 4.6 MB/s with no bench process alive |

Interpretation, kept to what is measured: the disk floor went below the guard during each run, and free disk was still falling at 17:45 with no benchmark process running. The machine was swapping heavily before and during the spike (Chrome renderers, several `claude` processes, omniroute and ego lite are resident). Formualizer's own memory during evaluation is NOT MEASURED, so I cannot state that the engine caused the swap growth.

### 7c. Guard weakness found (affects any retry)

- The watchdog's RSS check (`ps -o rss`) could not have seen the growth: swapped-out pages leave RSS. The only effective guard was the free-disk check, and at a 2 s poll it reacted late (first kill at 1.80 GB free, below the 2.5 GB floor).
- Recommendation for any retry: guard on `vm.swapusage` used plus free disk, poll at 1 s or less, and set the floor well above 2.5 GB (3.5 GB or more) because the observed drop between polls was up to about 0.7 GB.

## 8. Not done

- Full `evaluate_all()` timing and peak RSS: NOT MEASURED (killed; see 7b).
- Parity (step 5): NOT MEASURED (no computed values).
- Edit latency (step 6): NOT MEASURED.
- IronCalc (step 7): NOT RUN. The host did not have the swap/disk headroom to run a second engine safely while the machine was still recovering, and the first engine had already failed twice. IronCalc 0.8.3 is installed; its API was not explored.
- Probe attribution: the single-cell probe died during load, while the host was already under swap pressure. It does not show that a single-cell evaluation is expensive. A clean single-cell evaluation measurement is still needed on a quiet machine.

## 9. What would make the spike complete

1. Quit or pause the heavy apps (Chrome with many tabs, extra `claude` sessions, ego lite) and confirm `vm.swapusage` used is under 2 GB and free disk over 5 GB before starting.
2. Run `run ondemand` alone with a 1 s watchdog and a 3.5 GB floor. It answers whether single-cell evaluation is affordable.
3. Only if (2) is cheap, attempt `run full` (evaluate_all) once, then parity, then edits.
4. Run the IronCalc load and evaluate under the same guard, one process at a time.

## 10. Summary of what was measured

- Formualizer 0.11.1 installs and loads the 53 MB `.xlsm` directly: 4.0 to 5.8 s, 103 to 152 MB peak RSS after load.
- openpyxl extracted 12,054 cached formula values (0.8 s, 44 MB) for the six target sheets.
- Full recalculation, parity, edit latency and IronCalc: NOT MEASURED. The host ran out of disk and swap headroom during evaluation, the watchdog stopped each attempt, and I stopped rather than push the machine further.
