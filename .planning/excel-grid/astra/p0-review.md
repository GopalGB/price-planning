Reviewed the requested diff, SPEC.md and PLAN-P0.md. **10 defects found**, with [HIGH] confidence from source review and synthetic probes.

1. **HIGH: Renamed workbooks bypass the guard.** [content_guard.py:18](</Users/gopalmacbook/Desktop/Max HQ/price-planning/ops/content_guard.py:18>)  
   A workbook under 5 MB renamed to `.txt` passes, including one containing denylisted strings. ZIP compression and the NUL-byte check defeat scanning. **Fix:** inspect ZIP members and Office/OLE/PDF/database signatures independently of extensions; restrict binary exceptions to hash-verified vendor files. Add a real ZIP-based regression.

2. **HIGH: Encoding defeats denylist matching.** [content_guard.py:27](</Users/gopalmacbook/Desktop/Max HQ/price-planning/ops/content_guard.py:27>)  
   UTF-16 text containing a forbidden phrase passes because NUL bytes disable scanning. A UTF-8 BOM in `.content-denylist` also prevents its first entry matching. Both reproduced. **Fix:** recognize BOM-marked UTF-16/32, read the denylist with `utf-8-sig`, normalize entries, and reject unsupported encodings instead of silently ignoring bytes.

3. **HIGH: File-type changes escape staged checks.** [content_guard.py:47](</Users/gopalmacbook/Desktop/Max HQ/price-planning/ops/content_guard.py:47>)  
   Replacing a tracked benign symlink with a regular confidential file produces Git status `T`, omitted by `ACMR`. Git’s filter behavior was reproduced. **Fix:** include `T`, or inspect every non-deleted staged entry and its blob/mode. Add a symlink-to-file regression.

4. **HIGH: Private paths and logs are not prohibited.** [content_guard.py:14](</Users/gopalmacbook/Desktop/Max HQ/price-planning/ops/content_guard.py:14>)  
   Force-adding `results-private/bench.json` containing cell values passes; `DEBUG.LOG` also passes. `.gitignore` is the only protection. **Fix:** explicitly reject normalized private/runtime paths and log/database/VBA file types, including case variants, regardless of ignore status.

5. **HIGH: Private reports can overwrite public reports.** [bench_real.py:37](</Users/gopalmacbook/Desktop/Max HQ/price-planning/ops/bench_real.py:37>)  
   When output directories are identical or resolve through symlinks to the same location, the private write overwrites the public filename with mismatch details. Reproduced in memory. **Fix:** resolve destinations before processing; reject equality and private destinations inside the public tree.

6. **MEDIUM: Correct engine errors become parity mismatches.** [typed.py:13](</Users/gopalmacbook/Desktop/Max HQ/price-planning/ops/typed.py:13>)  
   Formualizer 0.11.1 returns errors such as `{"type":"Error","kind":"Div"}`; openpyxl returns `"#DIV/0!"`. The comparator classifies the dictionary as text. **Fix:** normalize engine error kinds to canonical error tags. Test actual engine return shapes; current tests supply strings only and miss this failure.

7. **MEDIUM: Ignored protection formatting creates false inputs.** [find_inputs.py:12](</Users/gopalmacbook/Desktop/Max HQ/price-planning/ops/find_inputs.py:12>)  
   An XF with `applyProtection="0"` and `locked="0"` is incorrectly classified as unlocked. Microsoft specifies that [`applyProtection` controls whether that formatting applies](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellformat.applyprotection?view=openxml-3.0.1). **Fix:** resolve effective protection using `cellXfs`, `xfId` and `cellStyleXfs`; do not apply explicitly disabled protection. Add handcrafted XML cases.

8. **MEDIUM: An incomplete `.venv` cannot recover.** [install_env.py:19](</Users/gopalmacbook/Desktop/Max HQ/price-planning/ops/install_env.py:19>)  
   A matching stamp with a missing interpreter prints a nonexistent executable. An existing interpreter without pip fails repeatedly because creation/bootstrap is skipped. Both branches reproduced with in-memory filesystem simulation. **Fix:** validate interpreter and pip before accepting the stamp; recreate an unusable environment or bootstrap `ensurepip`, then install and stamp only after success. Current tests exercise stamp hashes, not installation.

9. **MEDIUM: Hooks reject a supported launcher-only Windows setup.** [.githooks/pre-commit:3](</Users/gopalmacbook/Desktop/Max HQ/price-planning/.githooks/pre-commit:3>)  
   `run.bat` supports `py -3`, but the hook searches only `python3` and `python`. With only `py` available, every commit fails before scanning. Reproduced with shell command discovery. **Fix:** support a validated venv interpreter and a separate `py -3` invocation; test launcher-only and Max-HQ-absent environments together.

10. **LOW: Apostrophes corrupt input fanout ranking.** [find_inputs.py:43](</Users/gopalmacbook/Desktop/Max HQ/price-planning/ops/find_inputs.py:43>)  
    `O'Brien!A1` causes the pinned engine’s address parser to fail; the benchmark silently records fanout `-1`. **Fix:** emit `'O''Brien'!A1` and unescape doubled apostrophes in `bench_real.split()`.

No watchdog defect is claimed: the Mac evaluation probe allowed the monitoring thread to run. Native Windows watchdog behavior remains unverified. Shared and inline strings worked.

```text
STATUS:  DONE_WITH_CONCERNS (review complete)
CHANGED: none
TESTS:   13 scoped unittest tests passed; synthetic probes reproduced defects.
         python3 -B ops/content_guard.py --all: exit 0.
         Full suite and native Windows execution not run.
GATE:    n/a, approved read-only review
LESSON:  none
NEXT:    Fix findings before P0 approval.
```