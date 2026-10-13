# P0 fix red/green evidence

Date: 2026-10-10. All fixtures are synthetic. No real workbook or VBA audit opened.

Each finding was run independently against the original code before any production edits. All ten reproduced; none is NOT REPRODUCED. Full original outputs remain in `/tmp/price-planning-p0-evidence/`. Lines below are the failing test lines at the time of the RED run.

Engine probe: formualizer 0.11.1 returned exactly `{"type": "Error", "kind": KIND}` for `Div`, `Na`, `Value`, `Ref`, `Name`, `Num`. The NA category is spelled `Na` by this engine. Formulas: `=1/0`, `=NA()`, `=1+"x"`, `=#REF!`, `=MissingName`, `=SQRT(-1)`.

## Finding 01

Command: `.venv/bin/python -m unittest tests.test_p0_regressions.Finding01 -v`

RED: exit 1. Full output: `/tmp/price-planning-p0-evidence/finding-01-red.txt`.

```text
Traceback (most recent call last):
  File "/Users/gopalmacbook/Desktop/Max HQ/price-planning/tests/test_p0_regressions.py", line 55, in test_only_hash_verified_vendor_wheels_get_binary_exception
    self.assertTrue(guard({name: data}, allow={name: 'wrong'}))
AssertionError: [] is not true
```

GREEN: exit 0. Full output: `/tmp/price-planning-p0-evidence/finding-01-green.txt`.

```text

----------------------------------------------------------------------
Ran 2 tests in 0.001s

OK
```

## Finding 02

Command: `.venv/bin/python -m unittest tests.test_p0_regressions.Finding02 -v`

RED: exit 1. Full output: `/tmp/price-planning-p0-evidence/finding-02-red.txt`.

```text
Traceback (most recent call last):
  File "/Users/gopalmacbook/Desktop/Max HQ/price-planning/tests/test_p0_regressions.py", line 63, in test_bom_text_and_invalid_encodings_cannot_bypass_denylist
    self.assertTrue(guard({'note.txt': 'Forbidden Phrase'.encode(encoding)},
AssertionError: [] is not true
```

GREEN: exit 0. Full output: `/tmp/price-planning-p0-evidence/finding-02-green.txt`.

```text

----------------------------------------------------------------------
Ran 2 tests in 0.003s

OK
```

## Finding 03

Command: `.venv/bin/python -m unittest tests.test_p0_regressions.Finding03 -v`

RED: exit 1. Full output: `/tmp/price-planning-p0-evidence/finding-03-red.txt`.

```text
Traceback (most recent call last):
  File "/Users/gopalmacbook/Desktop/Max HQ/price-planning/tests/test_p0_regressions.py", line 92, in test_staged_type_change_is_scanned_from_index
    self.assertEqual(cg.main(['--staged']), 1)
AssertionError: 0 != 1
```

GREEN: exit 0. Full output: `/tmp/price-planning-p0-evidence/finding-03-green.txt`.

```text

----------------------------------------------------------------------
Ran 1 test in 0.002s

OK
```

## Finding 04

Command: `.venv/bin/python -m unittest tests.test_p0_regressions.Finding04 -v`

RED: exit 1. Full output: `/tmp/price-planning-p0-evidence/finding-04-red.txt`.

```text
Traceback (most recent call last):
  File "/Users/gopalmacbook/Desktop/Max HQ/price-planning/tests/test_p0_regressions.py", line 102, in test_force_added_private_paths_are_blocked_case_insensitively
    self.assertTrue(guard({name: b'synthetic'}))
AssertionError: [] is not true
```

GREEN: exit 0. Full output: `/tmp/price-planning-p0-evidence/finding-04-green.txt`.

```text

----------------------------------------------------------------------
Ran 1 test in 0.000s

OK
```

## Finding 05

Command: `.venv/bin/python -m unittest tests.test_p0_regressions.Finding05 -v`

RED: exit 1. Full output: `/tmp/price-planning-p0-evidence/finding-05-red.txt`.

```text
Traceback (most recent call last):
  File "/Users/gopalmacbook/Desktop/Max HQ/price-planning/tests/test_p0_regressions.py", line 118, in test_reports_refuse_equal_nested_and_symlinked_destinations
    with self.assertRaises(ValueError):
AssertionError: ValueError not raised
```

GREEN: exit 0. Full output: `/tmp/price-planning-p0-evidence/finding-05-green.txt`.

```text

----------------------------------------------------------------------
Ran 1 test in 0.010s

OK
```

## Finding 06

Command: `.venv/bin/python -m unittest tests.test_p0_regressions.Finding06 -v`

RED: exit 1. Full output: `/tmp/price-planning-p0-evidence/finding-06-red.txt`.

```text
Traceback (most recent call last):
  File "/Users/gopalmacbook/Desktop/Max HQ/price-planning/tests/test_p0_regressions.py", line 141, in test_real_engine_error_dicts_compare_to_excel_cached_tags
    self.assertEqual(typed(actual), ('error', tag))
AssertionError: Tuples differ: ('text', "{'type': 'Error', 'kind': 'Div'}") != ('error', '#DIV/0!')
```

GREEN: exit 0. Full output: `/tmp/price-planning-p0-evidence/finding-06-green.txt`.

```text

----------------------------------------------------------------------
Ran 1 test in 0.066s

OK
```

## Finding 07

Command: `.venv/bin/python -m unittest tests.test_p0_regressions.Finding07 -v`

RED: exit 1. Full output: `/tmp/price-planning-p0-evidence/finding-07-red.txt`.

```text
Traceback (most recent call last):
  File "/Users/gopalmacbook/Desktop/Max HQ/price-planning/tests/test_p0_regressions.py", line 158, in test_protection_uses_style_inheritance_and_apply_flag
    self.assertEqual(find_inputs._unlocked_style_ids(archive), {'1', '2', '4'})
AssertionError: Items in the first set but not the second:
```

GREEN: exit 0. Full output: `/tmp/price-planning-p0-evidence/finding-07-green.txt`.

```text

----------------------------------------------------------------------
Ran 1 test in 0.001s

OK
```

## Finding 08

Command: `.venv/bin/python -m unittest tests.test_p0_regressions.Finding08 -v`

RED: exit 1. Full output: `/tmp/price-planning-p0-evidence/finding-08-red.txt`.

```text
Traceback (most recent call last):
  File "/Users/gopalmacbook/Desktop/Max HQ/price-planning/tests/test_p0_regressions.py", line 225, in test_failed_install_does_not_leave_stamp
    self.run_installer(Path(directory), interpreter=False, install_fails=True)
  File "/Users/gopalmacbook/Desktop/Max HQ/price-planning/tests/test_p0_regressions.py", line 201, in run_installer
    with self.assertRaises(subprocess.CalledProcessError):
AssertionError: CalledProcessError not raised
```

GREEN: exit 0. Full output: `/tmp/price-planning-p0-evidence/finding-08-green.txt`.

```text

----------------------------------------------------------------------
Ran 4 tests in 0.010s

OK
```

## Finding 09

Command: `.venv/bin/python -m unittest tests.test_p0_regressions.Finding09 -v`

RED: exit 1. Full output: `/tmp/price-planning-p0-evidence/finding-09-red-clean.txt`.

```text
Traceback (most recent call last):
  File "/Users/gopalmacbook/Desktop/Max HQ/price-planning/tests/test_p0_regressions.py", line 247, in test_hook_supports_python_launcher_without_maxhq
    self.assertEqual(result.returncode, code, result.stderr)
AssertionError: 1 != 0 : /Users/gopalmacbook/Desktop/Max HQ/price-planning/.githooks/pre-commit: line 4: : command not found
```

GREEN: exit 0. Full output: `/tmp/price-planning-p0-evidence/finding-09-green.txt`.

```text

----------------------------------------------------------------------
Ran 1 test in 3.841s

OK
```

## Finding 10

Command: `.venv/bin/python -m unittest tests.test_p0_regressions.Finding10 -v`

RED: exit 1. Full output: `/tmp/price-planning-p0-evidence/finding-10-red.txt`.

```text
Traceback (most recent call last):
  File "/Users/gopalmacbook/Desktop/Max HQ/price-planning/tests/test_p0_regressions.py", line 275, in test_quoted_apostrophe_input_roundtrips_and_has_engine_fanout
    self.assertEqual(refs, ["'O''Brien'!A1"])
AssertionError: Lists differ: ["O'Brien!A1"] != ["'O''Brien'!A1"]
```

GREEN: exit 0. Full output: `/tmp/price-planning-p0-evidence/finding-10-green.txt`.

```text

----------------------------------------------------------------------
Ran 2 tests in 0.083s

OK
```

## Fixture corrections

The first launcher negative-path assertion tried to read a missing trace file. Before changing the hook, the test was corrected to assert that the trace exists and rerun RED with two assertion failures and no errors. The clean RED is cited above.

The apostrophe fanout fixture initially omitted `evaluate_all()` before inspection. The benchmark already performs evaluation before dependency inspection. Adding the same lifecycle step to the fixture preserved the one-dependent assertion and passed. The real engine rejects `O'Brien!A1` with `InvalidInspectionAddressError: unterminated sheet quote` and accepts `'O''Brien'!A1` after evaluation.

## Finding 01 compatibility follow-up

The first whole-repository guard rejected four existing ZIP wheels not listed in the original allowlist. Add a test that independently verifies their SHA-256 hashes occur in requirements.lock or requirements-dev.lock, then checks that the guard accepts the exact bytes.

RED command: `.venv/bin/python -m unittest tests.test_p0_regressions.Finding01.test_locked_tracked_vendor_wheels_keep_their_hash_exceptions -v`, exit 1, four failures. Full output: `/tmp/price-planning-p0-evidence/vendor-exceptions-red.txt`.

```text
Traceback (most recent call last):
  File "/Users/gopalmacbook/Desktop/Max HQ/price-planning/tests/test_p0_regressions.py", line 49, in test_locked_tracked_vendor_wheels_keep_their_hash_exceptions
    self.assertEqual(guard({name: data}, allow=allow), [])
AssertionError: Lists differ: ['BLOCKED vendor/waitress-3.0.2-py3-none-any.whl: ZIP content'] != []
```

Minimal fix: allow exact hash-matched .whl files under vendor/, including the legacy Waitress wheel used by app/runtime.py. Add only the four already-tracked wheel hashes verified against the existing locks. No wheels, versions or dependency locks changed.

GREEN: `.venv/bin/python -m unittest tests.test_p0_regressions.Finding01 -v`, exit 0, three tests OK. Full output: `/tmp/price-planning-p0-evidence/vendor-exceptions-green.txt`.

## Scoped verification and remaining gates

`.venv/bin/python -m unittest tests.test_p0_regressions tests.test_content_guard tests.test_typed tests.test_install_env tests.test_bench -v`: exit 0, 34 tests OK. Full output: `/tmp/price-planning-p0-evidence/focused-final.txt`.

`git --no-optional-locks diff --check -- ops tests .githooks .planning/excel-grid/sol`: exit 0. Local Gitleaks: no leaks. Three local Semgrep rules: zero findings in five production Python files; this is a limited SAST check. Scanner outputs: `/tmp/price-planning-p0-evidence/secrets-final.txt`, `/tmp/price-planning-p0-evidence/sast-final.txt`.

Named blockers: the mandatory exact `claude-opus-5-5` review and online dependency vulnerability audit require network, prohibited in this session. No review or vulnerability clearance is claimed. The full-suite runtime launcher test requires a localhost bind prohibited by this sandbox, so full-suite success is not claimed. No test was skipped or weakened.

Final command: `.venv/bin/python -m unittest discover -s tests`, exit 1. Full output: `/tmp/price-planning-p0-evidence/full-suite-final.txt`.

```text
........................................................E......
======================================================================
ERROR: test_first_admin_restart_and_real_waitress_launcher (test_runtime.Runtime.test_first_admin_restart_and_real_waitress_launcher)
----------------------------------------------------------------------
Traceback (most recent call last):
  File "/Users/gopalmacbook/Desktop/Max HQ/price-planning/tests/test_runtime.py", line 75, in test_first_admin_restart_and_real_waitress_launcher
    sock.bind(('127.0.0.1', 0))
PermissionError: [Errno 1] Operation not permitted

----------------------------------------------------------------------
Ran 63 tests in 6.718s

FAILED (errors=1)
```

Final command: `python3 ops/content_guard.py --all`, exit 0. Full output: `/tmp/price-planning-p0-evidence/content-guard-final.txt`.

```text
(no output)
```
