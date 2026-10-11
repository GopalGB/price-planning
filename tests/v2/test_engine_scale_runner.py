import io
import itertools
import json
import os
import signal
import subprocess
import sys
import tempfile
import time
import unittest
from contextlib import redirect_stderr
from pathlib import Path
from unittest import mock

from ops.bench_engine_scale import (
    ChildContext,
    EditContext,
    MonitorContext,
    ParentContext,
    _arguments,
    _check_delta,
    _parent_signal_handlers,
    _resource_check,
    _run_edits,
    _stop_owned,
    _stop_reason,
    _write_report,
    collect_child_result,
    monitor_owned_child,
    parse_swap_mib,
    run_child,
    run_parent,
)
from ops.engine_scale_fixture import build_manifest, expected_formula_values


class _FakeChild:
    def __init__(self, code=None):
        self.code = code
        self.terminated = False
        self.killed = False
        self.waited = False

    def poll(self):
        return self.code

    def terminate(self):
        self.terminated = True
        self.code = -15

    def kill(self):
        self.killed = True
        self.code = -9

    def wait(self, timeout=None):
        self.waited = True
        return self.code


class _RacedExitChild(_FakeChild):
    def terminate(self):
        self.code = 0
        raise ProcessLookupError("child exited before signal")


class _InterruptingChild(_FakeChild):
    def __init__(self):
        super().__init__()
        self.polls = 0

    def poll(self):
        self.polls += 1
        if self.polls == 1:
            raise KeyboardInterrupt
        return self.code


class _LateInterruptingChild(_FakeChild):
    pid = 999999999

    def __init__(self):
        super().__init__()
        self.polls = 0

    def poll(self):
        self.polls += 1
        if self.polls == 3:
            raise KeyboardInterrupt
        return self.code


def _monitor(timeout, baseline, swap_reader, disk_reader=lambda _: 4096.0, clock=None):
    return MonitorContext(
        timeout_seconds=timeout,
        swap_baseline_mib=baseline,
        swap_reader=swap_reader,
        disk_reader=disk_reader,
        sleep=lambda _: None,
        clock=clock or itertools.count(0.0, 0.5).__next__,
    )


class EngineScaleRunnerTests(unittest.TestCase):
    def test_check_delta_fails_closed_on_count_coordinate_and_value_mismatch(self):
        expected = [(0, 0, 0, 8.0), (0, 0, 1, 9.0)]

        self.assertTrue(_check_delta(list(expected), iter(expected), 2))
        self.assertFalse(_check_delta(expected[:1], iter(expected), 2))
        self.assertFalse(_check_delta([(0, 0, 1, 8.0), expected[1]], iter(expected), 2))
        self.assertFalse(_check_delta([expected[0], (0, 0, 1, 10.0)], iter(expected), 2))

    def test_parse_swap_mib_accepts_decimal_kilobyte_megabyte_and_gigabyte_values(self):
        self.assertEqual(parse_swap_mib("total = 8.00G  used = 1.50G"), 1536.0)
        self.assertEqual(parse_swap_mib("Swap: 1024K used = 512.5M"), 512.5)

    def test_resource_guards_use_owner_boundaries_from_common_baseline(self):
        with tempfile.TemporaryDirectory() as tmp:
            at_limits = _resource_check(
                _monitor(600, 100.0, lambda: 1636.0, disk_reader=lambda _: 2500.0),
                Path(tmp),
            )
            over_swap = _resource_check(
                _monitor(600, 100.0, lambda: 1636.1, disk_reader=lambda _: 2500.0),
                Path(tmp),
            )
            under_disk = _resource_check(
                _monitor(600, 100.0, lambda: 1636.0, disk_reader=lambda _: 2499.9),
                Path(tmp),
            )

        self.assertNotIn("failure_stage", at_limits)
        self.assertEqual(over_swap["failure_stage"], "swap_watchdog")
        self.assertEqual(under_disk["failure_stage"], "disk_watchdog")
        self.assertEqual(_stop_reason("swap_watchdog"), "swap_growth_exceeded_1536_mib")
        self.assertEqual(_stop_reason("disk_watchdog"), "disk_free_below_2500_mib")

    def test_watchdog_kills_only_its_child_when_swap_growth_exceeds_limit(self):
        with tempfile.TemporaryDirectory() as tmp:
            progress = Path(tmp) / "progress.json"
            progress.write_text(json.dumps({"stage": "load", "load_seconds": 1.25}))
            child = _FakeChild()
            readings = iter([100.0, 2148.1])
            context = _monitor(600, 100.0, lambda: next(readings))
            result = monitor_owned_child(child, progress, context)
        self.assertTrue(child.terminated)
        self.assertTrue(child.waited)
        self.assertEqual(result["status"], "failed")
        self.assertEqual(result["failure_stage"], "swap_watchdog")
        self.assertEqual(result["latest_progress"]["stage"], "load")
        self.assertEqual(result["attempt_elapsed_seconds"], 1.0)

    @unittest.skipIf(os.name == "nt", "POSIX process-group behavior")
    def test_repeated_stop_does_not_resignal_a_reused_process_group(self):
        child = _FakeChild(0)
        child.pid = 4321
        child._engine_scale_pgid = 4321

        with mock.patch("ops.bench_engine_scale.os.getpgid") as getpgid, \
                mock.patch("ops.bench_engine_scale.os.killpg") as killpg:
            _stop_owned(child)
            _stop_owned(child)

        killpg.assert_called_once_with(4321, signal.SIGKILL)
        getpgid.assert_not_called()

    @unittest.skipIf(os.name == "nt", "POSIX process-group behavior")
    def test_watchdog_kills_real_child_and_preserves_all_resource_extrema(self):
        child = subprocess.Popen(
            [sys.executable, "-c", "import time; time.sleep(60)"], start_new_session=True
        )
        with tempfile.TemporaryDirectory() as tmp:
            progress = Path(tmp) / "progress.json"
            progress.write_text(json.dumps({"stage": "load"}))
            swaps, disks, rss = iter([100.0, 2200.1]), iter([4096.0, 3072.0]), iter([1234, 5678])
            context = MonitorContext(
                timeout_seconds=600, swap_baseline_mib=100.0,
                swap_reader=lambda: next(swaps), disk_reader=lambda _: next(disks),
                sleep=lambda _: None, clock=itertools.count(0.0, 0.5).__next__,
                rss_reader=lambda _: next(rss),
            )

            result = monitor_owned_child(child, progress, context)

        self.assertIsNotNone(child.returncode)
        self.assertEqual(result["failure_stage"], "swap_watchdog")
        self.assertEqual(result["swap_max_mib"], 2200.1)
        self.assertEqual(result["disk_min_free_mib"], 3072.0)
        self.assertEqual(result["child_rss_max_bytes"], 5678)

    def test_timeout_and_failed_child_cannot_be_reported_completed(self):
        with tempfile.TemporaryDirectory() as tmp:
            progress = Path(tmp) / "progress.json"
            progress.write_text(json.dumps({"stage": "apply_ordinary", "status": "running"}))
            timed_out = _FakeChild()
            context = _monitor(1, 100.0, lambda: 100.0,
                               clock=iter([0.0, 1.1]).__next__)
            timeout_result = monitor_owned_child(timed_out, progress, context)
            failed_result = collect_child_result(-1, progress)
        self.assertEqual(timeout_result["status"], "failed")
        self.assertEqual(timeout_result["failure_stage"], "timeout")
        self.assertNotEqual(failed_result["status"], "completed")

    def test_only_a_successful_completed_child_may_report_completed(self):
        with tempfile.TemporaryDirectory() as tmp:
            progress = Path(tmp) / "progress.json"
            progress.write_text(json.dumps({"stage": "finished", "status": "completed",
                                             "load_seconds": 0.1,
                                             "native_edits_correct": True, "correctness": {
                                                 "ordinary": {"all_values_match": True},
                                                 "broad": {"all_values_match": True},
                                             }}))
            completed = collect_child_result(0, progress)
            progress.write_text(json.dumps({"stage": "finished", "status": "completed"}))
            incomplete = collect_child_result(0, progress)
        self.assertEqual(completed["status"], "completed")
        self.assertEqual(completed["load_seconds"], 0.1)
        self.assertNotEqual(incomplete["status"], "completed")

    def test_swap_read_failure_is_fail_closed_and_reaps_child(self):
        with tempfile.TemporaryDirectory() as tmp:
            progress = Path(tmp) / "progress.json"
            progress.write_text(json.dumps({"stage": "load", "load_seconds": 1.25}))
            child = _FakeChild()
            context = _monitor(600, 100.0, lambda: (_ for _ in ()).throw(OSError("swap")))

            result = monitor_owned_child(child, progress, context)

        self.assertEqual(result["failure_stage"], "resource_read")
        self.assertTrue(child.terminated)
        self.assertTrue(child.waited)
        self.assertEqual(result["latest_progress"]["load_seconds"], 1.25)

    def test_monitor_failure_preserves_prior_resource_maxima(self):
        with tempfile.TemporaryDirectory() as tmp:
            progress = Path(tmp) / "progress.json"
            progress.write_text(json.dumps({"stage": "load"}))
            child = _FakeChild()
            swaps = iter([100.0, OSError("swap")])

            def swap_reader():
                value = next(swaps)
                if isinstance(value, Exception):
                    raise value
                return value

            context = _monitor(600, 90.0, swap_reader)
            result = monitor_owned_child(child, progress, context)

        self.assertEqual(result["failure_stage"], "resource_read")
        self.assertEqual(result["swap_max_mib"], 100.0)
        self.assertEqual(result["disk_min_free_mib"], 4096.0)

    def test_ordinary_correctness_is_durable_before_broad_edit(self):
        manifest = build_manifest(0.0001, 20261011)
        ordinary = next(expected_formula_values(manifest, 7, 4))

        class Engine:
            calls = 0

            def apply(self, _changes):
                self.calls += 1
                if self.calls == 1:
                    return [ordinary]
                raise RuntimeError("broad interrupted")

        with tempfile.TemporaryDirectory() as tmp:
            progress = Path(tmp) / "progress.json"
            metrics = {}
            with self.assertRaisesRegex(RuntimeError, "broad interrupted"):
                _run_edits(Engine(), EditContext(manifest, metrics, progress))
            saved = json.loads(progress.read_text())

        self.assertTrue(saved["correctness"]["ordinary"]["all_values_match"])
        self.assertEqual(saved["correctness"]["ordinary"]["actual_changed"], 1)
        self.assertEqual(saved["stage"], "apply_broad")

    def test_raced_child_exit_is_still_waited_and_reaped(self):
        with tempfile.TemporaryDirectory() as tmp:
            progress = Path(tmp) / "progress.json"
            progress.write_text(json.dumps({"stage": "load"}))
            child = _RacedExitChild()
            context = _monitor(600, 100.0, lambda: 2200.1)

            result = monitor_owned_child(child, progress, context)

        self.assertEqual(result["failure_stage"], "swap_watchdog")
        self.assertTrue(child.waited)

    @unittest.skipIf(os.name == "nt", "POSIX process-group behavior")
    def test_non_session_child_falls_back_without_signaling_parent_group(self):
        with tempfile.TemporaryDirectory() as tmp:
            marker = Path(tmp) / "survived"
            script = Path(tmp) / "fallback.py"
            script.write_text(
                "import subprocess, sys, time\n"
                "from pathlib import Path\n"
                "sys.path.insert(0, sys.argv[2])\n"
                "from ops.bench_engine_scale import _stop_owned\n"
                "child = subprocess.Popen([sys.executable, '-c', 'import time; time.sleep(60)'])\n"
                "_stop_owned(child)\n"
                "Path(sys.argv[1]).write_text('yes')\n"
            )
            result = subprocess.run(
                [sys.executable, str(script), str(marker), str(Path(__file__).parents[2])],
                cwd=Path(__file__).parents[2],
                capture_output=True, text=True, timeout=10, start_new_session=True,
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertEqual(marker.read_text(), "yes")

    @unittest.skipIf(os.name == "nt", "POSIX signal behavior")
    def test_parent_signals_preserve_partial_report_and_stop_owned_group(self):
        unrelated = subprocess.Popen(
            [sys.executable, "-c", "import time; time.sleep(300)"], start_new_session=True
        )
        try:
            for parent_signal in (signal.SIGINT, signal.SIGTERM, signal.SIGHUP):
                with self.subTest(parent_signal=parent_signal), tempfile.TemporaryDirectory() as tmp:
                    self._assert_interrupted_parent(Path(tmp), parent_signal)
                self.assertIsNone(unrelated.poll())
        finally:
            unrelated.terminate()
            unrelated.wait(timeout=5)

    def test_tiny_real_child_checks_cache_native_and_both_edits(self):
        with tempfile.TemporaryDirectory() as tmp:
            workspace = Path(tmp)
            progress = workspace / "progress.json"

            code = run_child(ChildContext(0.0001, 20261011, workspace, progress, True))
            result = collect_child_result(code, progress)

        self.assertEqual(result["status"], "completed")
        self.assertEqual(result["cached_baseline"]["cached_match_count"], 150)
        self.assertEqual(result["cached_baseline"]["native_match_count"], 150)
        self.assertEqual(result["correctness"]["ordinary"]["actual_changed"], 1)
        self.assertEqual(result["correctness"]["broad"]["actual_changed"], 150)

    def test_load_failure_preserves_generated_manifest_summary(self):
        manifest = build_manifest(0.0001, 20261011)
        manifest.pop("formula_encoding_plan")
        with tempfile.TemporaryDirectory() as tmp:
            workspace = Path(tmp)
            progress = workspace / "progress.json"
            with mock.patch("ops.bench_engine_scale.generate", return_value=manifest), \
                    mock.patch(
                        "pp.engine.formualizer.FormualizerAdapter.load",
                        side_effect=RuntimeError("load failed"),
                    ):
                code = run_child(ChildContext(0.0001, 20261011, workspace, progress))
            result = collect_child_result(code, progress)

        self.assertEqual(result["failure_stage"], "load")
        self.assertEqual(result["actual_totals"]["cells"], manifest["cell_count"])
        self.assertEqual(result["actual_totals"]["formulas"], manifest["formula_count"])
        self.assertEqual(
            result["actual_totals"]["formula_encodings"],
            manifest["formula_encoding_counts"],
        )
        self.assertEqual(result["actual_totals"]["sheets"], manifest["sheets"])

    def test_disk_guard_failure_is_fail_closed_and_preserves_progress(self):
        with tempfile.TemporaryDirectory() as tmp:
            progress = Path(tmp) / "progress.json"
            progress.write_text(json.dumps({"stage": "apply_broad", "peak_rss_bytes": 1234}))
            child = _FakeChild()
            context = _monitor(600, 100.0, lambda: 100.0, disk_reader=lambda _: 1023.9)

            result = monitor_owned_child(child, progress, context)

        self.assertEqual(result["failure_stage"], "disk_watchdog")
        self.assertEqual(result["latest_progress"]["peak_rss_bytes"], 1234)

    def test_parent_checks_common_baseline_and_disk_before_spawning(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            spawned = []
            context = ParentContext(
                scales=(0.0001, 0.0002), seed=20261011,
                output=root / "report.json", timeout_seconds=60,
                swap_baseline_mib=6971.62, private_root=root,
                swap_reader=lambda: 6971.62,
                disk_reader=lambda _: 1000.0,
                spawn_child=lambda *args: spawned.append(args),
                physical_ram_reader=lambda: 8 * 1024 ** 3,
            )

            report = run_parent(context)

        self.assertEqual(spawned, [])
        self.assertEqual(report["attempted_scales"], [])
        self.assertEqual(report["skipped_scales"], [0.0001, 0.0002])
        self.assertEqual(report["stop_reason"], "disk_free_below_2500_mib")

    def test_parent_cleans_artifacts_after_every_completed_scale(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            outside = root / "outside.bin"
            outside.write_bytes(b"keep")
            workspaces = []

            def spawn(context):
                workspaces.append(context.workspace)
                context.progress_path.write_text(json.dumps({
                    "stage": "finished", "status": "completed",
                    "native_edits_correct": True,
                    "correctness": {
                        "ordinary": {"all_values_match": True},
                        "broad": {"all_values_match": True},
                    },
                }))
                (context.workspace / "child.log").write_text("recoverable log")
                (context.workspace / "fixture.xlsx").write_bytes(b"fixture")
                scratch = context.workspace / "scratch"
                scratch.mkdir()
                (scratch / "large.bin").write_bytes(b"temporary")
                return _FakeChild(0)

            context = ParentContext(
                scales=(0.0001, 0.0002), seed=20261011,
                output=root / "report.json", timeout_seconds=60,
                swap_baseline_mib=100.0, private_root=root,
                swap_reader=lambda: 100.0, disk_reader=lambda _: 4096.0,
                spawn_child=spawn, physical_ram_reader=lambda: 1024,
            )

            report = run_parent(context)
            remaining = [
                {path.name for path in workspace.iterdir()} for workspace in workspaces
            ]
            outside_contents = outside.read_bytes()

        self.assertEqual(report["completed_scales"], [0.0001, 0.0002])
        self.assertEqual(remaining, [
            {"child.log", "progress.json"}, {"child.log", "progress.json"},
        ])
        self.assertEqual(outside_contents, b"keep")

    def test_parent_stops_future_scales_after_child_failure(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            spawned = []

            def spawn(context):
                spawned.append(context.scale)
                context.progress_path.write_text(json.dumps({
                    "stage": "failed", "status": "failed", "failure_stage": "apply_broad"
                }))
                (context.workspace / "child.log").write_text("recoverable log")
                (context.workspace / "fixture.xlsx").write_bytes(b"fixture")
                (context.workspace / "temporary.bin").write_bytes(b"temporary")
                return _FakeChild(1)

            context = ParentContext(
                scales=(0.0001, 0.0002), seed=20261011,
                output=root / "report.json", timeout_seconds=60,
                swap_baseline_mib=6971.62, private_root=root,
                swap_reader=lambda: 6971.62,
                disk_reader=lambda _: 4096.0, spawn_child=spawn,
                physical_ram_reader=lambda: 8 * 1024 ** 3,
            )

            report = run_parent(context)
            workspace = next(root.glob("engine-scale-*"))
            remaining = {path.name for path in workspace.iterdir()}

        self.assertEqual(spawned, [0.0001])
        self.assertEqual(report["attempted_scales"], [0.0001])
        self.assertEqual(report["skipped_scales"], [0.0002])
        self.assertEqual(report["stop_reason"], "child_failed")
        self.assertEqual(remaining, {"child.log", "progress.json"})

    def test_monitor_keyboard_interrupt_preserves_completed_scale_and_partial_report(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            spawned = []

            def spawn(context):
                spawned.append(context.scale)
                (context.workspace / "child.log").write_text("recoverable log")
                (context.workspace / "fixture.xlsx").write_bytes(b"fixture")
                (context.workspace / "temporary.bin").write_bytes(b"temporary")
                if len(spawned) == 1:
                    context.progress_path.write_text(json.dumps({
                        "stage": "finished", "status": "completed",
                        "native_edits_correct": True,
                        "correctness": {
                            "ordinary": {"all_values_match": True},
                            "broad": {"all_values_match": True},
                        },
                    }))
                    return _FakeChild(0)
                context.progress_path.write_text(json.dumps({"stage": "load"}))
                return _InterruptingChild()

            output = root / "report.json"
            context = ParentContext(
                scales=(0.0001, 0.0002, 0.0003), seed=20261011, output=output,
                timeout_seconds=60, swap_baseline_mib=0.0, private_root=root,
                swap_reader=lambda: 0.0, disk_reader=lambda _: 4096.0,
                spawn_child=spawn, physical_ram_reader=lambda: 1024,
            )
            with self.assertRaises(KeyboardInterrupt):
                run_parent(context)
            report = json.loads(output.read_text())
            remaining = [
                {path.name for path in workspace.iterdir()}
                for workspace in root.glob("engine-scale-*")
            ]

        self.assertEqual(report["completed_scales"], [0.0001])
        self.assertEqual(report["attempted_scales"], [0.0001, 0.0002])
        self.assertEqual(report["results"][1]["failure_stage"], "interrupted")
        self.assertEqual(report["skipped_scales"], [0.0003])
        self.assertEqual(report["stop_reason"], "interrupted")
        self.assertEqual(remaining, [
            {"child.log", "progress.json"}, {"child.log", "progress.json"},
        ])

    def test_preflight_interrupt_accounts_for_all_unstarted_scales(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)

            def interrupt():
                raise KeyboardInterrupt

            context = ParentContext(
                scales=(0.0001, 0.0002, 0.0003), seed=20261011,
                output=root / "report.json", timeout_seconds=60,
                swap_baseline_mib=0.0, private_root=root,
                swap_reader=interrupt, disk_reader=lambda _: 4096.0,
                physical_ram_reader=lambda: 1024,
            )
            with self.assertRaises(KeyboardInterrupt):
                run_parent(context)
            report = json.loads(context.output.read_text())

        self.assertEqual(report["attempted_scales"], [])
        self.assertEqual(report["skipped_scales"], list(context.scales))
        self.assertTrue(all(result["status"] == "skipped" for result in report["results"]))
        self.assertEqual(report["stop_reason"], "interrupted")

    def test_workspace_creation_interrupt_accounts_for_all_unstarted_scales(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            context = ParentContext(
                scales=(0.0001, 0.0002), seed=20261011,
                output=root / "report.json", timeout_seconds=60,
                swap_baseline_mib=0.0, private_root=root,
                swap_reader=lambda: 0.0, disk_reader=lambda _: 4096.0,
                physical_ram_reader=lambda: 1024,
            )
            with mock.patch(
                "ops.bench_engine_scale.tempfile.mkdtemp", side_effect=KeyboardInterrupt
            ), self.assertRaises(KeyboardInterrupt):
                run_parent(context)
            report = json.loads(context.output.read_text())

        self.assertEqual(report["attempted_scales"], [])
        self.assertEqual(report["skipped_scales"], list(context.scales))
        self.assertEqual(report["stop_reason"], "interrupted")

    def test_post_monitor_interrupt_preserves_completed_result_and_timings(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)

            def spawn(context):
                context.progress_path.write_text(json.dumps({
                    "stage": "finished", "status": "completed", "load_seconds": 1.25,
                    "native_edits_correct": True,
                    "correctness": {
                        "ordinary": {"all_values_match": True},
                        "broad": {"all_values_match": True},
                    },
                }))
                return _FakeChild(0)

            context = ParentContext(
                scales=(0.0001, 0.0002), seed=20261011,
                output=root / "report.json", timeout_seconds=60,
                swap_baseline_mib=0.0, private_root=root,
                swap_reader=lambda: 0.0, disk_reader=lambda _: 4096.0,
                spawn_child=spawn, physical_ram_reader=lambda: 1024,
            )
            with mock.patch(
                "ops.bench_engine_scale.platform.platform", side_effect=KeyboardInterrupt
            ), self.assertRaises(KeyboardInterrupt):
                run_parent(context)
            report = json.loads(context.output.read_text())

        self.assertEqual(report["attempted_scales"], [0.0001])
        self.assertEqual(report["completed_scales"], [0.0001])
        self.assertEqual(report["results"][0]["load_seconds"], 1.25)
        self.assertEqual(report["skipped_scales"], [0.0002])
        self.assertEqual(report["stop_reason"], "interrupted")

    def test_monitor_interrupt_preserves_running_resource_extrema_in_report(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            swaps = iter([90.0, 100.0, 500.0])
            rss = iter([1234, 5678])

            def spawn(context):
                context.progress_path.write_text(json.dumps({"stage": "load"}))
                return _LateInterruptingChild()

            context = ParentContext(
                scales=(0.0001, 0.0002), seed=20261011,
                output=root / "report.json", timeout_seconds=60,
                swap_baseline_mib=90.0, private_root=root,
                swap_reader=lambda: next(swaps), disk_reader=lambda _: 4096.0,
                spawn_child=spawn, physical_ram_reader=lambda: 1024,
                rss_reader=lambda _: next(rss),
            )
            with mock.patch("ops.bench_engine_scale.POLL_SECONDS", 0), self.assertRaises(
                KeyboardInterrupt
            ):
                run_parent(context)
            report = json.loads(context.output.read_text())

        result = report["results"][0]
        self.assertEqual(result["swap_max_mib"], 500.0)
        self.assertEqual(result["disk_min_free_mib"], 4096.0)
        self.assertEqual(result["child_rss_max_bytes"], 5678)

    @unittest.skipUnless(hasattr(signal, "SIGHUP"), "SIGHUP behavior")
    def test_parent_signal_handlers_preserve_ignored_sighup(self):
        previous = signal.getsignal(signal.SIGHUP)
        try:
            signal.signal(signal.SIGHUP, signal.SIG_IGN)
            with _parent_signal_handlers():
                self.assertIs(signal.getsignal(signal.SIGHUP), signal.SIG_IGN)
        finally:
            signal.signal(signal.SIGHUP, previous)

    def test_interrupted_report_write_preserves_last_completed_report(self):
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "report.json"
            previous = '{"completed_scales": [0.0001]}\n'
            output.write_text(previous)
            context = ParentContext((), 1, output, 60, 0.0, Path(tmp))

            def interrupting_write(path, value):
                with path.open("w") as stream:
                    stream.write(value[:5])
                raise KeyboardInterrupt

            with mock.patch.object(Path, "write_text", interrupting_write), self.assertRaises(
                KeyboardInterrupt
            ):
                _write_report(context, {"completed_scales": [0.0001, 0.0002]})

            self.assertEqual(output.read_text(), previous)
            self.assertFalse(output.with_suffix(".json.tmp").exists())

    def test_timeout_rejects_zero_and_negative_values(self):
        for value in ("0", "-1"):
            argv = [
                "bench_engine_scale", "--scales", "0.0001", "--seed", "1",
                "--output", "unused.json", "--timeout", value,
            ]
            with self.subTest(value=value), mock.patch.object(sys, "argv", argv), \
                    redirect_stderr(io.StringIO()), self.assertRaises(SystemExit):
                _arguments()

    def test_exact_scales_flag_rejects_scale_abbreviation(self):
        result = subprocess.run(
            [sys.executable, "-m", "ops.bench_engine_scale", "--scale", "0.0001",
             "--seed", "20261011", "--output", "unused.json"],
            capture_output=True, text=True,
        )

        self.assertNotEqual(result.returncode, 0)
        self.assertIn("unrecognized arguments: --scale", result.stderr)

    def _assert_interrupted_parent(self, root, parent_signal):
        report, pids = root / "report.json", root / "pids.json"
        script = root / "parent.py"
        script.write_text(self._interrupt_script())
        parent = subprocess.Popen(
            ["/bin/sh", "-c", 'trap "" INT; exec "$@"', "sh", sys.executable,
             str(script), str(report), str(pids), str(root), str(Path(__file__).parents[2])],
            cwd=Path(__file__).parents[2], stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            text=True, start_new_session=True,
        )
        child_pgid = None
        try:
            self._wait_for_path(pids)
            child_pgid = json.loads(pids.read_text())["leader"]
            os.kill(parent.pid, parent_signal)
            _, stderr = parent.communicate(timeout=10)
            self.assertNotEqual(parent.returncode, 0)
            self.assertTrue(report.exists())
            data = json.loads(report.read_text())
            self.assertEqual(data["stop_reason"], "interrupted", (data, stderr))
            self.assertEqual(data["attempted_scales"], [0.0001])
            self.assertEqual(data["results"][0]["failure_stage"], "interrupted")
            self.assertEqual(data["skipped_scales"], [0.0002])
            self._wait_for_group_exit(child_pgid)
        finally:
            if parent.poll() is None:
                parent.kill()
            if child_pgid is not None:
                try:
                    os.killpg(child_pgid, signal.SIGKILL)
                except ProcessLookupError:
                    pass
            parent.communicate()

    def _interrupt_script(self):
        return """import json, os, signal, subprocess, sys\nfrom pathlib import Path\nsys.path.insert(0, sys.argv[4])\nfrom ops.bench_engine_scale import ParentContext, run_parent\nsignal.signal(signal.SIGINT, signal.default_int_handler)\nsignal.signal(signal.SIGTERM, signal.SIG_DFL)\nsignal.signal(signal.SIGHUP, signal.SIG_DFL)\n\ndef spawn(context):\n    child = Path(sys.argv[3]) / 'child.py'\n    child.write_text(\"import json, os, signal, subprocess, sys, time\\nfrom pathlib import Path\\ngrandchild = subprocess.Popen([sys.executable, '-c', 'import time; time.sleep(300)'])\\ndef stop(*_):\\n    grandchild.terminate()\\n    grandchild.wait()\\n    raise SystemExit(143)\\nsignal.signal(signal.SIGTERM, stop)\\ntime.sleep(0.05)\\ntarget = Path(sys.argv[1])\\ntemporary = target.with_suffix('.tmp')\\ntemporary.write_text(json.dumps({'leader': os.getpid(), 'grandchild': grandchild.pid}))\\nos.replace(temporary, target)\\ntime.sleep(300)\\n\")\n    return subprocess.Popen([sys.executable, str(child), sys.argv[2]], start_new_session=True)\n\ncontext = ParentContext((0.0001, 0.0002), 1, Path(sys.argv[1]), 60, 0.0, Path(sys.argv[3]), lambda: 0.0, lambda _: 4096.0, spawn, lambda: 1024, rss_reader=lambda _: 1024)\nrun_parent(context)\n"""

    def _wait_for_path(self, path):
        for _ in range(200):
            if path.exists():
                return
            time.sleep(0.01)
        self.fail(f"timed out waiting for {path.name}")

    def _wait_for_group_exit(self, pgid):
        for _ in range(250):
            try:
                os.killpg(pgid, 0)
            except ProcessLookupError:
                return
            time.sleep(0.02)
        self.fail(f"owned process group {pgid} survived interruption")


if __name__ == "__main__":
    unittest.main()
