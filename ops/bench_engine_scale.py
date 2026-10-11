"""Run bounded, reproducible native-engine scale measurements in an owned child."""

import argparse
import importlib.metadata
import json
import os
import platform
import re
import resource
import shutil
import signal
import subprocess
import sys
import tempfile
import time
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

from ops.engine_scale_fixture import expected_formula_values, generate, validate_scale


POLL_SECONDS = 0.5
SWAP_GROWTH_LIMIT_MIB = 1536.0
DISK_FREE_MIN_MIB = 2500.0
MIB = 1024 * 1024


def parse_swap_mib(text):
    match = re.search(r"\bused\s*=\s*([0-9]+(?:\.[0-9]+)?)\s*([KMG])\b", text, re.I)
    if not match:
        raise ValueError("swap usage does not contain a used K/M/G value")
    amount, unit = float(match.group(1)), match.group(2).upper()
    return amount * {"K": 1 / 1024, "M": 1, "G": 1024}[unit]


def read_swap_mib():
    if sys.platform == "darwin":
        result = subprocess.run(
            ["sysctl", "vm.swapusage"], check=True, capture_output=True, text=True
        )
        return parse_swap_mib(result.stdout)
    if sys.platform.startswith("linux"):
        values = _linux_memory_values()
        return (values["SwapTotal"] - values["SwapFree"]) / 1024
    raise RuntimeError("swap telemetry is unsupported on this platform")


def _linux_memory_values():
    values = {}
    for line in Path("/proc/meminfo").read_text().splitlines():
        key, value = line.split(":", 1)
        values[key] = int(value.split()[0])
    return values


def read_disk_free_mib(path):
    return shutil.disk_usage(path).free / MIB


def physical_ram_bytes():
    if sys.platform.startswith("linux"):
        return _linux_memory_values()["MemTotal"] * 1024
    if sys.platform == "darwin":
        for command in (["sysctl", "-n", "hw.memsize"], ["hostinfo"]):
            try:
                output = subprocess.run(
                    command, check=True, capture_output=True, text=True
                ).stdout.strip()
                if command[0] == "sysctl":
                    return int(output)
                match = re.search(r"Primary memory available:\s*([0-9.]+)\s*gigabytes", output, re.I)
                if match:
                    return round(float(match.group(1)) * 1024 ** 3)
            except (OSError, subprocess.SubprocessError, ValueError):
                continue
    pages = os.sysconf("SC_PHYS_PAGES")
    size = os.sysconf("SC_PAGE_SIZE")
    return pages * size


def read_child_rss_bytes(pid):
    result = subprocess.run(
        ["ps", "-o", "rss=", "-p", str(pid)], capture_output=True, text=True, check=False
    )
    value = result.stdout.strip()
    if result.returncode or not value:
        return None
    return int(value) * 1024


@dataclass(frozen=True)
class MonitorContext:
    timeout_seconds: float
    swap_baseline_mib: float
    swap_reader: object = read_swap_mib
    disk_reader: object = read_disk_free_mib
    sleep: object = time.sleep
    clock: object = time.monotonic
    rss_reader: object = read_child_rss_bytes


@dataclass(frozen=True)
class ChildContext:
    scale: float
    seed: int
    workspace: Path
    progress_path: Path
    verify_native: bool = False


@dataclass(frozen=True)
class EditContext:
    manifest: dict
    metrics: dict
    progress_path: Path


def _write_progress(path, data):
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(data, sort_keys=True) + "\n")
    os.replace(temporary, path)


def _read_progress(path):
    try:
        return json.loads(path.read_text())
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def _rss_bytes():
    value = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    return value if sys.platform == "darwin" else value * 1024


def _number_matches(actual, expected):
    return isinstance(actual, (int, float)) and abs(actual - expected) < 1e-9


def _check_delta(changes, expected, expected_count):
    if len(changes) != expected_count:
        return False
    for actual, wanted in zip(changes, expected, strict=True):
        if actual[:3] != wanted[:3] or not _number_matches(actual[3], wanted[3]):
            return False
    return True


def _tiny_cache_verification(engine, manifest, verify_native):
    if not verify_native:
        return {"performed": False, "reason": "not_requested"}
    anchors = sorted(engine.formula_anchors)
    expected = list(expected_formula_values(manifest, 7, 3))
    cached_matches = sum(
        coordinate == wanted[:3] and _number_matches(engine.get(*coordinate), wanted[3])
        for coordinate, wanted in zip(anchors, expected, strict=True)
    )
    engine.workbook.evaluate_cells([
        (engine.names[sheet], row + 1, col + 1) for sheet, row, col in anchors
    ])
    native_matches = sum(
        coordinate == wanted[:3] and _number_matches(
            engine.workbook.get_value(engine.names[coordinate[0]], coordinate[1] + 1,
                                      coordinate[2] + 1), wanted[3]
        )
        for coordinate, wanted in zip(anchors, expected, strict=True)
    )
    return {"performed": True, "formula_count": len(anchors),
            "cached_match_count": cached_matches, "native_match_count": native_matches,
            "all_values_match": cached_matches == native_matches == len(anchors)}


def _stage(metrics, path, name):
    metrics.update({"stage": name, "status": "running",
                    "rss_at_stage_start_bytes": _rss_bytes()})
    _write_progress(path, metrics)


def _run_edits(engine, context):
    manifest, metrics = context.manifest, context.metrics
    _stage(metrics, context.progress_path, "apply_ordinary")
    started = time.perf_counter()
    ordinary = engine.apply([(0, 0, 1, 4)])
    metrics["ordinary_edit_seconds"] = time.perf_counter() - started
    _write_progress(context.progress_path, metrics)
    ordinary_expected = [next(expected_formula_values(manifest, 7, 4))]
    ordinary_ok = _check_delta(ordinary, ordinary_expected, 1)
    metrics["correctness"] = {
        "ordinary": {"expected_changed": 1, "actual_changed": len(ordinary),
                     "all_values_match": ordinary_ok},
    }
    _write_progress(context.progress_path, metrics)
    _stage(metrics, context.progress_path, "apply_broad")
    started = time.perf_counter()
    broad = engine.apply([(0, 0, 0, 8)])
    metrics["broad_edit_seconds"] = time.perf_counter() - started
    _write_progress(context.progress_path, metrics)
    broad_ok = _check_delta(
        broad, expected_formula_values(manifest, 8, 4), manifest["formula_count"]
    )
    metrics["correctness"]["broad"] = {
        "expected_changed": manifest["formula_count"],
        "actual_changed": len(broad), "all_values_match": broad_ok,
    }
    _write_progress(context.progress_path, metrics)
    return ordinary_ok and broad_ok


def run_child(context):
    """Generate, native-load, and apply two real edits without inventing timings."""
    context.workspace.mkdir(parents=True, exist_ok=True)
    metrics = {"stage": "generation", "status": "running"}
    try:
        _stage(metrics, context.progress_path, "generation")
        started = time.perf_counter()
        manifest = generate(context.workspace / "fixture.xlsx", context.scale, context.seed)
        metrics["generation_seconds"] = time.perf_counter() - started
        metrics["actual_totals"] = _manifest_summary(manifest)
        _write_progress(context.progress_path, metrics)
        _stage(metrics, context.progress_path, "load")
        from pp.engine.formualizer import FormualizerAdapter

        started = time.perf_counter()
        engine = FormualizerAdapter.load(context.workspace / "fixture.xlsx")
        metrics["load_seconds"] = time.perf_counter() - started
        _write_progress(context.progress_path, metrics)
        metrics["cached_baseline"] = _tiny_cache_verification(
            engine, manifest, context.verify_native
        )
        _write_progress(context.progress_path, metrics)
        if metrics["cached_baseline"].get("all_values_match") is False:
            raise ValueError("cached baseline differs from native evaluation")
        correct = _run_edits(engine, EditContext(manifest, metrics, context.progress_path))
        metrics.update(_child_summary(manifest, correct))
        _write_progress(context.progress_path, metrics)
        return 0 if correct else 1
    except Exception as exc:
        failure_stage = metrics.get("stage", "unknown")
        metrics.update({"stage": "failed", "status": "failed",
                        "failure_stage": failure_stage, "error_type": type(exc).__name__,
                        "child_peak_rss_bytes": _rss_bytes()})
        _write_progress(context.progress_path, metrics)
        return 1


def _child_summary(manifest, correct):
    return {
        "stage": "finished", "status": "completed" if correct else "failed",
        "child_peak_rss_bytes": _rss_bytes(), "report_date": _today(),
        "engine_version": importlib.metadata.version("formualizer"),
        "physical_ram_bytes": physical_ram_bytes(),
        "actual_totals": _manifest_summary(manifest),
    }


def _manifest_summary(manifest):
    return {
        "cells": manifest["cell_count"], "formulas": manifest["formula_count"],
        "formula_encodings": manifest["formula_encoding_counts"],
        "function_uses": manifest["function_use_counts"],
        "sheets": manifest["sheets"],
    }


def _today():
    return datetime.now(timezone.utc).date().isoformat()


def _failed_result(stage, progress, return_code=None):
    result = {key: value for key, value in progress.items()
              if key not in {"stage", "status", "failure_stage"}}
    result.update({"status": "failed", "failure_stage": stage,
                   "latest_progress": progress})
    if return_code is not None:
        result["return_code"] = return_code
    return result


def collect_child_result(return_code, progress_path):
    progress = _read_progress(progress_path)
    correctness = progress.get("correctness", {})
    complete = (return_code == 0 and progress.get("status") == "completed" and
                progress.get("stage") == "finished" and
                correctness.get("ordinary", {}).get("all_values_match") is True and
                correctness.get("broad", {}).get("all_values_match") is True)
    if complete:
        return dict(progress)
    stage = progress.get("failure_stage", progress.get("stage", "child_exit"))
    return _failed_result(stage, progress, return_code)


def _resource_check(context, path):
    try:
        swap = context.swap_reader()
        disk = context.disk_reader(path)
    except Exception as exc:
        return {"failure_stage": "resource_read", "error_type": type(exc).__name__}
    reading = {"swap_current_mib": swap, "disk_free_mib": disk}
    if swap - context.swap_baseline_mib > SWAP_GROWTH_LIMIT_MIB:
        reading["failure_stage"] = "swap_watchdog"
    elif disk < DISK_FREE_MIN_MIB:
        reading["failure_stage"] = "disk_watchdog"
    return reading


def _owned_pgid(process):
    pid = getattr(process, "pid", None)
    if not pid or os.name == "nt":
        return None
    if getattr(process, "_engine_scale_pgid", None) == pid:
        return pid
    try:
        pgid = os.getpgid(pid)
    except ProcessLookupError:
        return None
    if pgid != pid:
        return None
    process._engine_scale_pgid = pid
    return pid


def _signal_owned(process, force=False):
    if pgid := _owned_pgid(process):
        os.killpg(pgid, signal.SIGKILL if force else signal.SIGTERM)
    elif force and hasattr(process, "kill"):
        process.kill()
    else:
        process.terminate()


def _sweep_owned_group(process):
    pgid = getattr(process, "_engine_scale_pgid", None)
    process._engine_scale_pgid = None
    if not pgid or os.name == "nt":
        return
    try:
        os.killpg(pgid, signal.SIGKILL)
    except ProcessLookupError:
        pass


def _stop_owned(process):
    if process.poll() is None:
        _owned_pgid(process)
        try:
            _signal_owned(process)
        except ProcessLookupError:
            pass
    try:
        return_code = process.wait(timeout=5)
    except subprocess.TimeoutExpired:
        try:
            _signal_owned(process, force=True)
        except ProcessLookupError:
            pass
        return_code = process.wait()
    finally:
        _sweep_owned_group(process)
    return return_code


def _monitor_failure(process, progress_path, reading):
    _stop_owned(process)
    progress = _read_progress(progress_path)
    result = _failed_result(reading["failure_stage"], progress)
    result.update(reading)
    return result


def _monitor_reading(process, progress_path, context):
    reading = _resource_check(context, progress_path.parent)
    if reading.get("failure_stage") == "resource_read":
        return reading
    pid = getattr(process, "pid", None)
    if pid is None:
        return reading
    try:
        rss = context.rss_reader(pid)
    except Exception as exc:
        reading.setdefault("failure_stage", "resource_read")
        reading.setdefault("error_type", type(exc).__name__)
        return reading
    if rss is not None:
        reading["child_rss_bytes"] = rss
    return reading


def _update_maxima(maxima, reading):
    if "swap_current_mib" in reading:
        maxima["swap_max_mib"] = max(maxima["swap_max_mib"],
                                     reading["swap_current_mib"])
    if "disk_free_mib" in reading:
        current = maxima["disk_min_free_mib"]
        maxima["disk_min_free_mib"] = (reading["disk_free_mib"] if current is None
                                        else min(current, reading["disk_free_mib"]))
    if "child_rss_bytes" in reading:
        current = maxima["child_rss_max_bytes"] or 0
        maxima["child_rss_max_bytes"] = max(current, reading["child_rss_bytes"])


def monitor_owned_child(process, progress_path, context):
    started = context.clock()
    maxima = {"swap_baseline_mib": context.swap_baseline_mib,
              "swap_max_mib": context.swap_baseline_mib,
              "disk_min_free_mib": None, "child_rss_max_bytes": None}
    process._engine_scale_maxima = maxima
    try:
        while process.poll() is None:
            reading = _monitor_reading(process, progress_path, context)
            _update_maxima(maxima, reading)
            elapsed = max(context.clock() - started, 0.0)
            if "failure_stage" in reading:
                reading.update(maxima)
                result = _monitor_failure(process, progress_path, reading)
                result["attempt_elapsed_seconds"] = elapsed
                return result
            if elapsed > context.timeout_seconds:
                reading.update(maxima)
                reading["failure_stage"] = "timeout"
                result = _monitor_failure(process, progress_path, reading)
                result["attempt_elapsed_seconds"] = elapsed
                return result
            context.sleep(POLL_SECONDS)
        result = collect_child_result(process.wait(), progress_path)
        result.update(maxima)
        result["attempt_elapsed_seconds"] = max(context.clock() - started, 0.0)
        return result
    finally:
        _stop_owned(process)


def _spawn_child(context):
    log_path = context.workspace / "child.log"
    command = [sys.executable, "-m", "ops.bench_engine_scale", "--child",
               "--scales", str(context.scale), "--seed", str(context.seed),
               "--workspace", str(context.workspace), "--progress", str(context.progress_path)]
    if context.verify_native:
        command.append("--verify-native")
    process = None
    try:
        with log_path.open("w") as log:
            process = subprocess.Popen(
                command, stdout=log, stderr=subprocess.STDOUT, start_new_session=True
            )
        _owned_pgid(process)
        return process
    except BaseException:
        if process is not None:
            _stop_owned(process)
        raise


def _cleanup_workspace(workspace, private_root):
    if not workspace.exists():
        return
    root = private_root.resolve()
    if (workspace.is_symlink() or workspace.parent.resolve() != root or
            not workspace.name.startswith("engine-scale-")):
        raise ValueError("refusing to clean a non-owned benchmark workspace")
    for path in workspace.iterdir():
        if path.name in {"progress.json", "child.log"}:
            continue
        if path.is_symlink() or path.is_file():
            path.unlink()
        else:
            shutil.rmtree(path)


@dataclass(frozen=True)
class ParentContext:
    scales: tuple[float, ...]
    seed: int
    output: Path
    timeout_seconds: float
    swap_baseline_mib: float | None = None
    private_root: Path = Path("results-private")
    swap_reader: object = read_swap_mib
    disk_reader: object = read_disk_free_mib
    spawn_child: object = _spawn_child
    physical_ram_reader: object = physical_ram_bytes
    verify_native: bool = False
    rss_reader: object = read_child_rss_bytes


def _stop_reason(stage):
    return {
        "swap_watchdog": "swap_growth_exceeded_1536_mib",
        "disk_watchdog": "disk_free_below_2500_mib",
        "resource_read": "resource_monitor_unavailable",
        "timeout": "timeout",
    }.get(stage, "child_failed")


def _base_report(context, baseline):
    return {
        "report_date": _today(), "engine_version": importlib.metadata.version("formualizer"),
        "physical_ram_bytes": context.physical_ram_reader(), "seed": context.seed,
        "swap_baseline_mib": baseline, "attempted_scales": [], "completed_scales": [],
        "skipped_scales": [], "results": [], "stop_reason": None,
        "topology_limitations": [
            "Synthetic formulas use a bounded three-row lookup palette, not the real inventory.",
            "Synthetic formulas use one broad dependency hub.",
            "Shared and single-cell array encodings model OOXML topology, not a real dependency graph.",
        ],
    }


def _write_report(context, report):
    context.output.parent.mkdir(parents=True, exist_ok=True)
    temporary = context.output.with_suffix(context.output.suffix + ".tmp")
    try:
        temporary.write_text(json.dumps(report, indent=2, sort_keys=True) + "\n")
        os.replace(temporary, context.output)
    except BaseException:
        temporary.unlink(missing_ok=True)
        raise
    return report


def _skip_remaining(report, scales, reason):
    report["stop_reason"] = reason
    for scale in scales:
        report["skipped_scales"].append(scale)
        report["results"].append({"scale": scale, "status": "skipped", "reason": reason})


def _skip_unreported(report, scales, reason):
    reported = {result["scale"] for result in report["results"] if "scale" in result}
    _skip_remaining(report, (scale for scale in scales if scale not in reported), reason)


def _resolve_baseline(context):
    if context.swap_baseline_mib is not None:
        return context.swap_baseline_mib, None
    try:
        return context.swap_reader(), None
    except Exception as exc:
        return None, {"failure_stage": "resource_read", "error_type": type(exc).__name__}


class BenchmarkInterrupted(KeyboardInterrupt):
    pass


def _raise_parent_signal(signum, _frame):
    raise BenchmarkInterrupted(signum)


@contextmanager
def _parent_signal_handlers():
    previous = {}
    try:
        for name in ("SIGTERM", "SIGHUP"):
            signum = getattr(signal, name, None)
            if signum is not None:
                handler = signal.getsignal(signum)
                if handler is signal.SIG_IGN:
                    continue
                previous[signum] = handler
                signal.signal(signum, _raise_parent_signal)
        yield
    finally:
        for signum, handler in previous.items():
            signal.signal(signum, handler)


def _record_parent_failure(report, state, exc):
    stage = "interrupted" if isinstance(exc, KeyboardInterrupt) else "parent_error"
    progress = _read_progress(state["progress_path"])
    result = _failed_result(stage, progress)
    maxima = getattr(state.get("child"), "_engine_scale_maxima", {})
    result.update(maxima)
    result.update({"scale": state["scale"], "error_type": type(exc).__name__})
    report["results"].append(result)
    result["platform"] = platform.platform()


def _run_scales(context, report, monitor):
    for position, scale in enumerate(context.scales):
        state = {"scale": scale, "progress_path": None, "child": None}
        child = None
        workspace = None
        try:
            reading = _resource_check(monitor, context.private_root)
            if "failure_stage" in reading:
                reason = _stop_reason(reading["failure_stage"])
                _skip_remaining(report, context.scales[position:], reason)
                break
            workspace = Path(tempfile.mkdtemp(prefix="engine-scale-", dir=context.private_root))
            progress_path = workspace / "progress.json"
            state["progress_path"] = progress_path
            child_context = ChildContext(
                scale, context.seed, workspace, progress_path, context.verify_native
            )
            report["attempted_scales"].append(scale)
            child = context.spawn_child(child_context)
            state["child"] = child
            result = monitor_owned_child(child, progress_path, monitor)
            result["scale"] = scale
            report["results"].append(result)
            if result["status"] == "completed":
                report["completed_scales"].append(scale)
            result["platform"] = platform.platform()
            _write_report(context, report)
            if result["status"] == "completed":
                continue
            reason = _stop_reason(result.get("failure_stage"))
            _skip_remaining(report, context.scales[position + 1:], reason)
            break
        except BaseException as exc:
            accounted = any(item.get("scale") == scale for item in report["results"])
            if scale in report["attempted_scales"] and not accounted:
                _record_parent_failure(report, state, exc)
            raise
        finally:
            try:
                if child is not None:
                    _stop_owned(child)
            finally:
                if workspace is not None:
                    _cleanup_workspace(workspace, context.private_root)


def run_parent(context):
    context.private_root.mkdir(parents=True, exist_ok=True)
    baseline, failure = _resolve_baseline(context)
    report = _base_report(context, baseline)
    if failure:
        _skip_remaining(report, context.scales, _stop_reason(failure["failure_stage"]))
        return _write_report(context, report)
    monitor = MonitorContext(
        timeout_seconds=context.timeout_seconds, swap_baseline_mib=baseline,
        swap_reader=context.swap_reader, disk_reader=context.disk_reader,
        rss_reader=context.rss_reader,
    )
    try:
        with _parent_signal_handlers():
            _run_scales(context, report, monitor)
    except BaseException as exc:
        reason = "interrupted" if isinstance(exc, KeyboardInterrupt) else "parent_error"
        _skip_unreported(report, context.scales, reason)
        raise
    finally:
        _write_report(context, report)
    return report


def _positive_timeout(value):
    timeout = int(value)
    if timeout <= 0:
        raise argparse.ArgumentTypeError("timeout must be positive")
    return timeout


def _arguments():
    parser = argparse.ArgumentParser(allow_abbrev=False)
    parser.add_argument("--scales", nargs="+", type=validate_scale)
    parser.add_argument("--seed", type=int, required=True)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--timeout", type=_positive_timeout, default=600)
    parser.add_argument("--swap-baseline-mib", type=float)
    parser.add_argument("--child", action="store_true")
    parser.add_argument("--workspace", type=Path)
    parser.add_argument("--progress", type=Path)
    parser.add_argument("--verify-native", action="store_true")
    return parser, parser.parse_args()


def main():
    parser, args = _arguments()
    if args.child:
        if args.workspace is None or args.progress is None or args.scales is None or len(args.scales) != 1:
            parser.error("child requires one scale, workspace, and progress path")
        context = ChildContext(
            args.scales[0], args.seed, args.workspace, args.progress, args.verify_native
        )
        raise SystemExit(run_child(context))
    if not args.scales or args.output is None:
        parser.error("parent requires scales and output")
    context = ParentContext(
        tuple(args.scales), args.seed, args.output, args.timeout,
        args.swap_baseline_mib, verify_native=args.verify_native,
    )
    print(json.dumps(run_parent(context), sort_keys=True))


if __name__ == "__main__":
    main()
