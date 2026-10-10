import os
from pathlib import Path
import tempfile
import unittest
from unittest import mock

from ops.fixtures import make_semantic


ROOT = Path(__file__).resolve().parents[2]
TMP = ROOT / "results-private" / "tmp"


def die_on_request(connection, copy_id):
    connection.recv()
    os._exit(17)


class EnginePoolTests(unittest.TestCase):
    def test_calculation_runs_in_spawned_child_for_requested_revision(self):
        """Running calculations in the ASGI process must make this fail."""
        from pp.engine.pool import CalculationRequest, EnginePool

        TMP.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            source = Path(temp) / "semantic.xlsm"
            make_semantic(source)
            pool = EnginePool(max_workers=1)
            try:
                result = pool.apply(CalculationRequest("copy-a", 7, source, [(0, 1, 1, 12)]))
            finally:
                pool.close()

            self.assertNotEqual(result["pid"], os.getpid())
            self.assertEqual(result["revision"], 7)
            self.assertIn((2, 1, 1, 24), result["changes"])

    def test_lru_uses_one_spawned_process_per_active_copy(self):
        """Reusing one generic executor process for unrelated copies must fail."""
        from pp.engine.pool import CalculationRequest, EnginePool

        TMP.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            source = Path(temp) / "semantic.xlsm"
            make_semantic(source)
            pool = EnginePool(max_workers=1)
            try:
                first = pool.apply(CalculationRequest("copy-a", 1, source, [(0, 1, 1, 12)]))
                second = pool.apply(CalculationRequest("copy-b", 1, source, [(0, 1, 1, 14)]))
            finally:
                pool.close()

            self.assertNotEqual(first["pid"], second["pid"])
            self.assertEqual(second["copy_id"], "copy-b")

    def test_exact_revision_rebuild_does_not_reuse_latest_overrides(self):
        """Serving mutable latest engine state for an older revision must fail."""
        from pp.engine.pool import CalculationRequest, EnginePool

        TMP.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            source = Path(temp) / "semantic.xlsm"
            make_semantic(source)
            pool = EnginePool(max_workers=1)
            try:
                old = pool.apply(CalculationRequest("copy-a", 4, source, [(0, 1, 1, 12)]))
                latest = pool.apply(CalculationRequest("copy-a", 5, source, [(0, 1, 1, 15)]))
                replay = pool.apply(CalculationRequest("copy-a", 4, source, [(0, 1, 1, 12)]))
            finally:
                pool.close()

            def calculated(result):
                return next(value for sheet, row, col, value in result["changes"]
                            if (sheet, row, col) == (2, 1, 1))

            self.assertEqual((calculated(old), calculated(latest), calculated(replay)), (24, 30, 24))
            self.assertEqual(old["pid"], latest["pid"])
            self.assertEqual(old["pid"], replay["pid"])

    def test_dead_cached_worker_is_reaped_and_replaced(self):
        """A dead idle worker must not poison every later request for its copy."""
        from pp.engine.pool import CalculationRequest, EnginePool

        TMP.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            source = Path(temp) / "semantic.xlsm"
            make_semantic(source)
            pool = EnginePool(max_workers=1)
            self.addCleanup(pool.close)
            first = pool.apply(CalculationRequest("dead-copy", 1, source, [(0, 1, 1, 12)]))
            worker = pool._workers["dead-copy"]
            worker.process.terminate()
            worker.process.join(timeout=5)

            recovered = pool.apply(CalculationRequest(
                "dead-copy", 2, source, [(0, 1, 1, 14)]
            ))

            self.assertNotEqual(first["pid"], recovered["pid"])
            self.assertIn((2, 1, 1, 28), recovered["changes"])

    def test_worker_death_during_request_is_bounded_and_next_request_recovers(self):
        """A broken receive must evict its worker without requiring a pool restart."""
        from pp.engine.formualizer import EngineError
        from pp.engine.pool import CalculationRequest, EnginePool

        TMP.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            source = Path(temp) / "semantic.xlsm"
            make_semantic(source)
            pool = EnginePool(max_workers=1)
            self.addCleanup(pool.close)
            request = CalculationRequest("dying-copy", 3, source, [(0, 1, 1, 16)])
            with mock.patch("pp.engine.pool.worker_loop", die_on_request):
                with self.assertRaises(EngineError):
                    pool.apply(request)

            recovered = pool.apply(request)

            self.assertEqual(recovered["revision"], 3)
            self.assertIn((2, 1, 1, 32), recovered["changes"])


if __name__ == "__main__":
    unittest.main()
