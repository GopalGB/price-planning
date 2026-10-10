import os
from pathlib import Path
import tempfile
import unittest

from ops.fixtures import make_semantic


ROOT = Path(__file__).resolve().parents[2]
TMP = ROOT / "results-private" / "tmp"


class EnginePoolTests(unittest.TestCase):
    def test_calculation_runs_in_spawned_child_for_requested_revision(self):
        """Running calculations in the ASGI process must make this fail."""
        from pp.engine.pool import EnginePool

        TMP.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            source = Path(temp) / "semantic.xlsm"
            make_semantic(source)
            pool = EnginePool(max_workers=1)
            try:
                result = pool.apply("copy-a", 7, source, [(0, 1, 1, 12)])
            finally:
                pool.close()

            self.assertNotEqual(result["pid"], os.getpid())
            self.assertEqual(result["revision"], 7)
            self.assertIn((2, 1, 1, 24), result["changes"])

    def test_lru_uses_one_spawned_process_per_active_copy(self):
        """Reusing one generic executor process for unrelated copies must fail."""
        from pp.engine.pool import EnginePool

        TMP.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            source = Path(temp) / "semantic.xlsm"
            make_semantic(source)
            pool = EnginePool(max_workers=1)
            try:
                first = pool.apply("copy-a", 1, source, [(0, 1, 1, 12)])
                second = pool.apply("copy-b", 1, source, [(0, 1, 1, 14)])
            finally:
                pool.close()

            self.assertNotEqual(first["pid"], second["pid"])
            self.assertEqual(second["copy_id"], "copy-b")

    def test_exact_revision_rebuild_does_not_reuse_latest_overrides(self):
        """Serving mutable latest engine state for an older revision must fail."""
        from pp.engine.pool import EnginePool

        TMP.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            source = Path(temp) / "semantic.xlsm"
            make_semantic(source)
            pool = EnginePool(max_workers=1)
            try:
                old = pool.apply("copy-a", 4, source, [(0, 1, 1, 12)])
                latest = pool.apply("copy-a", 5, source, [(0, 1, 1, 15)])
                replay = pool.apply("copy-a", 4, source, [(0, 1, 1, 12)])
            finally:
                pool.close()

            def calculated(result):
                return next(value for sheet, row, col, value in result["changes"]
                            if (sheet, row, col) == (2, 1, 1))

            self.assertEqual((calculated(old), calculated(latest), calculated(replay)), (24, 30, 24))
            self.assertEqual(old["pid"], latest["pid"])
            self.assertEqual(old["pid"], replay["pid"])


if __name__ == "__main__":
    unittest.main()
