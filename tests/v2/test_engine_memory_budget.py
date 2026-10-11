"""One resident engine per active copy, bounded by memory and protected while edited.

The pool keeps a loaded engine per active copy. Idle copies are dropped least
recently used first when the pool is over its memory budget, but never while a
client has that copy open, because an eviction would force a reload under an
active editor.
"""

import sys
import tempfile
import unittest
from pathlib import Path

from pp.engine.pool import EnginePool, MemoryBudget, Protection


class MemoryBudgetTests(unittest.TestCase):
    def test_a_budget_admits_engines_until_the_limit_is_reached(self):
        budget = MemoryBudget(limit_bytes=100, used_bytes=0)

        self.assertTrue(budget.admit(60))
        self.assertEqual(budget.used_bytes, 60)
        self.assertFalse(budget.admit(60))
        self.assertEqual(budget.used_bytes, 60)

    def test_releasing_an_engine_returns_its_memory_to_the_budget(self):
        budget = MemoryBudget(limit_bytes=100, used_bytes=0)
        budget.admit(70)

        budget.release(70)

        self.assertEqual(budget.used_bytes, 0)
        self.assertTrue(budget.admit(70))

    def test_re_measuring_an_engine_replaces_its_previous_size(self):
        budget = MemoryBudget(limit_bytes=100, used_bytes=0)
        budget.admit(20)

        budget.resize(20, 50)

        self.assertEqual(budget.used_bytes, 50)

    def test_a_refused_engine_is_not_charged_to_the_budget(self):
        budget = MemoryBudget(limit_bytes=100, used_bytes=0)
        budget.admit(80)

        self.assertFalse(budget.admit(40))
        self.assertEqual(budget.used_bytes, 80)

    def test_an_unbounded_budget_never_refuses(self):
        budget = MemoryBudget(limit_bytes=None, used_bytes=0)

        self.assertTrue(budget.admit(10 ** 12))
        self.assertEqual(budget.used_bytes, 10 ** 12)


class ProtectionTests(unittest.TestCase):
    def test_a_protected_copy_is_never_a_victim(self):
        protection = Protection({"copy-a"})

        self.assertFalse(protection.can_evict("copy-a"))
        self.assertTrue(protection.can_evict("copy-b"))

    def test_protection_can_be_added_and_removed_as_editors_come_and_go(self):
        protection = Protection()
        protection.protect("copy-a")

        self.assertFalse(protection.can_evict("copy-a"))

        protection.release("copy-a")

        self.assertTrue(protection.can_evict("copy-a"))


class EvictionOrderTests(unittest.TestCase):
    def test_eviction_takes_the_least_recently_used_idle_copy_first(self):
        from pp.engine.pool import _eviction_candidates

        pool = _PoolStub(order=["copy-a", "copy-b", "copy-c"],
                         sizes={"copy-a": 10, "copy-b": 10, "copy-c": 10},
                         busy={"copy-a"}, protected=set())

        victims = _eviction_candidates(pool, needed_bytes=15, limit_bytes=100)

        self.assertEqual(victims, ["copy-b", "copy-c"])

    def test_eviction_never_returns_a_protected_copy(self):
        from pp.engine.pool import _eviction_candidates

        pool = _PoolStub(order=["copy-a", "copy-b"],
                         sizes={"copy-a": 10, "copy-b": 10},
                         busy=set(), protected={"copy-a", "copy-b"})

        victims = _eviction_candidates(pool, needed_bytes=10, limit_bytes=100)

        self.assertEqual(victims, [])

    def test_a_protected_copy_is_skipped_while_an_open_one_is_still_freed(self):
        from pp.engine.pool import _eviction_candidates

        pool = _PoolStub(order=["copy-a", "copy-b"],
                         sizes={"copy-a": 10, "copy-b": 10},
                         busy=set(), protected={"copy-a"})

        victims = _eviction_candidates(pool, needed_bytes=10, limit_bytes=100)

        self.assertEqual(victims, ["copy-b"])

    def test_a_busy_copy_is_skipped_even_when_it_is_the_oldest(self):
        from pp.engine.pool import _eviction_candidates

        pool = _PoolStub(order=["copy-a", "copy-b"],
                         sizes={"copy-a": 10, "copy-b": 10},
                         busy={"copy-a"}, protected=set())

        victims = _eviction_candidates(pool, needed_bytes=10, limit_bytes=100)

        self.assertEqual(victims, ["copy-b"])

    def test_nothing_is_evicted_while_the_pool_is_within_budget(self):
        from pp.engine.pool import _eviction_candidates

        pool = _PoolStub(order=["copy-a", "copy-b"],
                         sizes={"copy-a": 10, "copy-b": 10},
                         busy=set(), protected=set())

        self.assertEqual(_eviction_candidates(pool, needed_bytes=1, limit_bytes=None), [])

    def test_an_unbounded_pool_offers_no_victims(self):
        from pp.engine.pool import _eviction_candidates

        pool = _PoolStub(order=["copy-a"], sizes={"copy-a": 10}, busy=set(), protected=set())

        self.assertEqual(_eviction_candidates(pool, needed_bytes=10**6,
                                              limit_bytes=None), [])


class _PoolStub:
    def __init__(self, order, sizes, busy, protected):
        self._order, self.sizes, self.busy, self.protected = order, sizes, busy, protected

    def order(self):
        return self._order

    def resident_bytes(self, copy_id):
        return self.sizes[copy_id]

    def is_busy(self, copy_id):
        return copy_id in self.busy

    def has_editors(self, copy_id):
        return copy_id in self.protected


class PoolBudgetWiringTests(unittest.TestCase):
    def test_the_pool_defaults_to_no_memory_budget(self):
        pool = EnginePool(max_workers=2)

        self.assertIsNone(pool.memory_budget.limit_bytes)
        pool.close()

    def test_a_configured_budget_is_enforced_in_bytes(self):
        pool = EnginePool(max_workers=2, memory_budget_mb=16)

        self.assertEqual(pool.memory_budget.limit_bytes, 16 * 1024 * 1024)
        pool.close()

    def test_protecting_a_copy_records_it_and_closing_is_safe(self):
        pool = EnginePool(max_workers=1)
        pool.protect("copy-a")

        self.assertFalse(pool.protection.can_evict("copy-a"))

        pool.release_protection("copy-a")
        pool.close()

        self.assertTrue(pool.protection.can_evict("copy-a"))


class ResidentAccountingTests(unittest.TestCase):
    def test_a_real_worker_reports_a_resident_size_the_pool_charges_for(self):
        """Without a reported size the budget cannot bound anything."""
        from ops.fixtures import make_semantic
        from pp.engine.pool import CalculationRequest, EnginePool
        from tests.v2.support import TMP

        TMP.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            source = Path(temp) / "semantic.xlsm"
            make_semantic(source)
            pool = EnginePool(max_workers=1, memory_budget_mb=512)
            try:
                pool.apply(CalculationRequest("copy-a", 1, source, [(0, 1, 1, 12)]))
                charged = pool.memory_budget.used_bytes
            finally:
                pool.close()

        self.assertGreater(charged, 0)

    def test_closing_the_pool_returns_the_charged_memory(self):
        from ops.fixtures import make_semantic
        from pp.engine.pool import CalculationRequest, EnginePool
        from tests.v2.support import TMP

        TMP.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            source = Path(temp) / "semantic.xlsm"
            make_semantic(source)
            pool = EnginePool(max_workers=1, memory_budget_mb=512)
            pool.apply(CalculationRequest("copy-a", 1, source, [(0, 1, 1, 12)]))
            pool.close()

            self.assertEqual(pool.memory_budget.used_bytes, 0)


class EngineMemoryBudgetConfigTests(unittest.TestCase):
    def test_the_default_budget_is_smaller_on_a_development_host_than_the_office(self):
        from pp.config import _default_engine_memory_mb

        self.assertEqual(_default_engine_memory_mb(), 16_000 if sys.platform == "win32" else 3_000)

    def test_the_runtime_pool_is_built_with_the_configured_budget(self):
        source = (Path(__file__).resolve().parents[2] / "pp" / "runtime.py").read_text()

        self.assertIn("EnginePool(config.max_workers, config.engine_memory_budget_mb)",
                      source)

    def test_the_config_field_is_named_for_the_memory_budget(self):
        import dataclasses

        from pp.config import Config

        names = {field.name for field in dataclasses.fields(Config)}
        self.assertIn("engine_memory_budget_mb", names)
        self.assertNotIn("engine_ram_budget_mb", names)


if __name__ == "__main__":
    unittest.main()