"""The documented engine names must load and calculate a real fixture."""

import unittest

import pp.engine as engines
from tests.v2.support import BackendFixture


class PublicEngineContractTests(unittest.TestCase):
    def test_named_engine_load_apply_and_get_contract(self):
        engine_class = getattr(engines, "FormualizerEngine", None)
        self.assertIsNotNone(engine_class, "FormualizerEngine must be publicly available")
        fixture = BackendFixture()
        self.addCleanup(fixture.close)
        engine = engine_class.load(fixture.source)
        engine.apply([(0, 1, 1, 13)])
        self.assertEqual(engine.get(2, 1, 1), 26)
        self.assertIsInstance(engine, engines.Engine)
