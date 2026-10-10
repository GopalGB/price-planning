import unittest

from tests.v2.support import BackendFixture


class RecoveryContractTests(unittest.TestCase):
    def setUp(self):
        self.fixture = BackendFixture()
        from pp.ops import OperationService
        self.operations = OperationService(self.fixture.database, self.fixture.snapshot)

    def tearDown(self):
        self.fixture.close()

    def test_revision_snapshots_are_exact_and_calc_publication_is_monotonic(self):
        """A delayed older result must not replace revision N plus 1."""
        from pp.calc_service import CalculationService

        self.operations.apply("copy-a", "editor", {
            "op_id": "r1", "action": "edit",
            "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 12, "expected_version": 0}],
        })
        self.operations.apply("copy-a", "editor", {
            "op_id": "r2", "action": "edit",
            "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 15, "expected_version": 1}],
        })
        self.assertEqual(self.operations.revision_changes("copy-a", 1), [(0, 1, 1, 12)])
        self.assertEqual(self.operations.revision_changes("copy-a", 2), [(0, 1, 1, 15)])

        calculations = CalculationService(self.fixture.database)
        self.assertTrue(calculations.complete("copy-a", 2, [(2, 1, 1, 30)]))
        self.assertFalse(calculations.complete("copy-a", 1, [(2, 1, 1, 24)]))
        with self.fixture.database.connect() as connection:
            copy = connection.execute("SELECT calculated_revision FROM pp_copies WHERE id='copy-a'").fetchone()
            result = connection.execute(
                "SELECT value_json,calculated_revision FROM pp_results WHERE copy_id='copy-a' AND sheet=2 AND row=1 AND col=1"
            ).fetchone()
        self.assertEqual(copy[0], 2)
        self.assertEqual((result[0], result[1]), ("30", 2))

    def test_recovery_returns_and_finishes_only_unfinished_jobs(self):
        """Startup recovery must be idempotent and leave failures unfinished."""
        from pp.recovery import recover_unfinished

        self.operations.apply("copy-a", "editor", {
            "op_id": "pending", "action": "edit",
            "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 12, "expected_version": 0}],
        })
        calls = []

        def calculate(copy_id, revision):
            calls.append((copy_id, revision))
            return [(2, 1, 1, 24)]

        first = recover_unfinished(self.fixture.database, calculate)
        second = recover_unfinished(self.fixture.database, calculate)

        self.assertEqual(first, {"completed": 1, "failed": 0})
        self.assertEqual(second, {"completed": 0, "failed": 0})
        self.assertEqual(calls, [("copy-a", 1)])


if __name__ == "__main__":
    unittest.main()
