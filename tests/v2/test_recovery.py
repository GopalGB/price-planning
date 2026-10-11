import json
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
        from pp.calc_service import CalculationPublication, CalculationService

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
        self.assertTrue(calculations.complete(CalculationPublication("copy-a", 2, [(2, 1, 1, 30)])))
        self.assertFalse(calculations.complete(CalculationPublication("copy-a", 1, [(2, 1, 1, 24)])))
        with self.fixture.database.connect() as connection:
            copy = connection.execute("SELECT calculated_revision FROM pp_copies WHERE id='copy-a'").fetchone()
            result = connection.execute(
                "SELECT value_json,calculated_revision FROM pp_results WHERE copy_id='copy-a' AND sheet=2 AND row=1 AND col=1"
            ).fetchone()
        self.assertEqual(copy[0], 2)
        self.assertEqual((result[0], result[1]), ("30", 2))

    def test_full_results_survive_out_of_order_evaluation_and_matching_revision(self):
        """R3 must publish a full snapshot even when its engine delta is empty after stale R1."""
        from pp.calc_service import CalculationPublication, CalculationService
        from pp.engine.pool import CalculationRequest, EnginePool

        pool = EnginePool(max_workers=1)
        self.addCleanup(pool.close)
        for revision, value, expected in ((1, 12, 0), (2, 15, 1), (3, 12, 2)):
            self.operations.apply("copy-a", "editor", {
                "op_id": f"r{revision}", "action": "edit",
                "cells": [{"sheet": 0, "r": 1, "c": 1, "v": value,
                           "expected_version": expected}],
            })

        service = CalculationService(self.fixture.database)
        second = pool.apply(CalculationRequest(
            "copy-a", 2, self.fixture.source, self.operations.revision_changes("copy-a", 2)
        ))
        self.assertIn("results", second, "workers must expose an exact full result snapshot")
        self.assertTrue(service.complete(CalculationPublication("copy-a", 2, second["results"])))
        first = pool.apply(CalculationRequest(
            "copy-a", 1, self.fixture.source, self.operations.revision_changes("copy-a", 1)
        ))
        self.assertIn("results", first, "workers must expose an exact full result snapshot")
        self.assertFalse(service.complete(CalculationPublication("copy-a", 1, first["results"])))
        third = pool.apply(CalculationRequest(
            "copy-a", 3, self.fixture.source, self.operations.revision_changes("copy-a", 3)
        ))
        self.assertIn("results", third, "workers must expose an exact full result snapshot")
        self.assertEqual(third["changes"], [])
        self.assertTrue(service.complete(CalculationPublication("copy-a", 3, third["results"])))

        with self.fixture.database.connect() as connection:
            result = connection.execute(
                "SELECT value_json,calculated_revision FROM pp_results "
                "WHERE copy_id='copy-a' AND sheet=2 AND row=1 AND col=1"
            ).fetchone()
            event = connection.execute(
                "SELECT payload_json FROM pp_events WHERE copy_id='copy-a' "
                "AND event_type='calc' AND input_revision=3"
            ).fetchone()
        self.assertEqual((result["value_json"], result["calculated_revision"]), ("24.0", 3))
        changes = {(item["sheet"], item["r"], item["c"]): item["v"]
                   for item in json.loads(event["payload_json"])["changes"]}
        self.assertEqual(changes[(2, 1, 1)], 24.0)

    def test_full_results_preserve_unchanged_constant_formula_across_edits(self):
        """Replacing stored results with a delta must not erase unchanged formulas."""
        from pp.calc_service import CalculationPublication, CalculationService
        from pp.engine.pool import CalculationRequest, EnginePool

        pool = EnginePool(max_workers=1)
        self.addCleanup(pool.close)
        service = CalculationService(self.fixture.database)
        for revision, value, expected in ((1, 12, 0), (2, 15, 1)):
            self.operations.apply("copy-a", "editor", {
                "op_id": f"constant-{revision}", "action": "edit",
                "cells": [{"sheet": 0, "r": 1, "c": 1, "v": value,
                           "expected_version": expected}],
            })
            result = pool.apply(CalculationRequest(
                "copy-a", revision, self.fixture.source,
                self.operations.revision_changes("copy-a", revision),
            ))
            self.assertIn("results", result, "workers must expose an exact full result snapshot")
            self.assertTrue(service.complete(CalculationPublication("copy-a", revision, result["results"])))
            if revision == 1:
                pool.apply(CalculationRequest("evict-copy-a", 0, self.fixture.source, []))

        with self.fixture.database.connect() as connection:
            rows = connection.execute(
                "SELECT row,value_json,calculated_revision FROM pp_results "
                "WHERE copy_id='copy-a' AND sheet=2 AND row IN (1,16) ORDER BY row"
            ).fetchall()
        self.assertEqual([(row["row"], row["calculated_revision"]) for row in rows], [(1, 2), (16, 2)])
        self.assertEqual(rows[1]["value_json"], '"x b"')

    def test_full_result_removal_publishes_a_null_delta_for_stale_spill_cells(self):
        """A shrinking spill must tell connected clients to clear removed followers."""
        from pp.calc_service import CalculationPublication, CalculationService

        with self.fixture.database.transaction() as connection:
            connection.execute(
                "UPDATE pp_copies SET input_revision=2,calculated_revision=0 WHERE id='copy-a'"
            )
        service = CalculationService(self.fixture.database)
        first = [(0, 0, 1, 10), (0, 1, 1, 11)]
        second = [(0, 0, 1, 10)]
        self.assertTrue(service.complete(CalculationPublication("copy-a", 1, first)))
        self.assertTrue(service.complete(CalculationPublication("copy-a", 2, second)))

        with self.fixture.database.connect() as connection:
            event = connection.execute(
                "SELECT payload_json FROM pp_events WHERE copy_id='copy-a' "
                "AND event_type='calc' AND input_revision=2"
            ).fetchone()
        changes = json.loads(event["payload_json"])["changes"]
        self.assertIn({"sheet": 0, "r": 1, "c": 1, "v": None}, changes)

    def test_revision_zero_job_replays_after_create_crash(self):
        """Startup recovery must finish a durable baseline job before serving the copy."""
        from pp.copies import CopyService
        from pp.engine.pool import CalculationRequest, EnginePool
        from pp.recovery import recover_unfinished

        with self.fixture.database.transaction() as connection:
            connection.execute(
                "INSERT INTO pp_categories(id,name,created_by,created_at) VALUES(?,?,?,?)",
                ("category", "Category", "owner", "2026-10-10T00:00:00Z"),
            )
            connection.execute(
                "INSERT INTO pp_category_members(category_id,user_id,role) VALUES(?,?,?)",
                ("category", "owner", "manager"),
            )
        copy = CopyService(self.fixture.database, self.fixture.snapshot.hash).create("owner", {
            "category_id": "category", "name": "Recovered", "visibility": "private",
            "default_role": "viewer",
        })
        pool = EnginePool(max_workers=1)
        self.addCleanup(pool.close)

        def calculate(copy_id, revision):
            request = CalculationRequest(
                copy_id, revision, self.fixture.source,
                self.operations.revision_changes(copy_id, revision),
            )
            return pool.apply(request)

        self.assertEqual(recover_unfinished(self.fixture.database, calculate), {"completed": 1, "failed": 0})
        with self.fixture.database.connect() as connection:
            current = connection.execute(
                "SELECT calculated_revision FROM pp_copies WHERE id=?", (copy["id"],)
            ).fetchone()
            value = connection.execute(
                "SELECT value_json FROM pp_results WHERE copy_id=? AND sheet=2 AND row=1 AND col=1",
                (copy["id"],),
            ).fetchone()
        self.assertEqual(current["calculated_revision"], 0)
        self.assertEqual(value["value_json"], "20.0")

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
            values = [(2, 1, 1, 24)]
            return {"changes": values, "results": values}

        first = recover_unfinished(self.fixture.database, calculate)
        second = recover_unfinished(self.fixture.database, calculate)

        self.assertEqual(first, {"completed": 1, "failed": 0})
        self.assertEqual(second, {"completed": 0, "failed": 0})
        self.assertEqual(calls, [("copy-a", 1)])


if __name__ == "__main__":
    unittest.main()
