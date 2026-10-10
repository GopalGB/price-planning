"""Warm copy workers must reuse their engine and reconstruct exact snapshots."""

import json

from pathlib import Path
import tempfile
import unittest
from unittest import mock

from openpyxl import Workbook

from pp.engine.pool import CalculationRequest, EnginePool
from tests.v2.support import BackendFixture, TMP


class WarmEngineTests(unittest.TestCase):
    def test_warm_revision_does_not_reload_source_and_restores_removed_overrides(self):
        fixture = BackendFixture()
        self.addCleanup(fixture.close)
        pool = EnginePool(max_workers=1)
        self.addCleanup(pool.close)
        original = pool.apply(CalculationRequest("warm-copy", 0, fixture.source, []))
        fixture.source.rename(fixture.source.with_suffix(".parked"))
        edited = pool.apply(CalculationRequest("warm-copy", 1, fixture.source, [(0, 1, 1, 15)]))
        restored = pool.apply(CalculationRequest("warm-copy", 2, fixture.source, []))
        self.assertEqual((original["pid"], edited["pid"]), (restored["pid"], restored["pid"]))
        self.assertEqual(self.cell(edited), 30)
        self.assertEqual(self.cell(restored), 20)

    def cell(self, response):
        values = {(sheet, row, col): value for sheet, row, col, value in response["changes"]}
        return values[(2, 1, 1)]

    def test_warm_overrides_distinguish_boolean_true_from_number_one(self):
        """Typed override comparison must send a boolean after a numeric one."""
        TMP.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            source = Path(temp) / "typed-warm.xlsm"
            workbook = Workbook()
            inputs = workbook.active
            inputs.title = "Inputs"
            inputs["A1"] = 0
            formulas = workbook.create_sheet("Formulas")
            formulas["A1"] = "=Inputs!A1"
            workbook.save(source)
            pool = EnginePool(max_workers=1)
            self.addCleanup(pool.close)

            pool.apply(CalculationRequest("typed-copy", 1, source, [(0, 0, 0, 1)]))
            boolean = pool.apply(CalculationRequest("typed-copy", 2, source, [(0, 0, 0, True)]))

            self.assertIn("results", boolean, "workers must expose an exact full result snapshot")
            results = {(sheet, row, col): value for sheet, row, col, value in boolean["results"]}
            self.assertIs(results[(1, 0, 0)], True)

    def test_warm_overrides_preserve_small_numeric_change(self):
        """Parity tolerance must not suppress a real input override."""
        TMP.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            source = Path(temp) / "exact-warm.xlsm"
            workbook = Workbook()
            inputs = workbook.active
            inputs.title = "Inputs"
            inputs["A1"] = 0
            formulas = workbook.create_sheet("Formulas")
            formulas["A1"] = "=Inputs!A1"
            workbook.save(source)
            pool = EnginePool(max_workers=1)
            self.addCleanup(pool.close)

            pool.apply(CalculationRequest("exact-copy", 1, source, [(0, 0, 0, 1)]))
            changed = pool.apply(CalculationRequest(
                "exact-copy", 2, source, [(0, 0, 0, 1.0000000001)]
            ))

            results = {(sheet, row, col): value for sheet, row, col, value in changed["results"]}
            self.assertEqual(results[(1, 0, 0)], 1.0000000001)

    def test_error_to_identical_looking_text_updates_engine_and_publication(self):
        """An API error object must differ from text that renders like that object."""
        from pp.calc_service import CalculationPublication, CalculationService

        fixture = BackendFixture()
        self.addCleanup(fixture.close)
        pool = EnginePool(max_workers=1)
        self.addCleanup(pool.close)
        with fixture.database.transaction() as connection:
            connection.execute(
                "UPDATE pp_copies SET input_revision=2,calculated_revision=0 WHERE id='copy-a'"
            )
            connection.executemany(
                "INSERT INTO pp_calc_jobs(copy_id,revision,status,created_at,updated_at) "
                "VALUES('copy-a',?,'pending','2026-10-10T00:00:00Z','2026-10-10T00:00:00Z')",
                [(1,), (2,)],
            )
        service = CalculationService(fixture.database)
        error = pool.apply(CalculationRequest(
            "copy-a", 1, fixture.source, [(0, 1, 1, {"error": "#DIV/0!"})]
        ))
        self.assertTrue(service.complete(CalculationPublication("copy-a", 1, error["results"])))

        literal = pool.apply(CalculationRequest(
            "copy-a", 2, fixture.source, [(0, 1, 1, "{'error': '#DIV/0!'}")]
        ))
        self.assertTrue(service.complete(CalculationPublication("copy-a", 2, literal["results"])))

        results = {(sheet, row, col): value for sheet, row, col, value in literal["results"]}
        with fixture.database.connect() as connection:
            row = connection.execute(
                "SELECT payload_json FROM pp_events WHERE copy_id='copy-a' "
                "AND event_type='calc' AND input_revision=2"
            ).fetchone()
        changes = {(item["sheet"], item["r"], item["c"]): item["v"]
                   for item in json.loads(row["payload_json"])["changes"]}
        self.assertEqual(results[(2, 1, 1)], {"error": "#VALUE!"})
        self.assertEqual(changes[(2, 1, 1)], {"error": "#VALUE!"})

    def test_failed_apply_discards_partially_mutated_warm_engine(self):
        """A failed native write must not poison the next exact snapshot replay."""
        from pp.engine.formualizer import EngineError
        from pp.engine.worker import worker_loop

        class MutatingEngine:
            def __init__(self):
                self.value = 10

            def get(self, sheet, row, col):
                return self.value

            def apply(self, updates):
                for _, _, _, value in updates:
                    self.value = value
                if self.value == 13:
                    raise EngineError("failure after native mutation")
                return [(1, 0, 0, self.value * 2)]

            def result_snapshot(self):
                return [(1, 0, 0, self.value * 2)]

        class ScriptedConnection:
            def __init__(self):
                self.requests = iter([
                    (1, "synthetic.xlsm", [(0, 0, 0, 12)]),
                    (2, "synthetic.xlsm", [(0, 0, 0, 13)]),
                    (3, "synthetic.xlsm", [(0, 0, 0, 12)]),
                    None,
                ])
                self.responses = []

            def recv(self):
                return next(self.requests)

            def send(self, response):
                self.responses.append(response)

            def close(self):
                pass

        connection = ScriptedConnection()
        with mock.patch("pp.engine.worker.FormualizerAdapter.load",
                        side_effect=lambda path: MutatingEngine()):
            worker_loop(connection, "exact-copy")

        self.assertEqual([success for success, _ in connection.responses], [True, False, True])
        self.assertEqual(connection.responses[0][1]["results"], [(1, 0, 0, 24)])
        self.assertEqual(connection.responses[2][1]["results"], [(1, 0, 0, 24)])
