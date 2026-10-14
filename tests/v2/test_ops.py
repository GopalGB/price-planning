import math
from contextlib import contextmanager
import unittest

from pp.errors import AppError
from tests.v2.support import BackendFixture


def operation(op_id, value, expected=0, cells=None):
    return {
        "op_id": op_id,
        "action": "edit",
        "cells": cells or [{"sheet": 0, "r": 1, "c": 1, "v": value, "expected_version": expected}],
    }


class RacingConnection:
    def __init__(self, connection, callback):
        self.connection = connection
        self.callback = callback
        self.triggered = False

    def execute(self, sql, parameters=()):
        if "FROM pp_history h JOIN" in sql and not self.triggered:
            self.triggered = True
            self.callback()
        return self.connection.execute(sql, parameters)


class RacingDatabase:
    def __init__(self, database, callback):
        self.database = database
        self.callback = callback

    @contextmanager
    def connect(self):
        with self.database.connect() as connection:
            yield RacingConnection(connection, self.callback)


class OperationContractTests(unittest.TestCase):
    def setUp(self):
        self.fixture = BackendFixture()
        from pp.ops import OperationService
        self.service = OperationService(self.fixture.database, self.fixture.snapshot)

    def tearDown(self):
        self.fixture.close()

    def test_identical_retry_returns_stored_ack_and_changed_reuse_is_rejected(self):
        """Reexecuting or rebinding an op_id must fail this test."""
        payload = operation("op-1", 12)

        first = self.service.apply("copy-a", "editor", payload)
        retry = self.service.apply("copy-a", "editor", payload)

        self.assertEqual(first, {"op_id": "op-1", "input_revision": 1, "accepted": 1})
        self.assertEqual(retry, first)
        self.assertEqual(self.fixture.count("pp_history"), 1)
        self.assertEqual(self.fixture.count("pp_calc_jobs"), 1)
        for actor, changed in (("editor", operation("op-1", 13)),
                               ("owner", operation("op-1", 12))):
            with self.subTest(actor=actor):
                with self.assertRaises(AppError) as caught:
                    self.service.apply("copy-a", actor, changed)
                self.assertEqual((caught.exception.status, caught.exception.code), (409, "op_id_reused"))
        self.assertEqual(self.fixture.override()["value"], 12)
        self.assertEqual(self.fixture.copy()["input_revision"], 1)

    def test_locked_formula_and_formula_like_text_are_rejected_atomically(self):
        """Accepting formulas or formula-text injection into inputs must fail."""
        locked = operation("locked", 5, cells=[
            {"sheet": 2, "r": 1, "c": 1, "v": 5, "expected_version": 0}
        ])
        with self.assertRaises(AppError) as caught:
            self.service.apply("copy-a", "editor", locked)
        self.assertEqual((caught.exception.status, caught.exception.code), (422, "locked"))

        for prefix in "=+-@":
            with self.subTest(prefix=prefix):
                with self.assertRaises(AppError) as caught:
                    self.service.apply("copy-a", "editor", operation(f"text-{prefix}", prefix + "bad"))
                self.assertEqual((caught.exception.status, caught.exception.code), (422, "invalid"))
        self.assertEqual(self.fixture.copy()["input_revision"], 0)
        self.assertIsNone(self.fixture.override())

    def test_mixed_stale_batch_and_malformed_values_leave_no_partial_write(self):
        """Valid cells in an invalid batch must never be committed."""
        self.service.apply("copy-a", "editor", operation("seed", 12))
        batch = operation("stale", 0, cells=[
            {"sheet": 0, "r": 2, "c": 1, "v": 30, "expected_version": 0},
            {"sheet": 0, "r": 1, "c": 1, "v": 11, "expected_version": 0},
        ])
        with self.assertRaises(AppError) as caught:
            self.service.apply("copy-a", "editor", batch)
        self.assertEqual((caught.exception.status, caught.exception.code), (409, "conflict"))
        self.assertIsNone(self.fixture.override(row=2))
        self.assertEqual(self.fixture.copy()["input_revision"], 1)

        bad_values = [math.nan, math.inf, 2 ** 63, -(2 ** 63) - 1, [], {"other": "x"}]
        for index, value in enumerate(bad_values):
            with self.subTest(value=value):
                with self.assertRaises(AppError) as invalid:
                    self.service.apply("copy-a", "editor", operation(f"bad-{index}", value, expected=1))
                self.assertEqual(invalid.exception.code, "invalid")
        self.assertEqual(self.fixture.override()["version"], 1)

    def test_unknown_error_token_rejects_the_entire_operation(self):
        """An unsupported API error must be rejected before any cell is committed."""
        payload = operation("bad-error", 0, cells=[
            {"sheet": 0, "r": 1, "c": 1, "v": 12, "expected_version": 0},
            {"sheet": 0, "r": 2, "c": 1, "v": {"error": "#BOGUS!"}, "expected_version": 0},
        ])

        with self.assertRaises(AppError) as caught:
            self.service.apply("copy-a", "editor", payload)

        self.assertEqual((caught.exception.status, caught.exception.code), (422, "invalid"))
        self.assertEqual(self.fixture.copy()["input_revision"], 0)
        self.assertIsNone(self.fixture.override(row=1))
        self.assertIsNone(self.fixture.override(row=2))
        self.assertEqual(self.fixture.count("pp_history"), 0)
        self.assertEqual(self.fixture.count("pp_calc_jobs"), 0)

    def test_viewer_is_forbidden_and_batch_limit_is_enforced(self):
        """Authorization and bounded write checks belong inside the transaction."""
        with self.assertRaises(AppError) as caught:
            self.service.apply("copy-a", "viewer", operation("viewer", 12))
        self.assertEqual((caught.exception.status, caught.exception.code), (403, "forbidden"))
        too_many = [{"sheet": 0, "r": 1, "c": 1, "v": 1, "expected_version": 0}] * 10001
        with self.assertRaises(AppError) as caught:
            self.service.apply("copy-a", "owner", operation("large", 0, cells=too_many))
        self.assertEqual((caught.exception.status, caught.exception.code), (422, "invalid"))

    def test_numeric_validation_bounds_and_coordinate_bounds_are_enforced(self):
        """Parsing validation metadata without enforcing its operator must fail."""
        self.fixture.snapshot.sheets[0]["validations"] = [{
            "type": "whole", "operator": "between", "min": 5, "max": 15,
            "ranges": [[1, 1, 1, 1]], "showErrorMessage": True, "errorStyle": "stop",
        }]
        with self.assertRaises(AppError) as caught:
            self.service.apply("copy-a", "editor", operation("outside", 16))
        self.assertEqual((caught.exception.status, caught.exception.code), (422, "invalid"))
        with self.assertRaises(AppError) as caught:
            self.service.apply("copy-a", "editor", operation("outside-sheet", 10, cells=[
                {"sheet": 99, "r": 0, "c": 0, "v": 10, "expected_version": 0}
            ]))
        self.assertEqual((caught.exception.status, caught.exception.code), (422, "invalid"))
        self.assertEqual(self.fixture.copy()["input_revision"], 0)

    def test_static_date_validation_enforces_numeric_bounds(self):
        """Date validation must reject bad serials and accept an in-range serial."""
        self.fixture.snapshot.sheets[0]["validations"] = [{
            "type": "date", "operator": "between", "min": 5, "max": 15,
            "ranges": [[1, 1, 1, 1]], "showErrorMessage": True, "errorStyle": "stop",
        }]
        with self.assertRaises(AppError) as caught:
            self.service.apply("copy-a", "editor", operation("date-bad", 16))
        self.assertEqual((caught.exception.status, caught.exception.code), (422, "invalid"))
        self.assertEqual(self.service.apply("copy-a", "editor", operation("date-good", 10))["accepted"], 1)

    def test_static_time_validation_enforces_numeric_bounds(self):
        """Time validation must reject bad serials and accept an in-range serial."""
        self.fixture.snapshot.sheets[0]["validations"] = [{
            "type": "time", "operator": "between", "min": 0, "max": 1,
            "ranges": [[1, 1, 1, 1]], "showErrorMessage": True, "errorStyle": "stop",
        }]
        with self.assertRaises(AppError) as caught:
            self.service.apply("copy-a", "editor", operation("time-bad", 2))
        self.assertEqual((caught.exception.status, caught.exception.code), (422, "invalid"))
        self.assertEqual(self.service.apply("copy-a", "editor", operation("time-good", 0.5))["accepted"], 1)

    def test_text_length_validation_enforces_static_bounds(self):
        """Text-length validation must not silently accept oversized text."""
        self.fixture.snapshot.sheets[0]["validations"] = [{
            "type": "textLength", "operator": "lessThanOrEqual", "min": 3,
            "ranges": [[1, 1, 1, 1]], "showErrorMessage": True, "errorStyle": "stop",
        }]
        with self.assertRaises(AppError) as caught:
            self.service.apply("copy-a", "editor", operation("length-bad", "four"))
        self.assertEqual((caught.exception.status, caught.exception.code), (422, "invalid"))
        self.assertEqual(self.service.apply("copy-a", "editor", operation("length-good", "abc"))["accepted"], 1)

    def test_unevaluated_validations_accept_the_value(self):
        """Custom, range-list and formula-bound rules the app cannot evaluate must not block typing."""
        rules = [
            {"type": "list", "formula1": "'Lookup'!$A$1:$A$3"},
            {"type": "custom", "formula1": "=B2>0"},
            {"type": "decimal", "operator": "between", "min": "=Lookup!A1", "max": 5},
        ]
        for index, rule in enumerate(rules):
            with self.subTest(rule=rule["type"]):
                rule["ranges"] = [[1, 1, 1, 1]]
                rule["showErrorMessage"] = True
                rule["errorStyle"] = "stop"
                self.fixture.snapshot.sheets[0]["validations"] = [rule]
                result = self.service.apply("copy-a", "editor", operation(f"unevaluated-{index}", 10 + index, index))
                self.assertEqual(result["accepted"], 1)
                self.assertEqual(self.fixture.override()["value"], 10 + index)

    def test_whole_validation_accepts_integral_float(self):
        """A mathematically integral finite float is a valid whole number."""
        self.fixture.snapshot.sheets[0]["validations"] = [{
            "type": "whole", "operator": "between", "min": 5, "max": 15,
            "ranges": [[1, 1, 1, 1]], "showErrorMessage": True, "errorStyle": "stop",
        }]
        result = self.service.apply("copy-a", "editor", operation("whole-float", 10.0))
        self.assertEqual(result["accepted"], 1)
        self.assertEqual(self.fixture.override()["value"], 10.0)

    def test_only_stop_style_enabled_validations_reject_edits(self):
        """Warning, information, disabled, and unconstrained rules must allow edits."""
        allowed = [
            {"type": None, "showErrorMessage": True, "errorStyle": "stop"},
            {"type": "none", "showErrorMessage": True, "errorStyle": "stop"},
            {"type": "whole", "showErrorMessage": False, "errorStyle": "stop"},
            {"type": "whole", "showErrorMessage": True, "errorStyle": "warning"},
            {"type": "whole", "showErrorMessage": True, "errorStyle": "information"},
        ]
        version = 0
        for index, rule in enumerate(allowed):
            with self.subTest(rule=rule):
                rule.update({"operator": "between", "min": 1, "max": 5,
                             "ranges": [[1, 1, 1, 1]]})
                self.fixture.snapshot.sheets[0]["validations"] = [rule]
                result = self.service.apply(
                    "copy-a", "editor", operation(f"allowed-{index}", 10, expected=version)
                )
                self.assertEqual(result["accepted"], 1)
                version += 1
        for style in (None, "stop"):
            with self.subTest(blocking_style=style):
                rule = {"type": "whole", "operator": "between", "min": 1, "max": 5,
                        "ranges": [[1, 1, 1, 1]], "showErrorMessage": True}
                if style is not None:
                    rule["errorStyle"] = style
                self.fixture.snapshot.sheets[0]["validations"] = [rule]
                with self.assertRaises(AppError):
                    self.service.apply(
                        "copy-a", "editor", operation(f"blocked-{style}", 10, expected=version)
                    )

    def test_revision_reconstruction_uses_one_read_snapshot_during_newer_commit(self):
        """A concurrent R2 commit must not leak its current override into R1."""
        from pp.ops import OperationService

        self.service.apply("copy-a", "editor", operation("snapshot-r1", 12))

        def commit_r2():
            self.service.apply("copy-a", "editor", operation("snapshot-r2", 15, expected=1))

        racing = OperationService(
            RacingDatabase(self.fixture.database, commit_r2), self.fixture.snapshot
        )
        changes = racing.revision_changes("copy-a", 1)

        values = {(sheet, row, col): value for sheet, row, col, value in changes}
        self.assertEqual(values[(0, 1, 1)], 12)
        self.assertEqual(self.fixture.copy()["input_revision"], 2)


if __name__ == "__main__":
    unittest.main()
