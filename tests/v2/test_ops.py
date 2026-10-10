import math
import unittest

from pp.errors import AppError
from tests.v2.support import BackendFixture


def operation(op_id, value, expected=0, cells=None):
    return {
        "op_id": op_id,
        "action": "edit",
        "cells": cells or [{"sheet": 0, "r": 1, "c": 1, "v": value, "expected_version": expected}],
    }


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

        bad_values = [math.nan, math.inf, [], {"other": "x"}]
        for index, value in enumerate(bad_values):
            with self.subTest(value=value):
                with self.assertRaises(AppError) as invalid:
                    self.service.apply("copy-a", "editor", operation(f"bad-{index}", value, expected=1))
                self.assertEqual(invalid.exception.code, "invalid")
        self.assertEqual(self.fixture.override()["version"], 1)

    def test_viewer_is_forbidden_and_batch_limit_is_enforced(self):
        """Authorization and bounded write checks belong inside the transaction."""
        with self.assertRaises(AppError) as caught:
            self.service.apply("copy-a", "viewer", operation("viewer", 12))
        self.assertEqual((caught.exception.status, caught.exception.code), (403, "forbidden"))
        too_many = [{"sheet": 0, "r": 1, "c": 1, "v": 1, "expected_version": 0}] * 10001
        with self.assertRaises(AppError) as caught:
            self.service.apply("copy-a", "owner", operation("large", 0, cells=too_many))
        self.assertEqual((caught.exception.status, caught.exception.code), (422, "invalid"))


if __name__ == "__main__":
    unittest.main()

