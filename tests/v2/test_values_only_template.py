"""A values-only template shows Excel's saved results and refuses edits."""

import json
import unittest
from unittest import mock

from tests.v2.api_support import APIFixture


class ValuesOnlyTemplateTests(unittest.TestCase):
    def setUp(self):
        self.fixture = APIFixture()
        self.addCleanup(self.fixture.close)

    def test_import_skips_the_engine_and_is_cached_separately(self):
        from pp.template.importer import import_template
        data = self.fixture.root / "values-only"
        with mock.patch("pp.template.importer._synchronize_formula_results",
                        side_effect=AssertionError("engine must not run")):
            path = import_template(self.fixture.source, data, "cached")
        self.assertEqual(json.loads((path / "meta.json").read_text())["results"], "cached")
        with mock.patch("pp.template.importer._synchronize_formula_results") as engine:
            import_template(self.fixture.source, data, "engine")
        engine.assert_called_once()

    def test_edits_are_refused_and_the_workbook_reports_it(self):
        client, csrf, _ = self.fixture.signup("values@example.test")
        _, copy = self.fixture.make_copy(client, csrf)
        self.fixture.snapshot.results = "cached"
        meta = client.get(f"/api/copies/{copy['id']}/workbook").json()
        response = client.post(f"/api/copies/{copy['id']}/ops", {
            "op_id": "paused", "action": "edit",
            "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 12, "expected_version": 0}],
        }, {"x-csrf-token": csrf})
        self.assertEqual(meta["calculation"], "cached")
        self.assertEqual((response.status_code, response.json()["error"]),
                         (409, "calculation_paused"))


if __name__ == "__main__":
    unittest.main()
