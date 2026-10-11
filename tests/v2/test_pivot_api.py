import unittest

from ops.fixtures import make_dashboard
from tests.v2.api_support import APIFixture


def api_dashboard(path):
    make_dashboard(path, row_count=120)


class PivotAPIContractTests(unittest.TestCase):
    def setUp(self):
        self.fixture = APIFixture(api_dashboard)
        self.client, self.csrf, _ = self.fixture.signup("pivot-owner@example.test", "Pivot Owner")
        _, self.copy = self.fixture.make_copy(self.client, self.csrf)
        self.path = f"/api/copies/{self.copy['id']}/sheets/0/pivots/DashboardPivot/refresh"

    def tearDown(self):
        self.fixture.close()

    def test_refresh_contract_is_authorized_locked_and_coordinate_complete(self):
        self.assertEqual(self.fixture.app.state.services.pivots.cache_info()["sources"], 1)
        response = self.client.post(self.path, {})

        self.assertEqual(response.status_code, 200, response.body)
        body = response.json()
        self.assertEqual(set(body), {"pivot", "filters", "columns", "rows", "cells", "slicers"})
        self.assertEqual(body["pivot"]["location"]["ref"], "A5:F14")
        self.assertEqual(body["pivot"]["style"]["name"], "PivotStyleLight16")
        self.assertTrue(body["cells"])
        self.assertTrue(all(cell["locked"] for cell in body["cells"]))
        self.assertTrue(all({"r", "c", "v", "number_format", "style", "row_kind"} <= set(cell)
                            for cell in body["cells"]))
        self.assertEqual(body["rows"][-1]["caption"], "Grand Total")
        self.assertEqual(len(body["slicers"]), 4)

        stranger, _, _ = self.fixture.signup("pivot-stranger@example.test", "Stranger")
        self.assertEqual(stranger.post(self.path, {}).status_code, 404)

    def test_pending_calculation_returns_named_retryable_response(self):
        with self.fixture.database.transaction() as connection:
            connection.execute(
                "UPDATE pp_copies SET input_revision=1,calculated_revision=0 WHERE id=?", (self.copy["id"],),
            )

        response = self.client.post(self.path, {})

        self.assertEqual(response.status_code, 409)
        self.assertEqual(response.json(), {
            "error": "calculation_pending", "message": "Dashboard calculations are pending",
            "input_revision": 1, "calculated_revision": 0,
        })

    def test_refresh_reports_a_named_conflict_before_overwriting_an_input(self):
        pivot = self.fixture.snapshot.sheets[0]["pivots"][0]
        slicers = self.fixture.snapshot.sheets[0]["slicers"]
        block = self.fixture.snapshot.block(0, 0)
        block.setdefault("15", {})["0"] = {"v": "Keep me", "s": 0, "k": "i"}
        self.fixture.snapshot.input_mask = frozenset(
            {*self.fixture.snapshot.input_mask, "0:15:0"}
        )
        options = {
            "page": [{"field": item["field"], "items": None}
                     for item in pivot["page_fields"]],
            "slicers": [{"name": item["name"], "field": item["field"], "items": None}
                        for item in slicers],
            "collapsed": [],
        }

        response = self.client.post(self.path, options)

        self.assertEqual(response.status_code, 409, response.body)
        self.assertEqual(response.json()["error"], "pivot_overlap")
        self.assertEqual(response.json()["cells"], [{"sheet": 0, "r": 15, "c": 0}])

    def test_refresh_rejects_malformed_and_unimported_filter_shapes(self):
        malformed = [
            {"page": [{"field": [], "items": None}]},
            {"page": [{"field": True, "items": None}]},
            {"page": [{"field": 5, "items": None}]},
            {"page": [{"field": 1, "items": [
                {"kind": [], "value": "x"}
            ]}]},
            {"page": [{"field": 1, "items": [
                {"kind": "number", "value": 10**400}
            ]}]},
            {"slicers": [{"name": [], "field": 0, "items": None}]},
            {"slicers": [{"name": "Unknown", "field": 0, "items": None}]},
            {"collapsed": [{"field": {}, "item": 0}]},
        ]

        for index, body in enumerate(malformed):
            with self.subTest(index=index):
                response = self.client.post(self.path, body)
                self.assertEqual(response.status_code, 422, response.body)
                self.assertEqual(response.json()["error"], "invalid")

    def test_real_input_edit_changes_formula_backed_pivot_source_value(self):
        before = self.client.post(self.path, {}).json()
        before_price = next(cell["v"] for cell in before["cells"]
                            if cell["row_kind"] == "grand_total" and cell.get("data_field") == 9)

        edit = self.client.post(f"/api/copies/{self.copy['id']}/ops", {
            "op_id": "pivot-source-formula", "action": "edit",
            "cells": [{"sheet": 0, "r": 0, "c": 7, "v": 5, "expected_version": 0}],
        }, {"x-csrf-token": self.csrf})
        after = self.client.post(self.path, {}).json()
        after_price = next(cell["v"] for cell in after["cells"]
                           if cell["row_kind"] == "grand_total" and cell.get("data_field") == 9)

        self.assertEqual(edit.status_code, 200, edit.body)
        self.assertEqual(after["pivot"]["input_revision"], 1)
        self.assertEqual(after["pivot"]["calculated_revision"], 1)
        self.assertEqual(after_price - before_price, 6)


if __name__ == "__main__":
    unittest.main()
