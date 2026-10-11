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
