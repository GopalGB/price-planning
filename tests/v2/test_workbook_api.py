import unittest

from tests.v2.api_support import APIFixture
from tests.v2.support import QueryBoundaryDatabase
from tests.v2.test_template_importer import default_style_fixture


class WorkbookAPIContractTests(unittest.TestCase):
    def setUp(self):
        self.fixture = APIFixture()
        self.client, self.csrf, self.user = self.fixture.signup("workbook@example.test", "Workbook Owner")
        _, self.copy = self.fixture.make_copy(self.client, self.csrf)

    def tearDown(self):
        self.fixture.close()

    def test_workbook_rows_cell_and_history_shapes(self):
        """Formula text in row streams or missing API metadata must fail."""
        copy_id = self.copy["id"]
        workbook = self.client.get(f"/api/copies/{copy_id}/workbook")
        self.assertEqual(workbook.status_code, 200)
        body = workbook.json()
        self.assertEqual(set(body), {"template", "input_revision", "calculated_revision", "calculation", "sheets", "styles"})
        self.assertEqual([sheet["name"] for sheet in body["sheets"]], ["Inputs", "Lookup", "Calc"])
        self.assertEqual(set(body["sheets"][0]), {"index", "name", "hidden", "rows", "cols", "freeze",
                                                   "col_widths", "row_heights", "hidden_rows", "hidden_cols", "merges"})

        rows = self.client.get(f"/api/copies/{copy_id}/sheets/0/rows?start=1&count=2")
        self.assertEqual(rows.status_code, 200)
        input_cell = rows.json()["rows"]["1"]["1"]
        self.assertEqual(input_cell["k"], "i")
        self.assertEqual(input_cell["ver"], 0)
        self.assertNotIn("formula", input_cell)
        formula = self.client.get(f"/api/copies/{copy_id}/sheets/2/cell?r=1&c=1").json()
        self.assertEqual(formula["formula"], "=Inputs!B2*2")
        self.assertEqual(formula["k"], "f")
        self.assertEqual(self.client.get(f"/api/copies/{copy_id}/sheets/0/rows?start=0&count=501").status_code, 422)
        self.assertEqual(self.client.get(f"/api/copies/{copy_id}/history").json(), {"entries": []})

    def test_private_workbook_is_invisible_without_a_role(self):
        """Knowing a private copy id must not authorize reads."""
        outsider, _, _ = self.fixture.signup("outsider@example.test", "Outsider")
        response = outsider.get(f"/api/copies/{self.copy['id']}/workbook")
        self.assertEqual((response.status_code, response.json()["error"]), (404, "not_found"))

    def test_rows_use_one_snapshot_when_calculation_publishes_between_queries(self):
        """Reading results after newer metadata must not mislabel a mixed SQLite snapshot."""
        from pp.calc_service import CalculationPublication, CalculationService
        from pp.workbook_service import RowQuery, WorkbookService

        services = self.fixture.app.state.services
        services.operations.apply(self.copy["id"], self.user["id"], {
            "op_id": "pending-calc", "action": "edit",
            "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 12, "expected_version": 0}],
        })
        calculations = CalculationService(self.fixture.database)

        def publish():
            self.assertTrue(calculations.complete(CalculationPublication(
                self.copy["id"], 1, [(2, 1, 1, 222)],
            )))

        interleaved = QueryBoundaryDatabase(
            self.fixture.database, "SELECT * FROM pp_copies WHERE id=?", publish,
        )
        query = RowQuery(self.copy["id"], self.user["id"], 2, 1, 1)

        body = WorkbookService(interleaved, self.fixture.snapshot).rows(query)

        self.assertTrue(interleaved.triggered)
        self.assertEqual((body["calculated_revision"], body["rows"]["1"]["1"]["v"]), (0, 20))
        fresh = services.workbooks.rows(query)
        self.assertEqual((fresh["calculated_revision"], fresh["rows"]["1"]["1"]["v"]), (1, 222))


class DefaultStyleWorkbookAPIContractTests(unittest.TestCase):
    def setUp(self):
        self.fixture = APIFixture(default_style_fixture)
        self.client, self.csrf, _ = self.fixture.signup("styled@example.test", "Styled Owner")
        _, self.copy = self.fixture.make_copy(self.client, self.csrf)

    def tearDown(self):
        self.fixture.close()

    def test_blank_default_style_cell_renders_and_accepts_an_edit(self):
        copy_id = self.copy["id"]
        rows = self.client.get(f"/api/copies/{copy_id}/sheets/0/rows?start=1&count=1").json()
        response = self.client.post(f"/api/copies/{copy_id}/ops", {
            "op_id": "blank-style", "action": "edit",
            "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 12, "expected_version": 0}],
        }, {"x-csrf-token": self.csrf})

        self.assertEqual(rows["rows"]["1"]["1"], {"v": None, "s": 1, "k": "i", "ver": 0})
        self.assertEqual(response.status_code, 200)


if __name__ == "__main__":
    unittest.main()
