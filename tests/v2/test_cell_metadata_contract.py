import unittest

from tests.v2.api_support import APIFixture
from tests.v2.support import QueryBoundaryDatabase


class CellMetadataContractTests(unittest.TestCase):
    def setUp(self):
        self.fixture = APIFixture()
        self.client, self.csrf, self.user = self.fixture.signup(
            "cell-metadata@example.test", "Metadata Owner"
        )
        _, self.copy = self.fixture.make_copy(self.client, self.csrf)

    def tearDown(self):
        self.fixture.close()

    def test_cell_rows_and_history_share_authoritative_input_metadata(self):
        copy_id = self.copy["id"]
        untouched_cell = self.client.get(
            f"/api/copies/{copy_id}/sheets/0/cell?r=2&c=1"
        ).json()
        untouched_row = self.client.get(
            f"/api/copies/{copy_id}/sheets/0/rows?start=2&count=1"
        ).json()["rows"]["2"]["1"]

        response = self.client.post(
            f"/api/copies/{copy_id}/ops",
            {
                "op_id": "cell-metadata",
                "action": "edit",
                "cells": [
                    {
                        "sheet": 0,
                        "r": 1,
                        "c": 1,
                        "v": 12,
                        "expected_version": 0,
                    }
                ],
            },
            {"x-csrf-token": self.csrf},
        )
        cell = self.client.get(
            f"/api/copies/{copy_id}/sheets/0/cell?r=1&c=1"
        ).json()
        row = self.client.get(
            f"/api/copies/{copy_id}/sheets/0/rows?start=1&count=1"
        ).json()["rows"]["1"]["1"]
        history = self.client.get(
            f"/api/copies/{copy_id}/history?sheet=0&r=1&c=1"
        ).json()["entries"][0]

        self.assertEqual(response.status_code, 200)
        self.assertEqual((untouched_cell["ver"], untouched_row["ver"]), (0, 0))
        self.assertTrue(all(key not in untouched_cell for key in ("by", "at")))
        expected = {
            "v": 12,
            "k": "i",
            "ver": history["revision"],
            "by": history["by"],
            "at": history["at"],
        }
        self.assertEqual({key: cell[key] for key in expected}, expected)
        self.assertEqual({key: row[key] for key in expected}, expected)

    def test_cell_metadata_uses_one_read_snapshot(self):
        from pp.workbook_service import CellQuery, WorkbookService

        services = self.fixture.app.state.services
        services.operations.apply(
            self.copy["id"],
            self.user["id"],
            {
                "op_id": "snapshot-metadata",
                "action": "edit",
                "cells": [
                    {
                        "sheet": 0,
                        "r": 1,
                        "c": 1,
                        "v": 12,
                        "expected_version": 0,
                    }
                ],
            },
        )

        def rename_actor():
            with self.fixture.database.transaction() as connection:
                connection.execute(
                    "UPDATE pp_users SET name=? WHERE id=?",
                    ("Renamed Owner", self.user["id"]),
                )

        interleaved = QueryBoundaryDatabase(
            self.fixture.database, "SELECT * FROM pp_overrides", rename_actor
        )
        query = CellQuery(self.copy["id"], self.user["id"], 0, 1, 1)

        body = WorkbookService(interleaved, self.fixture.snapshot).cell(query)

        self.assertTrue(interleaved.triggered)
        self.assertEqual(body["by"]["name"], "Metadata Owner")
        fresh = services.workbooks.cell(query)
        self.assertEqual(fresh["by"]["name"], "Renamed Owner")


if __name__ == "__main__":
    unittest.main()
