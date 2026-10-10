import json
import unittest

from openpyxl import Workbook
from openpyxl.styles import Protection

from tests.v2.api_support import APIFixture


def occupied_spill_input(path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Occupied"
    sheet.protection.sheet = True
    sheet["B1"] = "=SEQUENCE(2)"
    sheet["B2"] = 7
    sheet["B2"].protection = Protection(locked=False)
    workbook.save(path)


def blank_spill_input(path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Blank"
    sheet.protection.sheet = True
    sheet["A1"] = 1
    sheet["A1"].protection = Protection(locked=False)
    sheet["B1"] = "=SEQUENCE(A1)"
    sheet["B2"].protection = Protection(locked=False)
    workbook.save(path)


class SpillInputMetadataTests(unittest.TestCase):
    def event_cell(self, session, revision, coordinate=(0, 1, 1)):
        fixture, _client, _csrf, copy_id = session
        with fixture.database.connect() as connection:
            row = connection.execute(
                "SELECT payload_json FROM pp_events WHERE copy_id=? "
                "AND event_type='calc' AND input_revision=?",
                (copy_id, revision),
            ).fetchone()
        changes = json.loads(row["payload_json"])["changes"]
        return next(
            item
            for item in changes
            if (item["sheet"], item["r"], item["c"]) == coordinate
        )

    def edit(self, session, request):
        _fixture, client, csrf, copy_id = session
        op_id, coordinate, value, version = request
        sheet, row, col = coordinate
        return client.post(
            f"/api/copies/{copy_id}/ops",
            {
                "op_id": op_id,
                "action": "edit",
                "cells": [
                    {
                        "sheet": sheet,
                        "r": row,
                        "c": col,
                        "v": value,
                        "expected_version": version,
                    }
                ],
            },
            {"x-csrf-token": csrf},
        )

    def test_unchanged_spill_result_republishes_touched_input_metadata(self):
        fixture = APIFixture(occupied_spill_input)
        self.addCleanup(fixture.close)
        client, csrf, user = fixture.signup("occupied-meta@example.test", "Owner")
        _, copy = fixture.make_copy(client, csrf)
        session = (fixture, client, csrf, copy["id"])
        coordinate = (0, 1, 1)

        first = self.edit(session, ("first-clear", coordinate, None, 0))
        rows = client.get(
            f"/api/copies/{copy['id']}/sheets/0/rows?start=0&count=2"
        ).json()
        detail = client.get(
            f"/api/copies/{copy['id']}/sheets/0/cell?r=1&c=1"
        ).json()

        self.assertEqual(first.status_code, 200, first.body)
        expected = {
            "v": 2,
            "k": "f",
            "input": True,
            "input_value": None,
            "ver": 1,
        }
        for record in (
            rows["rows"]["1"]["1"],
            detail,
            self.event_cell(session, 1),
        ):
            self.assertEqual({key: record[key] for key in expected}, expected)
            self.assertEqual(record["by"], {"id": user["id"], "name": user["name"]})
            self.assertTrue(record["at"])

        second = self.edit(session, ("fresh-clear", coordinate, None, 1))
        republished = self.event_cell(session, 2)

        self.assertEqual(second.status_code, 200, second.body)
        self.assertEqual(
            {key: republished[key] for key in expected},
            {**expected, "ver": 2},
        )

    def test_untouched_blank_input_under_spill_carries_version_zero(self):
        fixture = APIFixture(blank_spill_input)
        self.addCleanup(fixture.close)
        client, csrf, _ = fixture.signup("blank-meta@example.test", "Owner")
        _, copy = fixture.make_copy(client, csrf)
        session = (fixture, client, csrf, copy["id"])

        expanded = self.edit(session, ("expand", (0, 0, 0), 2, 0))
        row = client.get(
            f"/api/copies/{copy['id']}/sheets/0/rows?start=1&count=1"
        ).json()["rows"]["1"]["1"]
        detail = client.get(
            f"/api/copies/{copy['id']}/sheets/0/cell?r=1&c=1"
        ).json()
        event = self.event_cell(session, 1)

        self.assertEqual(expanded.status_code, 200, expanded.body)
        for record in (row, detail, event):
            self.assertEqual(
                {
                    key: record[key]
                    for key in ("v", "k", "input", "input_value", "ver")
                },
                {
                    "v": 2,
                    "k": "f",
                    "input": True,
                    "input_value": None,
                    "ver": 0,
                },
            )


if __name__ == "__main__":
    unittest.main()
