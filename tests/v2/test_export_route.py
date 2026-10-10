"""The real HTTP export route preserves the template package."""
import io
import re
import unittest
import zipfile

from openpyxl import load_workbook

from tests.v2.api_support import APIFixture
from tests.v2.test_template_importer import rewrite_member


class ExportRouteTests(unittest.TestCase):
    def test_edited_input_download_preserves_the_package(self):
        fixture = APIFixture()
        self.addCleanup(fixture.close)
        client, csrf, _ = fixture.signup("route@example.test", "Export Route")
        _, copy = fixture.make_copy(client, csrf)
        source = fixture.snapshot.source_path
        rewrite_member(source, "xl/workbook.xml", lambda data: re.sub(
            br'fullCalcOnLoad="[^"]*"', b'fullCalcOnLoad="0"', data))
        edited = client.post(f"/api/copies/{copy['id']}/ops", {
            "op_id": "route-edit", "action": "edit",
            "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 12, "expected_version": 0}],
        }, {"x-csrf-token": csrf})
        self.assertEqual(edited.status_code, 200)
        response = client.get(f"/api/copies/{copy['id']}/export.xlsm")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.headers["content-type"].lower(),
                         "application/vnd.ms-excel.sheet.macroenabled.12")
        with zipfile.ZipFile(source) as original, zipfile.ZipFile(io.BytesIO(response.body)) as exported:
            self.assertEqual(original.namelist(), exported.namelist())
            changed = {name for name in original.namelist() if original.read(name) != exported.read(name)}
            self.assertEqual(changed, {"xl/workbook.xml", "xl/worksheets/sheet1.xml"})
        workbook = load_workbook(io.BytesIO(response.body), keep_vba=True)
        self.addCleanup(workbook.close)
        self.assertEqual(workbook["Inputs"]["B2"].value, 12)
        self.assertEqual(workbook["Calc"]["B2"].value, "=Inputs!B2*2")
        self.assertTrue(workbook.calculation.fullCalcOnLoad)
