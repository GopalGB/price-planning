import io
import re
import threading
import unittest
from unittest import mock
import zipfile
import xml.etree.ElementTree as ET

from tests.v2.api_support import APIFixture
from tests.v2.test_template_importer import rewrite_member


NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


class ExportContractTests(unittest.TestCase):
    def setUp(self):
        self.fixture = APIFixture()
        self.client, self.csrf, _ = self.fixture.signup("export@example.test", "Export Owner")
        _, self.copy = self.fixture.make_copy(self.client, self.csrf)

    def tearDown(self):
        self.fixture.close()

    def test_export_preserves_package_and_patches_only_inputs_plus_calc_mode(self):
        """Rebuilding the workbook or dropping package members must fail."""
        copy_id = self.copy["id"]
        rewrite_member(
            self.fixture.source, "xl/workbook.xml",
            lambda data: re.sub(br'fullCalcOnLoad="[^"]*"', b'fullCalcOnLoad="0"', data),
        )
        self.client.post(f"/api/copies/{copy_id}/ops", {
            "op_id": "export-op", "action": "edit",
            "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 12, "expected_version": 0}],
        }, {"x-csrf-token": self.csrf})

        response = self.client.get(f"/api/copies/{copy_id}/export.xlsm")

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.headers["content-type"], "application/vnd.ms-excel.sheet.macroenabled.12")
        with zipfile.ZipFile(self.fixture.source) as original, zipfile.ZipFile(io.BytesIO(response.body)) as exported:
            self.assertEqual(set(original.namelist()), set(exported.namelist()))
            changed = {name for name in original.namelist() if original.read(name) != exported.read(name)}
            self.assertEqual(changed, {"xl/workbook.xml", "xl/worksheets/sheet1.xml"})
            sheet = ET.fromstring(exported.read("xl/worksheets/sheet1.xml"))
            cell = sheet.find(f".//{NS}c[@r='B2']")
            self.assertEqual(cell.findtext(f"{NS}v"), "12")
            workbook = ET.fromstring(exported.read("xl/workbook.xml"))
            calc = workbook.find(f"{NS}calcPr")
            self.assertEqual(calc.get("fullCalcOnLoad"), "1")

    def test_calc_property_byte_patch_preserves_markup_compatibility_namespaces(self):
        """Changing fullCalcOnLoad must not reserialize workbook namespace prefixes."""
        from pp.export import _calculation_mode

        data = (
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            b'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            b'xmlns:mc="urn:mc" xmlns:x15="urn:x15" xmlns:xr="urn:xr" '
            b'mc:Ignorable="x15 xr"><sheets/>'
            b'<calcPr calcId="191029" calcMode="manual" forceFullCalc="0" fullCalcOnLoad="0"/>'
            b'<mc:AlternateContent><mc:Choice Requires="x15"><x15:feature xr:uid="42"/>'
            b'</mc:Choice></mc:AlternateContent></workbook>'
        )
        expected = data.replace(b'fullCalcOnLoad="0"', b'fullCalcOnLoad="1"')

        self.assertEqual(_calculation_mode(data), expected)

    def test_missing_calc_property_is_inserted_after_defined_names(self):
        """A new calcPr must use schema order without reserializing the workbook."""
        from pp.export import _calculation_mode

        data = (
            b'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            b'<sheets/><definedNames><definedName name="x">Sheet1!A1</definedName></definedNames>'
            b'<oleSize ref="A1"/></workbook>'
        )
        expected = data.replace(
            b'</definedNames>', b'</definedNames><calcPr fullCalcOnLoad="1"/>'
        )

        self.assertEqual(_calculation_mode(data), expected)

    def test_exported_inline_text_preserves_boundary_whitespace(self):
        """Excel must retain leading and trailing spaces in exported text inputs."""
        copy_id = self.copy["id"]
        response = self.client.post(f"/api/copies/{copy_id}/ops", {
            "op_id": "spaced-export", "action": "edit",
            "cells": [{"sheet": 0, "r": 1, "c": 1,
                       "v": "  leading and trailing  ", "expected_version": 0}],
        }, {"x-csrf-token": self.csrf})
        self.assertEqual(response.status_code, 200)

        exported = self.client.get(f"/api/copies/{copy_id}/export.xlsm")

        with zipfile.ZipFile(io.BytesIO(exported.body)) as archive:
            sheet = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
        text = sheet.find(f".//{NS}c[@r='B2']/{NS}is/{NS}t")
        self.assertEqual(text.text, "  leading and trailing  ")
        self.assertEqual(text.get("{http://www.w3.org/XML/1998/namespace}space"), "preserve")

    def test_export_archive_work_runs_off_the_asgi_event_loop(self):
        """Large archive rewriting must not block edits and WebSocket traffic."""
        from pp.export import export_copy

        threads = []

        def observed(source, overrides, destination):
            threads.append(threading.get_ident())
            return export_copy(source, overrides, destination)

        with mock.patch("pp.server.export_copy", side_effect=observed):
            response = self.client.get(f"/api/copies/{self.copy['id']}/export.xlsm")

        self.assertEqual(response.status_code, 200)
        self.assertNotEqual(threads, [threading.get_ident()])


if __name__ == "__main__":
    unittest.main()
