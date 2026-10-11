import json
from io import BytesIO
from pathlib import Path
import shutil
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET

import ops.fixtures as fixtures


ROOT = Path(__file__).resolve().parents[2]
TMP = ROOT / "results-private" / "tmp"
MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
X14 = "{http://schemas.microsoft.com/office/spreadsheetml/2009/9/main}"
XDR = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"


def rewrite_member(source, member, transform):
    replacement = source.with_name("rewritten.xlsm")
    with zipfile.ZipFile(source) as source_zip, zipfile.ZipFile(replacement, "w") as target_zip:
        for info in source_zip.infolist():
            data = source_zip.read(info.filename)
            if info.filename == member:
                data = transform(data)
            target_zip.writestr(info, data)
    shutil.move(replacement, source)


def dashboard_fixture(workspace, rows=32):
    make_dashboard = getattr(fixtures, "make_dashboard", None)
    if make_dashboard is None:
        raise AssertionError("ops.fixtures.make_dashboard is missing")
    source = workspace / "dashboard.xlsm"
    make_dashboard(source, row_count=rows)
    return source


class DashboardFixtureTests(unittest.TestCase):
    def test_fixture_has_approved_pivot_slicer_and_source_shape(self):
        TMP.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            source = dashboard_fixture(Path(temp))
            with zipfile.ZipFile(source) as archive:
                pivot = ET.fromstring(archive.read("xl/pivotTables/pivotTable1.xml"))
                cache = ET.fromstring(archive.read("xl/pivotCache/pivotCacheDefinition1.xml"))
                drawing = ET.fromstring(archive.read("xl/drawings/drawing1.xml"))
                dashboard_sheet = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
                source_sheet = ET.fromstring(archive.read("xl/worksheets/sheet2.xml"))
                slicer_caches = [
                    ET.fromstring(archive.read(f"xl/slicerCaches/slicerCache{index}.xml"))
                    for index in range(1, 5)
                ]
                self.assertNotIn("xl/pivotCache/pivotCacheRecords1.xml", archive.namelist())

            self.assertEqual(pivot.find(f"{MAIN}location").get("ref"), "A5:F14")
            self.assertEqual(
                dashboard_sheet.find(f"{MAIN}sheetFormatPr").get("defaultRowHeight"),
                "15",
            )
            self.assertEqual(len(pivot.findall(f"{MAIN}pageFields/{MAIN}pageField")), 3)
            self.assertEqual([node.get("x") for node in pivot.findall(f"{MAIN}colFields/{MAIN}field")], ["-2"])
            self.assertEqual(len(pivot.findall(f"{MAIN}dataFields/{MAIN}dataField")), 5)
            fields = cache.findall(f"{MAIN}cacheFields/{MAIN}cacheField")
            self.assertEqual(len(fields), 19)
            self.assertEqual([len(field.findall(f"{MAIN}sharedItems/*")) for field in fields[:4]], [11, 79, 770, 9])
            source_node = cache.find(f"{MAIN}cacheSource/{MAIN}worksheetSource")
            self.assertEqual((source_node.get("sheet"), source_node.get("ref")), ("SourceData", "A1:S33"))
            source_rows = source_sheet.findall(f"{MAIN}sheetData/{MAIN}row")
            self.assertEqual(len(source_rows), 33)
            self.assertEqual(len(source_rows[0].findall(f"{MAIN}c")), 19)
            self.assertEqual([
                len(item.findall(f"{X14}data/{X14}tabular/{X14}items/{X14}i"))
                for item in slicer_caches
            ], [10, 78, 769, 52])
            anchors = drawing.findall(f"{XDR}twoCellAnchor")
            self.assertEqual(len(anchors), 4)
            self.assertIsNotNone(anchors[0].find(f"{XDR}graphicFrame"))
            frames = [(
                int(anchor.findtext(f"{XDR}from/{XDR}col")),
                int(anchor.findtext(f"{XDR}from/{XDR}colOff")),
                int(anchor.findtext(f"{XDR}to/{XDR}col")),
                int(anchor.findtext(f"{XDR}to/{XDR}colOff")),
            ) for anchor in anchors]
            self.assertTrue(all(2 <= start <= end <= 7 for start, _, end, _ in frames))
            self.assertEqual((min(item[0] for item in frames), max(item[2] for item in frames)), (2, 7))
            self.assertEqual(len(set(frames)), 4)
            positions = [(start * 1_000_000 + start_offset, end * 1_000_000 + end_offset)
                         for start, start_offset, end, end_offset in frames]
            self.assertTrue(all(first[1] <= second[0]
                                for first, second in zip(positions, positions[1:])))
            self.assertTrue(all(
                anchor.findtext(f"{XDR}from/{XDR}row") == "16"
                and anchor.findtext(f"{XDR}to/{XDR}row") == "29" for anchor in anchors
            ))


class DashboardImportTests(unittest.TestCase):
    def test_related_xml_rejects_dtds_and_entities_in_utf8_and_utf16(self):
        from pp.template.dashboard import DashboardMetadataError, _part_xml

        documents = [
            '<!DOCTYPE root [<!ENTITY a "expanded">]><root>&a;</root>',
            '<!DOCTYPE root SYSTEM "https://example.test/blocked"><root/>',
        ]
        for encoding in ("utf-8", "utf-16"):
            for document in documents:
                with self.subTest(encoding=encoding, document=document):
                    buffer = BytesIO()
                    with zipfile.ZipFile(buffer, "w") as archive:
                        archive.writestr("part.xml", document.encode(encoding))
                    with zipfile.ZipFile(buffer) as archive:
                        with self.assertRaisesRegex(DashboardMetadataError, "DTD"):
                            _part_xml(archive, "part.xml")

    def import_dashboard(self, workspace):
        from pp.template.importer import import_template

        source = dashboard_fixture(workspace)
        snapshot = import_template(source, workspace / "data")
        meta = json.loads((snapshot / "meta.json").read_text(encoding="utf-8"))
        return next(sheet for sheet in meta["sheets"] if sheet["name"] == "Dashboard")

    def test_import_preserves_pivot_contract_without_cache_records(self):
        TMP.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            sheet = self.import_dashboard(Path(temp))

        self.assertEqual(len(sheet.get("pivots", [])), 1)
        self.assertEqual(sheet["default_row_height"], 15)
        pivot = sheet["pivots"][0]
        self.assertEqual(pivot["name"], "DashboardPivot")
        self.assertEqual(pivot["source"], {"sheet": "SourceData", "ref": "A1:S33"})
        self.assertEqual(pivot["location"], {
            "ref": "A5:F14", "bounds": [4, 0, 13, 5],
            "first_header_row": 1, "first_data_row": 2, "first_data_col": 1,
        })
        self.assertEqual(pivot["captions"], {
            "row": "Row Labels", "column": "Column Labels",
            "data": "Values", "grand_total": "Grand Total",
        })
        self.assertEqual(pivot["grand_totals"], {"rows": True, "columns": True})
        self.assertEqual(pivot["row_fields"], [0])
        self.assertEqual(pivot["column_fields"], [-2])
        self.assertEqual([field["field"] for field in pivot["page_fields"]], [1, 2, 3])
        self.assertEqual(pivot["page_fields"][0], {
            "field": 1, "item": 0, "hierarchy": -1, "cache_index": 0,
            "name": "Department", "caption": "Department",
        })
        self.assertEqual([field["field"] for field in pivot["data_fields"]], [5, 6, 7, 8, 9])
        self.assertEqual([field["name"] for field in pivot["data_fields"]], [
            "Sum of Units", "Sum of Revenue", "Sum of Margin", "Sum of Cost", "Sum of Price",
        ])
        self.assertEqual(pivot["data_fields"][1]["number_format"], "$#,##0.00")
        self.assertEqual([len(field["shared_items"]) for field in pivot["fields"][:4]], [11, 79, 770, 9])
        self.assertEqual(pivot["fields"][0]["shared_items"][0], {
            "kind": "string", "value": "Region 01",
        })
        self.assertEqual(pivot["fields"][0]["shared_items"][-1], {
            "kind": "missing", "value": None,
        })
        self.assertEqual(pivot["fields"][0]["subtotals"], ["sum"])
        self.assertEqual(pivot["fields"][0].get("sort"), {
            "type": "descending", "rank_by": 6, "data_source": True,
        })
        self.assertEqual(pivot["fields"][0]["items"][3], {
            "cache_index": 3, "kind": "data", "hidden": True,
        })
        self.assertEqual([item["indexes"] for item in pivot.get("row_items", [])[:3]], [[2], [0], [1]])
        self.assertEqual([item["data_index"] for item in pivot.get("column_items", [])], [1, 0, 2, 3, 4])
        self.assertEqual(pivot["style"], {
            "name": "PivotStyleLight16", "show_row_headers": True,
            "show_column_headers": True, "show_row_stripes": False,
            "show_column_stripes": False, "show_last_column": False,
        })

    def test_import_preserves_four_slicer_caches_views_and_anchors(self):
        TMP.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            sheet = self.import_dashboard(Path(temp))

        self.assertEqual([slicer["field"] for slicer in sheet.get("slicers", [])], [0, 1, 2, 4])
        self.assertEqual([len(slicer["items"]) for slicer in sheet["slicers"]], [10, 78, 769, 52])
        self.assertEqual([item["selected"] for item in sheet["slicers"][0]["items"]], [True] + [False] * 9)
        for slicer in sheet["slicers"][1:]:
            self.assertTrue(all(item["selected"] for item in slicer["items"]))
        first = sheet["slicers"][0]
        self.assertEqual(first["caption"], "Region")
        self.assertEqual(first["style"], "SlicerStyleLight1")
        self.assertEqual(first["column_count"], 1)
        self.assertEqual(first["row_height_emu"], 241300)
        self.assertEqual(first["sort_order"], "ascending")
        self.assertIs(first["custom_list_sort"], True)
        self.assertIs(first["show_missing"], True)
        self.assertEqual(first["cross_filter"], "showItemsWithDataAtTop")
        anchors = [slicer["anchor"] for slicer in sheet["slicers"]]
        self.assertEqual({anchor["from"]["row"] for anchor in anchors}, {16})
        self.assertEqual({anchor["to"]["row"] for anchor in anchors}, {29})
        self.assertEqual(min(anchor["from"]["column"] for anchor in anchors), 2)
        self.assertEqual(max(anchor["to"]["column"] for anchor in anchors), 7)
        self.assertEqual(len({json.dumps(anchor, sort_keys=True) for anchor in anchors}), 4)

    def test_omitted_slicer_view_style_uses_workbook_default(self):
        from pp.template.importer import import_template

        TMP.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            workspace = Path(temp)
            source = dashboard_fixture(workspace)

            def remove_style(data):
                root = ET.fromstring(data)
                root.find(f"{X14}slicer").attrib.pop("style")
                return ET.tostring(root, encoding="utf-8", xml_declaration=True)

            rewrite_member(source, "xl/slicers/slicer1.xml", remove_style)
            snapshot = import_template(source, workspace / "data")
            meta = json.loads((snapshot / "meta.json").read_text(encoding="utf-8"))
            sheet = next(item for item in meta["sheets"] if item["name"] == "Dashboard")

        self.assertEqual(sheet["slicers"][0]["style"], "SlicerStyleLight1")

    def assert_unsafe_member_rejected(self, member, transform):
        from pp.template.importer import TemplateImportError, import_template

        TMP.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            workspace = Path(temp)
            source = dashboard_fixture(workspace)
            rewrite_member(source, member, transform)
            with self.assertRaises(TemplateImportError):
                import_template(source, workspace / "data")

    def test_external_and_escaping_relationship_targets_are_rejected(self):
        rels = "xl/worksheets/_rels/sheet1.xml.rels"

        def external(data):
            root = ET.fromstring(data)
            relation = next(item for item in root if item.get("Id") == "rIdPivot")
            relation.set("Target", "https://example.invalid/pivot.xml")
            relation.set("TargetMode", "External")
            return ET.tostring(root, encoding="utf-8", xml_declaration=True)

        def escaping(data):
            root = ET.fromstring(data)
            relation = next(item for item in root if item.get("Id") == "rIdPivot")
            relation.set("Target", "../../../../escape.xml")
            return ET.tostring(root, encoding="utf-8", xml_declaration=True)

        for transform in (external, escaping):
            with self.subTest(transform=transform.__name__):
                self.assert_unsafe_member_rejected(rels, transform)

    def test_malformed_related_part_is_rejected(self):
        self.assert_unsafe_member_rejected(
            "xl/pivotTables/pivotTable1.xml", lambda _: b"<pivotTableDefinition",
        )

    def test_invalid_pivot_and_slicer_item_indices_are_rejected(self):
        def bad_pivot(data):
            root = ET.fromstring(data)
            root.find(f"{MAIN}pivotFields/{MAIN}pivotField/{MAIN}items/{MAIN}item").set("x", "999")
            return ET.tostring(root, encoding="utf-8", xml_declaration=True)

        def bad_slicer(data):
            root = ET.fromstring(data)
            root.find(f"{X14}data/{X14}tabular/{X14}items/{X14}i").set("x", "999")
            return ET.tostring(root, encoding="utf-8", xml_declaration=True)

        for member, transform in (
            ("xl/pivotTables/pivotTable1.xml", bad_pivot),
            ("xl/slicerCaches/slicerCache1.xml", bad_slicer),
        ):
            with self.subTest(member=member):
                self.assert_unsafe_member_rejected(member, transform)

    def test_out_of_grid_ranges_and_anchors_are_rejected(self):
        def bad_source(data):
            root = ET.fromstring(data)
            root.find(f"{MAIN}cacheSource/{MAIN}worksheetSource").set("ref", "A1:XFE33")
            return ET.tostring(root, encoding="utf-8", xml_declaration=True)

        def bad_anchor(data):
            root = ET.fromstring(data)
            root.find(f"{XDR}twoCellAnchor/{XDR}from/{XDR}col").text = "16384"
            return ET.tostring(root, encoding="utf-8", xml_declaration=True)

        for member, transform in (
            ("xl/pivotCache/pivotCacheDefinition1.xml", bad_source),
            ("xl/drawings/drawing1.xml", bad_anchor),
        ):
            with self.subTest(member=member):
                self.assert_unsafe_member_rejected(member, transform)

    def test_nonfinite_shared_number_is_rejected_before_json(self):
        def nonfinite(data):
            root = ET.fromstring(data)
            item = root.find(f"{MAIN}cacheFields/{MAIN}cacheField/{MAIN}sharedItems/{MAIN}s")
            item.tag = f"{MAIN}n"
            item.set("v", "1E400")
            return ET.tostring(root, encoding="utf-8", xml_declaration=True)

        self.assert_unsafe_member_rejected("xl/pivotCache/pivotCacheDefinition1.xml", nonfinite)


if __name__ == "__main__":
    unittest.main()
