import json
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from ops.fixtures import TINY_PNG, make_drawings
from pp.template.importer import TemplateImportError, import_template
from tests.v2.api_support import APIFixture


class DrawingImporterTests(unittest.TestCase):
    def test_imports_internal_picture_shapes_macro_and_slicer_at_exact_anchors(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "drawings.xlsx"
            make_drawings(source)
            snapshot = import_template(source, Path(directory) / "data")
            meta = json.loads((snapshot / "meta.json").read_text())
            drawings = meta["sheets"][0]["drawings"]

            self.assertEqual((meta["sheets"][0]["rows"], meta["sheets"][0]["cols"]), (6, 8))
            self.assertEqual([item["type"] for item in drawings], [
                "picture", "textbox", "macro", "slicer",
            ])
            self.assertEqual(drawings[0]["anchor"], {
                "from": {"col": 1, "col_offset": 95250, "row": 1, "row_offset": 190500},
                "extent": {"width": 914400, "height": 457200},
            })
            self.assertEqual(drawings[1]["text"], "Review price plan\nSecond line")
            self.assertEqual(drawings[1]["anchor"]["to"], {
                "col": 4, "col_offset": 47625, "row": 4, "row_offset": 95250,
            })
            self.assertEqual(drawings[2]["macro"], "Module1.Run Report")
            self.assertEqual(drawings[3]["name"], "Region slicer")
            asset = snapshot / "assets" / drawings[0]["asset_id"]
            self.assertEqual(asset.read_bytes(), TINY_PNG)
            self.assertNotIn("External blocked", {item["name"] for item in drawings})
            self.assertNotIn("Traversal blocked", {item["name"] for item in drawings})
            self.assertNotIn("Wrong relation blocked", {item["name"] for item in drawings})

    def test_rejects_an_internal_picture_over_the_decoded_asset_cap(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "large-picture.xlsx"
            make_drawings(source, TINY_PNG + b"x" * 64)
            with mock.patch("pp.template.importer.DRAWING_ASSET_BYTES", 64, create=True):
                with self.assertRaisesRegex(TemplateImportError, "drawing asset exceeds"):
                    import_template(source, Path(directory) / "data")

    def test_rejects_cumulative_bytes_and_drawing_counts_over_budget(self):
        limits = (
            ("DRAWING_TOTAL_BYTES", 1),
            ("DRAWING_ASSET_COUNT", 0),
            ("DRAWING_COUNT", 0),
        )
        for setting, value in limits:
            with self.subTest(setting=setting), tempfile.TemporaryDirectory() as directory:
                source = Path(directory) / "budget.xlsx"
                make_drawings(source)
                with mock.patch(f"pp.template.importer.{setting}", value):
                    with self.assertRaisesRegex(TemplateImportError, "import limit"):
                        import_template(source, Path(directory) / "data")


class DrawingAPIContractTests(unittest.TestCase):
    def setUp(self):
        self.fixture = APIFixture(make_drawings)
        self.owner, self.owner_csrf, _ = self.fixture.signup(
            "drawing-owner@example.test", "Drawing Owner"
        )
        _, self.copy = self.fixture.make_copy(self.owner, self.owner_csrf)
        self.viewer, self.viewer_csrf, _ = self.fixture.signup(
            "drawing-viewer@example.test", "Drawing Viewer"
        )
        self.owner.put(
            f"/api/copies/{self.copy['id']}/shares",
            {"email": "drawing-viewer@example.test", "role": "viewer"},
            {"x-csrf-token": self.owner_csrf},
        )

    def tearDown(self):
        self.fixture.close()

    def test_workbook_metadata_and_asset_require_copy_access(self):
        copy_id = self.copy["id"]
        workbook = self.owner.get(f"/api/copies/{copy_id}/workbook").json()
        picture = workbook["sheets"][0]["drawings"][0]
        path = f"/api/copies/{copy_id}/assets/{picture['asset_id']}"

        owner_asset = self.owner.get(path)
        viewer_asset = self.viewer.get(path)
        outsider, _, _ = self.fixture.signup("drawing-outsider@example.test")
        hidden = outsider.get(path)

        self.assertEqual(owner_asset.status_code, 200)
        self.assertEqual(owner_asset.headers["content-type"], "image/png")
        self.assertEqual(owner_asset.body, TINY_PNG)
        self.assertEqual(viewer_asset.status_code, 200)
        self.assertEqual((hidden.status_code, hidden.json()["error"]), (404, "not_found"))
        self.assertEqual(self.owner.get(f"/api/copies/{copy_id}/assets/missing.png").status_code, 404)

    def test_macro_action_is_authorized_but_explicitly_unavailable(self):
        copy_id = self.copy["id"]
        action = "Module1.Run Report"
        path = f"/api/copies/{copy_id}/actions/{action}"
        outsider, outsider_csrf, _ = self.fixture.signup("action-outsider@example.test")

        self.assertEqual(self.owner.post(path).status_code, 403)
        viewer = self.viewer.post(path, headers={"x-csrf-token": self.viewer_csrf})
        hidden = outsider.post(path, headers={"x-csrf-token": outsider_csrf})
        unknown = self.owner.post(
            f"/api/copies/{copy_id}/actions/Unknown",
            headers={"x-csrf-token": self.owner_csrf},
        )
        accepted = self.owner.post(path, headers={"x-csrf-token": self.owner_csrf})

        self.assertEqual((viewer.status_code, viewer.json()["error"]), (403, "forbidden"))
        self.assertEqual((hidden.status_code, hidden.json()["error"]), (404, "not_found"))
        self.assertEqual((unknown.status_code, unknown.json()["error"]), (404, "not_found"))
        self.assertEqual((accepted.status_code, accepted.json()["error"]), (501, "action_unavailable"))
        self.assertIn("not available yet", accepted.json()["message"])


if __name__ == "__main__":
    unittest.main()
