import json
from pathlib import Path
import unittest

from openpyxl import Workbook
from openpyxl.styles import Font

from tests.v2.api_support import APIFixture


ROOT = Path(__file__).resolve().parents[2]
FONT_FIXTURE = ROOT / "web" / "test" / "fixtures" / "imported-font-metadata.json"
FONT_NAMES = ("Arial", "Aptos Display", 'Bad", serif', "Bad\nName")
FONT_FIELDS = ("font_name", "font_size", "bold", "italic", "nf")


def font_fixture(path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Fonts"
    for row, name in enumerate(FONT_NAMES, start=1):
        sheet.cell(row=row, column=1, value=7 / 12).font = Font(
            name=name, size=14 if row == 1 else 11,
            bold=row == 1, italic=row == 1,
        )
    workbook.save(path)


def font_aggregate(styles):
    by_name = {style.get("font_name"): style for style in styles}
    return {
        "styles": [
            {field: by_name[name][field] for field in FONT_FIELDS}
            for name in FONT_NAMES
        ]
    }


class ImportedFontMetadataTests(unittest.TestCase):
    def setUp(self):
        self.fixture = APIFixture(font_fixture)
        self.client, self.csrf, _ = self.fixture.signup("fonts@example.test", "Font Owner")
        _, self.copy = self.fixture.make_copy(self.client, self.csrf)

    def tearDown(self):
        self.fixture.close()

    def expected_aggregate(self):
        return json.loads(FONT_FIXTURE.read_text(encoding="utf-8"))

    def test_imported_api_fonts_match_committed_renderer_fixture(self):
        response = self.client.get(f"/api/copies/{self.copy['id']}/workbook")

        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            font_aggregate(response.json()["styles"]), self.expected_aggregate()
        )

    def test_version_four_cache_without_font_names_is_reimported(self):
        from pp.template.importer import import_template

        meta_path = self.fixture.snapshot_path / "meta.json"
        styles_path = self.fixture.snapshot_path / "styles.json"
        meta = json.loads(meta_path.read_text(encoding="utf-8"))
        meta["import_version"] = 4
        meta_path.write_text(json.dumps(meta), encoding="utf-8")
        cached_styles = json.loads(styles_path.read_text(encoding="utf-8"))
        for style in cached_styles:
            style.pop("font_name", None)
        styles_path.write_text(json.dumps(cached_styles), encoding="utf-8")
        self.assertTrue(all("font_name" not in style for style in cached_styles))

        refreshed = import_template(self.fixture.source, self.fixture.root / "data")

        current_meta = json.loads(
            (refreshed / "meta.json").read_text(encoding="utf-8")
        )
        current_styles = json.loads(
            (refreshed / "styles.json").read_text(encoding="utf-8")
        )
        self.assertEqual(refreshed, self.fixture.snapshot_path)
        self.assertGreater(current_meta["import_version"], 4)
        self.assertEqual(font_aggregate(current_styles), self.expected_aggregate())


if __name__ == "__main__":
    unittest.main()
