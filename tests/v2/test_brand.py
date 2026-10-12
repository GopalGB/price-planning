"""Optional brand: name, colours and logo/favicon served before sign-in."""

import json
import tempfile
import unittest
from pathlib import Path

from pp.config import load_config

TMP = Path(__file__).resolve().parents[2] / "results-private" / "tmp"
TMP.mkdir(parents=True, exist_ok=True)
PNG = b"\x89PNG\r\n\x1a\n" + b"\0" * 32
BRAND = {"name": "Example Retail Co", "dir": "brand", "logo": "logo.png", "favicon": "fav.ico",
         "colors": {"primary": "#123456", "accent": "#abcdef", "highlight": "#fedcba"}}


def configured(brand):
    root = Path(tempfile.mkdtemp(dir=TMP))
    (root / "brand").mkdir()
    (root / "brand" / "logo.png").write_bytes(PNG)
    (root / "brand" / "fav.ico").write_bytes(b"\0\0\1\0")
    (root / "settings.json").write_text(json.dumps({"demo_mode": True, "brand": brand}))
    return root


class BrandConfigTests(unittest.TestCase):
    def test_valid_brand_resolves_assets(self):
        brand = load_config(configured(BRAND)).brand
        self.assertEqual(brand["name"], "Example Retail Co")
        self.assertEqual(brand["colors"]["accent"], "#abcdef")
        self.assertEqual(sorted(brand["assets"]), ["favicon", "logo"])

    def test_brand_is_optional(self):
        root = Path(tempfile.mkdtemp(dir=TMP))
        (root / "settings.json").write_text(json.dumps({"demo_mode": True}))
        self.assertIsNone(load_config(root).brand)

    def test_invalid_brands_are_rejected(self):
        bad = [{**BRAND, "logo": "../logo.png"}, {**BRAND, "logo": "logo.exe"},
               {**BRAND, "logo": "missing.png"}, {**BRAND, "colors": {"primary": "red"}},
               {**BRAND, "colors": {"other": "#123456"}}, {**BRAND, "name": ""}, "Example"]
        for brand in bad:
            with self.subTest(brand=brand), self.assertRaises(ValueError):
                load_config(configured(brand))

    def test_oversized_asset_is_rejected(self):
        root = configured(BRAND)
        (root / "brand" / "logo.png").write_bytes(PNG + b"\0" * 1_000_001)
        with self.assertRaises(ValueError):
            load_config(root)


class BrandApiTests(unittest.TestCase):
    def fixture(self, brand):
        from tests.v2.api_support import APIFixture
        fixture = APIFixture()
        self.addCleanup(fixture.close)
        fixture.config.brand = brand
        return fixture.client()

    def test_brand_and_assets_are_public_without_paths(self):
        brand = load_config(configured(BRAND)).brand
        client = self.fixture(brand)
        body = client.get("/api/brand").json()
        self.assertEqual(body, {"name": "Example Retail Co", "colors": BRAND["colors"],
                                "assets": ["favicon", "logo"]})
        logo = client.get("/brand/logo")
        self.assertEqual((logo.status_code, logo.body), (200, PNG))
        self.assertEqual(logo.headers["content-type"], "image/png")
        self.assertEqual(logo.headers["x-content-type-options"], "nosniff")
        self.assertEqual(client.get("/brand/touch_icon").status_code, 404)
        self.assertEqual(client.get("/brand/..%2Fsettings.json").status_code, 404)

    def test_unbranded_install(self):
        client = self.fixture(None)
        self.assertEqual(client.get("/api/brand").json(), {})
        self.assertEqual(client.get("/brand/logo").status_code, 404)


if __name__ == "__main__":
    unittest.main()
