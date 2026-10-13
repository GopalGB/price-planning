"""JSON travels gzipped; a workbook download is already a zip and stays untouched."""
import gzip
import json
import unittest

from ops.fixtures import make_conditional_formatting
from pp.server import GZIP_MINIMUM_BYTES
from tests.v2.api_support import APIFixture

GZIP = {"accept-encoding": "gzip"}


class GzipBoundaryTests(unittest.TestCase):
    def setUp(self):
        self.fixture = APIFixture(make_conditional_formatting)
        self.client, self.csrf, _ = self.fixture.signup("gzip@example.test", "Gzip Owner")
        _, self.copy = self.fixture.make_copy(self.client, self.csrf)
        self.rows = f"/api/copies/{self.copy['id']}/sheets/0/rows?start=0&count=256"

    def tearDown(self):
        self.fixture.close()

    def test_a_json_block_is_gzipped_when_the_client_asks(self):
        plain = self.client.get(self.rows)
        packed = self.client.get(self.rows, GZIP)
        self.assertEqual((plain.status_code, packed.status_code), (200, 200))
        self.assertGreater(len(plain.body), GZIP_MINIMUM_BYTES)
        self.assertEqual(packed.headers["content-encoding"], "gzip")
        self.assertEqual(packed.headers["vary"].lower(), "accept-encoding")
        self.assertLess(len(packed.body), len(plain.body))
        self.assertEqual(gzip.decompress(packed.body), plain.body)
        self.assertEqual(json.loads(gzip.decompress(packed.body)), plain.json())

    def test_a_client_that_does_not_ask_is_answered_plain(self):
        plain = self.client.get(self.rows)
        packed = self.client.get(self.rows, GZIP)
        self.assertNotIn("content-encoding", plain.headers)
        self.assertEqual(gzip.decompress(packed.body), plain.body)

    def test_a_small_json_body_is_left_alone(self):
        response = self.client.get("/api/me", GZIP)
        self.assertEqual(response.status_code, 200)
        self.assertLess(len(response.body), GZIP_MINIMUM_BYTES)
        self.assertNotIn("content-encoding", response.headers)

    def test_a_workbook_download_is_not_compressed(self):
        for suffix in ("export.xlsm", "export.xlsx"):
            with self.subTest(suffix=suffix):
                response = self.client.get(f"/api/copies/{self.copy['id']}/{suffix}", GZIP)
                self.assertEqual(response.status_code, 200, response.body)
                self.assertNotIn("content-encoding", response.headers)
                self.assertTrue(response.body.startswith(b"PK"), "an unzipped workbook package")


if __name__ == "__main__":
    unittest.main()