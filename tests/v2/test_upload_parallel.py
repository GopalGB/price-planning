"""Big uploads are read one sheet per worker process; results and safety checks must match the in-process read."""

import unittest
from unittest import mock

from openpyxl import load_workbook

from tests.v2.api_support import APIFixture
from tests.v2.test_transfer_upload import _rewrite_zip, _without_provenance


class ParallelUploadTests(unittest.TestCase):
    def setUp(self):
        self.fixture = APIFixture()
        self.addCleanup(self.fixture.close)
        self.owner, self.csrf, _ = self.fixture.signup("parallel@example.test", "Parallel")
        _, self.copy = self.fixture.make_copy(self.owner, self.csrf)
        exported = self.owner.get(f"/api/copies/{self.copy['id']}/export.xlsx")
        self.assertEqual(exported.status_code, 200, exported.body)
        source, changed = self.fixture.root / "base.xlsx", self.fixture.root / "changed.xlsx"
        source.write_bytes(exported.body)
        workbook = load_workbook(source)
        workbook["Inputs"]["B2"] = 4321  # a typed input
        workbook["Calc"]["H40"] = "typed on an open sheet"
        workbook.save(changed)
        workbook.close()
        self.changed = changed.read_bytes()

    def preview(self, data):
        return self.owner.raw(
            "POST", f"/api/copies/{self.copy['id']}/uploads/preview", data,
            {"x-csrf-token": self.csrf, "x-upload-filename": "changed.xlsx",
             "content-type": "application/octet-stream"},
        )

    def both(self, data):
        import pp.upload
        sequential = self.preview(data)
        pool = mock.Mock(wraps=pp.upload.ProcessPoolExecutor)
        with mock.patch("pp.upload.PARALLEL_MIN_BYTES", 0), \
                mock.patch("pp.upload._scan_workers", return_value=2), \
                mock.patch("pp.upload.ProcessPoolExecutor", pool):
            parallel = self.preview(data)
        self.assertTrue(pool.called, "the parallel read never started worker processes")
        return sequential, parallel

    def test_parallel_read_matches_the_in_process_read(self):
        for data in (self.changed, _without_provenance(self.changed, self.fixture.root)):
            sequential, parallel = self.both(data)
            self.assertEqual(sequential.status_code, 200, sequential.body)
            self.assertEqual(parallel.status_code, 200, parallel.body)
            keys = ("changes", "conflicts", "ignored", "identification")
            self.assertEqual({key: parallel.json()[key] for key in keys},
                             {key: sequential.json()[key] for key in keys})
            self.assertTrue(sequential.json()["changes"])

    def test_a_worker_screens_its_sheet_before_reading_it(self):
        def poison(name, content):
            if name == "xl/worksheets/sheet1.xml":
                return content.replace(b"<worksheet", b'<!DOCTYPE x [<!ENTITY e "boom">]><worksheet', 1)
            return content

        sequential, parallel = self.both(_rewrite_zip(self.changed, self.fixture.root, poison))
        self.assertEqual((sequential.status_code, parallel.status_code), (422, 422), parallel.body)

    def test_an_unreferenced_sheet_part_is_still_screened(self):
        stray = ("xl/worksheets/sheet99.xml", b'<!DOCTYPE x [<!ENTITY e "boom">]><worksheet/>')
        sequential, parallel = self.both(_rewrite_zip(self.changed, self.fixture.root, additions=(stray,)))
        self.assertEqual((sequential.status_code, parallel.status_code), (422, 422), parallel.body)

    def test_a_worker_checks_the_links_its_sheet_points_at(self):
        def dangling(name, content):
            if name == "xl/worksheets/sheet1.xml":
                return content.replace(b"<sheetData", b'<sheetData xmlns:q="http://schemas.openxmlformats.org/'
                                       b'officeDocument/2006/relationships" q:id="rId999"', 1)
            return content

        sequential, parallel = self.both(_rewrite_zip(self.changed, self.fixture.root, dangling))
        self.assertEqual((sequential.status_code, parallel.status_code), (422, 422), parallel.body)

    def test_a_dangling_link_in_a_utf16_part_is_still_found(self):
        part = '<item xmlns:q="http://schemas.openxmlformats.org/officeDocument/2006/relationships" q:id="rId9"/>'
        added = (("customXml/item9.xml", part.encode("utf-16")),)
        response = self.preview(_rewrite_zip(self.changed, self.fixture.root, additions=added))
        self.assertEqual(response.status_code, 422, response.body)

    def test_a_dead_reader_process_falls_back_to_an_in_process_read(self):
        from concurrent.futures import Future
        from concurrent.futures.process import BrokenProcessPool

        class DeadPool:
            def __init__(self, *args, **kwargs):
                pass

            def __enter__(self):
                return self

            def __exit__(self, *exc):
                return False

            def submit(self, *args):
                future = Future()
                future.set_exception(BrokenProcessPool("reader died"))
                return future

        sequential = self.preview(self.changed)
        with mock.patch("pp.upload.PARALLEL_MIN_BYTES", 0), \
                mock.patch("pp.upload._scan_workers", return_value=2), \
                mock.patch("pp.upload.ProcessPoolExecutor", DeadPool):
            fallback = self.preview(self.changed)
        self.assertEqual(fallback.status_code, 200, fallback.body)
        self.assertEqual(fallback.json()["changes"], sequential.json()["changes"])


if __name__ == "__main__":
    unittest.main()
