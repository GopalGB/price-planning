"""Upload limits must admit the real price-plan template."""

import unittest

from pp import upload


class RealTemplateBudgetTests(unittest.TestCase):
    def test_real_sized_workbook_fits_the_expansion_budget(self):
        """The real template unpacks to about 650 MB; 500 MB rejected it as a zip bomb."""
        self.assertGreaterEqual(upload.MAX_UNCOMPRESSED, 1024 * 1024 * 1024)
        self.assertLessEqual(upload.MAX_RATIO, 200)  # bombs are still caught by their ratio
        # ~157k styled row/column records and ~15 MB of their attributes in the real template
        self.assertGreaterEqual(upload.MAX_STRUCTURE_RECORDS, 500_000)
        self.assertGreaterEqual(upload.MAX_STRUCTURE_BYTES, 32 * 1024 * 1024)


if __name__ == "__main__":
    unittest.main()
