import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


class RepoDocs(unittest.TestCase):
    def test_no_em_dashes_in_docs(self):
        for path in ('README.md', 'docs/DEPLOY.md', 'docs/API.md'):
            self.assertNotIn('—', (ROOT / path).read_text(encoding='utf-8'), path)
