import hashlib
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / 'ops'))
import content_guard as cg


class CheckTests(unittest.TestCase):
    def run_check(self, files, deny=(), allow=None):
        return cg.check(list(files), lambda p: files[p], list(deny), allow or {})

    def test_spreadsheet_extension_blocked(self):
        self.assertTrue(self.run_check({'a/Plan.XLSM': b'x'}))

    def test_big_file_blocked_unless_allowlisted(self):
        big = b'0' * (cg.MAX_BYTES + 1)
        self.assertTrue(self.run_check({'vendor/wheels/x.whl': big}))
        allow = {'vendor/wheels/x.whl': hashlib.sha256(big).hexdigest()}
        self.assertEqual(self.run_check({'vendor/wheels/x.whl': big}, allow=allow), [])

    def test_allowlist_hash_mismatch_blocked(self):
        big = b'0' * (cg.MAX_BYTES + 1)
        self.assertTrue(self.run_check({'v.whl': big}, allow={'v.whl': 'deadbeef'}))

    def test_denylisted_string_in_txt_blocked(self):
        self.assertTrue(self.run_check({'notes.txt': b'Group 7 BU Recap'}, deny=['bu recap']))

    def test_clean_files_pass(self):
        self.assertEqual(self.run_check({'app/x.py': b'print(1)'}, deny=['secret label']), [])


class HookTests(unittest.TestCase):
    def test_hook_runs_without_maxhq(self):
        with tempfile.TemporaryDirectory() as d:
            repo = Path(d) / 'r'
            subprocess.run(['git', 'init', '-q', str(repo)], check=True)
            (repo / 'ops').mkdir()
            (repo / '.githooks').mkdir()
            for rel in ('ops/content_guard.py', 'ops/large_file_allowlist.json', '.githooks/pre-commit'):
                (repo / rel).write_bytes((ROOT / rel).read_bytes())
            (repo / 'bad.xlsx').write_bytes(b'x')
            subprocess.run(['git', '-C', str(repo), 'add', '-A'], check=True)
            result = subprocess.run(['sh', '.githooks/pre-commit'], cwd=repo, capture_output=True, text=True)
            self.assertEqual(result.returncode, 1, result.stdout + result.stderr)
            self.assertIn('BLOCKED bad.xlsx', result.stdout)


if __name__ == '__main__':
    unittest.main()
