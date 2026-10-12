"""The server imports on Windows, where Unix-only modules such as `resource` do not exist."""

import subprocess
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
UNIX_ONLY = ("resource", "fcntl", "termios", "pwd", "grp")


class WindowsImports(unittest.TestCase):
    def test_server_modules_import_without_unix_only_modules(self):
        code = ("import sys\n"
                f"for name in {UNIX_ONLY!r}: sys.modules[name] = None\n"
                "import pp.runtime, pp.engine.worker, pp.engine.excel, pp.live, pp.workbook_service\n")
        res = subprocess.run([sys.executable, "-c", code], cwd=ROOT, capture_output=True, text=True, timeout=120)
        self.assertEqual(res.returncode, 0, res.stderr)


if __name__ == "__main__":
    unittest.main()
