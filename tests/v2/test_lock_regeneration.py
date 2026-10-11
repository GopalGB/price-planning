"""Vendored wheels must survive the content guard after regeneration."""

import hashlib
import json
from pathlib import Path
import runpy
import tempfile
import unittest
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[2]
TMP = ROOT / "results-private" / "tmp"


class LockRegenerationTests(unittest.TestCase):
    def test_small_wheels_are_hash_allowlisted_after_regeneration(self):
        TMP.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=TMP) as temporary:
            root = Path(temporary)
            (root / "ops").mkdir()
            script = root / "ops" / "lock_deps.py"
            script.write_bytes((ROOT / "ops" / "lock_deps.py").read_bytes())
            allow_path = root / "ops" / "large_file_allowlist.json"
            allow_path.write_text(json.dumps({"vendor/legacy.whl": "preserved"}))
            locked = "exceptiongroup==1.3.1 ; python_version < '3.11' \\\n    --hash=sha256:synthetic\n"
            for name in ("requirements.lock", "requirements-dev.lock"):
                (root / name).write_text(locked)
            calls = []

            def dependency_boundary(arguments, **kwargs):
                calls.append(arguments)
                if "download" in arguments:
                    requirement_path = Path(arguments[arguments.index("-r") + 1])
                    self.assertNotIn("python_version", requirement_path.read_text())
                    self.assertIn("--hash=sha256:synthetic", requirement_path.read_text())
                    (root / "vendor" / "wheels" / "synthetic.whl").write_bytes(b"synthetic wheel")

            with patch("subprocess.run", side_effect=dependency_boundary):
                runpy.run_path(str(script))

            allow = json.loads(allow_path.read_text())
            self.assertEqual(allow.get("vendor/wheels/synthetic.whl"),
                             hashlib.sha256(b"synthetic wheel").hexdigest())
            self.assertEqual(allow["vendor/legacy.whl"], "preserved")
            compile_calls = [call for call in calls if "compile" in call]
            for call in compile_calls:
                self.assertIn("--python-version", call)
                self.assertEqual(call[call.index("--python-version") + 1], "3.10")
