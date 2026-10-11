"""The runtime lock must support importing the engine and importer."""

import importlib.abc
import importlib.metadata
from pathlib import Path
import re
import subprocess
import sys
import unittest


ROOT = Path(__file__).resolve().parents[2]


def check_runtime_imports():
    names = re.findall(r"^([a-zA-Z0-9_-]+)==", (ROOT / "requirements.lock").read_text(), re.M)
    allowed = {name.replace("_", "-").lower() for name in names}
    distributions = importlib.metadata.packages_distributions()

    class RuntimeOnlyImports(importlib.abc.MetaPathFinder):
        def find_spec(self, fullname, path=None, target=None):
            owners = distributions.get(fullname.split(".")[0], [])
            if owners and not any(owner.replace("_", "-").lower() in allowed for owner in owners):
                raise ModuleNotFoundError(f"runtime lock does not install {fullname}")
            return None

    sys.path.insert(0, str(ROOT))
    sys.meta_path.insert(0, RuntimeOnlyImports())
    from pp.engine import FormualizerAdapter
    from pp.template import import_template
    _ = FormualizerAdapter, import_template


class RuntimeDependencyTests(unittest.TestCase):
    def test_runtime_lock_can_import_engine_and_template_without_dev_packages(self):
        result = subprocess.run([sys.executable, __file__, "--runtime-check"], cwd=ROOT,
                                capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stderr)


if __name__ == "__main__":
    if "--runtime-check" in sys.argv:
        check_runtime_imports()
    else:
        unittest.main()
