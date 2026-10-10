import unittest
from pathlib import Path


ROOT = Path(__file__).parents[2]


class DependencyContractTests(unittest.TestCase):
    def test_runtime_pins_are_exact(self):
        text = (ROOT / "requirements.in").read_text()
        for requirement in ("formualizer==0.11.1", "waitress==3.0.2", "starlette==", "uvicorn==", "websockets=="):
            self.assertIn(requirement, text)

    def test_dev_lock_has_hashes_for_runtime_packages(self):
        text = (ROOT / "requirements-dev.lock").read_text()
        for package in ("starlette==", "uvicorn==", "websockets=="):
            start = text.index(package)
            end = text.find("\n    # via", start)
            self.assertIn("--hash=sha256:", text[start:end])

    def test_python310_conditional_dependency_is_locked_and_vendored(self):
        for filename in ("requirements.lock", "requirements-dev.lock"):
            self.assertIn("exceptiongroup==", (ROOT / filename).read_text())
        self.assertTrue(list((ROOT / "vendor/wheels").glob("exceptiongroup-*.whl")),
                        "Python 3.10 offline installs need AnyIO's conditional dependency")


if __name__ == "__main__":
    unittest.main()
