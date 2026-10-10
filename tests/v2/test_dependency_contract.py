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


if __name__ == "__main__":
    unittest.main()
