import json
import os
import tempfile
import unittest
from pathlib import Path
from unittest import mock


from pp.config import load_config

TMP = Path(__file__).resolve().parents[2] / "results-private" / "tmp"
TMP.mkdir(parents=True, exist_ok=True)

class ConfigContractTests(unittest.TestCase):
    def test_defaults_and_repo_template_precedence(self):
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            root = Path(temp)
            template_dir = root / "template"
            template_dir.mkdir()
            (template_dir / "z.xlsm").write_bytes(b"z")
            (template_dir / "a.xlsm").write_bytes(b"a")

            with mock.patch.dict(os.environ, {}, clear=True):
                config = load_config(root)

            self.assertEqual(config.bind, "127.0.0.1")
            self.assertEqual(config.port, 8000)
            self.assertEqual(config.data_dir, root / "data")
            self.assertFalse(config.demo_mode)
            self.assertEqual(config.template_path, template_dir / "a.xlsm")
            self.assertEqual(config.max_workers, 1)

    def test_settings_json_overrides_defaults_and_env_template_wins(self):
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            root = Path(temp)
            settings_template = root / "settings.xlsm"
            env_template = root / "env.xlsm"
            settings_template.write_bytes(b"settings")
            env_template.write_bytes(b"env")
            (root / "settings.json").write_text(
                json.dumps(
                    {
                        "bind": "0.0.0.0",
                        "port": 8123,
                        "data_dir": "runtime",
                        "demo_mode": True,
                        "template_path": str(settings_template),
                        "max_workers": 3,
                    }
                )
            )

            with mock.patch.dict(os.environ, {"PP_TEMPLATE_PATH": str(env_template)}, clear=True):
                config = load_config(root)

            self.assertEqual(config.bind, "0.0.0.0")
            self.assertEqual(config.port, 8123)
            self.assertEqual(config.data_dir, root / "runtime")
            self.assertTrue(config.demo_mode)
            self.assertEqual(config.template_path, env_template)
            self.assertEqual(config.max_workers, 3)

    def test_invalid_settings_are_rejected(self):
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            root = Path(temp)
            (root / "settings.json").write_text(json.dumps({"port": 0}))
            with self.assertRaises(ValueError):
                load_config(root)

    def test_database_path_prefers_environment_then_settings_then_default(self):
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            root = Path(temp)
            settings = root / "alternate.json"
            settings.write_text(json.dumps({
                "data_dir": "settings-data", "database_path": "settings-db/custom.db",
            }))

            with mock.patch.dict(os.environ, {}, clear=True):
                self.assertEqual(load_config(root, settings).database_path,
                                 root / "settings-db" / "custom.db")
            with mock.patch.dict(os.environ, {"DATA_DIR": "environment-data"}, clear=True):
                self.assertEqual(load_config(root, settings).database_path,
                                 root / "environment-data" / "planning.db")
            with mock.patch.dict(os.environ, {"PP_DATABASE_PATH": "environment-db/live.db"}, clear=True):
                self.assertEqual(load_config(root, settings).database_path,
                                 root / "environment-db" / "live.db")
            with mock.patch.dict(os.environ, {}, clear=True):
                self.assertEqual(load_config(root, root / "missing.json").database_path,
                                 root / "data" / "planning.db")

    def test_explicit_settings_keeps_relative_paths_rooted_at_repo(self):
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            root = Path(temp)
            alternate = root / "config" / "settings.json"
            alternate.parent.mkdir()
            alternate.write_text(json.dumps({"data_dir": "relative-data"}))

            with mock.patch.dict(os.environ, {}, clear=True):
                config = load_config(root, alternate)

            self.assertEqual(config.data_dir, root / "relative-data")
            self.assertEqual(config.database_path, root / "relative-data" / "planning.db")


if __name__ == "__main__":
    unittest.main()
