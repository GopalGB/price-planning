import unittest
from unittest import mock
from pathlib import Path
from types import SimpleNamespace


class RunV2ContractTests(unittest.TestCase):
    def test_server_uses_configured_uvicorn_asyncio_and_websockets(self):
        """Starting another WSGI or calculation process model must fail."""
        import run_v2

        runtime = mock.Mock()
        runtime.config.bind = "0.0.0.0"
        runtime.config.port = 8123
        with mock.patch.object(run_v2, "create_runtime", return_value=runtime):
            with mock.patch.object(run_v2.uvicorn, "run") as run:
                result = run_v2.main([])

        self.assertEqual(result, 0)
        run.assert_called_once_with(runtime.app, host="0.0.0.0", port=8123,
                                    loop="asyncio", ws="websockets", workers=1)
        runtime.close.assert_called_once_with()

    def test_import_template_mode_does_not_start_server(self):
        """Import-only mode must not initialize the database, pool, or recovery."""
        import run_v2

        self.assertTrue(hasattr(run_v2, "load_config"), "import-only mode needs side-effect-free config")
        self.assertTrue(hasattr(run_v2, "import_template"), "import-only mode must call the importer directly")
        config = SimpleNamespace(template_path=Path(__file__), data_dir=Path("data"))
        with mock.patch.object(run_v2, "load_config", return_value=config), \
                mock.patch.object(run_v2, "import_template") as import_template, \
                mock.patch.object(run_v2, "create_runtime") as create_runtime, \
                mock.patch.object(run_v2.uvicorn, "run") as run:
            result = run_v2.main(["--import-template"])

        self.assertEqual(result, 0)
        run.assert_not_called()
        create_runtime.assert_not_called()
        import_template.assert_called_once_with(config.template_path, config.data_dir)

    def test_template_directory_is_gitignored(self):
        """Configured workbook files must stay ignored even before an extension is chosen."""
        root = Path(__file__).resolve().parents[2]
        patterns = set((root / ".gitignore").read_text(encoding="utf-8").splitlines())
        self.assertIn("/template/", patterns)


if __name__ == "__main__":
    unittest.main()
