"""The v2 runner creates only its first-start server settings."""

import contextlib
import importlib.util
import io
import json
import socket
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from types import ModuleType, SimpleNamespace
from unittest import mock

TMP = Path(__file__).resolve().parents[2] / "results-private" / "tmp"
TMP.mkdir(parents=True, exist_ok=True)


class RunnerContractTests(unittest.TestCase):
    def runner_module(self):
        runtime_module = ModuleType("pp.runtime")
        runtime_module.create_runtime = None
        spec = importlib.util.spec_from_file_location(
            "test_run_v2_module", Path(__file__).resolve().parents[2] / "run_v2.py"
        )
        module = importlib.util.module_from_spec(spec)
        with mock.patch.dict(sys.modules, {"pp.runtime": runtime_module}):
            spec.loader.exec_module(module)
        # keep tests independent of whatever else listens on this machine
        module._port_in_use = lambda port: False
        return module

    def runtime(self):
        return SimpleNamespace(
            app=object(),
            config=SimpleNamespace(
                bind="0.0.0.0",
                port=8000,
                allowed_origins=(
                    "http://localhost:8000",
                    "http://127.0.0.1:8000",
                    "http://10.0.0.25:8000",
                    "http://192.168.1.5:8000",
                ),
            ),
            close=mock.Mock(),
        )

    def test_first_server_start_prints_route_fallback_lan_url(self):
        """Using only interface CLI results for sharing would hide the fallback URL."""
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            runner, root = self.runner_module(), Path(temp)
            output = io.StringIO()
            created = []

            class Route:
                def __enter__(self):
                    return self

                def __exit__(self, *_args):
                    return None

                def connect(self, target):
                    if target[0] == "192.0.2.1":
                        raise OSError("documentation route unavailable")

                def getsockname(self):
                    return ("10.0.0.25", 0)

            def runtime_from_config(runtime_root):
                runtime = SimpleNamespace(
                    app=object(),
                    config=runner.load_config(runtime_root),
                    close=mock.Mock(),
                )
                created.append(runtime)
                return runtime

            with mock.patch.object(runner, "ROOT", root), mock.patch.object(
                runner, "create_runtime", side_effect=runtime_from_config
            ), mock.patch(
                "pp.config.subprocess.run",
                return_value=subprocess.CompletedProcess(["interfaces"], 0, "", ""),
            ), mock.patch(
                "pp.config.socket.gethostname", return_value="office-host"
            ), mock.patch(
                "pp.config.socket.getaddrinfo",
                return_value=[(None, None, None, None, ("192.168.99.44", 0))],
            ), mock.patch(
                "pp.config.socket.socket", side_effect=lambda *_args: Route()
            ), mock.patch.object(runner.uvicorn, "run") as serve, contextlib.redirect_stdout(output):
                self.assertEqual(runner.main([]), 0)

            settings = root / "settings.json"
            self.assertTrue(settings.exists(), "server startup must persist first-start defaults")
            self.assertEqual(json.loads(settings.read_text()), {
                "bind": "0.0.0.0", "port": 8000, "demo_mode": True,
            })
            runtime = created[0]
            self.assertEqual(runtime.config.bind, "0.0.0.0")
            self.assertTrue(runtime.config.demo_mode)
            self.assertIn("http://10.0.0.25:8000", runtime.config.allowed_origins)
            self.assertNotIn("http://192.168.99.44:8000", runtime.config.allowed_origins)
            serve.assert_called_once_with(runtime.app, host="0.0.0.0", port=8000,
                                         loop="asyncio", ws="websockets", workers=1)
            self.assertEqual(
                [
                    line
                    for line in output.getvalue().splitlines()
                    if line.startswith("Share with your team:")
                ],
                ["Share with your team: http://10.0.0.25:8000"],
            )
            runtime.close.assert_called_once_with()

    def test_share_urls_prefer_10_then_private_addresses(self):
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            runner, root, runtime = self.runner_module(), Path(temp), self.runtime()
            runtime.config.allowed_origins = (
                "http://office-host:8000",
                "http://100.69.34.119:8000",
                "http://192.168.1.5:8000",
                "http://10.0.0.25:8000",
                "http://localhost:8000",
            )
            output = io.StringIO()
            with mock.patch.object(runner, "ROOT", root), mock.patch.object(
                runner, "create_runtime", return_value=runtime
            ), mock.patch.object(
                runner,
                "_detected_ipv4",
                return_value={"100.69.34.119", "192.168.1.5", "10.0.0.25"},
                create=True,
            ), mock.patch.object(runner.uvicorn, "run"), contextlib.redirect_stdout(output):
                self.assertEqual(runner.main([]), 0)

            share_lines = [
                line for line in output.getvalue().splitlines()
                if line.startswith("Share with your team:")
            ]
            self.assertGreaterEqual(len(share_lines), 2, "server startup must print team URLs")
            self.assertEqual(share_lines[0], "Share with your team: http://10.0.0.25:8000")
            self.assertEqual(share_lines[1], "Share with your team: http://192.168.1.5:8000")
            self.assertNotIn("office-host", output.getvalue())

    def test_share_urls_require_an_allowed_assigned_and_reachable_ipv4(self):
        cases = (
            (
                "loopback",
                "127.0.0.1",
                ("http://10.0.0.25:8000",),
                [],
            ),
            (
                "specific",
                "192.168.1.5",
                ("http://10.0.0.25:8000", "http://192.168.1.5:8000"),
                ["Share with your team: http://192.168.1.5:8000"],
            ),
            (
                "explicit external",
                "0.0.0.0",
                ("https://proxy.example",),
                [],
            ),
            (
                "ipv6 wildcard",
                "::",
                ("http://10.0.0.25:8000",),
                [],
            ),
        )
        for name, bind, origins, expected in cases:
            with self.subTest(name=name), tempfile.TemporaryDirectory(dir=TMP) as temp:
                runner, root, runtime = self.runner_module(), Path(temp), self.runtime()
                runtime.config.bind = bind
                runtime.config.allowed_origins = origins
                output = io.StringIO()
                with mock.patch.object(runner, "ROOT", root), mock.patch.object(
                    runner, "create_runtime", return_value=runtime
                ), mock.patch.object(
                    runner,
                    "_detected_ipv4",
                    return_value={"10.0.0.25", "192.168.1.5"},
                    create=True,
                ), mock.patch.object(runner.uvicorn, "run"), contextlib.redirect_stdout(output):
                    self.assertEqual(runner.main([]), 0)

                share_lines = [
                    line
                    for line in output.getvalue().splitlines()
                    if line.startswith("Share with your team:")
                ]
                self.assertEqual(share_lines, expected)

    def test_hostname_bind_prints_only_resolved_assigned_allowed_ipv4(self):
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            runner, root, runtime = self.runner_module(), Path(temp), self.runtime()
            runtime.config.bind = "office-host"
            runtime.config.allowed_origins = (
                "http://10.0.0.25:8000",
                "http://192.168.1.5:8000",
                "http://office-host:8000",
            )
            output = io.StringIO()
            resolved = [(None, None, None, None, ("10.0.0.25", 0))]
            with mock.patch.object(runner, "ROOT", root), mock.patch.object(
                runner, "create_runtime", return_value=runtime
            ), mock.patch.object(
                runner,
                "_detected_ipv4",
                return_value={"10.0.0.25", "192.168.1.5"},
                create=True,
            ), mock.patch.object(
                runner.socket, "getaddrinfo", return_value=resolved
            ), mock.patch.object(runner.uvicorn, "run"), contextlib.redirect_stdout(output):
                self.assertEqual(runner.main([]), 0)

            share_lines = [
                line
                for line in output.getvalue().splitlines()
                if line.startswith("Share with your team:")
            ]
            self.assertEqual(
                share_lines,
                ["Share with your team: http://10.0.0.25:8000"],
            )

    def test_port_in_use_detects_a_real_loopback_listener(self):
        runner = self.runner_module()
        spec = importlib.util.spec_from_file_location(
            "test_run_v2_probe", Path(__file__).resolve().parents[2] / "run_v2.py"
        )
        real = importlib.util.module_from_spec(spec)
        with mock.patch.dict(sys.modules, {"pp.runtime": ModuleType("pp.runtime")}):
            sys.modules["pp.runtime"].create_runtime = None
            spec.loader.exec_module(real)
        with socket.socket() as listener:
            listener.bind(("127.0.0.1", 0))
            listener.listen()
            port = listener.getsockname()[1]
            self.assertTrue(real._port_in_use(port))
        self.assertFalse(real._port_in_use(port))
        self.assertFalse(runner._port_in_use(port))

    def test_occupied_port_refuses_to_start_with_an_actionable_message(self):
        """Another app on the port (e.g. the legacy v1 server on 8000) must not be shadowed."""
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            runner, root, runtime = self.runner_module(), Path(temp), self.runtime()
            runner._port_in_use = lambda port: port == 8000
            output = io.StringIO()
            with mock.patch.object(runner, "ROOT", root), mock.patch.object(
                runner, "create_runtime", return_value=runtime
            ), mock.patch.object(
                runner, "_detected_ipv4", return_value=set(), create=True
            ), mock.patch.object(runner.uvicorn, "run") as serve, contextlib.redirect_stdout(output):
                self.assertEqual(runner.main([]), 1)

            serve.assert_not_called()
            self.assertIn("Port 8000 is already in use", output.getvalue())
            self.assertIn("settings.json", output.getvalue())
            self.assertNotIn("Share with your team", output.getvalue())

    def test_existing_server_settings_are_never_overwritten(self):
        """Replacing a configured server setting on restart would be data loss."""
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            runner, root, runtime = self.runner_module(), Path(temp), self.runtime()
            settings = root / "settings.json"
            original = '{"bind":"127.0.0.1","port":8123,"demo_mode":false}\n'
            settings.write_text(original, encoding="utf-8")
            with mock.patch.object(runner, "ROOT", root), mock.patch.object(
                runner, "create_runtime", return_value=runtime
            ), mock.patch.object(
                runner, "_detected_ipv4", return_value=set(), create=True
            ), mock.patch.object(runner.uvicorn, "run"):
                self.assertEqual(runner.main([]), 0)

            self.assertEqual(settings.read_text(encoding="utf-8"), original)

    def test_import_template_mode_is_read_only_and_does_not_start_server(self):
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            runner, root = self.runner_module(), Path(temp)
            template = root / "synthetic.xlsm"
            template.write_bytes(b"synthetic")
            config = SimpleNamespace(template_path=template, data_dir=root / "data")
            with mock.patch.object(runner, "ROOT", root), mock.patch.object(
                runner, "load_config", return_value=config
            ), mock.patch.object(runner, "import_template") as import_template, mock.patch.object(
                runner, "create_runtime"
            ) as create_runtime, mock.patch.object(runner.uvicorn, "run") as serve:
                self.assertEqual(runner.main(["--import-template"]), 0)

            self.assertFalse((root / "settings.json").exists())
            import_template.assert_called_once_with(template, root / "data", "engine")
            create_runtime.assert_not_called()
            serve.assert_not_called()

    def test_template_directory_is_gitignored(self):
        root = Path(__file__).resolve().parents[2]
        patterns = set((root / ".gitignore").read_text(encoding="utf-8").splitlines())
        self.assertIn("/template/", patterns)


if __name__ == "__main__":
    unittest.main()
