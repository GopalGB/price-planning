import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock


from pp.config import _interface_ipv4, load_config

TMP = Path(__file__).resolve().parents[2] / "results-private" / "tmp"
TMP.mkdir(parents=True, exist_ok=True)

class ConfigContractTests(unittest.TestCase):
    def test_missing_settings_keeps_conservative_defaults_without_writing(self):
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            root = Path(temp)
            template_dir = root / "template"
            template_dir.mkdir()
            (template_dir / "z.xlsm").write_bytes(b"z")
            (template_dir / "a.xlsm").write_bytes(b"a")

            class Route:
                def __enter__(self):
                    return self

                def __exit__(self, *_args):
                    return None

                def connect(self, target):
                    self.source = {
                        "192.168.1.9": "192.168.1.9",
                    }.get(target[0], "10.0.0.25")

                def getsockname(self):
                    return (self.source, 0)

            with mock.patch.dict(os.environ, {}, clear=True), mock.patch(
                "pp.config.socket.gethostname", return_value="test-host"
            ), mock.patch("pp.config.socket.getaddrinfo", return_value=[
                (None, None, None, None, ("10.0.0.25", 0)),
                (None, None, None, None, ("192.168.1.9", 0)),
            ]), mock.patch("pp.config._interface_ipv4", return_value=set()), mock.patch(
                "pp.config.socket.socket", side_effect=lambda *_args: Route()
            ):
                config = load_config(root)

            self.assertEqual(config.bind, "127.0.0.1")
            self.assertEqual(config.port, 8000)
            self.assertEqual(config.data_dir, root / "data")
            self.assertFalse(config.demo_mode)
            self.assertTrue(config.cookie_secure)
            self.assertEqual(
                config.allowed_origins,
                (
                    "http://localhost:8000",
                    "http://127.0.0.1:8000",
                    "http://test-host:8000",
                    "http://10.0.0.25:8000",
                    "http://192.168.1.9:8000",
                ),
            )
            self.assertFalse((root / "settings.json").exists())
            self.assertEqual(config.template_path, template_dir / "a.xlsm")
            self.assertEqual(config.max_workers, 1)
            with mock.patch.dict(os.environ, {"PP_ENGINE": "excel"}):
                self.assertEqual(load_config(root).max_workers, 6)  # one Excel per open copy

    def test_interface_enumeration_parses_only_assigned_darwin_and_linux_addresses(self):
        cases = (
            (
                "darwin",
                "en0: flags=8863<UP>\n"
                "\tinet 192.168.1.5 netmask 0xffffff00 broadcast 192.168.1.255\n"
                "\tdefault gateway 192.168.1.1\n",
                {"192.168.1.5"},
            ),
            (
                "linux",
                "2: eth0    inet 10.8.0.7/24 brd 10.8.0.255 scope global eth0\n"
                "default via 10.8.0.1 dev eth0\n",
                {"10.8.0.7"},
            ),
        )
        for platform, output, expected in cases:
            with self.subTest(platform=platform), mock.patch(
                "pp.config.sys.platform", platform
            ), mock.patch("pp.config.subprocess.run", return_value=subprocess.CompletedProcess(
                ["interfaces"], 0, output, ""
            )):
                self.assertEqual(_interface_ipv4(), expected)

    def test_interface_enumeration_replaces_undecodable_output_bytes(self):
        command = [
            sys.executable,
            "-c",
            "import os; os.write(1, b'\\xff\\n\\tinet 10.24.3.8 netmask 0xffffff00\\n')",
        ]
        with mock.patch(
            "pp.config._interface_sources", return_value=((command, "ifconfig"),)
        ):
            try:
                addresses = _interface_ipv4()
            except UnicodeDecodeError as exc:
                self.fail(f"interface discovery aborted on undecodable output: {exc}")

        self.assertEqual(addresses, {"10.24.3.8"})

    def test_windows_prefers_unicast_address_output_over_ipconfig_tokens(self):
        commands = []

        def run(command, **_kwargs):
            commands.append(command)
            if command[0].lower().startswith("powershell"):
                return subprocess.CompletedProcess(command, 0, "10.24.3.8\n169.254.2.4\n", "")
            return subprocess.CompletedProcess(
                command,
                0,
                "IPv4 Address. . . . . . . . . . . : 192.168.1.20\n"
                "Default Gateway . . . . . . . . . : 192.168.1.1\n",
                "",
            )

        with mock.patch("pp.config.sys.platform", "win32"), mock.patch(
            "pp.config.subprocess.run", side_effect=run
        ):
            self.assertEqual(_interface_ipv4(), {"10.24.3.8"})
        self.assertTrue(commands[0][0].lower().startswith("powershell"))

    def test_windows_ipconfig_fallback_accepts_only_ipv4_address_records(self):
        def run(command, **_kwargs):
            if command[0].lower().startswith("powershell"):
                raise OSError("PowerShell unavailable")
            return subprocess.CompletedProcess(
                command,
                0,
                "IPv4 Address. . . . . . . . . . . : 192.168.1.20\n"
                "Subnet Mask . . . . . . . . . . . : 255.255.255.0\n"
                "Default Gateway . . . . . . . . . : 192.168.1.1\n",
                "",
            )

        with mock.patch("pp.config.sys.platform", "win32"), mock.patch(
            "pp.config.subprocess.run", side_effect=run
        ):
            self.assertEqual(_interface_ipv4(), {"192.168.1.20"})

    def test_detected_origins_enumerate_interfaces_when_hostname_only_has_link_local_address(self):
        """Removing interface enumeration would hide an otherwise usable LAN URL."""
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            root = Path(temp)
            command = mock.Mock(return_value=subprocess.CompletedProcess(
                ["ifconfig"], 0,
                "en0: flags=1\n\tinet 192.168.1.5 netmask 0xffffff00\n"
                "utun4: flags=1\n\tinet 100.69.34.119 netmask 0xffffffff\n",
                "",
            ))
            with mock.patch.dict(os.environ, {}, clear=True), mock.patch(
                "pp.config.socket.gethostname", return_value="office-mac"
            ), mock.patch("pp.config.socket.getaddrinfo", return_value=[
                (None, None, None, None, ("169.254.220.114", 0)),
            ]), mock.patch("subprocess.run", command), mock.patch("pp.config.sys.platform", "darwin"):
                config = load_config(root)

            self.assertEqual(config.allowed_origins[-2:], (
                "http://192.168.1.5:8000", "http://100.69.34.119:8000",
            ))
            command.assert_called_once()

    def test_detected_origins_exclude_stale_dns_when_interfaces_are_available(self):
        with tempfile.TemporaryDirectory(dir=TMP) as temp, mock.patch.dict(
            os.environ, {}, clear=True
        ), mock.patch(
            "pp.config._interface_ipv4", return_value={"10.0.0.25"}
        ), mock.patch(
            "pp.config.socket.gethostname", return_value="office-host"
        ), mock.patch(
            "pp.config.socket.getaddrinfo",
            return_value=[(None, None, None, None, ("192.168.99.44", 0))],
        ) as resolve:
            config = load_config(Path(temp))

        self.assertIn("http://10.0.0.25:8000", config.allowed_origins)
        self.assertNotIn("http://192.168.99.44:8000", config.allowed_origins)
        resolve.assert_not_called()

    def test_detected_origins_use_route_when_name_and_interface_lookup_have_no_usable_address(self):
        """Removing the UDP route fallback leaves a reachable LAN server unshareable."""
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            route = mock.MagicMock()
            route.__enter__.return_value = route
            route.getsockname.return_value = ("10.0.0.25", 0)
            with mock.patch.dict(os.environ, {}, clear=True), mock.patch(
                "pp.config.socket.getaddrinfo", return_value=[
                    (None, None, None, None, ("169.254.220.114", 0)),
                ]
            ), mock.patch("pp.config.subprocess.run", return_value=subprocess.CompletedProcess(
                ["ifconfig"], 0, "", ""
            )), mock.patch("pp.config.socket.socket", return_value=route), mock.patch(
                "pp.config.sys.platform", "darwin"
            ):
                config = load_config(Path(temp))

            self.assertIn("http://10.0.0.25:8000", config.allowed_origins)

    def test_derived_origins_include_every_local_address_and_specific_bind(self):
        """Changing a specific bind must not drop other valid local browser origins."""
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            root = Path(temp)
            (root / "settings.json").write_text(json.dumps({"bind": "10.2.3.4", "port": 8123}))
            with mock.patch.dict(os.environ, {}, clear=True), mock.patch(
                "pp.config.socket.gethostname", return_value="Office-PC"
            ), mock.patch("pp.config.socket.getaddrinfo", return_value=[]), mock.patch(
                "pp.config._interface_ipv4",
                return_value={"192.168.1.5", "100.69.34.119"},
            ):
                config = load_config(root)
            self.assertEqual(config.allowed_origins, (
                "http://localhost:8123", "http://127.0.0.1:8123",
                "http://office-pc:8123", "http://10.2.3.4:8123",
                "http://192.168.1.5:8123", "http://100.69.34.119:8123",
            ))
            self.assertFalse(config.demo_mode)
            self.assertTrue(config.cookie_secure)

    def test_derived_http_origins_use_browser_canonical_port_80(self):
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            root = Path(temp)
            (root / "settings.json").write_text(json.dumps({"port": 80}))
            with mock.patch(
                "pp.config.socket.gethostname", return_value="office-host"
            ), mock.patch(
                "pp.config._interface_ipv4", return_value={"10.0.0.25"}
            ):
                config = load_config(root)

        self.assertIn("http://localhost", config.allowed_origins)
        self.assertIn("http://10.0.0.25", config.allowed_origins)
        self.assertNotIn("http://localhost:80", config.allowed_origins)

    def test_derived_http_origin_brackets_specific_ipv6_bind(self):
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            root = Path(temp)
            (root / "settings.json").write_text(json.dumps({"bind": "::1"}))
            with mock.patch(
                "pp.config.socket.gethostname", return_value="office-host"
            ), mock.patch("pp.config._interface_ipv4", return_value=set()), mock.patch(
                "pp.config.socket.getaddrinfo", return_value=[]
            ), mock.patch("pp.config.socket.socket", side_effect=OSError):
                config = load_config(root)

        self.assertIn("http://[::1]:8000", config.allowed_origins)
        self.assertNotIn("http://::1:8000", config.allowed_origins)

    def test_derived_http_origin_lowercases_specific_hostname_bind(self):
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            root = Path(temp)
            (root / "settings.json").write_text(json.dumps({"bind": "PlanSrv"}))
            with mock.patch(
                "pp.config.socket.gethostname", return_value="office-host"
            ), mock.patch("pp.config._interface_ipv4", return_value={"10.0.0.25"}):
                config = load_config(root)

        self.assertIn("http://plansrv:8000", config.allowed_origins)
        self.assertNotIn("http://PlanSrv:8000", config.allowed_origins)

    def test_explicit_origins_win_and_foreign_origin_is_rejected_exactly(self):
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            root = Path(temp)
            (root / "settings.json").write_text(json.dumps({
                "allowed_origins": ["http://trusted.example:8000/"],
            }))
            with mock.patch.dict(os.environ, {}, clear=True):
                config = load_config(root)
            self.assertEqual(config.allowed_origins, ("http://trusted.example:8000",))
            from pp.auth import require_origin
            from pp.errors import AppError
            require_origin("http://trusted.example:8000/", config.allowed_origins)
            with self.assertRaises(AppError) as rejected:
                require_origin("http://trusted.example:8000.evil/", config.allowed_origins)
            self.assertEqual((rejected.exception.status, rejected.exception.code), (403, "origin"))

    def test_derived_origin_rejects_foreign_and_prefix_matches(self):
        with tempfile.TemporaryDirectory(dir=TMP) as temp, mock.patch.dict(
            os.environ, {}, clear=True
        ), mock.patch("pp.config.socket.gethostname", return_value="office-host"), mock.patch(
            "pp.config.socket.getaddrinfo", return_value=[]
        ), mock.patch("pp.config._interface_ipv4", return_value={"10.0.0.25"}):
            config = load_config(Path(temp))

        from pp.auth import require_origin
        from pp.errors import AppError
        require_origin("http://10.0.0.25:8000", config.allowed_origins)
        for value in ("http://10.0.0.25:8000.evil", "http://10.0.0.250:8000"):
            with self.subTest(origin=value), self.assertRaises(AppError):
                require_origin(value, config.allowed_origins)

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

    def test_engine_environment_wins_over_an_older_settings_file(self):
        with tempfile.TemporaryDirectory(dir=TMP) as temp:
            root = Path(temp)
            (root / "settings.json").write_text(json.dumps({"engine": "formualizer"}), encoding="utf-8")
            with mock.patch.dict(os.environ, {"PP_ENGINE": "excel"}, clear=True):
                self.assertEqual(load_config(root).engine, "excel")
            with mock.patch.dict(os.environ, {}, clear=True):
                self.assertEqual(load_config(root).engine, "formualizer")

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
