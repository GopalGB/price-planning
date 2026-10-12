"""Pure configuration loading for the v2 backend."""

from dataclasses import dataclass
import json
import ipaddress
import os
from pathlib import Path
import re
import socket
import subprocess
import sys
from urllib.parse import urlparse


@dataclass(frozen=True)
class Config:
    bind: str
    port: int
    data_dir: Path
    demo_mode: bool
    template_path: Path | None
    max_workers: int
    database_path: Path
    allowed_origins: tuple[str, ...]
    cookie_secure: bool
    replay_window: int
    engine_memory_budget_mb: int
    brand: dict | None = None
    template_results: str = "engine"
    engine: str = "formualizer"
    upload_max_mb: int = 100


BRAND_ASSETS = ("logo", "favicon", "touch_icon")
BRAND_TYPES = {".png": "image/png", ".svg": "image/svg+xml", ".ico": "image/x-icon"}
BRAND_COLORS = ("primary", "accent", "highlight")
BRAND_MAX_BYTES = 1_000_000
HEX_COLOR = re.compile(r"#[0-9A-Fa-f]{6}")


def _brand(root: Path, value: object) -> dict | None:
    """Validate the optional company brand; assets must be small images inside brand.dir."""
    if value is None:
        return None
    if not isinstance(value, dict):
        raise ValueError("brand must be an object")
    name = value.get("name")
    if not isinstance(name, str) or not 0 < len(name.strip()) <= 60:
        raise ValueError("brand.name must be 1 to 60 characters")
    directory = _path(root, value["dir"], "brand.dir") if "dir" in value else None
    assets = {}
    for key in BRAND_ASSETS:
        file = value.get(key)
        if file is None:
            continue
        valid = (directory is not None and isinstance(file, str) and Path(file).name == file
                 and Path(file).suffix.lower() in BRAND_TYPES)
        path = directory / file if valid else None
        if path is None or not path.is_file() or path.stat().st_size > BRAND_MAX_BYTES:
            raise ValueError(f"brand.{key} must be a PNG, SVG or ICO under 1 MB in brand.dir")
        assets[key] = path
    colors = value.get("colors", {})
    if (not isinstance(colors, dict) or set(colors) - set(BRAND_COLORS)
            or not all(isinstance(item, str) and HEX_COLOR.fullmatch(item) for item in colors.values())):
        raise ValueError("brand.colors takes primary, accent and highlight as #RRGGBB")
    return {"name": name.strip(), "assets": assets, "colors": dict(colors)}


def _read_settings(root: Path, settings_path: str | Path | None = None) -> dict:
    path = Path(settings_path) if settings_path is not None else root / "settings.json"
    if not path.exists():
        return {}
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError("invalid settings.json") from exc
    if not isinstance(value, dict):
        raise ValueError("settings.json must contain an object")
    return value


def _positive(value: object, name: str, upper: int) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or not 0 < value <= upper:
        raise ValueError(f"{name} must be an integer from 1 to {upper}")
    return value


def _origin(value: object) -> str:
    if not isinstance(value, str):
        raise ValueError("allowed_origins must contain strings")
    parsed = urlparse(value)
    if parsed.scheme not in {"http", "https"} or not parsed.netloc or parsed.path not in {"", "/"}:
        raise ValueError("allowed origin must be an http(s) origin")
    return value.rstrip("/")


def _usable_ipv4(value: str) -> str | None:
    address = ipaddress.ip_address(value)
    if address.version != 4 or any((
        address.is_loopback,
        address.is_link_local,
        address.is_multicast,
        address.is_reserved,
        address.is_unspecified,
    )):
        return None
    return str(address)


def http_origin(host: str, port: int) -> str:
    try:
        address = ipaddress.ip_address(host)
    except ValueError:
        formatted_host = host.lower()
    else:
        formatted_host = f"[{address}]" if address.version == 6 else str(address)
    port_suffix = "" if port == 80 else f":{port}"
    return f"http://{formatted_host}{port_suffix}"


_IPV4_PATTERN = r"(?:\d{1,3}\.){3}\d{1,3}"
_POWERSHELL_UNICAST = (
    "$ErrorActionPreference='Stop'; "
    "[System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces() | "
    "ForEach-Object { $_.GetIPProperties().UnicastAddresses } | "
    "ForEach-Object { if ($_.Address.AddressFamily -eq "
    "[System.Net.Sockets.AddressFamily]::InterNetwork) { $_.Address.IPAddressToString } }"
)


def _interface_sources() -> tuple[tuple[list[str], str], ...]:
    if sys.platform == "win32":
        return (
            (["powershell.exe", "-NoProfile", "-NonInteractive", "-Command",
              _POWERSHELL_UNICAST], "powershell"),
            (["ipconfig"], "ipconfig"),
        )
    if sys.platform == "darwin":
        return ((["ifconfig"], "ifconfig"),)
    return ((["ip", "-4", "-o", "addr", "show"], "ip"),
            (["ifconfig"], "ifconfig"))


def _address_records(output: str, kind: str) -> tuple[str, ...]:
    if kind == "powershell":
        return tuple(line.strip() for line in output.splitlines()
                     if re.fullmatch(_IPV4_PATTERN, line.strip()))
    if kind == "ipconfig":
        pattern = rf"(?im)^\s*(?:IPv4 Address|IP Address)[^:]*:\s*({_IPV4_PATTERN})"
    elif kind == "ip":
        pattern = rf"(?m)^\s*\d+:\s+\S+\s+inet\s+({_IPV4_PATTERN})/\d+\b"
    else:
        pattern = rf"(?m)^\s*inet(?:\s+addr:|\s+)({_IPV4_PATTERN})(?:/\d+)?\b"
    return tuple(re.findall(pattern, output))


def _interface_ipv4() -> set[str]:
    addresses = set()
    for command, kind in _interface_sources():
        try:
            result = subprocess.run(command, capture_output=True, check=False,
                                    text=True, errors="replace", timeout=2)
        except (OSError, subprocess.TimeoutExpired):
            continue
        for candidate in _address_records(result.stdout, kind):
            try:
                address = _usable_ipv4(candidate)
            except ValueError:
                continue
            if address:
                addresses.add(address)
        if addresses:
            break
    return addresses


def _route_destinations() -> tuple[str, ...]:
    destinations = ["192.0.2.1"]
    try:
        results = socket.getaddrinfo(
            socket.gethostname(), None, socket.AF_INET, socket.SOCK_STREAM
        )
    except OSError:
        return tuple(destinations)
    for result in results:
        try:
            address = ipaddress.ip_address(result[4][0])
        except ValueError:
            continue
        if address.version == 4 and not address.is_unspecified and not address.is_multicast:
            destinations.append(str(address))
    return tuple(dict.fromkeys(destinations))


def _route_ipv4() -> set[str]:
    addresses = set()
    for destination in _route_destinations():
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as route:
                route.connect((destination, 9))
                address = _usable_ipv4(route.getsockname()[0])
        except (OSError, ValueError):
            continue
        if address:
            addresses.add(address)
    return addresses


def _detected_ipv4() -> tuple[str, ...]:
    addresses = _interface_ipv4() or _route_ipv4()
    return tuple(sorted(
        addresses,
        key=lambda value: (
            not value.startswith("10."),
            not ipaddress.ip_address(value).is_private,
            int(ipaddress.ip_address(value)),
        ),
    ))


def _derived_origins(bind: str, port: int) -> tuple[str, ...]:
    hostname = socket.gethostname().lower()
    origins = [http_origin("localhost", port), http_origin("127.0.0.1", port)]
    if hostname and all(char.isascii() and (char.isalnum() or char in ".-") for char in hostname):
        origins.append(http_origin(hostname, port))
    if bind not in {"0.0.0.0", "::"}:
        origins.append(http_origin(bind, port))
    origins.extend(http_origin(address, port) for address in _detected_ipv4())
    return tuple(dict.fromkeys(origins))


def _path(root: Path, value: object, name: str) -> Path:
    if not isinstance(value, str) or not value:
        raise ValueError(f"{name} must be a non-empty path")
    path = Path(value)
    if str(path).startswith("\\\\") or path.is_absolute() and path.anchor.startswith("\\\\"):
        raise ValueError(f"{name} cannot be a UNC path")
    return path if path.is_absolute() else root / path


def _default_engine_memory_mb() -> int:
    """The development Mac holds fewer resident engines than the office server."""
    return 16_000 if sys.platform == "win32" else 3_000


def _data_dir(root: Path, settings: dict) -> Path:
    return _path(root, os.environ.get("DATA_DIR", settings.get("data_dir", "data")), "data_dir")


def resolve_database_path(root: str | Path, settings_path: str | Path | None = None) -> Path:
    """Resolve the shared database location without validating unrelated server settings."""
    root = Path(root)
    settings = _read_settings(root, settings_path)
    if database := os.environ.get("PP_DATABASE_PATH"):
        return _path(root, database, "PP_DATABASE_PATH")
    if "DATA_DIR" in os.environ:
        return _path(root, os.environ["DATA_DIR"], "DATA_DIR") / "planning.db"
    if database := settings.get("database_path"):
        return _path(root, database, "database_path")
    return _data_dir(root, settings) / "planning.db"


def load_config(root: str | Path, settings_path: str | Path | None = None) -> Config:
    root = Path(root)
    settings = _read_settings(root, settings_path)
    if sys.version_info < (3, 10):
        raise ValueError("Python 3.10 or newer is required")
    bind = settings.get("bind", "127.0.0.1")
    if not isinstance(bind, str) or not bind:
        raise ValueError("bind must be a non-empty string")
    # PP_PORT: run_v2 moves to the next free port when the configured one is taken
    port = _positive(int(os.environ["PP_PORT"]) if os.environ.get("PP_PORT", "").isdigit()
                     else settings.get("port", 8000), "port", 65535)
    data_dir = _data_dir(root, settings)
    database_path = resolve_database_path(root, settings_path)
    max_workers = _positive(settings.get("max_workers", 1), "max_workers", 128)
    replay_window = _positive(settings.get("replay_window", 1000), "replay_window", 1_000_000)
    engine_memory = _positive(settings.get("engine_memory_budget_mb", _default_engine_memory_mb()),
                              "engine_memory_budget_mb", 1_000_000)
    upload_max = _positive(settings.get("upload_max_mb", 100), "upload_max_mb", 10_000)
    demo_mode = settings.get("demo_mode", False)
    if not isinstance(demo_mode, bool):
        raise ValueError("demo_mode must be boolean")
    configured_origins = settings.get("allowed_origins", ())
    origins = (
        tuple(_origin(item) for item in configured_origins)
        if configured_origins
        else _derived_origins(bind, port)
    )
    cookie_secure = settings.get("cookie_secure", not demo_mode)
    if not isinstance(cookie_secure, bool):
        raise ValueError("cookie_secure must be boolean")
    if not demo_mode and not cookie_secure:
        raise ValueError("cookie_secure must be enabled outside demo mode")
    template_value = os.environ.get("PP_TEMPLATE_PATH", settings.get("template_path"))
    template_path = _path(root, template_value, "template_path") if template_value else None
    if template_path is None:
        candidates = sorted((root / "template").glob("*.xlsm"))
        template_path = candidates[0] if candidates else None
    return Config(bind, port, data_dir, demo_mode, template_path, max_workers, database_path,
                  origins, cookie_secure, replay_window, engine_memory,
                  _brand(root, settings.get("brand")), _template_results(settings), _engine(settings),
                  upload_max)


def _engine(settings) -> str:
    # "excel" calculates copies in Microsoft Excel (Windows server with Excel installed).
    value = settings.get("engine", "formualizer")
    if value not in ("formualizer", "excel"):
        raise ValueError("engine must be 'formualizer' or 'excel'")
    return value


def _template_results(settings) -> str:
    # "cached" shows Excel's saved results and pauses editing until the engine can hold the template.
    value = os.environ.get("PP_TEMPLATE_RESULTS", settings.get("template_results", "engine"))
    if value not in ("engine", "cached"):
        raise ValueError("template_results must be 'engine' or 'cached'")
    return value
