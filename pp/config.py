"""Pure configuration loading for the v2 backend."""

from dataclasses import dataclass
import json
import os
from pathlib import Path
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
    engine_ram_budget_mb: int


def _read_settings(root: Path) -> dict:
    path = root / "settings.json"
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


def _path(root: Path, value: object, name: str) -> Path:
    if not isinstance(value, str) or not value:
        raise ValueError(f"{name} must be a non-empty path")
    path = Path(value)
    if str(path).startswith("\\\\") or path.is_absolute() and path.anchor.startswith("\\\\"):
        raise ValueError(f"{name} cannot be a UNC path")
    return path if path.is_absolute() else root / path


def load_config(root: str | Path) -> Config:
    root = Path(root)
    settings = _read_settings(root)
    if sys.version_info < (3, 10):
        raise ValueError("Python 3.10 or newer is required")
    bind = settings.get("bind", "127.0.0.1")
    if not isinstance(bind, str) or not bind:
        raise ValueError("bind must be a non-empty string")
    port = _positive(settings.get("port", 8000), "port", 65535)
    data_dir = _path(root, settings.get("data_dir", "data"), "data_dir")
    database_path = _path(root, settings.get("database_path", str(data_dir / "planning.db")), "database_path")
    max_workers = _positive(settings.get("max_workers", 1), "max_workers", 128)
    replay_window = _positive(settings.get("replay_window", 1000), "replay_window", 1_000_000)
    engine_ram = _positive(settings.get("engine_ram_budget_mb", 1024), "engine_ram_budget_mb", 1_000_000)
    demo_mode = settings.get("demo_mode", False)
    if not isinstance(demo_mode, bool):
        raise ValueError("demo_mode must be boolean")
    origins = tuple(_origin(item) for item in settings.get("allowed_origins", ()))
    cookie_secure = settings.get("cookie_secure", not demo_mode)
    if not isinstance(cookie_secure, bool) or cookie_secure and demo_mode:
        raise ValueError("cookie_secure is unsafe for demo mode")
    template_value = os.environ.get("PP_TEMPLATE_PATH", settings.get("template_path"))
    template_path = _path(root, template_value, "template_path") if template_value else None
    if template_path is None:
        candidates = sorted((root / "template").glob("*.xlsm"))
        template_path = candidates[0] if candidates else None
    return Config(bind, port, data_dir, demo_mode, template_path, max_workers, database_path,
                  origins, cookie_secure, replay_window, engine_ram)
