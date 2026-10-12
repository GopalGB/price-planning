"""Opt-in Starlette/Uvicorn entry point for the Track A backend."""

import argparse
import ipaddress
import json
import logging
import os
from pathlib import Path
import socket

import uvicorn

from pp.auto_backup import KEEP, AutoBackup
from pp.config import _detected_ipv4, http_origin, load_config
from pp.runtime import create_runtime
from pp.template.importer import import_template


ROOT = Path(__file__).resolve().parent
FIRST_START_SETTINGS = {"bind": "0.0.0.0", "port": 8000, "demo_mode": True}


def _create_first_start_settings(root: Path) -> bool:
    path = root / "settings.json"
    try:
        with path.open("x", encoding="utf-8") as handle:
            json.dump(FIRST_START_SETTINGS, handle, indent=2)
            handle.write("\n")
    except FileExistsError:
        return False
    return True


def _share_url_key(value: str) -> tuple[bool, bool, int]:
    address = ipaddress.ip_address(value)
    return (
        not value.startswith("10."),
        not address.is_private,
        int(address),
    )


def _bound_ipv4(bind: str, assigned: set[str]) -> set[str]:
    if bind == "0.0.0.0":
        return assigned
    try:
        address = ipaddress.ip_address(bind)
    except ValueError:
        try:
            results = socket.getaddrinfo(bind, None, socket.AF_INET, socket.SOCK_STREAM)
        except OSError:
            return set()
        resolved = set()
        for result in results:
            try:
                address = ipaddress.ip_address(result[4][0])
            except ValueError:
                continue
            if address.version == 4:
                resolved.add(str(address))
        return assigned & resolved
    if address.version != 4 or address.is_loopback:
        return set()
    value = str(address)
    return {value} if value in assigned else set()


def _share_urls(bind: str, port: int, origins: tuple[str, ...]) -> tuple[str, ...]:
    allowed = set(origins)
    assigned = set(_detected_ipv4())
    reachable = _bound_ipv4(bind, assigned)
    urls = (http_origin(address, port) for address in sorted(reachable, key=_share_url_key))
    return tuple(url for url in urls if url in allowed)


def _print_share_urls(bind: str, port: int, origins: tuple[str, ...]) -> None:
    for origin in _share_urls(bind, port, origins):
        print(f"Share with your team: {origin}", flush=True)


def _port_in_use(port: int) -> bool:
    # A wildcard bind can coexist with a loopback listener on macOS, silently shadowing it for LAN users.
    try:
        with socket.create_connection(("127.0.0.1", port), timeout=0.5):
            return True
    except OSError:
        return False


def _free_port(port: int, tries: int = 50) -> int | None:
    return next((candidate for candidate in range(port, min(port + tries, 65536))
                 if not _port_in_use(candidate)), None)


def _move_off_busy_port() -> bool:
    """Use the next free port when the configured one is taken; False when none is free.

    The choice is made before the runtime exists, so the derived browser origins use the port actually served.
    """
    config = load_config(ROOT)
    port = _free_port(config.port)
    if port is None:
        return False
    if port != config.port:
        if json.loads((ROOT / "settings.json").read_text(encoding="utf-8")).get("allowed_origins"):
            return False  # fixed origins name the configured port; moving would lock every browser out
        print(f"Port {config.port} is already in use by another program, so this app uses port {port}.",
              flush=True)
        os.environ["PP_PORT"] = str(port)
    return True


def _start_auto_backup(config) -> AutoBackup | None:
    """Back up the live database on a timer while the server serves; settings.json can turn it off."""
    if not config.auto_backup:
        return None
    dest = ROOT / "backups"
    auto_backup = AutoBackup(config.database_path, dest).start()  # the database the server opened
    print(f"Automatic backup: daily into {dest}, keeping the newest {KEEP}.", flush=True)
    return auto_backup


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Price Planning v2 backend")
    parser.add_argument("--import-template", action="store_true",
                        help="Import or reuse the configured template snapshot, then exit")
    args = parser.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s %(message)s")
    if args.import_template:
        config = load_config(ROOT)
        if config.template_path is None or not config.template_path.is_file():
            raise ValueError("template_path must name an existing .xlsm file")
        from pp.runtime import _import_results
        import_template(config.template_path, config.data_dir, _import_results(config))
        return 0
    _create_first_start_settings(ROOT)
    if not _move_off_busy_port():
        print(f"Port {load_config(ROOT).port} is already in use by another program. "
              f"Change \"port\" in {ROOT / 'settings.json'} (for example to 8765) and start again.", flush=True)
        return 1
    runtime = create_runtime(ROOT)
    auto_backup = _start_auto_backup(runtime.config)
    try:
        if _port_in_use(runtime.config.port):
            print(f"Port {runtime.config.port} is already in use by another program. "
                  f"Change \"port\" in {ROOT / 'settings.json'} (for example to 8765) and start again.",
                  flush=True)
            return 1
        _print_share_urls(
            runtime.config.bind,
            runtime.config.port,
            runtime.config.allowed_origins,
        )
        uvicorn.run(runtime.app, host=runtime.config.bind, port=runtime.config.port,
                    loop="asyncio", ws="websockets", workers=1)
        return 0
    finally:
        if auto_backup is not None:
            auto_backup.stop()
        runtime.close()


if __name__ == "__main__":
    raise SystemExit(main())
