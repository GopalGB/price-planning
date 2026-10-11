"""Atomic storage helpers for imported template snapshots."""

import os
from pathlib import Path
import shutil
import tempfile


def staging_directory(parent: Path) -> Path:
    parent.mkdir(parents=True, exist_ok=True)
    return Path(tempfile.mkdtemp(prefix=".staging-", dir=parent))


def publish(staging: Path, destination: Path) -> Path:
    if destination.exists():
        shutil.rmtree(staging)
        return destination
    os.replace(staging, destination)
    return destination


def replace_snapshot(staging: Path, destination: Path) -> Path:
    backup = Path(tempfile.mkdtemp(prefix=".replaced-", dir=destination.parent))
    backup.rmdir()
    os.replace(destination, backup)
    try:
        os.replace(staging, destination)
    except Exception:
        os.replace(backup, destination)
        raise
    shutil.rmtree(backup)
    return destination
