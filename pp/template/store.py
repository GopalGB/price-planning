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
