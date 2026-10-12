"""Daily automatic backups of the live database while the v2 server runs."""

import logging
from pathlib import Path
import shutil
import threading

from .backup import backup

LOGGER = logging.getLogger(__name__)
FIRST_DELAY = 300
INTERVAL = 86400
KEEP = 7


def _siblings(database: Path) -> tuple[Path, Path]:
    """The record and template bundle that belong to one backup .db."""
    stem = str(database)[: -len(".db")]
    return Path(stem + ".json"), Path(stem + ".templates")


def _remove(path: Path) -> None:
    if path.is_dir() and not path.is_symlink():
        shutil.rmtree(path)
    else:
        path.unlink(missing_ok=True)


def prune(dest_dir, keep=KEEP):
    """Delete all but the newest `keep` backup sets; other names, including .partial files, stay."""
    dest_dir = Path(dest_dir)
    if not dest_dir.is_dir():
        return []
    # the backup timestamp sorts lexicographically, so the file name is the backup time
    sets = sorted(path for path in dest_dir.glob("planning-*.db")
                  if path.is_file() and not path.is_symlink())
    removed = sets[: max(0, len(sets) - keep)]
    for database in removed:
        record, templates = _siblings(database)
        for path in (database, record, templates):
            if path.exists() or path.is_symlink():
                _remove(path)
    return removed


class AutoBackup:
    """One daemon thread: a backup shortly after start, then one every interval, each followed by a prune."""

    def __init__(self, db_path, dest_dir, first_delay=FIRST_DELAY, interval=INTERVAL, keep=KEEP):
        self.db_path = Path(db_path)
        self.dest_dir = Path(dest_dir)
        self.first_delay = first_delay
        self.interval = interval
        self.keep = keep
        self.thread = None
        self._stop = threading.Event()

    def start(self):
        if self.thread is None:
            self.thread = threading.Thread(target=self._loop, name="pp-auto-backup", daemon=True)
            self.thread.start()
        return self

    def _run_once(self):
        try:
            backup(self.db_path, self.dest_dir)
            prune(self.dest_dir, self.keep)
        except Exception:  # a failed backup is logged and retried at the next tick, never fatal
            LOGGER.exception("automatic backup of %s failed", self.db_path)

    def _loop(self):
        if not self._stop.wait(self.first_delay):
            self._run_once()
        while not self._stop.wait(self.interval):
            self._run_once()

    def stop(self, timeout=30):
        self._stop.set()
        if self.thread is not None:
            self.thread.join(timeout)