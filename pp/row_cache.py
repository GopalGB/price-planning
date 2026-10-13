"""Bounded, in-process cache of rendered row blocks.

A row block is the most expensive thing the server builds: a protected sheet with thousands of
rows sends megabytes of JSON per window, and conditional formatting evaluates through a formula
engine. Re-visiting a sheet the copy has not changed answers from here instead of rendering again.

The caller owns the key and must include every input the rendered body depends on; see
``pp.workbook_service._row_cache_key``. A cached body is handed out as the same mapping, so a
reader must never mutate it.
"""

from __future__ import annotations

from collections import OrderedDict
import json
import threading

DEFAULT_MAX_BYTES = 256 * 1024 * 1024
DEFAULT_MAX_ENTRIES = 512


class RowBlockCache:
    """Least recently used cache of rendered bodies, bounded by bytes and by entry count."""

    def __init__(self, max_bytes: int = DEFAULT_MAX_BYTES, max_entries: int = DEFAULT_MAX_ENTRIES):
        self.max_bytes = max_bytes
        self.max_entries = max_entries
        self.hits = 0
        self.misses = 0
        self._entries: OrderedDict[tuple, tuple[dict, int]] = OrderedDict()
        self._bytes = 0
        self._lock = threading.Lock()

    def get(self, key: tuple) -> dict | None:
        with self._lock:
            entry = self._entries.get(key)
            if entry is None:
                self.misses += 1
                return None
            self._entries.move_to_end(key)
            self.hits += 1
            return entry[0]

    def put(self, key: tuple, body: dict) -> None:
        # ASCII-escaped JSON, so the character count is the byte count the response would send.
        size = len(json.dumps(body, separators=(",", ":")))
        with self._lock:
            self._discard(key)
            if size > self.max_bytes:
                return
            self._entries[key] = (body, size)
            self._bytes += size
            while len(self._entries) > self.max_entries or self._bytes > self.max_bytes:
                self._discard(next(iter(self._entries)))

    def clear(self) -> None:
        with self._lock:
            self._entries.clear()
            self._bytes = 0

    @property
    def entries(self) -> int:
        with self._lock:
            return len(self._entries)

    @property
    def bytes_held(self) -> int:
        with self._lock:
            return self._bytes

    def _discard(self, key: tuple) -> None:
        entry = self._entries.pop(key, None)
        if entry is not None:
            self._bytes -= entry[1]