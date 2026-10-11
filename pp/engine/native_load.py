"""Hand the engine a template whose shared formulas are already explicit.

Formualizer 0.11.1 expands a shared-formula follower by translating the master
text with the follower offset, and that translation also shifts a cross-sheet
qualifier whose sheet name parses as a cell address: ``S01!$A$1`` is rewritten
as ``T1!$A$1``, so every follower of the family reads a sheet that does not
exist and evaluates to ``#REF!``. Excel never translates the sheet part.

Correcting the followers after load works but costs one engine write per cell,
which is the documented startup bottleneck. This module rewrites the package
instead: every follower carries its own explicit formula, so native load stays
a single read and the engine never runs the defective translation.
"""

from __future__ import annotations

import hashlib
import os
import re
import tempfile
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from xml.sax.saxutils import escape, unescape

from openpyxl.formula.translate import Translator


CHUNK_BYTES = 1 << 20
TAIL_GUARD_BYTES = 1 << 20
WORKSHEET_PREFIX = "xl/worksheets/"
CACHE_DIRECTORY = "pp-engine-cache"
CELL_ELEMENT = re.compile(rb"<c\b[^>]*?(?:/>|>.*?</c>)", re.DOTALL)
CELL_START = re.compile(rb"<c\b")
CELL_REFERENCE = re.compile(rb'\br\s*=\s*"([^"]*)"')
SHARED_TYPE = re.compile(rb'\bt\s*=\s*"shared"')
SHARED_INDEX = re.compile(rb'\bsi\s*=\s*"([^"]*)"')


class SharedFormulaError(RuntimeError):
    """A shared formula cannot be expanded without inventing a formula."""


@dataclass
class ExpansionStats:
    sheets: int = 0
    masters: int = 0
    followers: int = 0
    parts: int = 0

    def as_dict(self):
        return {"sheets": self.sheets, "masters": self.masters,
                "followers": self.followers, "parts": self.parts}


@dataclass
class _Family:
    origin: str
    text: str
    members: list = field(default_factory=list)


def translate_shared(master_text: str, origin: str, target: str) -> str:
    """Translate a shared master formula from its master cell onto a follower.

    The sheet qualifier, absolute anchors, defined names and string literals are
    preserved; only relative cell and range references move.
    """
    body = master_text.lstrip("=")
    return Translator(f"={body}", origin=origin).translate_formula(target).lstrip("=")


def expand_shared_formulas(source, destination) -> ExpansionStats:
    """Write a copy of ``source`` whose shared followers hold explicit formulas."""
    source, destination = Path(source), Path(destination)
    stats = ExpansionStats()
    with zipfile.ZipFile(source) as archive:
        names = archive.namelist()
        if "xl/workbook.xml" not in names:
            raise SharedFormulaError("package has no xl/workbook.xml")
        with zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as output:
            for info in archive.infolist():
                stats.parts += 1
                entry = _entry(info)
                if info.filename.startswith(WORKSHEET_PREFIX):
                    with archive.open(info) as reader, output.open(entry, "w") as writer:
                        _expand_sheet(reader, writer, stats)
                else:
                    output.writestr(entry, archive.read(info.filename))
    return stats


def prepared_template(source, cache_dir=None) -> Path:
    """Return a cached package whose shared formulas are explicit.

    The expansion is a pure function of the immutable template, so it is keyed
    by the template's identity and reused by every copy that loads it.
    """
    source = Path(source)
    try:
        status = source.stat()
    except OSError as exc:
        raise SharedFormulaError(f"template is unreadable: {exc}") from exc
    directory = Path(cache_dir) if cache_dir is not None else default_cache_dir()
    directory.mkdir(parents=True, exist_ok=True)
    identity = f"{source.resolve()}|{status.st_size}|{status.st_mtime_ns}"
    digest = hashlib.sha256(identity.encode("utf-8")).hexdigest()[:16]
    target = directory / f"{source.stem}.shared-{digest}{source.suffix}"
    if target.exists():
        return target
    pending = target.with_name(f"{target.name}.{os.getpid()}.tmp")
    try:
        expand_shared_formulas(source, pending)
        os.replace(pending, target)
    except BaseException:
        pending.unlink(missing_ok=True)
        raise
    return target


def default_cache_dir() -> Path:
    """Keep prepared packages outside every tree that is scanned or shipped."""
    override = os.environ.get("PP_ENGINE_CACHE_DIR")
    if override:
        return Path(override)
    return Path(tempfile.gettempdir()) / CACHE_DIRECTORY


def _entry(info):
    entry = zipfile.ZipInfo(info.filename, info.date_time)
    entry.compress_type = info.compress_type
    entry.create_system = info.create_system
    entry.external_attr = info.external_attr
    entry.internal_attr = info.internal_attr
    return entry


def _expand_sheet(reader, writer, stats):
    """Stream one worksheet, replacing shared formulas as they are read.

    Cells are visited in document order, so a shared master is always seen
    before the followers that inherit from it.
    """
    families: dict[bytes, _Family] = {}
    stats.sheets += 1
    buffer, finished = b"", False
    while not finished:
        chunk = reader.read(CHUNK_BYTES)
        if chunk:
            buffer += chunk
        else:
            finished = True
        limit = len(buffer) if finished else max(len(buffer) - TAIL_GUARD_BYTES, 0)
        position = 0
        while True:
            match = CELL_ELEMENT.search(buffer, position, limit)
            if match is None:
                break
            writer.write(buffer[position:match.start()])
            writer.write(_expand_cell(match.group(0), families, stats))
            position = match.end()
        if finished:
            writer.write(buffer[position:])
            buffer = b""
            continue
# Carry from the first cell that cannot be completed inside the window,
        # so a cell split across two reads is still examined rather than copied.
        # The search runs past the window so a cell starting on its last byte is
        # still seen; the flush itself never passes the window.
        boundary = CELL_START.search(buffer, position)
        stop = min(boundary.start(), limit) if boundary is not None else limit
        writer.write(buffer[position:stop])
        buffer = buffer[stop:]


def _expand_cell(cell, families, stats):
    if b"<f" not in cell:
        return cell
    start = cell.index(b"<f")
    attributes, body, end = _split_formula(cell, start)
    if not SHARED_TYPE.search(attributes):
        return cell
    index = SHARED_INDEX.search(attributes)
    if index is None:
        raise SharedFormulaError("shared formula has no si index")
    reference = CELL_REFERENCE.search(cell)
    if reference is None:
        raise SharedFormulaError("shared formula cell has no reference")
    address = reference.group(1).decode("utf-8")
    key = index.group(1)
    if body:
        # XML parsers turn raw CR/CRLF (Alt+Enter in Excel) into LF; openpyxl's tokenizer only accepts LF.
        text = unescape(body.decode("utf-8")).replace("\r\n", "\n").replace("\r", "\n")
        families[key] = _Family(address, text)
        stats.masters += 1
        return cell[:start] + b"<f>" + body + b"</f>" + cell[end:]
    family = families.get(key)
    if family is None:
        raise SharedFormulaError(f"shared formula si={key.decode()} has no master in this sheet")
    expanded = escape(translate_shared(family.text, family.origin, address))
    stats.followers += 1
    return cell[:start] + b"<f>" + expanded.encode("utf-8") + b"</f>" + cell[end:]


def _split_formula(cell, start):
    """Return the formula's open tag, its body and the offset just past it."""
    close = cell.index(b">", start)
    attributes = cell[start:close + 1]
    if cell[close - 1:close] == b"/":
        return attributes, b"", close + 1
    end = cell.index(b"</f>", close) + len(b"</f>")
    return attributes, cell[close + 1:end - len(b"</f>")], end