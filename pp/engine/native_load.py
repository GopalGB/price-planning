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
import posixpath
import re
import xml.etree.ElementTree as ET
import tempfile
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from xml.sax.saxutils import escape, unescape

from openpyxl.formula.translate import Translator
from openpyxl.utils.cell import get_column_letter, range_boundaries


CHUNK_BYTES = 1 << 20
TAIL_GUARD_BYTES = 1 << 20
WORKSHEET_PREFIX = "xl/worksheets/"
CACHE_DIRECTORY = "pp-engine-cache"
CELL_ELEMENT = re.compile(rb"<c\b[^>]*?(?:/>|>.*?</c>)", re.DOTALL)
CELL_START = re.compile(rb"<c\b")
CELL_REFERENCE = re.compile(rb'\br\s*=\s*"([^"]*)"')
SHARED_TYPE = re.compile(rb'\bt\s*=\s*"shared"')
SHARED_INDEX = re.compile(rb'\bsi\s*=\s*"([^"]*)"')
DEFINED_NAME = re.compile(rb"(<definedName\b[^>]*>)(.*?)(</definedName>)", re.DOTALL)
# Bump when the prepared package changes shape, so stale cached copies are not reused.
PREPARE_VERSION = 2
MAIN_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
OOXML_ESCAPE = re.compile(r"_x([0-9A-Fa-f]{4})_")
NAME_START = re.compile(r"[A-Za-z_\\]")
NAME_CHARS = re.compile(r"[A-Za-z0-9_.\\]*")


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


@dataclass
class Table:
    """An Excel table: its sheet, full range, header and totals rows, column names."""
    sheet: str
    ref: str
    headers: int
    totals: int
    columns: list


def rewrite_structured(text: str, tables: dict) -> str:
    """Replace ``Table[...]`` references with the absolute range they denote.

    Strings and quoted sheet names are copied untouched; table names match
    case-insensitively like Excel. Forms that depend on the calling row
    (``[@Col]``, ``[#This Row]``) have no static range and raise.
    """
    if not tables:
        return text
    out, i, n = [], 0, len(text)
    while i < n:
        char = text[i]
        if char in "\"'":
            end = i + 1
            while end < n and not (text[end] == char and text[end + 1:end + 2] != char):
                end += 2 if text[end] == char else 1
            out.append(text[i:end + 1])
            i = end + 1
        elif NAME_START.match(char):
            end = NAME_CHARS.match(text, i + 1).end()
            table = tables.get(text[i:end].casefold())
            following = text[end:end + 1]
            if table is None or following in ("!", "("):
                out.append(text[i:end])
                i = end
            elif following == "[":
                close = _bracket_end(text, end)
                out.append(_resolve(table, text[end + 1:close]))
                i = close + 1
            else:
                out.append(_resolve(table, ""))
                i = end
        else:
            out.append(char)
            i += 1
    return "".join(out)


def _bracket_end(text, start):
    depth, i = 0, start
    while i < len(text):
        char = text[i]
        if char == "'":
            i += 2
            continue
        depth += {"[": 1, "]": -1}.get(char, 0)
        if depth == 0:
            return i
        i += 1
    raise SharedFormulaError("unterminated structured reference")


def _items(content):
    """Split the inside of ``Table[...]`` into specifiers, column names and ':'."""
    if not content.startswith("["):
        return [content] if content else []
    items, i = [], 0
    while i < len(content):
        char = content[i]
        if char == "[":
            close = _bracket_end(content, i)
            items.append(content[i + 1:close])
            i = close + 1
        elif char == ":":
            items.append(":")
            i += 1
        elif char in ", ":
            i += 1
        else:
            raise SharedFormulaError("malformed structured reference")
    return items


def _unescape_item(item):
    return re.sub(r"'(.)", r"\1", item)


def _resolve(table, content):
    items = _items(content)
    sections = {item.casefold() for item in items if item.startswith("#")}
    names = [item if item == ":" else _unescape_item(item)
             for item in items if not item.startswith("#")]
    if any(name.startswith("@") for name in names) or "#this row" in sections:
        raise SharedFormulaError("row-relative structured reference")
    first_col, top, last_col, bottom = range_boundaries(table.ref)
    spans = {"#all": (top, bottom), "#data": (top + table.headers, bottom - table.totals),
             "#headers": (top, top + table.headers - 1) if table.headers else None,
             "#totals": (bottom - table.totals + 1, bottom) if table.totals else None}
    rows = [spans.get(section, False) for section in sections or {"#data"}]
    if any(span is None or span is False for span in rows):
        raise SharedFormulaError("unsupported structured reference section")
    row_from, row_to = min(span[0] for span in rows), max(span[1] for span in rows)
    if names in ([], [":"]):
        col_from, col_to = first_col, last_col
    else:
        picked = [_column(table, name) + first_col for name in names if name != ":"]
        if len(picked) > 2 or (len(picked) == 2 and ":" not in names):
            raise SharedFormulaError("unsupported structured column list")
        col_from, col_to = min(picked), max(picked)
    corner = f"${get_column_letter(col_from)}${row_from}"
    other = f"${get_column_letter(col_to)}${row_to}"
    sheet = "'" + table.sheet.replace("'", "''") + "'"
    return f"{sheet}!{corner}" if corner == other else f"{sheet}!{corner}:{other}"


def _column(table, name):
    wanted = name.strip().casefold()
    for index, column in enumerate(table.columns):
        if column.casefold() == wanted:
            return index
    raise SharedFormulaError("structured reference names an unknown table column")


def read_tables(archive) -> dict:
    """Map each table's display name (casefolded) to its sheet and layout."""
    names = archive.namelist()
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {item.get("Id"): item.get("Target") for item in rels}
    tables = {}
    for sheet in workbook.iter(f"{MAIN_NS}sheet"):
        part = _part("xl", targets.get(sheet.get(RID), ""))
        sheet_rels = posixpath.join(posixpath.dirname(part), "_rels",
                                    posixpath.basename(part) + ".rels")
        if sheet_rels not in names:
            continue
        for rel in ET.fromstring(archive.read(sheet_rels)).iter(f"{REL_NS}Relationship"):
            if not rel.get("Type", "").endswith("/table"):
                continue
            table_part = _part(posixpath.dirname(part), rel.get("Target", ""))
            node = ET.fromstring(archive.read(table_part))
            columns = [OOXML_ESCAPE.sub(lambda m: chr(int(m.group(1), 16)), column.get("name", ""))
                       for column in node.iter(f"{MAIN_NS}tableColumn")]
            tables[node.get("displayName", node.get("name", "")).casefold()] = Table(
                sheet.get("name", ""), node.get("ref", ""),
                int(node.get("headerRowCount", "1")), int(node.get("totalsRowCount", "0")), columns)
    return tables


def _part(base, target):
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(base, target))


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
        tables = read_tables(archive)
        with zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as output:
            for info in archive.infolist():
                stats.parts += 1
                entry = _entry(info)
                if info.filename.startswith(WORKSHEET_PREFIX):
                    with archive.open(info) as reader, output.open(entry, "w") as writer:
                        _expand_sheet(reader, writer, stats, tables)
                elif info.filename == "xl/workbook.xml" and tables:
                    output.writestr(entry, DEFINED_NAME.sub(
                        lambda m: m.group(1) + _rewrite_body(m.group(2), tables) + m.group(3),
                        archive.read(info.filename)))
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
    identity = f"{PREPARE_VERSION}|{source.resolve()}|{status.st_size}|{status.st_mtime_ns}"
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


def _expand_sheet(reader, writer, stats, tables=None):
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
            writer.write(_expand_cell(match.group(0), families, stats, tables))
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


def _rewrite_body(body, tables):
    text = unescape(body.decode("utf-8"))
    rewritten = rewrite_structured(text, tables)
    return body if rewritten == text else escape(rewritten).encode("utf-8")


def _expand_cell(cell, families, stats, tables=None):
    if b"<f" not in cell:
        return cell
    start = cell.index(b"<f")
    attributes, body, end = _split_formula(cell, start)
    if not SHARED_TYPE.search(attributes):
        if not body or not tables:
            return cell
        rewritten = _rewrite_body(body, tables)
        return cell if rewritten is body else cell[:start] + attributes + rewritten + b"</f>" + cell[end:]
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
        text = rewrite_structured(text, tables or {})
        families[key] = _Family(address, text)
        stats.masters += 1
        return cell[:start] + b"<f>" + escape(text).encode("utf-8") + b"</f>" + cell[end:]
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