"""Bounded-memory OOXML importer for synthetic workbook templates."""

from __future__ import annotations

import hashlib
import json
import logging
import math
from dataclasses import dataclass
from pathlib import Path
import re
import shutil
import sqlite3
import tempfile
import zipfile
import xml.etree.ElementTree as ET

from openpyxl.formula.translate import Translator
from openpyxl.styles.colors import COLOR_INDEX
from openpyxl.styles.numbers import BUILTIN_FORMATS
from openpyxl.utils.cell import get_column_letter, range_boundaries
from openpyxl.utils.escape import unescape as unescape_ooxml

from .store import publish, replace_snapshot, staging_directory


LOGGER = logging.getLogger(__name__)


NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RNS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
DRAWING_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
CELL_REF = re.compile(r"([A-Z]+)([1-9][0-9]*)")
IMPORT_VERSION = 8
CF_RULE_TYPES = frozenset({"expression", "cellIs", "containsText", "duplicateValues", "beginsWith", "containsErrors"})
CF_BORDER_STYLES = {
    "thin": 1, "hair": 2, "dotted": 3, "dashed": 4,
    "dashDot": 5, "dashDotDot": 6, "double": 7, "medium": 8,
    "mediumDashed": 9, "mediumDashDot": 10, "mediumDashDotDot": 11,
    "slantDashDot": 12, "thick": 13,
}


class TemplateImportError(Exception):
    """A source archive could not be safely imported."""


@dataclass(frozen=True)
class StyleDefinition:
    index: int
    node: ET.Element
    base: ET.Element | None


@dataclass(frozen=True)
class CellContext:
    database: sqlite3.Connection
    sheet: int
    unlocked: set[int]


@dataclass(frozen=True)
class SheetContext:
    archive: zipfile.ZipFile
    database: sqlite3.Connection
    unlocked: set[int]
    dxf_styles: list[dict]
    destination: Path


@dataclass(frozen=True)
class SheetState:
    rows: dict
    metadata: dict
    cells: CellContext


@dataclass(frozen=True)
class FormulaSync:
    destination: Path
    database: sqlite3.Connection
    metadata: list[dict]


@dataclass(frozen=True)
class CapturedSource:
    path: Path
    digest: str
    name: str


def _source_hash(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _capture_source(source: Path, parent: Path) -> CapturedSource:
    parent.mkdir(parents=True, exist_ok=True)
    handle = tempfile.NamedTemporaryFile(prefix=".captured-", dir=parent, delete=False)
    temporary = Path(handle.name)
    try:
        with source.open("rb") as reader, handle:
            shutil.copyfileobj(reader, handle, length=1024 * 1024)
        return CapturedSource(temporary, _source_hash(temporary), source.name)
    except Exception:
        temporary.unlink(missing_ok=True)
        raise


def _write_json(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, separators=(",", ":"), sort_keys=True), encoding="utf-8")


def _preserve_source(source: Path, destination: Path) -> None:
    target = destination / "source.xlsm"
    if target.exists():
        return
    destination.mkdir(parents=True, exist_ok=True)
    handle = tempfile.NamedTemporaryFile(prefix=".source-", dir=destination, delete=False)
    temporary = Path(handle.name)
    try:
        with source.open("rb") as reader, handle:
            shutil.copyfileobj(reader, handle, length=1024 * 1024)
        temporary.replace(target)
    except Exception:
        temporary.unlink(missing_ok=True)
        raise


def _cell_ref(ref: str) -> tuple[int, int]:
    match = CELL_REF.fullmatch(ref.replace("$", ""))
    if match is None:
        raise TemplateImportError("invalid cell coordinate")
    col = 0
    for letter in match.group(1):
        col = col * 26 + ord(letter) - 64
    return int(match.group(2)) - 1, col - 1


def _bounds(ref: str) -> list[int]:
    start, _, end = ref.partition(":")
    row1, col1 = _cell_ref(start)
    row2, col2 = _cell_ref(end or start)
    return [min(row1, row2), min(col1, col2), max(row1, row2), max(col1, col2)]


def _canonical_color(value: str | None) -> str | None:
    if not value:
        return None
    value = value.lstrip("#")
    if len(value) == 8:
        value = value[2:]
    return f"#{value.upper()}" if len(value) == 6 else None


def _tinted(color: str | None, tint: str | None) -> str | None:
    if color is None or tint is None:
        return color
    amount = float(tint)
    channels = [int(color[index:index + 2], 16) for index in (1, 3, 5)]
    adjusted = [round(value * (1 + amount) if amount < 0 else value * (1 - amount) + 255 * amount)
                for value in channels]
    return "#" + "".join(f"{max(0, min(255, value)):02X}" for value in adjusted)


def _color(node: ET.Element | None, theme: dict[int, str] | None = None) -> str | None:
    if node is None:
        return None
    color = _canonical_color(node.get("rgb"))
    if color is None and node.get("theme") and theme:
        color = theme.get(int(node.get("theme")))
    if color is None and node.get("indexed"):
        index = int(node.get("indexed"))
        color = _canonical_color(COLOR_INDEX[index] if index < len(COLOR_INDEX) else None)
    if color is None and node.get("auto") in {"1", "true"}:
        color = "#000000"
    return _tinted(color, node.get("tint"))


def _theme_colors(archive: zipfile.ZipFile) -> dict[int, str]:
    if "xl/theme/theme1.xml" not in archive.namelist():
        return {}
    root = ET.fromstring(archive.read("xl/theme/theme1.xml"))
    scheme = root.find(f".//{DRAWING_NS}clrScheme")
    if scheme is None:
        return {}
    nodes = {item.tag.rsplit("}", 1)[-1]: item for item in scheme}
    order = ("lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink")
    colors = {}
    for index, name in enumerate(order):
        child = next(iter(nodes.get(name, ())), None)
        value = None if child is None else child.get("lastClr") or child.get("val")
        if color := _canonical_color(value):
            colors[index] = color
    return colors


def _font_name(font: ET.Element | None) -> str | None:
    if font is None:
        return None
    name = font.find(f"{NS}name")
    return name.get("val") if name is not None else None


def _component(items: list[ET.Element], index: int) -> ET.Element | None:
    return items[index] if 0 <= index < len(items) else None


def _component_id(xf: ET.Element, base: ET.Element | None, name: str) -> int:
    apply_name = "applyNumberFormat" if name == "numFmt" else f"apply{name[0].upper()}{name[1:]}"
    apply = xf.get(apply_name)
    source = base if apply in {"0", "false"} else xf
    return int(source.get(f"{name}Id", "0")) if source is not None else 0


def _locked(xf: ET.Element | None, inherited: bool = True) -> bool:
    if xf is None or xf.get("applyProtection") in {"0", "false"}:
        return inherited
    protection = xf.find(f"{NS}protection")
    if protection is None:
        return inherited
    return protection.get("locked", "1") not in {"0", "false"}


def _style_record(style: StyleDefinition, root: ET.Element, theme: dict[int, str]) -> dict:
    index, xf, base = style.index, style.node, style.base
    fonts = list(root.findall(f"{NS}fonts/{NS}font"))
    fills = list(root.findall(f"{NS}fills/{NS}fill"))
    borders = list(root.findall(f"{NS}borders/{NS}border"))
    font = _component(fonts, _component_id(xf, base, "font"))
    fill = _component(fills, _component_id(xf, base, "fill"))
    border = _component(borders, _component_id(xf, base, "border"))
    alignment = None if xf.get("applyAlignment") in {"0", "false"} else xf.find(f"{NS}alignment")
    if alignment is None and base is not None:
        alignment = base.find(f"{NS}alignment")
    num_id = _component_id(xf, base, "numFmt")
    custom = {int(item.get("numFmtId", "0")): item.get("formatCode", "General")
              for item in root.findall(f"{NS}numFmts/{NS}numFmt")}
    def side(name):
        node = border.find(f"{NS}{name}") if border is not None else None
        return node.get("style") if node is not None else None
    return {
        "id": index, "nf": custom.get(num_id, BUILTIN_FORMATS.get(num_id, "General")),
        "bold": font is not None and font.find(f"{NS}b") is not None,
        "italic": font is not None and font.find(f"{NS}i") is not None,
        "font_name": _font_name(font),
        "color": _color(font.find(f"{NS}color") if font is not None else None, theme),
        "fill": _color(fill.find(f"{NS}patternFill/{NS}fgColor") if fill is not None else None, theme),
        "align": alignment.get("horizontal") if alignment is not None else None,
        "valign": alignment.get("vertical") if alignment is not None else None,
        "wrap": alignment is not None and alignment.get("wrapText") in {"1", "true"},
        "font_size": float(font.find(f"{NS}sz").get("val")) if font is not None and font.find(f"{NS}sz") is not None else None,
        "locked": _locked(xf, _locked(base)),
        "border": {"t": side("top"), "r": side("right"), "b": side("bottom"), "l": side("left")},
    }


def _style_data(archive: zipfile.ZipFile) -> tuple[list[dict], set[int]]:
    if "xl/styles.xml" not in archive.namelist():
        return [_default_style()], set()
    root = ET.fromstring(archive.read("xl/styles.xml"))
    theme = _theme_colors(archive)
    bases = list(root.findall(f"{NS}cellStyleXfs/{NS}xf"))
    styles, unlocked = [], set()
    for index, xf in enumerate(root.findall(f"{NS}cellXfs/{NS}xf")):
        parent = int(xf.get("xfId", "0"))
        base = bases[parent] if parent < len(bases) else None
        styles.append(_style_record(StyleDefinition(index, xf, base), root, theme))
        if not _locked(xf, _locked(base)):
            unlocked.add(index)
    return styles or [_default_style()], unlocked


def _dxf_data(archive: zipfile.ZipFile) -> list[dict]:
    if "xl/styles.xml" not in archive.namelist():
        return []
    root = ET.fromstring(archive.read("xl/styles.xml"))
    theme = _theme_colors(archive)
    return [_dxf_style(node, theme) for node in root.findall(f"{NS}dxfs/{NS}dxf")]


def _dxf_style(node: ET.Element, theme: dict[int, str]) -> dict:
    style, font, fill = {}, node.find(f"{NS}font"), node.find(f"{NS}fill")
    if font is not None:
        bold, color = font.find(f"{NS}b"), _color(font.find(f"{NS}color"), theme)
        if bold is not None:
            style["bold"] = bold.get("val", "1").lower() not in {"0", "false"}
        if color is not None:
            style["color"] = color
    if fill is not None:
        pattern = fill.find(f"{NS}patternFill")
        if pattern is not None and pattern.get("patternType") == "none":
            style["fill"] = None
        elif pattern is not None:
            color = _color(pattern.find(f"{NS}bgColor"), theme)
            color = color or _color(pattern.find(f"{NS}fgColor"), theme)
            if color is not None:
                style["fill"] = color
    number = node.find(f"{NS}numFmt")
    if number is not None and number.get("formatCode") is not None:
        style["nf"] = number.get("formatCode")
    border = node.find(f"{NS}border")
    if border is not None:
        edges = {}
        for source, target in (("top", "t"), ("right", "r"), ("bottom", "b"), ("left", "l")):
            side = border.find(f"{NS}{source}")
            if side is None:
                continue
            border_style = CF_BORDER_STYLES.get(side.get("style"))
            if border_style is None:
                continue
            color = _color(side.find(f"{NS}color"), theme)
            edge = {"s": border_style}
            if color is not None:
                edge["cl"] = {"rgb": color}
            edges[target] = edge
        if edges:
            style["border"] = edges
    return style


def _default_style() -> dict:
    return {"id": 0, "nf": "General", "bold": False, "italic": False, "color": None,
            "fill": None, "align": None, "valign": None, "wrap": False, "font_size": None,
            "font_name": None, "locked": True,
            "border": {"t": None, "r": None, "b": None, "l": None}}


def _workbook_data(archive: zipfile.ZipFile) -> tuple[list[dict], list[dict]]:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {item.get("Id"): item.get("Target") for item in rels}
    sheets = []
    for index, node in enumerate(workbook.findall(f"{NS}sheets/{NS}sheet")):
        target = targets[node.get(f"{RNS}id")].lstrip("/")
        state = node.get("state", "visible")
        sheets.append({"index": index, "name": node.get("name", ""),
                       "part": target if target.startswith("xl/") else f"xl/{target}",
                       "hidden": False if state == "visible" else state})
    names = [{"name": node.get("name", ""), "formula": node.text or "",
              "local_sheet_id": int(node.get("localSheetId")) if node.get("localSheetId") else None,
              "hidden": node.get("hidden") in {"1", "true"}}
             for node in workbook.findall(f"{NS}definedNames/{NS}definedName")]
    return sheets, names


def _shared_strings(archive: zipfile.ZipFile, database: sqlite3.Connection) -> None:
    database.execute("CREATE TABLE strings (id INTEGER PRIMARY KEY, value TEXT NOT NULL)")
    if "xl/sharedStrings.xml" not in archive.namelist():
        return
    with archive.open("xl/sharedStrings.xml") as stream:
        stack, index = [], 0
        for event, element in ET.iterparse(stream, events=("start", "end")):
            if event == "start":
                stack.append(element)
                continue
            if element.tag == f"{NS}si":
                value = _visible_text(element)
                database.execute("INSERT INTO strings VALUES (?, ?)", (index, value))
                index += 1
                if len(stack) > 1:
                    stack[-2].remove(element)
            stack.pop()
    database.commit()


def _visible_text(element):
    if element is None:
        return ""
    parts = []
    for child in element:
        if child.tag == f"{NS}t":
            parts.append(child.text or "")
        elif child.tag == f"{NS}r":
            text = child.find(f"{NS}t")
            parts.append(text.text or "" if text is not None else "")
    return _decode_text("".join(unescape_ooxml(part) for part in parts))


def _decode_text(value):
    output, index = [], 0
    while index < len(value):
        high = ord(value[index])
        if 0xD800 <= high <= 0xDBFF and index + 1 < len(value):
            low = ord(value[index + 1])
            if 0xDC00 <= low <= 0xDFFF:
                output.append(chr(0x10000 + ((high - 0xD800) << 10) + low - 0xDC00))
                index += 2
                continue
        output.append(value[index])
        index += 1
    return "".join(output)


def _value(cell: ET.Element, database: sqlite3.Connection):
    kind, value = cell.get("t"), cell.findtext(f"{NS}v")
    if kind == "s" and value is not None:
        row = database.execute("SELECT value FROM strings WHERE id = ?", (int(value),)).fetchone()
        return row[0] if row else ""
    if kind == "b":
        return value == "1"
    if kind == "inlineStr":
        return _visible_text(cell.find(f"{NS}is"))
    if kind == "e":
        return {"error": value or "#VALUE!"}
    if value is None:
        return None
    try:
        return float(value) if any(marker in value for marker in ".eE") else int(value)
    except ValueError:
        return value


def _formula(cell: ET.Element, database: sqlite3.Connection, coordinate: tuple[int, int, int]) -> str | None:
    sheet, row, col = coordinate
    node = cell.find(f"{NS}f")
    if node is None:
        return None
    text, shared_id = node.text or "", node.get("si")
    if node.get("t") != "shared" or shared_id is None:
        return "=" + text.lstrip("=")
    if text:
        database.execute("INSERT OR REPLACE INTO formulas VALUES (?, ?, ?, ?, ?)",
                         (sheet, int(shared_id), row, col, "=" + text.lstrip("=")))
    return "@shared:" + shared_id


def _validation(node: ET.Element) -> dict:
    formulas = [node.findtext(f"{NS}formula1"), node.findtext(f"{NS}formula2")]
    ranges = [_bounds(ref) for ref in (node.get("sqref") or "").split()]
    result = {"type": node.get("type"), "ranges": ranges,
              "showErrorMessage": node.get("showErrorMessage") in {"1", "true"}}
    if node.get("errorStyle"):
        result["errorStyle"] = node.get("errorStyle")
    if node.get("operator"):
        result["operator"] = node.get("operator")
    if result["type"] == "list":
        if formulas[0] and formulas[0].startswith('"'):
            result["list"] = [_list_member(item) for item in formulas[0].strip('"').split(",")]
        elif formulas[0] is not None:
            result["formula1"] = formulas[0]
    elif formulas[0] is not None:
        result["min"] = _literal(formulas[0])
        if formulas[1] is not None:
            result["max"] = _literal(formulas[1])
    return result


def _list_member(item: str):
    # Excel enters a picked member as if typed: "1" becomes the number 1, "true" the boolean.
    if item.strip().upper() in {"TRUE", "FALSE"}:
        return item.strip().upper() == "TRUE"
    value = _literal(item)
    return value if not isinstance(value, float) or math.isfinite(value) else item


def _literal(value: str):
    try:
        return float(value) if any(marker in value for marker in ".eE") else int(value)
    except ValueError:
        return value


def _flush_block(directory: Path, block: int | None, rows: dict) -> None:
    if block is not None:
        _write_json(directory / f"{block}.json", rows)


def _record_cell(element: ET.Element, context: CellContext) -> tuple[int, int, dict]:
    reference = element.get("r")
    if reference is None:
        raise TemplateImportError("cell coordinate is missing")
    row, col = _cell_ref(reference)
    style = int(element.get("s", "0"))
    formula = _formula(element, context.database, (context.sheet, row, col))
    record = {"v": _value(element, context.database), "s": style, "k": "f" if formula is not None else "c"}
    if formula is not None:
        record["formula"] = formula
    elif style in context.unlocked:
        context.database.execute("INSERT INTO candidates VALUES (?, ?, ?)", (context.sheet, row, col))
    return row, col, record


def _sheet_pass(context: SheetContext, sheet_info) -> tuple[dict, bool]:
    index = sheet_info["index"]
    metadata = {**sheet_info, "rows": 0, "cols": 0, "freeze": {"rows": 0, "cols": 0},
                "col_widths": {}, "row_heights": {}, "hidden_rows": [], "hidden_cols": [],
                "col_styles": [], "row_styles": {}, "merges": [], "validations": [],
                "conditional_formats": [], "has_conditional_formats": False}
    metadata.pop("part")
    block_index, block_rows, protected = None, {}, False
    sheet_dir = context.destination / "sheets" / str(index)
    state = SheetState(block_rows, metadata, CellContext(context.database, index, context.unlocked))
    with context.archive.open(sheet_info["part"]) as stream:
        stack = []
        for event, element in ET.iterparse(stream, events=("start", "end")):
            if event == "start":
                stack.append(element)
                continue
            if element.tag == f"{NS}row":
                row_number = int(element.get("r", "1")) - 1
                next_block = row_number // 256
                if block_index is not None and next_block != block_index:
                    _flush_block(sheet_dir, block_index, block_rows)
                    block_rows = {}
                    state = SheetState(block_rows, metadata, state.cells)
                block_index = next_block
                _record_row(element, row_number, state)
            else:
                protected = _sheet_metadata(element, metadata, protected, context.dxf_styles)
            if len(stack) > 1 and element.tag == f"{NS}row":
                stack[-2].remove(element)
            stack.pop()
    _flush_block(sheet_dir, block_index, block_rows)
    sheet_dir.mkdir(parents=True, exist_ok=True)
    metadata["protected"] = protected
    return metadata, protected


def _record_row(element, row_number, state: SheetState) -> None:
    row_cells = {}
    for cell in element.findall(f"{NS}c"):
        row, col, record = _record_cell(cell, state.cells)
        row_cells[str(col)] = record
        state.metadata["rows"] = max(state.metadata["rows"], row + 1)
        state.metadata["cols"] = max(state.metadata["cols"], col + 1)
    if row_cells:
        state.rows[str(row_number % 256)] = row_cells
    if element.get("ht"):
        state.metadata["row_heights"][str(row_number)] = float(element.get("ht"))
    if element.get("hidden") in {"1", "true"}:
        state.metadata["hidden_rows"].append(row_number)
    if element.get("customFormat") in {"1", "true"} and element.get("s") is not None:
        state.metadata["row_styles"][str(row_number)] = int(element.get("s"))


def _sheet_metadata(element, metadata, protected, dxf_styles):
    if element.tag == f"{NS}dimension" and element.get("ref"):
        _, _, row2, col2 = _bounds(element.get("ref"))
        metadata["rows"] = max(metadata["rows"], row2 + 1)
        metadata["cols"] = max(metadata["cols"], col2 + 1)
    elif element.tag == f"{NS}sheetProtection":
        protected = element.get("sheet") in {"1", "true"}
    elif element.tag == f"{NS}pane" and element.get("state") == "frozen":
        metadata["freeze"] = {"rows": int(float(element.get("ySplit", "0"))),
                              "cols": int(float(element.get("xSplit", "0")))}
    elif element.tag == f"{NS}col":
        _column_metadata(element, metadata)
    elif element.tag == f"{NS}mergeCell" and element.get("ref"):
        metadata["merges"].append(_bounds(element.get("ref")))
    elif element.tag == f"{NS}dataValidation":
        metadata["validations"].append(_validation(element))
    elif element.tag == f"{NS}conditionalFormatting":
        metadata["conditional_formats"].extend(_conditional_formats(element, dxf_styles))
        metadata["has_conditional_formats"] = bool(metadata["conditional_formats"])
    return protected


def _conditional_formats(element, dxf_styles: list[dict]) -> list[dict]:
    ranges = [_cf_bounds(value) for value in (element.get("sqref") or "").split()]
    top_left = ranges[0][:2] if ranges else [0, 0]
    output = []
    for rule_index, rule in enumerate(element.findall(f"{NS}cfRule")):
        kind = rule.get("type")
        if kind not in CF_RULE_TYPES:
            continue
        try:
            dxf_id, priority = int(rule.get("dxfId", "-1")), int(rule.get("priority", str(2**31)))
        except ValueError:
            continue
        formulas = [item.text or "" for item in rule.findall(f"{NS}formula")]
        dxf = dxf_styles[dxf_id] if 0 <= dxf_id < len(dxf_styles) else {}
        for area_index, bounds in enumerate(ranges):
            output.append({"type": kind, "range": bounds, "top_left": top_left, "priority": priority,
                           "stop_if_true": rule.get("stopIfTrue") in {"1", "true"},
                           "operator": rule.get("operator"), "text": rule.get("text"),
                           "formulas": formulas, "dxf": dxf,
                           "rule_id": f"cf-{priority}-{rule_index}-{area_index}"})
    return output


def _cf_bounds(reference: str) -> list[int]:
    min_col, min_row, max_col, max_row = range_boundaries(reference)
    return [(min_row or 1) - 1, (min_col or 1) - 1,
            (max_row or 1_048_576) - 1, (max_col or 16_384) - 1]


def _column_metadata(element, metadata):
    first, last = int(element.get("min", "1")) - 1, int(element.get("max", "1")) - 1
    for col in range(first, last + 1):
        if element.get("width"):
            metadata["col_widths"][str(col)] = float(element.get("width"))
        if element.get("hidden") in {"1", "true"}:
            metadata["hidden_cols"].append(col)
    if element.get("style") is not None:
        metadata["col_styles"].append([first, last, int(element.get("style"))])
    metadata["cols"] = max(metadata["cols"], last + 1)


def _expand_shared_formulas(destination: Path, database: sqlite3.Connection, sheet: int) -> None:
    directory = destination / "sheets" / str(sheet)
    for path in sorted(directory.glob("*.json"), key=lambda item: int(item.stem)):
        rows = json.loads(path.read_text(encoding="utf-8"))
        changed = False
        for row_key, columns in rows.items():
            for col_key, record in columns.items():
                formula = record.get("formula", "")
                if not formula.startswith("@shared:"):
                    continue
                shared_id = int(formula.split(":", 1)[1])
                master = database.execute("SELECT row, col, formula FROM formulas WHERE sheet=? AND id=?",
                                          (sheet, shared_id)).fetchone()
                if master is None:
                    raise TemplateImportError("shared formula master is missing")
                origin = f"{get_column_letter(master[1] + 1)}{master[0] + 1}"
                target = f"{get_column_letter(int(col_key) + 1)}{int(path.stem) * 256 + int(row_key) + 1}"
                record["formula"] = Translator(master[2], origin=origin).translate_formula(target)
                changed = True
        if changed:
            _write_json(path, rows)


def _write_formula_block(sync: FormulaSync, key, cells) -> None:
    sheet, block = key
    path = sync.destination / "sheets" / str(sheet) / f"{block}.json"
    data = json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}
    for row, col, value, anchor in cells:
        row_data = data.setdefault(str(row % 256), {})
        record = row_data.setdefault(str(col), {"v": None, "s": 0, "k": "c"})
        record.update({"v": value, "k": "f"})
        if not anchor:
            record.pop("formula", None)
    _write_json(path, data)


def _synchronize_formula_results(source: Path, sync: FormulaSync) -> None:
    from ..engine.formualizer import EngineError, FormualizerAdapter

    try:
        engine = FormualizerAdapter.load(source)
    except EngineError as exc:
        raise TemplateImportError(f"template formula capability error: {exc}") from exc
    grouped = {}
    for sheet, row, col, value in engine.result_snapshot():
        anchor = (sheet, row, col) in engine.formula_anchors
        grouped.setdefault((sheet, row // 256), []).append((row, col, value, anchor))
        sync.database.execute("DELETE FROM candidates WHERE sheet=? AND row=? AND col=?",
                              (sheet, row, col))
        sync.metadata[sheet]["rows"] = max(sync.metadata[sheet]["rows"], row + 1)
        sync.metadata[sheet]["cols"] = max(sync.metadata[sheet]["cols"], col + 1)
    for key, cells in grouped.items():
        _write_formula_block(sync, key, cells)


def _apply_input_mask(destination: Path, database: sqlite3.Connection, protected: set[int]) -> None:
    for sheet in protected:
        rows = database.execute("SELECT row, col FROM candidates WHERE sheet=? ORDER BY row,col", (sheet,))
        by_block: dict[int, list[tuple[int, int]]] = {}
        for row, col in rows:
            by_block.setdefault(row // 256, []).append((row, col))
        for block, cells in by_block.items():
            path = destination / "sheets" / str(sheet) / f"{block}.json"
            data = json.loads(path.read_text(encoding="utf-8"))
            for row, col in cells:
                data[str(row % 256)][str(col)]["k"] = "i"
            _write_json(path, data)


def _write_input_mask(path: Path, database: sqlite3.Connection, protected: set[int]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as stream:
        stream.write("[")
        first = True
        for sheet in sorted(protected):
            for row, col in database.execute("SELECT row,col FROM candidates WHERE sheet=? ORDER BY row,col", (sheet,)):
                if not first:
                    stream.write(",")
                stream.write(json.dumps(f"{sheet}:{row}:{col}"))
                first = False
        stream.write("]")


def _import_archive(source: CapturedSource, staging: Path) -> None:
    database = sqlite3.connect(staging / "import.sqlite")
    database.execute("CREATE TABLE candidates (sheet INTEGER, row INTEGER, col INTEGER, PRIMARY KEY(sheet,row,col))")
    database.execute("CREATE TABLE formulas (sheet INTEGER, id INTEGER, row INTEGER, col INTEGER, formula TEXT, PRIMARY KEY(sheet,id))")
    try:
        with zipfile.ZipFile(source.path) as archive:
            _shared_strings(archive, database)
            styles, unlocked = _style_data(archive)
            dxf_styles = _dxf_data(archive)
            sheet_specs, names = _workbook_data(archive)
            metadata, protected = [], set()
            context = SheetContext(archive, database, unlocked, dxf_styles, staging)
            for spec in sheet_specs:
                info, is_protected = _sheet_pass(context, spec)
                metadata.append(info)
                if is_protected:
                    protected.add(spec["index"])
                _write_json(staging / "progress.json", {"complete": False,
                            "sheets_completed": len(metadata), "sheets_total": len(sheet_specs)})
                LOGGER.info("Template import progress: sheet %d/%d", len(metadata), len(sheet_specs))
            database.commit()
        for info in metadata:
            _expand_shared_formulas(staging, database, info["index"])
        _synchronize_formula_results(source.path, FormulaSync(staging, database, metadata))
        _apply_input_mask(staging, database, protected)
        _write_input_mask(staging / "input_mask.json", database, protected)
        _write_json(staging / "styles.json", styles)
        _write_json(staging / "meta.json", {"hash": source.digest, "name": source.name,
                    "import_version": IMPORT_VERSION, "sheets": metadata, "defined_names": names})
        _write_json(staging / "progress.json", {"complete": True,
                    "sheets_completed": len(metadata), "sheets_total": len(metadata)})
    finally:
        database.close()
        (staging / "import.sqlite").unlink(missing_ok=True)


def _snapshot_current(destination: Path) -> bool:
    try:
        meta = json.loads((destination / "meta.json").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return False
    return (meta.get("hash") == destination.name
            and meta.get("import_version") == IMPORT_VERSION
            and (destination / "source.xlsm").is_file())


def _build_snapshot(source: CapturedSource, destination: Path) -> Path:
    staging = staging_directory(destination.parent)
    try:
        _import_archive(source, staging)
        source.path.replace(staging / "source.xlsm")
        if destination.exists():
            return replace_snapshot(staging, destination)
        return publish(staging, destination)
    except Exception as exc:
        shutil.rmtree(staging, ignore_errors=True)
        if isinstance(exc, TemplateImportError):
            raise
        if isinstance(exc, (OSError, ValueError, IndexError, sqlite3.Error,
                            ET.ParseError, zipfile.BadZipFile, KeyError)):
            raise TemplateImportError("template import failed") from exc
        raise


def import_template(source: str | Path, data_dir: str | Path) -> Path:
    source, data_dir = Path(source), Path(data_dir)
    captured = _capture_source(source, data_dir / "templates")
    destination = data_dir / "templates" / captured.digest
    try:
        if _snapshot_current(destination):
            return destination
        return _build_snapshot(captured, destination)
    finally:
        captured.path.unlink(missing_ok=True)
