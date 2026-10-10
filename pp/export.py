"""Streaming patch export that preserves the original OOXML package."""

from __future__ import annotations

import html
from pathlib import Path
import re
import shutil
import tempfile
import zipfile
import xml.etree.ElementTree as ET

from openpyxl.utils.cell import get_column_letter


NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RNS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
CELL_START = re.compile(br"<c\b[^>]*\br=\"([^\"]+)\"[^>]*(?:/>|>)")
CALC_PR = re.compile(br"<(?P<prefix>[A-Za-z_][\w.-]*:)?calcPr\b[^>]*>")
FULL_CALC = re.compile(br"(\bfullCalcOnLoad\s*=\s*)([\"'])(?:[^\"']*)\2")


def export_copy(source: Path, overrides: dict, output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    handle = tempfile.NamedTemporaryFile(prefix="export-", suffix=".xlsm", dir=output_dir, delete=False)
    output = Path(handle.name)
    handle.close()
    by_sheet = _by_sheet(overrides)
    try:
        with zipfile.ZipFile(source) as original, zipfile.ZipFile(output, "w") as target:
            parts = _sheet_parts(original)
            patch_parts = {parts[index]: values for index, values in by_sheet.items()}
            for info in original.infolist():
                with original.open(info) as reader, target.open(info, "w") as writer:
                    if info.filename in patch_parts:
                        _patch_sheet(reader, writer, patch_parts[info.filename])
                    elif info.filename == "xl/workbook.xml":
                        writer.write(_calculation_mode(reader.read()))
                    else:
                        shutil.copyfileobj(reader, writer, length=1024 * 1024)
        return output
    except Exception:
        output.unlink(missing_ok=True)
        raise


def _sheet_parts(archive):
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {node.get("Id"): node.get("Target") for node in rels}
    parts = {}
    for index, sheet in enumerate(workbook.findall(f"{NS}sheets/{NS}sheet")):
        target = targets[sheet.get(f"{RNS}id")].lstrip("/")
        parts[index] = target if target.startswith("xl/") else "xl/" + target
    return parts


def _by_sheet(overrides):
    grouped = {}
    for (sheet, row, col), value in overrides.items():
        reference = f"{get_column_letter(col + 1)}{row + 1}".encode()
        grouped.setdefault(sheet, {})[reference] = value
    return grouped


def _patch_sheet(reader, writer, updates):
    buffer = b""
    while True:
        chunk = reader.read(64 * 1024)
        buffer += chunk
        final = not chunk
        while True:
            match = CELL_START.search(buffer)
            if match is None:
                keep = 0 if final else max(0, len(buffer) - 256)
                writer.write(buffer[:keep])
                buffer = buffer[keep:]
                break
            writer.write(buffer[:match.start()])
            cell_end = match.end() if match.group(0).endswith(b"/>") else buffer.find(b"</c>", match.end())
            if cell_end < 0 and not final:
                buffer = buffer[match.start():]
                break
            if not match.group(0).endswith(b"/>"):
                cell_end += len(b"</c>")
            fragment = buffer[match.start():cell_end]
            writer.write(_patch_cell(fragment, updates[match.group(1)]) if match.group(1) in updates else fragment)
            buffer = buffer[cell_end:]
        if final:
            writer.write(buffer)
            return


def _patch_cell(fragment, value):
    start_end = fragment.find(b">") + 1
    start = fragment[:start_end]
    if start.endswith(b"/>"):
        start = start[:-2] + b">"
        content = b""
    else:
        content = fragment[start_end:-4]
    content = re.sub(br"<v>.*?</v>|<is>.*?</is>", b"", content, flags=re.S)
    cell_type, payload = _cell_value(value)
    start = re.sub(br"\s+t=\"[^\"]*\"", b"", start[:-1])
    if cell_type:
        start += b' t="' + cell_type + b'"'
    return start + b">" + content + payload + b"</c>"


def _cell_value(value):
    if value is None:
        return b"", b""
    if isinstance(value, bool):
        return b"b", b"<v>" + (b"1" if value else b"0") + b"</v>"
    if isinstance(value, (int, float)):
        return b"", f"<v>{value}</v>".encode()
    if isinstance(value, dict):
        return b"e", f"<v>{html.escape(value['error'])}</v>".encode()
    return b"inlineStr", f'<is><t xml:space="preserve">{html.escape(value)}</t></is>'.encode()


def _calculation_mode(data):
    match = CALC_PR.search(data)
    if match is None:
        return _insert_calc_property(data)
    tag = _full_calc_property(match.group(0))
    return data[:match.start()] + tag + data[match.end():]


def _full_calc_property(tag):
    match = FULL_CALC.search(tag)
    if match is not None:
        replacement = match.group(1) + match.group(2) + b"1" + match.group(2)
        return tag[:match.start()] + replacement + tag[match.end():]
    closing = tag.rfind(b"/>")
    closing = closing if closing >= 0 else tag.rfind(b">")
    return tag[:closing] + b' fullCalcOnLoad="1"' + tag[closing:]


def _insert_calc_property(data):
    for name in (b"definedNames", b"sheets"):
        pattern = re.compile(br"</(?P<prefix>[A-Za-z_][\w.-]*:)?" + name + br"\s*>")
        match = pattern.search(data)
        if match is not None:
            prefix = match.group("prefix") or b""
            tag = b"<" + prefix + b'calcPr fullCalcOnLoad="1"/>'
            return data[:match.end()] + tag + data[match.end():]
    closing = re.search(br"</(?P<prefix>[A-Za-z_][\w.-]*:)?workbook\s*>", data)
    if closing is None:
        raise ValueError("workbook root closing tag is missing")
    prefix = closing.group("prefix") or b""
    tag = b"<" + prefix + b'calcPr fullCalcOnLoad="1"/>'
    return data[:closing.start()] + tag + data[closing.start():]
