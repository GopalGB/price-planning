"""Streaming, byte-preserving OOXML transforms for workbook downloads."""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from html import unescape
import math
import os
import posixpath
import re
import shutil
import tempfile
from pathlib import Path
import xml.etree.ElementTree as ET
from xml.parsers import expat
from xml.sax.saxutils import escape
import zipfile

from openpyxl import load_workbook

from .xlsm import CHUNK, REL_NS, col_index, export_xlsm


CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
CUSTOM_NS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
CUSTOM_REL = f"{REL_NS}/custom-properties"
PIVOT_REL_SUFFIXES = ("/pivotCacheDefinition", "/slicerCache", "/slicerCacheDefinition")
MAX_CLEANUP_XML = 64 * 1024 * 1024
REL_ATTRIBUTE_NAMESPACES = {
    REL_NS,
    "http://purl.oclc.org/ooxml/officeDocument/relationships",
}
SHEET_TAG = re.compile(rb"<(?P<prefix>[A-Za-z_][\w.-]*:)?sheet\b[^>]*(?:/>|>.*?</(?P=prefix)sheet\s*>)", re.S)
NAME_TAG = re.compile(rb"<(?:[A-Za-z_][\w.-]*:)?definedName\b[^>]*>.*?</(?:[A-Za-z_][\w.-]*:)?definedName\s*>", re.S)
CELL_START = re.compile(rb"<(?P<prefix>[A-Za-z_][\w.-]*:)?c\b[^>]*>")
SHEET_DATA_END = re.compile(
    rb"(?:<(?:[A-Za-z_][\w.-]*:)?sheetData\b[^>]*/\s*>|"
    rb"</(?:[A-Za-z_][\w.-]*:)?sheetData\s*>)",
)
RELATIONSHIP_TAG = re.compile(
    rb"<(?P<prefix>[A-Za-z_][\w.-]*:)?Relationship\b[^>]*(?:/>|>.*?</(?P=prefix)Relationship\s*>)",
    re.S,
)
OVERRIDE_TAG = re.compile(
    rb"<(?P<prefix>[A-Za-z_][\w.-]*:)?Override\b[^>]*(?:/>|>.*?</(?P=prefix)Override\s*>)",
    re.S,
)
REMOVABLE_OBJECT_PREFIXES = (
    "xl/pivotTables/", "xl/pivotCache/", "xl/slicers/", "xl/slicerCaches/", "xl/charts/",
)


@dataclass(frozen=True)
class TransferRequest:
    template_path: Path
    cells: object
    out_path: Path
    template_hash: str
    copy_id: str
    input_revision: int
    fmt: str = "xlsx"
    selected: object = None
    results: object = None


@dataclass(frozen=True)
class DependencyContext:
    archive: object
    parts: set
    graph: dict
    workbook_xml: bytes
    sheets: list
    kept: list
    fmt: str


@dataclass(frozen=True)
class FormulaContext:
    sheets: list
    removed_names: set
    definitions: dict
    affected_names: set
    results: dict


@dataclass(frozen=True)
class RewriteContext:
    removed_names: set
    old_to_new: dict
    sheets: list


@dataclass(frozen=True)
class CleanupOptions:
    dropped_ids: set
    worksheet: bool = False


@dataclass
class RelationshipFrame:
    start: int
    tag_end: int
    self_closing: bool
    anchor: bool
    direct: bool
    descendant: bool = False


@dataclass(frozen=True)
class WritePlan:
    archive: object
    destination: Path
    removed: set
    replacements: dict
    frozen: dict
    cleanups: dict
    custom: bytes


@dataclass
class DependencyState:
    removed: set
    dropped: dict
    kept_parts: set
    removed_sheet_parts: set
    roots: set
    removed_identifiers: set


@dataclass
class FreezeState:
    shared: dict
    arrays: dict
    targets: dict


@dataclass
class CustomState:
    data: bytes
    prefix: bytes
    vt_prefix: bytes
    property_tag: object
    pids: list


@dataclass(frozen=True)
class TransformState:
    parts: set
    graph: dict
    workbook_xml: bytes
    sheets: list
    kept: list
    removed: set
    dropped: dict
    rewritten_workbook: bytes
    frozen: dict


@dataclass(frozen=True)
class ValidationState:
    names: set
    defaults: dict
    overrides: dict
    graph: dict


def _clone(info):
    clone = zipfile.ZipInfo(info.filename, info.date_time)
    clone.compress_type = info.compress_type
    clone.external_attr = info.external_attr
    clone.create_system = info.create_system
    return clone


def _local(tag):
    return tag.rsplit("}", 1)[-1]


def _attribute(tag, name):
    match = re.search(rb"\s(?:[A-Za-z_][\w.-]*:)?" + re.escape(name.encode()) + rb'="([^"]*)"', tag)
    return unescape(match.group(1).decode("utf-8")) if match else None


def _set_attribute(tag, name, value):
    pattern = rb"\s" + re.escape(name.encode()) + rb'="[^"]*"'
    encoded = escape(str(value), {'"': "&quot;"}).encode("utf-8")
    if re.search(pattern, tag):
        return re.sub(pattern, b" " + name.encode() + b'="' + encoded + b'"', tag, count=1)
    return tag[:-1] + b" " + name.encode() + b'="' + encoded + b'">'


def _relationship_source(name):
    if name == "_rels/.rels":
        return ""
    match = re.fullmatch(r"(?:(.*)/)?_rels/([^/]+)\.rels", name)
    if match is None:
        return None
    return posixpath.join(match.group(1) or "", match.group(2))


def _relationship_target(source, target):
    if target.startswith("/"):
        resolved = posixpath.normpath(target).lstrip("/")
    else:
        resolved = posixpath.normpath(posixpath.join(posixpath.dirname(source), target))
    if resolved in {"", ".", ".."} or resolved.startswith("../"):
        raise ValueError(f"unsafe OOXML relationship target: {target}")
    return resolved


def _relationship_graph(archive):
    graph = defaultdict(list)
    for info in archive.infolist():
        source = _relationship_source(info.filename)
        if source is None:
            continue
        with archive.open(info) as stream:
            root = ET.parse(stream).getroot()
        for node in root:
            if _local(node.tag) != "Relationship":
                continue
            target = None
            if node.get("TargetMode") != "External":
                target = _relationship_target(source, node.get("Target", ""))
            graph[source].append({
                "id": node.get("Id"), "type": node.get("Type", ""), "target": target,
            })
    return graph


def _sheet_records(workbook_xml, graph):
    root = ET.fromstring(workbook_xml)
    relationships = {item["id"]: item for item in graph["xl/workbook.xml"]}
    records = []
    for index, node in enumerate(item for item in root.iter() if _local(item.tag) == "sheet"):
        relationship = relationships.get(node.get(f"{{{REL_NS}}}id"))
        if relationship is None or relationship["target"] is None:
            raise ValueError("workbook sheet is missing its relationship")
        records.append({
            "index": index,
            "name": node.get("name"),
            "hidden": node.get("state") in {"hidden", "veryHidden"},
            "relationship": relationship["id"],
            "part": relationship["target"],
        })
    return records


def _closure(starts, graph):
    seen, todo = set(), list(starts)
    while todo:
        source = todo.pop()
        if source in seen:
            continue
        seen.add(source)
        todo.extend(
            item["target"] for item in graph.get(source, [])
            if item["target"] is not None
        )
    return seen


def _object_sources_removed(context, removed_names, affected_names):
    removed = set()
    folded = {name.casefold() for name in removed_names}
    affected = set(affected_names)
    for name in context.parts:
        if not name.endswith(".xml") or not name.startswith(("xl/pivotCache/", "xl/charts/")):
            continue
        with context.archive.open(name) as stream:
            root = ET.parse(stream).getroot()
        for node in root.iter():
            source_name = (node.get("name") or "").casefold()
            if (_local(node.tag) == "worksheetSource"
                    and ((node.get("sheet") or "").casefold() in folded
                         or source_name in affected)):
                removed.add(name)
                break
            if (_local(node.tag) == "f"
                    and (_sheet_names_in_formula(node.text or "", context.sheets) & removed_names
                         or _definition_references(node.text or "", affected))):
                removed.add(name)
                break
    return removed


def _dependency_plan(context):
    state = _dependency_state(context)
    _remove_dependent_objects(context, state)
    _finish_dependency_plan(context, state)
    return state.removed, state.dropped, state.removed_identifiers


def _dependency_state(context):
    kept_parts = {sheet["part"] for sheet in context.kept}
    state = DependencyState(
        {"xl/calcChain.xml"} & context.parts, defaultdict(set), kept_parts, set(), set(), set(),
    )
    for sheet in context.sheets:
        if sheet["part"] not in kept_parts:
            state.removed.add(sheet["part"])
            state.removed_sheet_parts.add(sheet["part"])
            state.dropped["xl/workbook.xml"].add(sheet["relationship"])
    state.roots.update(state.removed_sheet_parts)
    for source, relationships in context.graph.items():
        for relationship in relationships:
            if context.fmt == "xlsx" and relationship["type"].endswith("/vbaProject"):
                if relationship["target"]:
                    state.removed.add(relationship["target"])
                    state.roots.add(relationship["target"])
                state.dropped[source].add(relationship["id"])
    return state


def _remove_dependent_objects(context, state):
    if state.removed_sheet_parts:
        removed_names = {
            sheet["name"] for sheet in context.sheets
            if sheet["part"] in state.removed_sheet_parts
        }
        table_names = _removed_table_names(context, state)
        state.removed_identifiers.update(removed_names | table_names)
        _, affected = _definition_state(
            context.workbook_xml, context.sheets, state.removed_identifiers,
        )
        protected = _closure(state.kept_parts, context.graph)
        state.removed.update(_closure(state.roots, context.graph) - protected)
        state.removed.update(_object_sources_removed(
            context, removed_names, {name for _, name in affected} | table_names,
        ))
        changed = True
        while changed:
            changed = False
            for source, relationships in context.graph.items():
                if source in state.removed or not source.startswith(REMOVABLE_OBJECT_PREFIXES):
                    continue
                if any(item["target"] in state.removed for item in relationships):
                    state.removed.add(source)
                    changed = True


def _removed_table_names(context, state):
    names = set()
    for part in state.removed_sheet_parts:
        for relationship in context.graph.get(part, ()):
            target = relationship["target"]
            if not relationship["type"].endswith("/table") or target not in context.parts:
                continue
            with context.archive.open(target) as stream:
                root = ET.parse(stream).getroot()
            name = root.get("displayName") or root.get("name")
            if name:
                names.add(name.casefold())
    return names


def _finish_dependency_plan(context, state):
    if state.roots:
        protected = _closure(state.kept_parts, context.graph)
        state.removed.update(_closure(state.roots, context.graph) - protected)
    for source, relationships in context.graph.items():
        for relationship in relationships:
            if source in state.removed or relationship["target"] in state.removed:
                state.dropped[source].add(relationship["id"])
    for name in context.parts:
        source = _relationship_source(name)
        if source in state.removed:
            state.removed.add(name)


def _formula_code(formula):
    text = unescape(formula or "")
    output, index = [], 0
    while index < len(text):
        if text[index] != '"':
            output.append(text[index])
            index += 1
            continue
        output.append(" ")
        index += 1
        while index < len(text):
            if text[index] != '"':
                output.append(" ")
                index += 1
            elif index + 1 < len(text) and text[index + 1] == '"':
                output.extend((" ", " "))
                index += 2
            else:
                output.append(" ")
                index += 1
                break
    return "".join(output)


def _sheet_names_in_formula(formula, sheets):
    names, result = [sheet["name"] for sheet in sheets], set()
    indexes = {name.casefold(): index for index, name in enumerate(names)}
    pattern = re.compile(r"(?:'((?:[^']|'')+)'|([\w\\][\w.]*(?::[\w\\][\w.]*)?))!")
    for match in pattern.finditer(_formula_code(formula)):
        value = (match.group(1) or match.group(2)).replace("''", "'").casefold()
        endpoints = value.split(":", 1)
        if len(endpoints) == 2 and all(item in indexes for item in endpoints):
            start, end = sorted((indexes[endpoints[0]], indexes[endpoints[1]]))
            result.update(names[start:end + 1])
        elif value in indexes:
            result.add(names[indexes[value]])
    return result


def _definition_references(formula, known):
    code = _formula_code(formula)
    return {
        name for name in known
        if re.search(rf"(?i)(?<![A-Za-z0-9_.]){re.escape(name)}(?![A-Za-z0-9_.(])", code)
    }


def _definition_state(data, sheets, removed_names):
    definitions = {}
    for match in NAME_TAG.finditer(data):
        tag = match.group()
        name, local_id = _attribute(tag, "name"), _attribute(tag, "localSheetId")
        if not name:
            continue
        scope = int(local_id) if local_id is not None else None
        formula = unescape(re.sub(rb"^<[^>]*>|</[^>]*>$", b"", tag, flags=re.S).decode("utf-8"))
        definitions[(scope, name.casefold())] = formula
    known = {name for _, name in definitions}
    removed_indexes = {sheet["index"] for sheet in sheets if sheet["name"] in removed_names}
    affected, visiting = set(), set()

    def resolve(scope, name):
        return (scope, name) if (scope, name) in definitions else (None, name)

    def is_affected(key):
        if key in affected:
            return True
        if key in visiting or key not in definitions:
            return False
        visiting.add(key)
        scope, _ = key
        formula = definitions[key]
        broken = scope in removed_indexes or bool(
            _sheet_names_in_formula(formula, sheets) & removed_names
            or _definition_references(formula, {name.casefold() for name in removed_names})
        )
        if not broken:
            broken = any(is_affected(resolve(scope, name)) for name in _definition_references(formula, known))
        visiting.remove(key)
        if broken:
            affected.add(key)
        return broken

    for key in definitions:
        is_affected(key)
    return definitions, affected


def _formula_needs_freeze(formula, sheet_index, context):
    if _sheet_names_in_formula(formula, context.sheets) & context.removed_names:
        return True
    if _definition_references(
            formula, {name.casefold() for name in context.removed_names}):
        return True
    known = {name for _, name in context.definitions}
    for name in _definition_references(formula, known):
        key = ((sheet_index, name) if (sheet_index, name) in context.definitions
               else (None, name))
        if key in context.affected_names:
            return True
    return False


def _rewrite_workbook(data, dropped_relationships, context):
    definitions, affected = _definition_state(data, context.sheets, context.removed_names)

    def replace_sheet(match):
        return b"" if _attribute(match.group(), "id") in dropped_relationships else match.group()

    data = SHEET_TAG.sub(replace_sheet, data)

    def replace_name(match):
        tag = match.group()
        name, local_id = _attribute(tag, "name"), _attribute(tag, "localSheetId")
        local_index = int(local_id) if local_id is not None else None
        if name and (local_index, name.casefold()) in affected:
            return b""
        if local_index is not None:
            tag = _set_attribute(tag, "localSheetId", context.old_to_new[local_index])
        return tag

    data = NAME_TAG.sub(replace_name, data)
    if dropped_relationships:
        pivot = re.compile(rb"<(?P<prefix>[A-Za-z_][\w.-]*:)?(?:pivotCache|slicerCache)\b[^>]*(?:/>|>.*?</(?P=prefix)(?:pivotCache|slicerCache)\s*>)", re.S)
        data = pivot.sub(lambda match: b"" if _attribute(match.group(), "id") in dropped_relationships else match.group(), data)
    return data, definitions, affected


def _cell_formula(cell):
    for child in cell:
        if _local(child.tag) == "f":
            return child
    return None


def _iter_cells(stream):
    stack = []
    for event, element in ET.iterparse(stream, events=("start", "end")):
        if event == "start":
            stack.append(element)
            continue
        if _local(element.tag) == "c":
            yield element
        if len(stack) > 1 and _local(element.tag) in {"c", "row"}:
            stack[-2].remove(element)
        stack.pop()


def _range_references(reference):
    match = re.fullmatch(
        r"\$?([A-Z]+)\$?([1-9][0-9]*)(?::\$?([A-Z]+)\$?([1-9][0-9]*))?",
        reference or "",
    )
    if match is None:
        raise ValueError(f"unsupported array formula range: {reference}")
    end_col, end_row = match.group(3) or match.group(1), match.group(4) or match.group(2)
    col1, row1, col2, row2 = (
        col_index(match.group(1)), int(match.group(2)) - 1,
        col_index(end_col), int(end_row) - 1,
    )
    return {
        (f"{letters}{row + 1}", row, col)
        for row in range(min(row1, row2), max(row1, row2) + 1)
        for col in range(min(col1, col2), max(col1, col2) + 1)
        for letters in (_column_letters(col),)
    }


def _column_letters(column):
    value, output = column + 1, ""
    while value:
        value, remainder = divmod(value - 1, 26)
        output = chr(65 + remainder) + output
    return output


def _freeze_targets(archive, kept, context):
    state = FreezeState(defaultdict(set), defaultdict(set), defaultdict(dict))
    for sheet in kept:
        with archive.open(sheet["part"]) as source:
            for cell in _iter_cells(source):
                formula = _cell_formula(cell)
                if formula is not None and formula.text and _formula_needs_freeze(
                    formula.text, sheet["index"], context
                ):
                    if formula.get("t") == "shared":
                        state.shared[sheet["part"]].add(formula.get("si"))
                    if formula.get("t") == "array":
                        state.arrays[sheet["part"]].update(_range_references(formula.get("ref")))
                    if formula.get("t") not in {"shared", "array"} and cell.get("r"):
                        _record_freeze_target(state.targets, sheet, (cell.get("r"), context.results))
    for sheet in kept:
        _expand_freeze_groups(archive, sheet, (state, context))
    return state.targets


def _expand_freeze_groups(archive, sheet, payload):
    state, context = payload
    expected_array = {reference for reference, _, _ in state.arrays[sheet["part"]]}
    if not expected_array and not state.shared[sheet["part"]]:
        return
    found_array = set()
    with archive.open(sheet["part"]) as source:
        for cell in _iter_cells(source):
            formula, reference = _cell_formula(cell), cell.get("r")
            array_cell = reference in expected_array
            cell_affected = array_cell or (formula is not None and (
                formula.get("si") in state.shared[sheet["part"]] or bool(
                    formula.text and _formula_needs_freeze(
                        formula.text, sheet["index"], context,
                    )
                )
            ))
            if not cell_affected or not reference:
                continue
            if array_cell:
                found_array.add(reference)
            _record_freeze_target(state.targets, sheet, (reference, context.results))
    missing = expected_array - found_array
    if missing:
        reference = sorted(missing)[0]
        raise ValueError(f"array formula cell missing from worksheet: {sheet['name']}!{reference}")


def _record_freeze_target(targets, sheet, payload):
    reference, results = payload
    match = re.fullmatch(r"([A-Z]+)([0-9]+)", reference.replace("$", ""))
    if match is None:
        raise ValueError(f"invalid formula address: {reference}")
    key = (sheet["index"], int(match.group(2)) - 1, col_index(match.group(1)))
    if key not in results:
        raise ValueError(f"current calculation result missing for {sheet['name']}!{reference}")
    targets[sheet["part"]][reference] = results[key]


def _value_xml(value):
    if value is None:
        return None, b""
    if isinstance(value, bool):
        return "b", b"1" if value else b"0"
    if isinstance(value, dict) and set(value) == {"error"}:
        return "e", escape(value["error"]).encode("utf-8")
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if isinstance(value, float) and not math.isfinite(value):
            raise ValueError("current calculation result is non-finite")
        return None, repr(value).encode("utf-8")
    if isinstance(value, str):
        return "str", escape(value.replace("\r", "_x000D_")).encode("utf-8")
    raise ValueError("current calculation result has an unsupported type")


def _freeze_cell(cell, value):
    opening = CELL_START.match(cell)
    if opening is None:
        raise ValueError("formula cell XML is malformed")
    prefix = opening.group("prefix") or b""
    closing = b"</" + prefix + b"c>"
    if not cell.endswith(closing):
        raise ValueError("formula cell XML is malformed")
    inner = cell[opening.end():-len(closing)]
    inner = re.sub(rb"<" + re.escape(prefix) + rb"f\b[^>]*(?:/>|>.*?</" + re.escape(prefix) + rb"f\s*>)", b"", inner, flags=re.S)
    inner = re.sub(rb"<" + re.escape(prefix) + rb"v\b[^>]*(?:/>|>.*?</" + re.escape(prefix) + rb"v\s*>)", b"", inner, flags=re.S)
    tag = re.sub(rb'\s+t="[^"]*"', b"", opening.group())
    tag = re.sub(rb'\s+cm="[^"]*"', b"", tag)
    value_type, content = _value_xml(value)
    if value_type is not None:
        tag = _set_attribute(tag, "t", value_type)
    value_xml = b"" if value is None else b"<" + prefix + b"v>" + content + b"</" + prefix + b"v>"
    return tag + value_xml + inner + closing


def _freeze_sheet(source, target, values):
    buffer = b""
    while True:
        match = CELL_START.search(buffer)
        if match is None:
            chunk = source.read(CHUNK)
            if not chunk:
                target.write(buffer)
                return
            buffer += chunk
            if len(buffer) > CHUNK:
                target.write(buffer[:-CHUNK])
                buffer = buffer[-CHUNK:]
            continue
        prefix = match.group("prefix") or b""
        closing = b"</" + prefix + b"c>"
        self_closing = match.group().rstrip().endswith(b"/>")
        end = match.end() if self_closing else buffer.find(closing, match.end())
        if not self_closing:
            if end < 0:
                chunk = source.read(CHUNK)
                if not chunk:
                    raise ValueError("worksheet XML ended inside a cell")
                buffer += chunk
                continue
            end += len(closing)
        cell = buffer[match.start():end]
        target.write(buffer[:match.start()])
        reference = _attribute(match.group(), "r")
        if self_closing and reference in values:
            cell = re.sub(rb"/\s*>$", b">", cell) + closing
        target.write(_freeze_cell(cell, values[reference]) if reference in values else cell)
        buffer = buffer[end:]


def _filter_relationships(data, dropped_ids):
    if not dropped_ids:
        return data
    return RELATIONSHIP_TAG.sub(
        lambda match: b"" if _attribute(match.group(), "Id") in dropped_ids else match.group(), data
    )


def _xml_tag_end(data, start):
    quote = None
    for index in range(start, len(data)):
        value = data[index]
        if quote is not None:
            if value == quote:
                quote = None
        elif value in {34, 39}:
            quote = value
        elif value == 62:
            return index + 1
    raise ValueError("relationship-bearing XML is malformed")


def _dropped_relationship(attributes, dropped_ids):
    for name, value in attributes.items():
        namespace, separator, local = name.rpartition("}")
        if (separator and namespace in REL_ATTRIBUTE_NAMESPACES
                and local in {"id", "embed", "link"} and value in dropped_ids):
            return True
    return False


def _merged_ranges(ranges):
    merged = []
    for start, end in sorted(ranges):
        if merged and start <= merged[-1][1]:
            merged[-1] = merged[-1][0], max(merged[-1][1], end)
        else:
            merged.append((start, end))
    return merged


class RelationshipRangeParser:
    def __init__(self, data, dropped_ids):
        self.data, self.dropped_ids = data, dropped_ids
        self.parser = expat.ParserCreate(namespace_separator="}")
        self.stack, self.ranges = [], []

    def parse(self):
        self.parser.StartElementHandler = self.start
        self.parser.EndElementHandler = self.end
        try:
            self.parser.Parse(self.data, True)
        except expat.ExpatError as exc:
            raise ValueError("relationship-bearing XML is malformed") from exc
        return _merged_ranges(self.ranges)

    def start(self, name, attributes):
        start = self.parser.CurrentByteIndex
        tag_end = _xml_tag_end(self.data, start)
        frame = RelationshipFrame(
            start, tag_end, self.data[start:tag_end].rstrip().endswith(b"/>"),
            name.rpartition("}")[2] in {"oneCellAnchor", "twoCellAnchor", "absoluteAnchor"},
            _dropped_relationship(attributes, self.dropped_ids),
        )
        self.stack.append(frame)
        if frame.direct:
            for ancestor in self.stack:
                if ancestor.anchor:
                    ancestor.descendant = True

    def end(self, _name):
        frame = self.stack.pop()
        end = (frame.tag_end if frame.self_closing
               else _xml_tag_end(self.data, self.parser.CurrentByteIndex))
        if (frame.anchor and frame.descendant) or frame.direct:
            self.ranges.append((frame.start, end))


def _drop_relationship_references(data, dropped_ids, root=b""):
    if not dropped_ids:
        return data
    offset = len(root)
    ranges = RelationshipRangeParser(root + data, dropped_ids).parse()
    for start, end in reversed(ranges):
        if start >= offset:
            data = data[:start - offset] + data[end - offset:]
    empty = re.compile(
        rb"<(?P<prefix>(?:[A-Za-z_][\w.-]*:)?)"
        rb"(?P<name>pivotTableParts|oleObjects|controls)\b[^>]*>\s*"
        rb"</(?P=prefix)(?P=name)\s*>",
        re.S,
    )
    return empty.sub(b"", data)


def _custom_properties(existing, provenance):
    state = _custom_state(existing)
    for name, value in provenance.items():
        _set_custom_property(state, name, value)
    return state.data


def _custom_state(existing):
    if not existing:
        existing = (
            f'<?xml version="1.0" encoding="utf-8"?><cp:Properties xmlns:cp="{CUSTOM_NS}" '
            f'xmlns:vt="{VT_NS}"></cp:Properties>'
        ).encode()
    ET.fromstring(existing)
    root = re.search(rb"<(?P<prefix>[A-Za-z_][\w.-]*:)?Properties\b[^>]*>", existing)
    if root is None:
        raise ValueError("custom properties root is malformed")
    prefix = root.group("prefix") or b""
    start = root.group()
    vt_match = re.search(rb'xmlns:([A-Za-z_][\w.-]*)="' + re.escape(VT_NS.encode()) + rb'"', start)
    if vt_match is None:
        start = start[:-1] + f' xmlns:vt="{VT_NS}">'.encode()
        existing = existing[:root.start()] + start + existing[root.end():]
        vt_prefix = b"vt:"
    else:
        vt_prefix = vt_match.group(1) + b":"
    property_tag = re.compile(
        rb"<(?P<prefix>[A-Za-z_][\w.-]*:)?property\b[^>]*(?:/>|>.*?</(?P=prefix)property\s*>)",
        re.S,
    )
    pids = [
        int(value) for match in property_tag.finditer(existing)
        if (value := _attribute(match.group(), "pid")) and value.isdigit()
    ]
    return CustomState(existing, prefix, vt_prefix, property_tag, pids)


def _set_custom_property(state, name, value):
    encoded = escape(str(value)).encode("utf-8")
    match = next(
        (item for item in state.property_tag.finditer(state.data)
         if _attribute(item.group(), "name") == name),
        None,
    )
    content = b"<" + state.vt_prefix + b"lpwstr>" + encoded + b"</" + state.vt_prefix + b"lpwstr>"
    if match is not None:
        opening = match.group().find(b">") + 1
        closing = match.group().rfind(b"</")
        replacement = match.group()[:opening] + content + match.group()[closing:]
        state.data = state.data[:match.start()] + replacement + state.data[match.end():]
        return
    pid = max(state.pids, default=1) + 1
    state.pids.append(pid)
    addition = (
        b"<" + state.prefix + b'property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="'
        + str(pid).encode() + b'" name="' + name.encode() + b'">' + content
        + b"</" + state.prefix + b"property>"
    )
    closing = re.search(rb"</" + re.escape(state.prefix) + rb"Properties\s*>", state.data)
    if closing is None:
        raise ValueError("custom properties root is malformed")
    state.data = state.data[:closing.start()] + addition + state.data[closing.start():]


def _ensure_custom_relationship(data):
    for match in RELATIONSHIP_TAG.finditer(data):
        if _attribute(match.group(), "Type") == CUSTOM_REL:
            tag = _set_attribute(match.group(), "Target", "docProps/custom.xml")
            return data[:match.start()] + tag + data[match.end():]
    ids = {_attribute(match.group(), "Id") for match in RELATIONSHIP_TAG.finditer(data)}
    number = 1
    while f"rId{number}" in ids:
        number += 1
    root = re.search(rb"<(?P<prefix>[A-Za-z_][\w.-]*:)?Relationships\b[^>]*>", data)
    if root is None:
        raise ValueError("package relationships root is malformed")
    prefix = root.group("prefix") or b""
    addition = (b"<" + prefix + f'Relationship Id="rId{number}" Type="{CUSTOM_REL}" '
                'Target="docProps/custom.xml"/>'.encode("utf-8"))
    closing = re.search(rb"</" + re.escape(prefix) + rb"Relationships\s*>", data)
    if closing is None:
        raise ValueError("package relationships root is malformed")
    return data[:closing.start()] + addition + data[closing.start():]


def _rewrite_content_types(data, removed, fmt):
    data = OVERRIDE_TAG.sub(
        lambda match: b"" if (_attribute(match.group(), "PartName") or "").lstrip("/") in removed else match.group(),
        data,
    )
    if fmt == "xlsx":
        data = re.sub(
            rb'(<(?:[A-Za-z_][\w.-]*:)?Override\b(?=[^>]*\bPartName="/xl/workbook\.xml")[^>]*\bContentType=")[^"]*',
            rb'\1application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml', data,
        )
    if b'PartName="/docProps/custom.xml"' not in data:
        root = re.search(rb"<(?P<prefix>[A-Za-z_][\w.-]*:)?Types\b[^>]*>", data)
        if root is None:
            raise ValueError("content types root is malformed")
        prefix = root.group("prefix") or b""
        addition = (b"<" + prefix + b'Override PartName="/docProps/custom.xml" '
                    b'ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>')
        closing = re.search(rb"</" + re.escape(prefix) + rb"Types\s*>", data)
        if closing is None:
            raise ValueError("content types root is malformed")
        data = data[:closing.start()] + addition + data[closing.start():]
    return data


def _copy_cleaned_xml(source, target, options):
    if options.worksheet:
        buffer, root = b"", b""
        while True:
            chunk = source.read(CHUNK)
            if not chunk:
                raise ValueError("worksheet XML has no closing sheetData")
            buffer += chunk
            if not root:
                opening = re.search(rb"<(?:[A-Za-z_][\w.-]*:)?worksheet\b[^>]*>", buffer)
                root = opening.group() if opening is not None else b""
            match = SHEET_DATA_END.search(buffer)
            if match is not None:
                target.write(buffer[:match.end()])
                tail = buffer[match.end():] + source.read()
                if not root:
                    raise ValueError("worksheet XML root is malformed")
                target.write(_drop_relationship_references(tail, options.dropped_ids, root))
                return
            if len(buffer) > CHUNK * 2:
                target.write(buffer[:-CHUNK])
                buffer = buffer[-CHUNK:]
    data = source.read(MAX_CLEANUP_XML + 1)
    if len(data) > MAX_CLEANUP_XML:
        raise ValueError("relationship-bearing XML part is too large for surgical cleanup")
    target.write(_drop_relationship_references(data, options.dropped_ids))


def _write_transformed(plan):
    archive, destination = plan.archive, plan.destination
    fd, temporary = tempfile.mkstemp(suffix=".part", dir=Path(destination).parent)
    os.close(fd)
    try:
        with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as target:
            existing = set()
            for info in archive.infolist():
                name = info.filename
                if name in plan.removed:
                    continue
                existing.add(name)
                if name in plan.replacements:
                    target.writestr(_clone(info), plan.replacements[name])
                elif name in plan.frozen and name in plan.cleanups:
                    with archive.open(info) as source, tempfile.TemporaryFile() as staged:
                        _freeze_sheet(source, staged, plan.frozen[name])
                        staged.seek(0)
                        with target.open(_clone(info), "w") as stream:
                            _copy_cleaned_xml(
                                staged, stream, CleanupOptions(plan.cleanups[name], True)
                            )
                elif name in plan.frozen:
                    with archive.open(info) as source, target.open(_clone(info), "w") as stream:
                        _freeze_sheet(source, stream, plan.frozen[name])
                elif name in plan.cleanups:
                    with archive.open(info) as source, target.open(_clone(info), "w") as stream:
                        _copy_cleaned_xml(
                            source, stream,
                            CleanupOptions(plan.cleanups[name], name.startswith("xl/worksheets/")),
                        )
                else:
                    with archive.open(info) as source, target.open(_clone(info), "w") as stream:
                        shutil.copyfileobj(source, stream, CHUNK)
            if "docProps/custom.xml" not in existing:
                target.writestr("docProps/custom.xml", plan.custom)
        os.replace(temporary, destination)
    except BaseException:
        Path(temporary).unlink(missing_ok=True)
        raise


def _transform_package(source, destination, request):
    if request.fmt not in {"xlsx", "xlsm"}:
        raise ValueError("unsupported workbook format")
    with zipfile.ZipFile(source) as archive:
        state = _transform_state(archive, request)
        replacements, custom = _transform_replacements(archive, state, request)
        cleanups = _transform_cleanups(state)
        _write_transformed(WritePlan(
            archive, destination, state.removed, replacements, state.frozen, cleanups, custom,
        ))
    validate_package(destination, require_provenance=True)
    return [sheet["name"] for sheet in state.kept]


def _select_transfer_sheets(sheets, selected):
    if selected is None:
        return sheets
    selected = set(selected)
    names = {sheet["name"] for sheet in sheets}
    unknown = selected - names
    if not selected:
        raise ValueError("at least one sheet must be selected")
    if unknown:
        raise ValueError(f"unknown sheet selection: {', '.join(sorted(unknown))}")
    kept = [sheet for sheet in sheets if sheet["name"] in selected]
    if not any(not sheet["hidden"] for sheet in kept):
        raise ValueError("at least one visible sheet must be selected")
    return kept


def _transform_state(archive, request):
    parts = {info.filename for info in archive.infolist()}
    graph = _relationship_graph(archive)
    workbook_xml = archive.read("xl/workbook.xml")
    sheets = _sheet_records(workbook_xml, graph)
    kept = _select_transfer_sheets(sheets, request.selected)
    removed_names = {sheet["name"] for sheet in sheets} - {sheet["name"] for sheet in kept}
    dependency = DependencyContext(
        archive, parts, graph, workbook_xml, sheets, kept, request.fmt,
    )
    removed, dropped, removed_identifiers = _dependency_plan(dependency)
    old_to_new = {sheet["index"]: index for index, sheet in enumerate(kept)}
    rewritten, definitions, affected = _rewrite_workbook(
        workbook_xml, dropped["xl/workbook.xml"],
        RewriteContext(removed_identifiers, old_to_new, sheets),
    )
    formula = FormulaContext(
        sheets, removed_identifiers, definitions, affected, request.results or {},
    )
    frozen = _freeze_targets(archive, kept, formula) if removed_names else {}
    return TransformState(
        parts, graph, workbook_xml, sheets, kept, removed, dropped, rewritten, frozen,
    )


def _transform_replacements(archive, state, request):
    root_rels = _ensure_custom_relationship(
        _filter_relationships(archive.read("_rels/.rels"), state.dropped[""])
    )
    replacements = {
        "xl/workbook.xml": state.rewritten_workbook,
        "xl/_rels/workbook.xml.rels": _filter_relationships(
            archive.read("xl/_rels/workbook.xml.rels"), state.dropped["xl/workbook.xml"]
        ),
        "_rels/.rels": root_rels,
        "[Content_Types].xml": _rewrite_content_types(
            archive.read("[Content_Types].xml"), state.removed, request.fmt
        ),
    }
    for name in state.parts:
        source = _relationship_source(name)
        if source is not None and state.dropped.get(source) and name not in state.removed:
            replacements[name] = _filter_relationships(archive.read(name), state.dropped[source])
    existing = archive.read("docProps/custom.xml") if "docProps/custom.xml" in state.parts else None
    custom = _custom_properties(existing, {
        "TemplateHash": request.template_hash, "CopyId": request.copy_id,
        "InputRevision": str(request.input_revision),
    })
    replacements["docProps/custom.xml"] = custom
    return replacements, custom


def _transform_cleanups(state):
    cleanups = {}
    for source, dropped_ids in state.dropped.items():
        if (source and source in state.parts and source not in state.removed
                and source != "xl/workbook.xml" and source.endswith(".xml")):
            cleanups[source] = dropped_ids
    return cleanups


def _content_types(archive):
    root = ET.fromstring(archive.read("[Content_Types].xml"))
    defaults, overrides = {}, {}
    for node in root:
        if _local(node.tag) == "Default":
            defaults[node.get("Extension")] = node.get("ContentType")
        elif _local(node.tag) == "Override":
            overrides[node.get("PartName", "").lstrip("/")] = node.get("ContentType")
    return defaults, overrides


def _relationship_references(stream):
    references, stack = set(), []
    for event, element in ET.iterparse(stream, events=("start", "end")):
        if event == "start":
            stack.append(element)
            continue
        for name, value in element.attrib.items():
            if not name.startswith("{"):
                continue
            namespace, local = name[1:].split("}", 1)
            if namespace in REL_ATTRIBUTE_NAMESPACES and local in {"id", "embed", "link"}:
                references.add(value)
        if len(stack) > 1:
            stack[-2].remove(element)
        element.clear()
        stack.pop()
    return references


def validate_package(path, require_provenance=False):
    """Reject archives with dangling relationships, missing content types, or unreadable XML."""
    with zipfile.ZipFile(path) as archive:
        state = _validation_state(archive)
        for info in archive.infolist():
            _validate_package_member(archive, info, state)
        _validate_package_graph(state, require_provenance)
    workbook = load_workbook(path, read_only=True, data_only=False, keep_links=False)
    workbook.close()


def _validation_state(archive):
    infos = archive.infolist()
    names = {info.filename for info in infos}
    required = {"[Content_Types].xml", "_rels/.rels", "xl/workbook.xml"}
    if len(names) != len(infos) or required - names:
        raise ValueError("invalid OOXML package structure")
    defaults, overrides = _content_types(archive)
    dangling = set(overrides) - names
    if dangling:
        raise ValueError(f"content type override targets a missing part: {sorted(dangling)[0]}")
    graph = _relationship_graph(archive)
    for source, relationships in graph.items():
        identifiers = [item["id"] for item in relationships]
        if None in identifiers or len(identifiers) != len(set(identifiers)):
            raise ValueError(f"duplicate relationship id for {source or 'package root'}")
    return ValidationState(names, defaults, overrides, graph)


def _validate_package_member(archive, info, state):
    name, references = info.filename, set()
    if name.endswith(".xml"):
        with archive.open(info) as stream:
            references = _relationship_references(stream)
    elif name.endswith(".rels"):
        with archive.open(info) as stream:
            for _, element in ET.iterparse(stream, events=("end",)):
                element.clear()
    if name != "[Content_Types].xml" and not name.endswith(".rels"):
        extension = name.rsplit(".", 1)[-1] if "." in name else ""
        if name not in state.overrides and extension not in state.defaults:
            raise ValueError(f"missing content type for {name}")
    if name.endswith(".xml"):
        known = {item["id"] for item in state.graph.get(name, [])}
        missing = references - known
        if missing:
            raise ValueError(
                f"relationship reference has no matching relationship: {name} {sorted(missing)[0]}"
            )


def _validate_package_graph(state, require_provenance):
    for source, relationships in state.graph.items():
        if source and source not in state.names:
            raise ValueError(f"relationship source is missing: {source}")
        for relationship in relationships:
            if relationship["target"] is not None and relationship["target"] not in state.names:
                raise ValueError(f"relationship target is missing: {relationship['target']}")
    custom = any(
        item["type"] == CUSTOM_REL and item["target"] == "docProps/custom.xml"
        for item in state.graph[""]
    )
    if (require_provenance and
            (not custom or "docProps/custom.xml" not in state.names
             or "docProps/custom.xml" not in state.overrides)):
        raise ValueError("MISSING CUSTOM PROPERTY ROOT RELATIONSHIP")


def export_transfer(request):
    """Patch inputs into the original template, then apply the download transform."""
    out_path = Path(request.out_path)
    with tempfile.TemporaryDirectory(dir=out_path.parent) as directory:
        staged = Path(directory) / "staged.xlsm"
        export_xlsm(request.template_path, request.cells, staged)
        _transform_package(staged, out_path, request)
    return out_path
