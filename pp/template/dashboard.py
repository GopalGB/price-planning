"""Safe OOXML pivot and slicer metadata import."""

from __future__ import annotations

import posixpath
import re
import math
from urllib.parse import unquote, urlsplit
import zipfile
import xml.etree.ElementTree as ET

from openpyxl.styles.numbers import BUILTIN_FORMATS


MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
PACKAGE_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
X14 = "{http://schemas.microsoft.com/office/spreadsheetml/2009/9/main}"
XDR = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"
XDR14 = "{http://schemas.microsoft.com/office/drawing/2010/slicer}"

PIVOT_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable"
CACHE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition"
DRAWING_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
SLICER_REL = "http://schemas.microsoft.com/office/2007/relationships/slicer"
SLICER_CACHE_REL = "http://schemas.microsoft.com/office/2007/relationships/slicerCache"
CELL_RANGE = re.compile(r"\$?[A-Z]{1,3}\$?[1-9][0-9]*(?::\$?[A-Z]{1,3}\$?[1-9][0-9]*)?")

SUBTOTALS = {
    "defaultSubtotal": "default", "sumSubtotal": "sum", "countASubtotal": "countA",
    "avgSubtotal": "average", "maxSubtotal": "max", "minSubtotal": "min",
    "productSubtotal": "product", "countSubtotal": "count", "stdDevSubtotal": "stdDev",
    "stdDevPSubtotal": "stdDevP", "varSubtotal": "variance", "varPSubtotal": "varianceP",
}


class DashboardMetadataError(ValueError):
    """Dashboard relationships or metadata are unsafe or inconsistent."""


class DashboardUnsupported(DashboardMetadataError):
    """A valid pivot this app cannot rebuild from cells; the template keeps Excel's saved values."""


class _DashboardXML(ET.TreeBuilder):
    def doctype(self, name, pubid, system):
        raise DashboardMetadataError("dashboard XML DTDs are forbidden")


def _bool(value, default=False):
    if value is None:
        return default
    if value in {"1", "true"}:
        return True
    if value in {"0", "false"}:
        return False
    raise DashboardMetadataError("dashboard boolean is invalid")


def _integer(value, label, minimum=0):
    try:
        result = int(value)
    except (TypeError, ValueError) as exc:
        raise DashboardMetadataError(f"{label} is invalid") from exc
    if result < minimum:
        raise DashboardMetadataError(f"{label} is invalid")
    return result


def _part_xml(archive, part):
    try:
        return ET.fromstring(archive.read(part), parser=ET.XMLParser(target=_DashboardXML()))
    except KeyError as exc:
        raise DashboardMetadataError("dashboard relationship target is missing") from exc
    except ET.ParseError as exc:
        raise DashboardMetadataError("dashboard related part is malformed") from exc


def _relationship_part(source_part):
    directory, name = posixpath.split(source_part)
    return posixpath.join(directory, "_rels", f"{name}.rels")


def _safe_target(archive, source_part, relationship):
    if relationship.get("TargetMode") == "External":
        raise DashboardMetadataError("external dashboard relationships are unsupported")
    raw = unquote(relationship.get("Target", ""))
    parsed = urlsplit(raw)
    if not raw or parsed.scheme or parsed.netloc or parsed.query or parsed.fragment or "\\" in raw:
        raise DashboardMetadataError("dashboard relationship target is invalid")
    joined = raw.lstrip("/") if raw.startswith("/") else posixpath.join(posixpath.dirname(source_part), raw)
    target = posixpath.normpath(joined)
    if target in {"", ".", ".."} or target.startswith("../") or target not in archive.namelist():
        raise DashboardMetadataError("dashboard relationship escapes or is missing")
    return target


def _related_parts(archive, source_part, relationship_type):
    rels_part = _relationship_part(source_part)
    if rels_part not in archive.namelist():
        return []
    root = _part_xml(archive, rels_part)
    relationships = root.findall(f"{PACKAGE_REL}Relationship")
    return [
        _safe_target(archive, source_part, relationship)
        for relationship in relationships if relationship.get("Type") == relationship_type
    ]


def _cell(cell):
    match = re.fullmatch(r"\$?([A-Z]{1,3})\$?([1-9][0-9]*)", cell)
    if match is None:
        raise DashboardMetadataError("dashboard cell range is invalid")
    column = 0
    for letter in match.group(1):
        column = column * 26 + ord(letter) - 64
    row = int(match.group(2))
    if row > 1_048_576 or column > 16_384:
        raise DashboardMetadataError("dashboard cell range exceeds the Excel grid")
    return row - 1, column - 1


def _bounds(reference):
    if not reference or CELL_RANGE.fullmatch(reference) is None:
        raise DashboardMetadataError("dashboard cell range is invalid")
    start, _, end = reference.partition(":")
    row1, col1 = _cell(start)
    row2, col2 = _cell(end or start)
    return [min(row1, row2), min(col1, col2), max(row1, row2), max(col1, col2)]


def _number(value):
    try:
        number = float(value)
    except (TypeError, ValueError) as exc:
        raise DashboardMetadataError("pivot cache number is invalid") from exc
    if not math.isfinite(number):
        raise DashboardMetadataError("pivot cache number must be finite")
    return int(number) if number.is_integer() else number


def _shared_item(node):
    kind = node.tag.removeprefix(MAIN)
    raw = node.get("v")
    if kind == "m":
        value = None
    elif kind == "n":
        value = _number(raw)
    elif kind == "b":
        value = _bool(raw)
    elif kind == "e":
        value = {"error": raw or "#VALUE!"}
    elif kind in {"s", "d"} and raw is not None:
        value = raw
    else:
        raise DashboardUnsupported("pivot cache item type is unsupported")
    return {"kind": {"s": "string", "n": "number", "d": "date", "b": "boolean",
                     "e": "error", "m": "missing"}[kind], "value": value}


def _item_caption(item):
    value = item["value"]
    if item["kind"] == "missing":
        return "(blank)"
    if item["kind"] == "boolean":
        return "TRUE" if value else "FALSE"
    if item["kind"] == "error":
        return value["error"]
    return str(value)


def _number_formats(archive):
    if "xl/styles.xml" not in archive.namelist():
        return {}
    root = _part_xml(archive, "xl/styles.xml")
    result = {}
    for node in root.findall(f"{MAIN}numFmts/{MAIN}numFmt"):
        identifier = _integer(node.get("numFmtId"), "number format id")
        result[identifier] = node.get("formatCode", "General")
    return result


def _cache_fields(cache, number_formats):
    output = []
    for index, node in enumerate(cache.findall(f"{MAIN}cacheFields/{MAIN}cacheField")):
        shared = node.find(f"{MAIN}sharedItems")
        items = [_shared_item(item) for item in shared] if shared is not None else []
        if shared is not None and shared.get("count") is not None:
            if _integer(shared.get("count"), "shared item count") != len(items):
                raise DashboardMetadataError("pivot cache shared item count does not match")
        format_id = _integer(node.get("numFmtId", "0"), "number format id")
        output.append({
            "index": index, "name": node.get("name", ""), "number_format_id": format_id,
            "number_format": number_formats.get(format_id, BUILTIN_FORMATS.get(format_id, "General")),
            "shared_items": items,
        })
    declared = cache.find(f"{MAIN}cacheFields")
    if declared is None or _integer(declared.get("count"), "cache field count") != len(output):
        raise DashboardUnsupported("pivot cache field count does not match")
    return output


def _pivot_item(node, shared_count):
    kind = node.get("t", "data")
    cache_index = None
    if kind == "data":
        cache_index = _integer(node.get("x"), "pivot item index")
        if cache_index >= shared_count:
            raise DashboardMetadataError("pivot item index is outside the cache field")
    return {"cache_index": cache_index, "kind": kind, "hidden": _bool(node.get("h"))}


def _pivot_fields(root, cache_fields):
    nodes = root.findall(f"{MAIN}pivotFields/{MAIN}pivotField")
    declared = root.find(f"{MAIN}pivotFields")
    if declared is None or _integer(declared.get("count"), "pivot field count") != len(cache_fields):
        raise DashboardMetadataError("pivot field count does not match the cache")
    if len(nodes) != len(cache_fields):
        raise DashboardMetadataError("pivot field definitions do not match the cache")
    fields = [_pivot_field(index, node, cache_fields[index]) for index, node in enumerate(nodes)]
    if any(field["sort"]["rank_by"] is not None
           and field["sort"]["rank_by"] >= len(fields) for field in fields):
        raise DashboardMetadataError("pivot sort field is outside the cache")
    return fields


def _pivot_field(index, node, cached):
    axis = {"axisRow": "row", "axisCol": "column", "axisPage": "page"}.get(node.get("axis"))
    items = node.find(f"{MAIN}items")
    parsed_items = [_pivot_item(item, len(cached["shared_items"])) for item in items] if items is not None else []
    if items is not None and _integer(items.get("count"), "pivot item count") != len(parsed_items):
        raise DashboardMetadataError("pivot item count does not match")
    return {
        **cached, "index": index, "caption": node.get("name", cached["name"]), "axis": axis,
        "data_field": _bool(node.get("dataField")), "outline": _bool(node.get("outline"), True),
        "subtotal_top": _bool(node.get("subtotalTop"), True), "show_all": _bool(node.get("showAll"), True),
        "subtotals": [name for attribute, name in SUBTOTALS.items() if _bool(node.get(attribute))],
        "sort": {"type": node.get("sortType", "manual"),
                 "rank_by": _integer(node.get("rankBy"), "pivot sort field") if node.get("rankBy") else None,
                 "data_source": _bool(node.get("dataSourceSort"))}, "items": parsed_items,
    }


def _field_references(root, section, field_count):
    result = []
    for node in root.findall(f"{MAIN}{section}/{MAIN}field"):
        index = _integer(node.get("x"), "pivot field reference", minimum=-2)
        if index != -2 and index >= field_count:
            raise DashboardMetadataError("pivot field reference is outside the cache")
        result.append(index)
    return result


def _page_fields(root, fields):
    result = []
    for node in root.findall(f"{MAIN}pageFields/{MAIN}pageField"):
        field = _integer(node.get("fld"), "page field index")
        if field >= len(fields):
            raise DashboardMetadataError("page field index is outside the cache")
        raw_item = node.get("item")
        item = _integer(raw_item, "page item index") if raw_item is not None else None
        if item is not None and item >= len(fields[field]["items"]):
            raise DashboardMetadataError("page item index is outside the pivot field")
        selected = fields[field]["items"][item] if item is not None else None
        result.append({
            "field": field, "item": item, "hierarchy": int(node.get("hier", "-1")),
            "cache_index": selected["cache_index"] if selected else None, "name": fields[field]["name"],
            "caption": fields[field]["caption"],
        })
    return result


def _data_fields(root, fields, number_formats):
    result = []
    for node in root.findall(f"{MAIN}dataFields/{MAIN}dataField"):
        field = _integer(node.get("fld"), "data field index")
        if field >= len(fields):
            raise DashboardMetadataError("data field index is outside the cache")
        format_id = _integer(node.get("numFmtId", str(fields[field]["number_format_id"])), "number format id")
        result.append({
            "field": field, "name": node.get("name", fields[field]["name"]),
            "subtotal": node.get("subtotal", "sum"), "number_format_id": format_id,
            "number_format": number_formats.get(format_id, BUILTIN_FORMATS.get(format_id, "General")),
        })
    return result


def _display_items(root, section):
    container = root.find(f"{MAIN}{section}")
    if container is None:
        return []
    output = []
    for node in container.findall(f"{MAIN}i"):
        data_index = node.get("i")
        output.append({
            "type": node.get("t", "data"), "repeat": _integer(node.get("r", "0"), "pivot item repeat"),
            "data_index": _integer(data_index, "pivot data index") if data_index is not None else 0,
            "indexes": [_integer(item.get("v"), "pivot display index")
                        for item in node.findall(f"{MAIN}x")],
        })
    if _integer(container.get("count"), "pivot display item count") != len(output):
        raise DashboardMetadataError("pivot display item count does not match")
    return output


def _pivot_style(root):
    node = root.find(f"{MAIN}pivotTableStyleInfo")
    if node is None:
        return None
    return {
        "name": node.get("name"), "show_row_headers": _bool(node.get("showRowHeaders")),
        "show_column_headers": _bool(node.get("showColHeaders")),
        "show_row_stripes": _bool(node.get("showRowStripes")),
        "show_column_stripes": _bool(node.get("showColStripes")),
        "show_last_column": _bool(node.get("showLastColumn")),
    }


def _cache_source(cache, sheet_names):
    source = cache.find(f"{MAIN}cacheSource/{MAIN}worksheetSource")
    if source is None or source.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"):
        raise DashboardUnsupported("external or missing pivot source is unsupported")
    sheet, reference = source.get("sheet"), source.get("ref")
    if sheet not in sheet_names:
        raise DashboardMetadataError("pivot source worksheet is missing")
    _bounds(reference)
    return {"sheet": sheet, "ref": reference}


def _pivot_location(root):
    node = root.find(f"{MAIN}location")
    if node is None:
        raise DashboardMetadataError("pivot location is missing")
    reference = node.get("ref")
    return {
        "ref": reference, "bounds": _bounds(reference),
        "first_header_row": _integer(node.get("firstHeaderRow", "0"), "first header row"),
        "first_data_row": _integer(node.get("firstDataRow", "0"), "first data row"),
        "first_data_col": _integer(node.get("firstDataCol", "0"), "first data column"),
    }


def _pivot_cache_id(cache, fallback):
    extension = cache.find(f".//{X14}pivotCacheDefinition")
    return _integer(extension.get("pivotCacheId"), "pivot cache id") if extension is not None else fallback


def _read_pivot(archive, part, context):
    sheet_names, number_formats = context
    root = _part_xml(archive, part)
    cache_parts = _related_parts(archive, part, CACHE_REL)
    if len(cache_parts) != 1:
        raise DashboardMetadataError("pivot table must have one cache definition")
    cache = _part_xml(archive, cache_parts[0])
    cache_id = _integer(root.get("cacheId"), "pivot cache id")
    cache_fields = cache.findall(f"{MAIN}cacheFields/{MAIN}cacheField")
    source = _cache_source(cache, sheet_names)
    start_row, start_col, _, end_col = _bounds(source["ref"])
    if end_col - start_col + 1 < len(cache_fields):
        raise DashboardUnsupported("pivot source width is narrower than cache fields")
    if any(field.get("databaseField") == "0" or field.get("fieldGroup") is not None
           for field in cache_fields):
        raise DashboardUnsupported("calculated and grouped pivot fields are unsupported")
    for field in root.findall(f"{MAIN}dataFields/{MAIN}dataField"):
        if field.get("subtotal", "sum") != "sum" or field.get("showDataAs") is not None:
            raise DashboardUnsupported("pivot data field feature is unsupported")
    fields = _pivot_fields(root, _cache_fields(cache, number_formats))
    return {
        "name": root.get("name", ""), "cache_id": _pivot_cache_id(cache, cache_id),
        "source": source, "location": _pivot_location(root),
        "captions": {"row": root.get("rowHeaderCaption", "Row Labels"),
                     "column": root.get("colHeaderCaption", "Column Labels"),
                     "data": root.get("dataCaption", "Values"),
                     "grand_total": root.get("grandTotalCaption", "Grand Total")},
        "grand_totals": {"rows": _bool(root.get("rowGrandTotals"), True),
                         "columns": _bool(root.get("colGrandTotals"), True)},
        "fields": fields, "row_fields": _field_references(root, "rowFields", len(fields)),
        "column_fields": _field_references(root, "colFields", len(fields)),
        "row_items": _display_items(root, "rowItems"),
        "column_items": _display_items(root, "colItems"),
        "page_fields": _page_fields(root, fields),
        "data_fields": _data_fields(root, fields, number_formats), "style": _pivot_style(root),
    }


def _marker(anchor, name):
    node = anchor.find(f"{XDR}{name}")
    if node is None:
        raise DashboardMetadataError("slicer drawing marker is missing")
    marker = {
        "row": _integer(node.findtext(f"{XDR}row"), "slicer anchor row"),
        "column": _integer(node.findtext(f"{XDR}col"), "slicer anchor column"),
        "row_offset": _integer(node.findtext(f"{XDR}rowOff", "0"), "slicer row offset"),
        "column_offset": _integer(node.findtext(f"{XDR}colOff", "0"), "slicer column offset"),
    }
    if marker["row"] >= 1_048_576 or marker["column"] >= 16_384:
        raise DashboardMetadataError("slicer anchor exceeds the Excel grid")
    return marker


def _drawing_anchors(archive, sheet_part):
    result = {}
    for part in _related_parts(archive, sheet_part, DRAWING_REL):
        root = _part_xml(archive, part)
        for anchor in root.findall(f"{XDR}twoCellAnchor"):
            slicer = anchor.find(f".//{XDR14}slicer")
            if slicer is not None and slicer.get("name"):
                result[slicer.get("name")] = {"from": _marker(anchor, "from"), "to": _marker(anchor, "to")}
    return result


def _slicer_cache(archive, part):
    root = _part_xml(archive, part)
    tabular = root.find(f"{X14}data/{X14}tabular")
    if tabular is None:
        raise DashboardMetadataError("only tabular slicer caches are supported")
    items = tabular.findall(f"{X14}items/{X14}i")
    declared = tabular.find(f"{X14}items")
    if declared is None or _integer(declared.get("count"), "slicer item count") != len(items):
        raise DashboardMetadataError("slicer item count does not match")
    return {
        "name": root.get("name", ""), "source_name": root.get("sourceName", ""),
        "pivot_cache_id": _integer(tabular.get("pivotCacheId"), "slicer pivot cache id"),
        "sort_order": tabular.get("sortOrder", "ascending"),
        "custom_list_sort": _bool(tabular.get("customListSort"), True),
        "show_missing": _bool(tabular.get("showMissing"), True),
        "cross_filter": tabular.get("crossFilter", "showItemsWithDataAtTop"),
        "pivot_tables": [{"sheet_id": _integer(node.get("tabId"), "slicer pivot sheet id"),
                          "name": node.get("name", "")}
                         for node in root.findall(f"{X14}pivotTables/{X14}pivotTable")],
        "items": [{"cache_index": _integer(node.get("x"), "slicer item index"),
                   "selected": _bool(node.get("s")), "no_data": _bool(node.get("nd"))}
                  for node in items],
    }


def _view_nodes(archive, sheet_part):
    nodes = []
    for part in _related_parts(archive, sheet_part, SLICER_REL):
        root = _part_xml(archive, part)
        nodes.extend(root.findall(f"{X14}slicer"))
    return nodes


def _matching_pivot(cache, pivots):
    names = {item["name"] for item in cache["pivot_tables"]}
    matches = [pivot for pivot in pivots
               if pivot["cache_id"] == cache["pivot_cache_id"] and (not names or pivot["name"] in names)]
    if not matches:
        raise DashboardMetadataError("slicer cache does not identify a pivot table")
    return matches[0]


def _slicer_items(cache, field):
    output = []
    shared = field["shared_items"]
    for item in cache["items"]:
        index = item["cache_index"]
        if index >= len(shared):
            raise DashboardMetadataError("slicer item index is outside the cache field")
        output.append({**item, "caption": _item_caption(shared[index])})
    return output


def _read_slicer(view, cache, linked):
    pivot, anchor, default_style = linked
    fields = [field for field in pivot["fields"] if field["name"] == cache["source_name"]]
    if len(fields) != 1:
        raise DashboardMetadataError("slicer field is missing or ambiguous")
    field = fields[0]
    return {
        "name": view.get("name", ""), "cache": cache["name"],
        "caption": view.get("caption", cache["source_name"]), "field": field["index"],
        "source_name": cache["source_name"], "pivot_cache_id": cache["pivot_cache_id"],
        "pivot_tables": cache["pivot_tables"], "items": _slicer_items(cache, field),
        "sort_order": cache["sort_order"], "custom_list_sort": cache["custom_list_sort"],
        "show_missing": cache["show_missing"], "cross_filter": cache["cross_filter"],
        "style": view.get("style", default_style),
        "column_count": _integer(view.get("columnCount", "1"), "slicer columns", 1),
        "row_height_emu": _integer(view.get("rowHeight"), "slicer row height", 1),
        "start_item": _integer(view.get("startItem", "0"), "slicer start item"),
        "show_caption": _bool(view.get("showCaption"), True),
        "locked_position": _bool(view.get("lockedPosition")), "anchor": anchor,
    }


def _sheet_slicers(archive, spec, context):
    caches, pivots, default_style = context
    anchors = _drawing_anchors(archive, spec["part"])
    output = []
    for view in _view_nodes(archive, spec["part"]):
        name, cache_name = view.get("name", ""), view.get("cache", "")
        if cache_name not in caches or name not in anchors:
            raise DashboardMetadataError("slicer view linkage is incomplete")
        cache = caches[cache_name]
        output.append(_read_slicer(view, cache, (
            _matching_pivot(cache, pivots), anchors[name], default_style,
        )))
    return output


def _default_slicer_style(archive):
    if "xl/styles.xml" not in archive.namelist():
        return "SlicerStyleLight1"
    root = _part_xml(archive, "xl/styles.xml")
    node = root.find(f".//{X14}slicerStyles")
    return node.get("defaultSlicerStyle", "SlicerStyleLight1") if node is not None else "SlicerStyleLight1"


def read_dashboard_metadata(archive: zipfile.ZipFile, sheet_specs: list[dict]) -> dict[int, dict]:
    """Return pivot and slicer metadata keyed by zero-based worksheet index."""
    names = archive.namelist()
    if len(names) != len(set(names)):
        raise DashboardMetadataError("duplicate package members are unsafe")
    sheet_names = {spec["name"] for spec in sheet_specs}
    formats = _number_formats(archive)
    default_style = _default_slicer_style(archive)
    result = {spec["index"]: {"pivots": [], "slicers": []} for spec in sheet_specs}
    for spec in sheet_specs:
        parts = _related_parts(archive, spec["part"], PIVOT_REL)
        result[spec["index"]]["pivots"] = [
            _read_pivot(archive, part, (sheet_names, formats)) for part in parts
        ]
    pivots = [pivot for metadata in result.values() for pivot in metadata["pivots"]]
    cache_parts = _related_parts(archive, "xl/workbook.xml", SLICER_CACHE_REL)
    parsed_caches = [_slicer_cache(archive, part) for part in cache_parts]
    caches = {cache["name"]: cache for cache in parsed_caches}
    if len(caches) != len(cache_parts):
        raise DashboardMetadataError("slicer cache names must be unique")
    for spec in sheet_specs:
        result[spec["index"]]["slicers"] = _sheet_slicers(
            archive, spec, (caches, pivots, default_style),
        )
    return result
