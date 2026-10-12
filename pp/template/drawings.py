"""Pictures and shapes from sheet drawings, drawn over the grid where Excel draws them."""

from __future__ import annotations

import base64
import posixpath
import xml.etree.ElementTree as ET
import zipfile

from .colors import Palette

DRAWING_XDR = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"
DRAWING_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
R_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PACKAGE_RELS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
DRAWING_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
VML_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"
EMU_PER_PIXEL = 9525
ANCHORS = {f"{DRAWING_XDR}twoCellAnchor", f"{DRAWING_XDR}oneCellAnchor", f"{DRAWING_XDR}absoluteAnchor"}
IMAGE_TYPES = {"png": "image/png", "jpeg": "image/jpeg", "jpg": "image/jpeg", "gif": "image/gif",
               "bmp": "image/bmp", "webp": "image/webp"}
# shortcut: pictures travel inline in the workbook payload; serve them from a route if a template carries big ones
MAX_PICTURE_BYTES = 4 * 1024 * 1024
LINE_WIDTHS = (0, 1, 2.67, 4)  # theme line styles 1-3 of the Office theme, in pixels
POINT_PIXELS = 96 / 72


def _emu_pixels(value: str | None) -> float:
    return round(int(value or 0) / EMU_PER_PIXEL, 2)


def _anchor_point(node: ET.Element) -> dict:
    return {"row": int(node.findtext(f"{DRAWING_XDR}row", "0")),
            "col": int(node.findtext(f"{DRAWING_XDR}col", "0")),
            "dx": _emu_pixels(node.findtext(f"{DRAWING_XDR}colOff")),
            "dy": _emu_pixels(node.findtext(f"{DRAWING_XDR}rowOff"))}


def anchor_box(anchor: ET.Element) -> dict:
    """Where an anchor sits: a start cell plus an end cell or a pixel size."""
    if anchor.tag == f"{DRAWING_XDR}absoluteAnchor":
        pos, ext = anchor.find(f"{DRAWING_XDR}pos"), anchor.find(f"{DRAWING_XDR}ext")
        start = {"row": 0, "col": 0, "dx": _emu_pixels(pos.get("x")), "dy": _emu_pixels(pos.get("y"))}
        return {"from": start, "size": {"w": _emu_pixels(ext.get("cx")), "h": _emu_pixels(ext.get("cy"))}}
    start = _anchor_point(anchor.find(f"{DRAWING_XDR}from"))
    if anchor.tag == f"{DRAWING_XDR}twoCellAnchor":
        return {"from": start, "to": _anchor_point(anchor.find(f"{DRAWING_XDR}to"))}
    ext = anchor.find(f"{DRAWING_XDR}ext")
    return {"from": start, "size": {"w": _emu_pixels(ext.get("cx")), "h": _emu_pixels(ext.get("cy"))}}


def anchors(root: ET.Element):
    """Top-level anchors as Excel shows them: the mc:Choice branch, never the fallback copy."""
    for node in root:
        if node.tag in ANCHORS:
            yield node
        elif node.tag == f"{MC}AlternateContent":
            choice = node.find(f"{MC}Choice")
            if choice is not None:
                yield from (child for child in choice if child.tag in ANCHORS)


def related_parts(archive: zipfile.ZipFile, part: str) -> dict[str, tuple[str, str, str]]:
    """Relationship id -> (type, resolved part or raw target, target mode) for one package part."""
    directory, name = posixpath.split(part)
    rels_part = posixpath.join(directory, "_rels", f"{name}.rels")
    if rels_part not in archive.namelist():
        return {}
    found = {}
    for item in ET.fromstring(archive.read(rels_part)).findall(f"{PACKAGE_RELS}Relationship"):
        target, mode = item.get("Target", ""), item.get("TargetMode", "Internal")
        if mode != "External" and not target.startswith("#"):
            target = target.lstrip("/") if target.startswith("/") else posixpath.normpath(
                posixpath.join(directory, target))
        found[item.get("Id", "")] = (item.get("Type", ""), target, mode)
    return found


def sheet_drawing_parts(archive: zipfile.ZipFile, sheet_part: str) -> list[tuple[str, str]]:
    return [(kind, target) for kind, target, _ in related_parts(archive, sheet_part).values()
            if kind in {DRAWING_REL, VML_REL} and target in archive.namelist()]


def _hidden(shape: ET.Element) -> bool:
    props = next(shape.iter(f"{DRAWING_XDR}cNvPr"), None)
    return props is not None and props.get("hidden") in {"1", "true"}


def _picture(shape: ET.Element, archive: zipfile.ZipFile, rels: dict) -> dict | None:
    blip = next(shape.iter(f"{DRAWING_A}blip"), None)
    _, target, _ = rels.get(blip.get(f"{R_ID}embed", "") if blip is not None else "", ("", "", ""))
    kind = IMAGE_TYPES.get(target.rsplit(".", 1)[-1].lower())
    if not kind or target not in archive.namelist() or archive.getinfo(target).file_size > MAX_PICTURE_BYTES:
        return None  # EMF/WMF and oversized pictures are left out
    return {"kind": "picture", "src": f"data:{kind};base64,{base64.b64encode(archive.read(target)).decode()}"}


def _style_ref(shape: ET.Element, name: str, palette: Palette) -> tuple[int, str | None]:
    ref = shape.find(f"{DRAWING_XDR}style/{DRAWING_A}{name}")
    if ref is None:
        return 0, None
    index = ref.get("idx", "0")  # fontRef names a theme font ("minor"/"major") instead of a number
    return (int(index) if index.isdigit() else int(index != "none")), palette.drawing(ref)


def _fill(props: ET.Element | None, shape: ET.Element, palette: Palette) -> str | None:
    for child in props if props is not None else ():
        tag = child.tag.removeprefix(DRAWING_A)
        if tag == "noFill":
            return None
        if tag == "solidFill":
            return palette.drawing(child)
        if tag == "gradFill":  # shortcut: a gradient shows its first stop color
            return palette.drawing(next(child.iter(f"{DRAWING_A}gs"), None))
    index, color = _style_ref(shape, "fillRef", palette)
    return color if index else None


def _line(props: ET.Element | None, shape: ET.Element, palette: Palette) -> dict | None:
    line = props.find(f"{DRAWING_A}ln") if props is not None else None
    index, theme_color = _style_ref(shape, "lnRef", palette)
    if line is not None and line.find(f"{DRAWING_A}noFill") is not None:
        return None
    color = palette.drawing(line.find(f"{DRAWING_A}solidFill")) if line is not None else None
    color = color or (theme_color if index else None)
    if not color:
        return None
    width = _emu_pixels(line.get("w")) if line is not None and line.get("w") else LINE_WIDTHS[min(index, 3)]
    return {"color": color, "width": max(width, 1)}


def _run(run: ET.Element, default: dict, palette: Palette) -> dict:
    props = run.find(f"{DRAWING_A}rPr")
    get = props.get if props is not None else (lambda key, fallback=None: fallback)
    color = palette.drawing(props.find(f"{DRAWING_A}solidFill")) if props is not None else None
    latin = props.find(f"{DRAWING_A}latin") if props is not None else None
    return {"t": "".join(text.text or "" for text in run.iter(f"{DRAWING_A}t")) if run.tag != f"{DRAWING_A}br" else "\n",
            "b": get("b", default["b"]) in {"1", "true"}, "i": get("i", default["i"]) in {"1", "true"},
            "sz": round(int(get("sz", default["sz"])) / 100 * POINT_PIXELS, 2),
            "color": color or default["color"],
            "font": latin.get("typeface") if latin is not None and not latin.get("typeface", "+").startswith("+")
            else None}


def _text(shape: ET.Element, palette: Palette) -> dict | None:
    body = shape.find(f"{DRAWING_XDR}txBody")
    if body is None:
        return None
    _, font_color = _style_ref(shape, "fontRef", palette)
    paragraphs = []
    for paragraph in body.findall(f"{DRAWING_A}p"):
        props = paragraph.find(f"{DRAWING_A}pPr")
        defaults = props.find(f"{DRAWING_A}defRPr") if props is not None else None
        default = {"b": "0", "i": "0", "sz": "1100", "color": font_color or "#000000"}
        if defaults is not None:
            default.update({key: defaults.get(key) for key in ("b", "i", "sz") if defaults.get(key)})
            default["color"] = palette.drawing(defaults.find(f"{DRAWING_A}solidFill")) or default["color"]
        runs = [_run(run, default, palette) for run in paragraph
                if run.tag in {f"{DRAWING_A}r", f"{DRAWING_A}fld", f"{DRAWING_A}br"}]
        paragraphs.append({"align": props.get("algn", "l") if props is not None else "l", "runs": runs})
    if not any(run["t"].strip() for paragraph in paragraphs for run in paragraph["runs"]):
        return None
    props = body.find(f"{DRAWING_A}bodyPr")
    get = props.get if props is not None else (lambda key, fallback=None: fallback)
    return {"anchor": get("anchor", "t"), "wrap": get("wrap", "square") != "none",
            "inset": [_emu_pixels(get(key, default)) for key, default in
                      (("lIns", "91440"), ("tIns", "45720"), ("rIns", "91440"), ("bIns", "45720"))],
            "paragraphs": paragraphs}


def _link(shape: ET.Element, rels: dict) -> dict | None:
    """An in-workbook hyperlink ("#'Sheet'!A1") on the shape; other links are not followed."""
    click = next(shape.iter(f"{DRAWING_A}hlinkClick"), None)
    _, target, _ = rels.get(click.get(f"{R_ID}id", "") if click is not None else "", ("", "", ""))
    if not target.startswith("#") or "!" not in target:
        return None
    sheet, cell = target[1:].rsplit("!", 1)
    if sheet.startswith("'") and sheet.endswith("'"):
        sheet = sheet[1:-1].replace("''", "'")
    return {"sheet": sheet, "cell": cell.replace("$", "")}


def _shape(shape: ET.Element, palette: Palette, rels: dict) -> dict | None:
    props = shape.find(f"{DRAWING_XDR}spPr")
    geometry = props.find(f"{DRAWING_A}prstGeom") if props is not None else None
    adjust = geometry.find(f"{DRAWING_A}avLst/{DRAWING_A}gd") if geometry is not None else None
    found = {"kind": "shape", "geom": geometry.get("prst", "rect") if geometry is not None else "rect",
             "fill": _fill(props, shape, palette), "line": _line(props, shape, palette),
             "shadow": props is not None and props.find(f"{DRAWING_A}effectLst/{DRAWING_A}outerShdw") is not None,
             "text": _text(shape, palette), "link": _link(shape, rels)}
    if adjust is not None and (adjust.get("fmla") or "").startswith("val "):
        found["adjust"] = int(adjust.get("fmla")[4:]) / 100000
    if not (found["fill"] or found["line"] or found["text"]):
        return None  # nothing Excel would paint
    return found


def drawing_objects(root: ET.Element, part: str, archive: zipfile.ZipFile, palette: Palette) -> list[dict]:
    """Visible pictures and shapes of one drawing part. Macro buttons, slicers and charts are drawn elsewhere."""
    rels = related_parts(archive, part)
    found = []
    for anchor in anchors(root):
        # shortcut: group shapes (xdr:grpSp) are skipped; draw their children if a template uses them
        for shape in anchor:
            if shape.tag not in {f"{DRAWING_XDR}pic", f"{DRAWING_XDR}sp"} or _hidden(shape) or shape.get("macro"):
                continue
            try:
                item = (_picture(shape, archive, rels) if shape.tag == f"{DRAWING_XDR}pic"
                        else _shape(shape, palette, rels))
                if item:
                    found.append({**anchor_box(anchor), **item})
            except (AttributeError, TypeError, ValueError):  # a malformed object is left out, not the template
                continue
    return found
