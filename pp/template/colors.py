"""Excel colors (ARGB, theme + tint, indexed palette) resolved to CSS hex at import."""

from __future__ import annotations

import colorsys
import posixpath
import xml.etree.ElementTree as ET
import defusedxml.ElementTree as SafeET
import zipfile

from defusedxml.common import DefusedXmlException
from openpyxl.styles.colors import COLOR_INDEX

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
PACKAGE_RELS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
THEME_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
# Excel numbers theme colors light-first: theme 0 is lt1 (bg1), 1 is dk1 (tx1), then lt2, dk2.
THEME_ORDER = ("lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5",
               "accent6", "hlink", "folHlink")
# Office 2013+ default theme, used when a workbook carries no theme part.
DEFAULT_THEME = ("#FFFFFF", "#000000", "#E7E6E6", "#44546A", "#4472C4", "#ED7D31", "#A5A5A5",
                 "#FFC000", "#5B9BD5", "#70AD47", "#0563C1", "#954F72")


def _hex(argb: str) -> str | None:
    value = argb.strip().upper()
    if len(value) == 8:
        value = value[2:]  # Excel stores ARGB; the alpha byte is not used for cells
    return f"#{value}" if len(value) == 6 and all(c in "0123456789ABCDEF" for c in value) else None


def tinted(color: str, tint: float) -> str:
    """Excel's tint: scale HLS lightness toward black (tint < 0) or white (tint > 0)."""
    if not tint:
        return color
    red, green, blue = (int(color[i:i + 2], 16) / 255 for i in (1, 3, 5))
    hue, light, saturation = colorsys.rgb_to_hls(red, green, blue)
    light = light * (1 + tint) if tint < 0 else light * (1 - tint) + tint
    return "#" + "".join(f"{round(part * 255):02X}" for part in colorsys.hls_to_rgb(hue, light, saturation))


def _scheme_color(node: ET.Element) -> str | None:
    for child in node:
        if child.tag == f"{A}srgbClr":
            return _hex(child.get("val", ""))
        if child.tag == f"{A}sysClr":
            return _hex(child.get("lastClr", ""))
    return None


def theme_colors(archive: zipfile.ZipFile) -> tuple[str, ...]:
    """The workbook theme's twelve colors in Excel's theme index order."""
    try:
        rels = SafeET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        target = next(item.get("Target", "") for item in rels.findall(f"{PACKAGE_RELS}Relationship")
                      if item.get("Type") == THEME_REL)
        part = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
        scheme = SafeET.fromstring(archive.read(part)).find(f"{A}themeElements/{A}clrScheme")
    except (KeyError, StopIteration, ET.ParseError, DefusedXmlException):
        return DEFAULT_THEME
    if scheme is None:
        return DEFAULT_THEME
    found = {child.tag.removeprefix(A): _scheme_color(child) for child in scheme}
    return tuple(found.get(name) or default for name, default in zip(THEME_ORDER, DEFAULT_THEME))


def indexed_palette(styles: ET.Element | None) -> list[str]:
    custom = [] if styles is None else [
        item.get("rgb", "") for item in styles.findall(f"{NS}colors/{NS}indexedColors/{NS}rgbColor")]
    return [_hex(value) or "#000000" for value in (custom or COLOR_INDEX)]


class Palette:
    """Resolves spreadsheetml `<color>` elements and DrawingML color choices to "#RRGGBB"."""

    def __init__(self, theme: tuple[str, ...] = DEFAULT_THEME, indexed: list[str] | None = None):
        self.theme, self.indexed = theme, indexed or indexed_palette(None)

    def cell(self, node: ET.Element | None) -> str | None:
        if node is None or node.get("auto") in {"1", "true"}:
            return None
        rgb, theme, indexed = node.get("rgb"), node.get("theme"), node.get("indexed")
        try:
            if rgb:
                color = _hex(rgb)
            elif theme is not None:
                color = self.theme[int(theme)] if 0 <= int(theme) < len(self.theme) else None
            elif indexed is not None:
                # 64 and 65 are the system foreground and background: automatic, like auto="1"
                color = self.indexed[int(indexed)] if 0 <= int(indexed) < min(64, len(self.indexed)) else None
            else:
                color = None
            return tinted(color, float(node.get("tint") or 0)) if color else None
        except ValueError:  # a malformed attribute shows as no color, like Excel's automatic color
            return None

    def drawing(self, node: ET.Element | None) -> str | None:
        """A DrawingML color choice (srgbClr, schemeClr, sysClr, prstClr) with lumMod/lumOff/tint/shade."""
        if node is None:
            return None
        for child in node:
            tag = child.tag.removeprefix(A)
            if tag == "srgbClr":
                color = _hex(child.get("val", ""))
            elif tag == "sysClr":
                color = _hex(child.get("lastClr", ""))
            elif tag == "schemeClr":
                color = self._scheme(child.get("val", ""))
            else:
                continue
            return _modified(color, child) if color else None
        return None

    def _scheme(self, name: str) -> str | None:
        alias = {"bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2"}.get(name, name)
        return self.theme[THEME_ORDER.index(alias)] if alias in THEME_ORDER else None


def _modified(color: str, node: ET.Element) -> str:
    red, green, blue = (int(color[i:i + 2], 16) / 255 for i in (1, 3, 5))
    hue, light, saturation = colorsys.rgb_to_hls(red, green, blue)
    for change in node:
        try:
            amount = int(change.get("val") or 0) / 100000
        except ValueError:
            continue
        tag = change.tag.removeprefix(A)
        if tag == "lumMod":
            light *= amount
        elif tag == "lumOff":
            light += amount
        elif tag == "tint":
            light = light * amount + (1 - amount)
        elif tag == "shade":
            light *= amount
    light = min(1.0, max(0.0, light))
    return "#" + "".join(f"{round(part * 255):02X}" for part in colorsys.hls_to_rgb(hue, light, saturation))
