"""The template as the calculator sees it: typed cells, formulas grouped into families, Excel's results.

A family is every formula whose text is the same once its relative references are written as offsets
from its own cell (Excel's R1C1 view). Each family is parsed and compiled once.
"""

from __future__ import annotations

import os
import pickle
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from defusedxml import ElementTree as SafeET

from ..engine.formualizer import NS, _cached_formula_map, _sheet_parts
from .values import XlError

STRING = re.compile(r'"(?:[^"]|"")*"')
SHEET = r"(?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!"
CELL = re.compile(r"(\$?)([A-Za-z]{1,3})(\$?)(\d+)")
REFERENCE = re.compile(
    r"(?<![\w.$'!\]])(" + SHEET + r")?(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?"
    r"|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d+:\$?\d+)(?![\w(.!\[])")
VERSION = 1


def column_number(letters: str) -> int:
    number = 0
    for letter in letters.upper():
        number = number * 26 + ord(letter) - 64
    return number


def relative_key(text: str, row: int, col: int) -> str:
    """Formula text with relative cell parts replaced by offsets from (row, col), both 0-based."""

    def cell(match):
        col_abs, letters, row_abs, digits = match.groups()
        c = column_number(letters) - 1
        r = int(digits) - 1
        return ("C" + str(c) if col_abs else f"C[{c - col}]") + ("R" + str(r) if row_abs else f"R[{r - row}]")

    def reference(match):
        prefix, body = match.group(1) or "", match.group(2)
        return prefix + CELL.sub(cell, body)

    parts, strings = STRING.split(text), STRING.findall(text)
    out = []
    for index, part in enumerate(parts):
        out.append(REFERENCE.sub(reference, part))
        if index < len(strings):
            out.append(strings[index])
    return "".join(out)


def scalar(value):
    """A cached workbook value in the calculator's form."""
    if isinstance(value, dict):
        return XlError(value.get("error", "#VALUE!"))
    if isinstance(value, bool) or value is None or isinstance(value, str):
        return value
    return float(value)


@dataclass
class Template:
    names: list
    formulas: dict  # (sheet, row, col) -> family id
    families: list  # family id -> (anchor (sheet, row, col), formula text, is_array)
    results: dict  # (sheet, row, col) -> Excel's saved result of a formula cell
    constants: dict  # (sheet, row, col) -> typed value
    formula_text: dict = field(repr=False, default_factory=dict)  # for the dependency index

    @classmethod
    def load(cls, prepared: Path) -> Template:
        cache = prepared.with_name(f"{prepared.name}.calc-{VERSION}.pickle")
        try:
            with cache.open("rb") as handle:
                return pickle.load(handle)  # nosec B301 # per-user engine cache written only by this app
        except FileNotFoundError:
            pass
        template = cls._build(prepared)
        pending = cache.with_name(f"{cache.name}.{os.getpid()}.tmp")
        with pending.open("wb") as handle:
            pickle.dump(template, handle, protocol=pickle.HIGHEST_PROTOCOL)
        os.replace(pending, cache)
        return template

    @classmethod
    def _build(cls, prepared: Path) -> Template:
        _, formulas, names, caches, _, _, _, arrays = _cached_formula_map(prepared)
        index, families, keyed = {}, [], {}
        for key in sorted(formulas):
            text = formulas[key]
            family_key = (key[0], key in arrays, relative_key(text, key[1], key[2]))
            family = keyed.get(family_key)
            if family is None:
                family = keyed[family_key] = len(families)
                families.append((key, text, key in arrays))
            index[key] = family
        results = {key: scalar(value) for key, value in caches.items()}
        return cls(names, index, families, results, _constants(prepared), formulas)


def _constants(prepared: Path) -> dict:
    values = {}
    with zipfile.ZipFile(prepared) as archive:
        shared = _shared_strings(archive)
        for sheet, (_, part) in enumerate(_sheet_parts(archive)):
            with archive.open(part) as stream:
                for _, element in SafeET.iterparse(stream):
                    if element.tag == f"{NS}c":
                        if element.find(f"{NS}f") is None:
                            value = _typed(element, shared)
                            if value is not None:
                                match = CELL.fullmatch(element.get("r", ""))
                                if match:
                                    values[(sheet, int(match.group(4)) - 1, column_number(match.group(2)) - 1)] = value
                        element.clear()
                    elif element.tag == f"{NS}row":
                        element.clear()
    return values


def _shared_strings(archive) -> list:
    try:
        root = SafeET.fromstring(archive.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    return ["".join(node.text or "" for node in item.iter(f"{NS}t")) for item in root.findall(f"{NS}si")]


def _typed(element, shared):
    kind = element.get("t")
    if kind == "inlineStr":
        return "".join(node.text or "" for node in element.iter(f"{NS}t"))
    raw = element.find(f"{NS}v")
    if raw is None or raw.text is None:
        return None
    text = raw.text
    if kind == "s":
        return shared[int(text)]
    if kind == "b":
        return text == "1"
    if kind == "e":
        return XlError(text)
    if kind == "str":
        return text
    return float(text)
