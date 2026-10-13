"""Which formulas can an edit reach: a static reverse-dependency index.

Built once per prepared template from formula text alone. References are parsed
statically (cells, ranges, whole rows and columns, sheet-qualified references and
defined names); formulas whose targets cannot be known statically (OFFSET,
INDIRECT, 3-D or unresolved structured references) and volatile functions are
always treated as affected, so the cone over-approximates and never misses.
Range membership is answered for many ranges at once with 2-D prefix sums.
"""

from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.sax.saxutils import unescape

import numpy as np

STRING = re.compile(r'"(?:[^"]|"")*"')
EXTERNAL = re.compile(
    r"(?:'\[\d+\](?:[^']|'')*'|\[\d+\][^!\s,()]*)!\$?[A-Za-z]{0,3}\$?\d*(?::\$?[A-Za-z]{0,3}\$?\d*)?")
SHEET = r"(?:'(?P<q>(?:[^']|'')+)'|(?P<u>[A-Za-z_ -￿][\w. -￿]*))!"
CELL = r"\$?[A-Za-z]{1,3}\$?\d+"
REFERENCE = re.compile(
    r"(?<![\w.$'!\]])(?:" + SHEET + r")?(?P<a>" + CELL + r"(?::" + CELL
    + r")?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d+:\$?\d+)(?![\w(.!\[])")
NAME = re.compile(r"(?<![\w.$'!\]])([A-Za-z_\\][\w.]*)(?![\w(.!\[])")
DYNAMIC = re.compile(r"\b(?:INDIRECT|OFFSET)\s*\(", re.I)
VOLATILE = re.compile(r"\b(?:TODAY|NOW|RAND|RANDBETWEEN|RANDARRAY|CELL|INFO)\s*\(", re.I)
THREE_D = re.compile(r"(?<![\w.$'])[A-Za-z_][\w.]*:[A-Za-z_][\w.]*!")
DEFINED = re.compile(r"<definedName\b([^>]*)>(.*?)</definedName>", re.S)
ADDRESS = re.compile(r"([A-Za-z]*)(\d*)$")
ENTITIES = {"&quot;": '"', "&apos;": "'"}
MAX_ROWS, MAX_COLS = 1_048_576, 16_384


def _column(letters: str) -> int:
    number = 0
    for letter in letters.upper():
        number = number * 26 + ord(letter) - 64
    return number - 1


def _address(text: str):
    match = ADDRESS.match(text.replace("$", ""))
    row, col = match.group(2), match.group(1)
    return (int(row) - 1 if row else None, _column(col) if col else None)


@dataclass
class _Parsed:
    references: list
    always: bool


class _Parser:
    def __init__(self, sheet_names: list[str], defined: dict):
        self.by_name = {name.casefold(): index for index, name in enumerate(sheet_names)}
        self.defined = defined
        self.resolved: dict = {}

    def parse(self, text: str, home: int, depth: int = 0) -> _Parsed:
        body = EXTERNAL.sub("0", STRING.sub('""', text))
        always = bool(DYNAMIC.search(body) or VOLATILE.search(body) or THREE_D.search(body)
                      or "[" in body)
        references = []
        for match in REFERENCE.finditer(body):
            name = match.group("q") or match.group("u")
            sheet = home
            if name is not None:
                sheet = self.by_name.get(name.replace("''", "'").casefold())
                if sheet is None:
                    always = True
                    continue
            parts = match.group("a").split(":")
            (r1, c1), (r2, c2) = _address(parts[0]), _address(parts[-1])
            references.append((sheet, r1, c1, r2, c2))
        for name in NAME.findall(REFERENCE.sub(" ", body)):
            key = self._lookup(name, home)
            if key is None:
                continue
            if depth > 20:
                always = True
                continue
            parsed = self._name(key, depth)
            references.extend(parsed.references)
            always |= parsed.always
        return _Parsed(references, always)

    def _lookup(self, name, home):
        name = name.casefold()
        if (name, home) in self.defined:
            return name, home
        return (name, None) if (name, None) in self.defined else None

    def _name(self, key, depth):
        if key not in self.resolved:
            self.resolved[key] = _Parsed([], True)  # cycle guard: assume the worst
            home = key[1] if key[1] is not None else 0
            self.resolved[key] = self.parse(self.defined[key], home, depth + 1)
        return self.resolved[key]


def _defined_names(archive: zipfile.ZipFile) -> dict:
    text = archive.read("xl/workbook.xml").decode("utf-8")
    names = {}
    for match in DEFINED.finditer(text):
        name = re.search(r'name="([^"]+)"', match.group(1))
        local = re.search(r'localSheetId="(\d+)"', match.group(1))
        if name:
            names[(unescape(name.group(1), ENTITIES).casefold(),
                   int(local.group(1)) if local else None)] = unescape(match.group(2), ENTITIES)
    return names


class DependencyIndex:
    """Reverse dependencies of every formula cell, answered with numpy."""

    def __init__(self, sheets, formulas, single, ranges, always):
        self.sheets = sheets  # [(rows, cols)] per sheet
        self.coordinates = formulas  # int32 (n, 3): sheet, row, col
        self.always = always  # bool (n,)
        self.single = single  # int32 (m, 4): formula id, sheet, row, col
        self.ranges = ranges  # (range bounds int64 (k, 5), links int32 (l, 2): formula id, range id)

    @classmethod
    def build(cls, prepared: str | Path, formulas: dict, sheet_names: list[str]) -> DependencyIndex:
        with zipfile.ZipFile(prepared) as archive:
            parser = _Parser(sheet_names, _defined_names(archive))
        keys = sorted(formulas)
        rows = [1] * len(sheet_names)
        cols = [1] * len(sheet_names)
        for sheet, row, col in keys:
            rows[sheet] = max(rows[sheet], row + 1)
            cols[sheet] = max(cols[sheet], col + 1)
        single, links, range_ids, always = [], [], {}, np.zeros(len(keys), dtype=bool)
        for fid, (sheet, row, col) in enumerate(keys):
            parsed = parser.parse(formulas[(sheet, row, col)].lstrip("="), sheet)
            always[fid] = parsed.always
            for target, r1, c1, r2, c2 in parsed.references:
                if r1 is not None and c1 is not None and (r1, c1) == (r2, c2):
                    single.append((fid, target, r1, c1))
                    continue
                key = (target, r1, c1, r2, c2)
                links.append((fid, range_ids.setdefault(key, len(range_ids))))
        bounds = np.zeros((len(range_ids), 5), dtype=np.int64)
        for (target, r1, c1, r2, c2), rid in range_ids.items():
            top, bottom = (0, MAX_ROWS - 1) if r1 is None else (min(r1, r2), max(r1, r2))
            left, right = (0, MAX_COLS - 1) if c1 is None else (min(c1, c2), max(c1, c2))
            bounds[rid] = (target, top, left, bottom, right)
        return cls(list(zip(rows, cols)), np.array(keys, dtype=np.int32).reshape(-1, 3),
                   np.array(single, dtype=np.int32).reshape(-1, 4),
                   (bounds, np.array(links, dtype=np.int32).reshape(-1, 2)), always)

    def cone(self, edited) -> set[tuple[int, int, int]]:
        """Every formula whose value may change when the given cells change."""
        bounds, links = self.ranges
        affected = self.always.copy()
        shapes = self.sheets
        heights = np.array([shape[0] for shape in shapes])
        widths = np.array([shape[1] for shape in shapes])
        live = [np.zeros(shape, dtype=bool) for shape in shapes]
        for sheet, row, col in edited:
            if row < shapes[sheet][0] and col < shapes[sheet][1]:
                live[sheet][row, col] = True
                continue
            # outside the formula grid: test the direct references and ranges themselves
            direct = self.single
            affected[direct[(direct[:, 1] == sheet) & (direct[:, 2] == row) & (direct[:, 3] == col), 0]] = True
            holds = ((bounds[:, 0] == sheet) & (bounds[:, 1] <= row) & (bounds[:, 3] >= row)
                     & (bounds[:, 2] <= col) & (bounds[:, 4] >= col))
            affected[links[holds[links[:, 1]], 0]] = True
        inside = (bounds[:, 1] < heights[bounds[:, 0]]) & (bounds[:, 2] < widths[bounds[:, 0]])
        bottom = np.minimum(bounds[:, 3], heights[bounds[:, 0]] - 1) + 1
        right = np.minimum(bounds[:, 4], widths[bounds[:, 0]] - 1) + 1
        single = self.single[(self.single[:, 2] < heights[self.single[:, 1]])
                             & (self.single[:, 3] < widths[self.single[:, 1]])]
        marked = np.zeros_like(affected)
        first = True
        while True:
            new = affected & ~marked
            if not first and not new.any():
                break
            first = False
            marked |= new
            for sheet, row, col in self.coordinates[new]:
                live[sheet][row, col] = True
            affected[single[live_lookup(live, single[:, 1:]), 0]] = True
            range_live = np.zeros(len(bounds), dtype=bool)
            for sheet, grid in enumerate(live):
                rows = inside & (bounds[:, 0] == sheet)
                if not rows.any() or not grid.any():
                    continue
                prefix = np.zeros((grid.shape[0] + 1, grid.shape[1] + 1), dtype=np.int32)
                prefix[1:, 1:] = grid.cumsum(0, dtype=np.int32).cumsum(1, dtype=np.int32)
                top, left = bounds[rows, 1], bounds[rows, 2]
                low, high = bottom[rows], right[rows]
                range_live[rows] = (prefix[low, high] - prefix[top, high]
                                    - prefix[low, left] + prefix[top, left]) > 0
            affected[links[range_live[links[:, 1]], 0]] = True
        return {(int(s), int(r), int(c)) for s, r, c in self.coordinates[affected]}


def live_lookup(live, cells):
    result = np.zeros(len(cells), dtype=bool)
    for sheet, grid in enumerate(live):
        rows = cells[:, 0] == sheet
        if rows.any():
            result[rows] = grid[cells[rows, 1], cells[rows, 2]]
    return result
