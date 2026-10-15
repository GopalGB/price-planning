"""Excel value model: scalars, error values, 2D areas, coercions and ordering.

Scalars are float (all numbers, dates as serials), str, bool, None (blank)
and XlError. An Area is an immutable 2D block of scalars.
"""

from __future__ import annotations

import math
import re
from typing import Iterator, Optional, Union

_ERROR_CODES = frozenset(
    {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#SPILL!", "#CALC!"}
)


class XlError:
    """An Excel error value. Equal by code, so XlError("#N/A") == NA."""

    __slots__ = ("code",)

    def __init__(self, code: str) -> None:
        if code not in _ERROR_CODES:
            raise ValueError(f"unknown Excel error code: {code!r}")
        self.code = code

    def __eq__(self, other: object) -> bool:
        return isinstance(other, XlError) and other.code == self.code

    def __hash__(self) -> int:
        return hash(("XlError", self.code))

    def __repr__(self) -> str:
        return f"XlError({self.code!r})"


NA = XlError("#N/A")
VALUE = XlError("#VALUE!")
REF = XlError("#REF!")
DIV0 = XlError("#DIV/0!")
NUM = XlError("#NUM!")
NAME = XlError("#NAME?")
NULL = XlError("#NULL!")
SPILL = XlError("#SPILL!")
CALC = XlError("#CALC!")

Scalar = Union[float, str, bool, None, XlError]


def _norm(v: object) -> Scalar:
    if v is None or isinstance(v, (bool, str, XlError)):
        return v
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, Area):
        raise TypeError("an Area cannot be nested inside another Area")
    raise TypeError(f"not an Excel scalar: {type(v).__name__}")


class Area:
    """Immutable 2D block of scalars. `memo` holds caches built by lookup functions."""

    __slots__ = ("_rows", "memo")

    def __init__(self, rows: list[list[object]]) -> None:
        if not rows or not rows[0]:
            raise ValueError("an Area needs at least one cell")
        width = len(rows[0])
        if any(len(row) != width for row in rows):
            raise ValueError("Area rows must have equal length")
        self._rows: tuple[tuple[Scalar, ...], ...] = tuple(
            tuple(_norm(v) for v in row) for row in rows
        )
        self.memo: dict = {}

    @classmethod
    def scalar(cls, value: object) -> "Area":
        return cls([[value]])

    @property
    def height(self) -> int:
        return len(self._rows)

    @property
    def width(self) -> int:
        return len(self._rows[0])

    @property
    def size(self) -> int:
        return self.height * self.width

    def get(self, r: int, c: int) -> Scalar:
        return self._rows[r][c]

    def cells(self) -> Iterator[Scalar]:
        for row in self._rows:
            yield from row

    def column(self, c: int) -> list:
        return [row[c] for row in self._rows]

    def row(self, r: int) -> list:
        return list(self._rows[r])

    def __eq__(self, other: object) -> bool:
        return isinstance(other, Area) and other._rows == self._rows

    def __hash__(self) -> int:
        return hash(self._rows)

    def __repr__(self) -> str:
        return f"Area({[list(r) for r in self._rows]!r})"


def scalar_of(v: object) -> Scalar:
    """A 1x1 Area is its value; a bigger Area where one scalar is needed is #VALUE!."""
    if isinstance(v, Area):
        return v.get(0, 0) if v.size == 1 else VALUE
    return _norm(v)


_NUMBER_BODY = re.compile(r"[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?")


def parse_number(text: str, decimal: str = ".", group: str = ",") -> Optional[float]:
    """Parse numeric text the way VALUE/NUMBERVALUE do. None when it is not a number.

    Group separators are removed wherever they appear (lenient: "1,2345" parses).
    """
    s = text.strip()
    if group:
        s = s.replace(group, "")
    percent = s.endswith("%")
    if percent:
        s = s[:-1].strip()
    s = s.replace(decimal, ".")
    if not _NUMBER_BODY.fullmatch(s):
        return None
    value = float(s)
    if not math.isfinite(value):
        return None
    return value / 100 if percent else value


def to_number(v: object) -> Union[float, XlError]:
    v = scalar_of(v)
    if isinstance(v, XlError):
        return v
    if v is None:
        return 0.0
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if isinstance(v, float):
        return v
    parsed = parse_number(v)
    return VALUE if parsed is None else parsed


def format_general(x: float) -> str:
    """Excel General text: 15 significant digits, scientific outside [1e-6, 1e15)."""
    if not math.isfinite(x):
        return NUM.code
    if x == 0:
        return "0"
    mantissa, exp_text = f"{x:.14e}".split("e")
    exp = int(exp_text)
    if exp >= 15 or exp < -6:
        digits = mantissa.rstrip("0").rstrip(".")
        sign = "+" if exp >= 0 else "-"
        return f"{digits}E{sign}{abs(exp):02d}"
    text = f"{x:.{14 - exp}f}"
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text


def to_text(v: object) -> str:
    v = scalar_of(v)
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float):
        return format_general(v)
    if isinstance(v, XlError):
        return v.code
    return v


def to_bool(v: object) -> Union[bool, XlError]:
    v = scalar_of(v)
    if isinstance(v, XlError):
        return v
    if v is None:
        return False
    if isinstance(v, bool):
        return v
    if isinstance(v, float):
        return v != 0
    upper = v.upper()
    if upper == "TRUE":
        return True
    if upper == "FALSE":
        return False
    return VALUE


def first_error(*values: object) -> Optional[XlError]:
    for v in values:
        items = v.cells() if isinstance(v, Area) else (v,)
        for item in items:
            if isinstance(item, XlError):
                return item
    return None


def _kind(v: Scalar) -> int:
    """Excel sort order across types: numbers < text < booleans < errors."""
    if isinstance(v, bool):
        return 2
    if isinstance(v, str):
        return 1
    if isinstance(v, XlError):
        return 3
    return 0


def _blank_as(other: Scalar) -> Scalar:
    if isinstance(other, bool):
        return False
    if isinstance(other, str):
        return ""
    return 0.0


def compare(a: object, b: object) -> int:
    """-1, 0 or 1 under Excel ordering. Text is case-insensitive; blank is 0, "" or FALSE."""
    a = scalar_of(a)
    b = scalar_of(b)
    if a is None and b is None:
        return 0
    if a is None:
        a = _blank_as(b)
    if b is None:
        b = _blank_as(a)
    ka, kb = _kind(a), _kind(b)
    if ka != kb:
        return -1 if ka < kb else 1
    if ka == 1:
        a, b = a.lower(), b.lower()
    elif ka == 3:
        a, b = a.code, b.code
    return (a > b) - (a < b)
