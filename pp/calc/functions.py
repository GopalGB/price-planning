"""Excel function library over already-evaluated arguments.

Each function takes scalars or Areas and returns a scalar or an Area. Errors are
returned as XlError values, never raised. Scalar-only functions are wrapped by
`_scalar`, which lifts them element-wise over Areas.
"""

from __future__ import annotations

import functools
import math
import re
from collections import Counter
from datetime import datetime, timedelta
from decimal import ROUND_CEILING, ROUND_DOWN, ROUND_FLOOR, ROUND_HALF_UP, ROUND_UP, Context, Decimal
from typing import Callable, Iterable, Optional, Union

from pp.calc.values import (
    CALC,
    DIV0,
    NA,
    NUM,
    REF,
    VALUE,
    Area,
    XlError,
    compare,
    first_error,
    parse_number,
    scalar_of,
    to_bool,
    to_number,
    to_text,
)

_MISSING = object()

# Excel never rounds to 28 digits; a wide precision context keeps big values exact.
_CTX = Context(prec=400)
_HALF = Context(prec=400, rounding=ROUND_HALF_UP)
_UP = Context(prec=400, rounding=ROUND_UP)
_FLOOR = Context(prec=400, rounding=ROUND_FLOOR)
_CEIL = Context(prec=400, rounding=ROUND_CEILING)
_DOWN = Context(prec=400, rounding=ROUND_DOWN)

LAZY = frozenset(
    {"IF", "IFERROR", "IFS", "SWITCH", "LET", "COLUMN", "ROW", "ISFORMULA",
     "CELL", "OFFSET", "INDIRECT", "AND", "OR"}
)


def _as_area(v: object) -> Area:
    return v if isinstance(v, Area) else Area.scalar(v)


def _at(arg: object, r: int, c: int) -> object:
    if isinstance(arg, Area):
        return arg.get(r, c) if arg.size > 1 else scalar_of(arg)
    return arg


def _scalar(propagate: bool = True) -> Callable:
    """Make a scalar function lift over Areas; with propagate, the first error wins."""

    def decorate(fn: Callable) -> Callable:
        def apply(vals: list) -> object:
            if propagate:
                err = first_error(*vals)
                if err is not None:
                    return err
            return fn(*vals)

        @functools.wraps(fn)
        def wrapper(*args: object) -> object:
            big = [a for a in args if isinstance(a, Area) and a.size > 1]
            if not big:
                return apply([scalar_of(a) for a in args])
            h, w = big[0].height, big[0].width
            if any(a.height != h or a.width != w for a in big):
                return NA
            return Area(
                [[apply([_at(a, r, c) for a in args]) for c in range(w)] for r in range(h)]
            )

        return wrapper

    return decorate


def _int_arg(v: object) -> Union[int, XlError]:
    n = to_number(v)
    if isinstance(n, XlError):
        return n
    if not math.isfinite(n):
        return NUM
    return int(n)


def _dec(x: float) -> Decimal:
    return Decimal(repr(x))


# ---------------------------------------------------------------- text

@_scalar()
def _substitute(text: str, old: str, new: str, instance: object = _MISSING) -> object:
    text, old, new = to_text(text), to_text(old), to_text(new)
    if old == "":
        return text
    if instance is _MISSING:
        return text.replace(old, new)
    n = _int_arg(instance)
    if isinstance(n, XlError):
        return n
    if n < 1:
        return VALUE
    start = 0
    pos = -1
    for _ in range(n):
        pos = text.find(old, start)
        if pos < 0:
            return text
        start = pos + len(old)
    return text[:pos] + new + text[pos + len(old):]


@_scalar()
def _trim(text: object) -> str:
    return re.sub(" +", " ", to_text(text)).strip(" ")


@_scalar()
def _numbervalue(text: object, decimal: object = ".", group: object = ",") -> object:
    s = to_text(text).replace(" ", "")
    if s == "":
        return 0.0
    dec, grp = to_text(decimal), to_text(group)
    if dec == grp:
        return VALUE
    parsed = parse_number(s, dec, grp)
    return VALUE if parsed is None else parsed


def _concat(*args: object) -> object:
    parts = []
    for arg in args:
        for item in (arg.cells() if isinstance(arg, Area) else (arg,)):
            if isinstance(item, XlError):
                return item
            parts.append(to_text(item))
    return "".join(parts)


@_scalar()
def _concatenate(*args: object) -> str:
    return "".join(to_text(a) for a in args)


def _count_arg(v: object, default: int) -> Union[int, XlError]:
    return default if v is _MISSING else _int_arg(v)


@_scalar()
def _left(text: object, count: object = _MISSING) -> object:
    n = _count_arg(count, 1)
    if isinstance(n, XlError):
        return n
    if n < 0:
        return VALUE
    return to_text(text)[:n]


@_scalar()
def _right(text: object, count: object = _MISSING) -> object:
    n = _count_arg(count, 1)
    if isinstance(n, XlError):
        return n
    if n < 0:
        return VALUE
    t = to_text(text)
    return t[max(len(t) - n, 0):]


@_scalar()
def _mid(text: object, start: object, count: object) -> object:
    s = _int_arg(start)
    n = _int_arg(count)
    for x in (s, n):
        if isinstance(x, XlError):
            return x
    if s < 1 or n < 0:
        return VALUE
    return to_text(text)[s - 1:s - 1 + n]


@_scalar()
def _find(needle: object, haystack: object, start: object = _MISSING) -> object:
    s = _count_arg(start, 1)
    if isinstance(s, XlError):
        return s
    h = to_text(haystack)
    if s < 1 or s > len(h):
        return VALUE
    pos = h.find(to_text(needle), s - 1)
    return VALUE if pos < 0 else float(pos + 1)


@functools.lru_cache(maxsize=512)
def _wild_pattern(pattern: str) -> re.Pattern:
    """Excel wildcards (* ? and ~ escapes) as a case-insensitive regex."""
    out = []
    i = 0
    while i < len(pattern):
        ch = pattern[i]
        if ch == "~" and i + 1 < len(pattern) and pattern[i + 1] in "?*~":
            out.append(re.escape(pattern[i + 1]))
            i += 2
            continue
        out.append(".*" if ch == "*" else "." if ch == "?" else re.escape(ch))
        i += 1
    return re.compile("".join(out), re.IGNORECASE | re.DOTALL)


@_scalar()
def _search(needle: object, haystack: object, start: object = _MISSING) -> object:
    s = _count_arg(start, 1)
    if isinstance(s, XlError):
        return s
    h = to_text(haystack)
    if s < 1 or s > len(h):
        return VALUE
    m = _wild_pattern(to_text(needle)).search(h, s - 1)
    return VALUE if m is None else float(m.start() + 1)


@_scalar()
def _len(text: object) -> float:
    return float(len(to_text(text)))


@_scalar()
def _value(text: object) -> object:
    if isinstance(text, bool):
        return VALUE
    return to_number(text)


_NUM_FORMAT = re.compile(r"(#,##)?(0+)(?:\.(0+))?(%?)")


def _fixed(n: float, places: int, min_int: int, grouped: bool, percent: bool) -> str:
    if not math.isfinite(n):
        return NUM.code
    d = _HALF.quantize(_dec(n), Decimal(1).scaleb(-places))
    sign = "-" if d < 0 else ""
    whole, _, frac = format(d.copy_abs(), "f").partition(".")
    whole = whole.zfill(min_int)
    if grouped:
        whole = f"{int(whole):,}"
    decimals = f".{frac}" if places else ""
    return f"{sign}{whole}{decimals}{'%' if percent else ''}"


def _serial_ymd(serial: float) -> Optional[tuple[int, int, int]]:
    days = int(math.floor(serial))
    if days < 1:
        return None
    if days == 60:  # Excel's phantom 1900-02-29
        return 1900, 2, 29
    base = datetime(1899, 12, 31) if days < 60 else datetime(1899, 12, 30)
    d = base + timedelta(days=days)
    return d.year, d.month, d.day


def _format_date(serial: float, fmt: str) -> Optional[str]:
    ymd = _serial_ymd(serial)
    if ymd is None:
        return None
    y, m, d = ymd
    out = (
        fmt.replace("yyyy", f"{y:04d}")
        .replace("yy", f"{y % 100:02d}")
        .replace("mm", f"{m:02d}")
        .replace("dd", f"{d:02d}")
    )
    return None if re.search(r"[A-Za-z]", out) else out


@_scalar()
def _text(value: object, fmt: object) -> str:
    f = to_text(fmt)
    if f == "@" or isinstance(value, bool):
        return to_text(value)
    n = to_number(value)
    if isinstance(n, XlError):
        return to_text(value)
    m = _NUM_FORMAT.fullmatch(f)
    if m:
        grouping, zeros, dec, pct = m.groups()
        return _fixed(n * 100 if pct else n, len(dec or ""), len(zeros),
                      bool(grouping), bool(pct))
    if re.search(r"[ymd]", f):
        formatted = _format_date(n, f)
        if formatted is not None:
            return formatted
    return to_text(value)


# ---------------------------------------------------------------- math

@_scalar()
def _int(v: object) -> object:
    n = to_number(v)
    if isinstance(n, XlError):
        return n
    return float(math.floor(n))


@_scalar()
def _floor(number: object, significance: object) -> object:
    n = to_number(number)
    s = to_number(significance)
    for x in (n, s):
        if isinstance(x, XlError):
            return x
    if s == 0 or n == 0:
        return 0.0
    if s < 0 and n > 0:
        return NUM
    dn, ds = _dec(n), _dec(s)
    if s > 0:
        q = _FLOOR.to_integral_value(_CTX.divide(dn, ds))
    else:
        q = _DOWN.to_integral_value(_CTX.divide(dn, ds))
    return float(_CTX.multiply(q, ds))


@_scalar()
def _floor_math(number: object, significance: object = _MISSING, mode: object = _MISSING) -> object:
    n = to_number(number)
    s = 1.0 if significance is _MISSING else to_number(significance)
    m = 0.0 if mode is _MISSING else to_number(mode)
    for x in (n, s, m):
        if isinstance(x, XlError):
            return x
    if s == 0 or n == 0:
        return 0.0
    ds = _dec(abs(s))
    rounding = _FLOOR if (n >= 0 or m == 0) else _CEIL
    q = rounding.to_integral_value(_CTX.divide(_dec(n), ds))
    return float(_CTX.multiply(q, ds))


@_scalar()
def _roundup(number: object, digits: object) -> object:
    n = to_number(number)
    d = _int_arg(digits)
    for x in (n, d):
        if isinstance(x, XlError):
            return x
    return float(_UP.quantize(_dec(n), Decimal(1).scaleb(-d)))


@_scalar()
def _mod(number: object, divisor: object) -> object:
    n = to_number(number)
    d = to_number(divisor)
    for x in (n, d):
        if isinstance(x, XlError):
            return x
    if d == 0:
        return DIV0
    dn, dd = _dec(n), _dec(d)
    q = _FLOOR.to_integral_value(_CTX.divide(dn, dd))
    return float(_CTX.subtract(dn, _CTX.multiply(dd, q)))


def _sum(*args: object) -> object:
    total = 0.0
    for arg in args:
        if isinstance(arg, Area):
            for v in arg.cells():
                if isinstance(v, XlError):
                    return v
                if type(v) is float:
                    total += v
        else:
            n = to_number(arg)
            if isinstance(n, XlError):
                return n
            total += n
    return total


def _extreme(args: tuple, pick: Callable) -> object:
    values = []
    for arg in args:
        if isinstance(arg, Area):
            for v in arg.cells():
                if isinstance(v, XlError):
                    return v
                if type(v) is float:
                    values.append(v)
        else:
            n = to_number(arg)
            if isinstance(n, XlError):
                return n
            values.append(n)
    return pick(values) if values else 0.0


def _max(*args: object) -> object:
    return _extreme(args, max)


def _min(*args: object) -> object:
    return _extreme(args, min)


def _sumproduct(*arrays: object) -> object:
    areas = [_as_area(a) for a in arrays]
    h, w = areas[0].height, areas[0].width
    if any(a.height != h or a.width != w for a in areas):
        return VALUE
    total = 0.0
    for r in range(h):
        for c in range(w):
            product = 1.0
            for a in areas:
                v = a.get(r, c)
                if isinstance(v, XlError):
                    return v
                product *= v if type(v) is float else 0.0
            total += product
    return total


@_scalar()
def _iseven(v: object) -> object:
    n = to_number(v)
    if isinstance(n, XlError):
        return n
    return int(math.trunc(n)) % 2 == 0


@_scalar()
def _days(end: object, start: object) -> object:
    e = to_number(end)
    s = to_number(start)
    for x in (e, s):
        if isinstance(x, XlError):
            return x
    return float(math.trunc(e) - math.trunc(s))


# ---------------------------------------------------------------- info

@_scalar(propagate=False)
def _iserror(v: object) -> bool:
    return isinstance(v, XlError)


@_scalar(propagate=False)
def _isnumber(v: object) -> bool:
    return type(v) is float


@_scalar(propagate=False)
def _isblank(v: object) -> bool:
    return v is None


@_scalar(propagate=False)
def _istext(v: object) -> bool:
    return isinstance(v, str)


@_scalar()
def _not(v: object) -> object:
    b = to_bool(v)
    return b if isinstance(b, XlError) else not b


def _logical(args: tuple, want_all: bool) -> object:
    seen = []
    for arg in args:
        if isinstance(arg, Area):
            for v in arg.cells():
                if isinstance(v, XlError):
                    return v
                if isinstance(v, bool):
                    seen.append(v)
                elif type(v) is float:
                    seen.append(v != 0)
        else:
            b = to_bool(arg)
            if isinstance(b, XlError):
                return b
            seen.append(b)
    if not seen:
        return VALUE
    return all(seen) if want_all else any(seen)


def _and(*args: object) -> object:
    return _logical(args, want_all=True)


def _or(*args: object) -> object:
    return _logical(args, want_all=False)


def _counta(*args: object) -> float:
    n = 0
    for arg in args:
        if isinstance(arg, Area):
            n += sum(1 for v in arg.cells() if v is not None)
        elif arg is not None:
            n += 1
    return float(n)


# ---------------------------------------------------------------- lookup

def _has_wildcard(text: str) -> bool:
    return any(ch in text for ch in "*?~")


def _positions(n: int, last: bool) -> Iterable[int]:
    return range(n - 1, -1, -1) if last else range(n)


def _vector(area: Area, tag: object, build: Callable[[], list]) -> list:
    key = ("vec", tag)
    vec = area.memo.get(key)
    if vec is None:
        vec = area.memo[key] = build()
    return vec


def _key(v: object) -> Optional[tuple]:
    """Exact-match identity: case-insensitive text; numbers and bools kept apart."""
    if isinstance(v, str):
        return ("t", v.lower())
    if isinstance(v, bool):
        return ("b", v)
    if isinstance(v, float):
        return ("n", v)
    return None


def _exact_index(area: Area, tag: object, vector: list, last: bool) -> dict:
    key = ("idx", tag, last)
    index = area.memo.get(key)
    if index is None:
        index = {}
        for pos in _positions(len(vector), last):
            k = _key(vector[pos])
            if k is not None:
                index.setdefault(k, pos)
        area.memo[key] = index
    return index


def _wild_scan(vector: list, lookup: str, last: bool) -> Optional[int]:
    pattern = _wild_pattern(lookup)
    for pos in _positions(len(vector), last):
        v = vector[pos]
        if isinstance(v, str) and pattern.fullmatch(v) is not None:
            return pos
    return None


def _approx_asc(vector: list, lookup: object) -> Optional[int]:
    """Largest position whose value is <= lookup (vector sorted ascending)."""
    lo, hi, best = 0, len(vector) - 1, None
    while lo <= hi:
        mid = (lo + hi) // 2
        if compare(vector[mid], lookup) <= 0:
            best = mid
            lo = mid + 1
        else:
            hi = mid - 1
    return best


def _approx_desc(vector: list, lookup: object) -> Optional[int]:
    """Last position whose value is >= lookup (vector sorted descending)."""
    lo, hi, best = 0, len(vector) - 1, None
    while lo <= hi:
        mid = (lo + hi) // 2
        if compare(vector[mid], lookup) >= 0:
            best = mid
            lo = mid + 1
        else:
            hi = mid - 1
    return best


def _position(area: Area, tag: object, build: Callable[[], list], lookup: object,
              mode: int, last: bool = False) -> Optional[int]:
    """0-based position of lookup in the vector. mode: 0 exact, 1 ascending, -1 descending."""
    vector = _vector(area, tag, build)
    if mode == 0:
        if isinstance(lookup, str) and _has_wildcard(lookup):
            return _wild_scan(vector, lookup, last)
        k = _key(lookup)
        if k is None:
            return None
        return _exact_index(area, tag, vector, last).get(k)
    if mode < 0:
        return _approx_desc(vector, lookup)
    return _approx_asc(vector, lookup)


def _key_arg(value: object) -> object:
    """A lookup value as Excel passes it: a blank cell looks up 0."""
    value = scalar_of(value)
    return 0.0 if value is None else value


def _match(lookup: object, lookup_array: object, match_type: object = _MISSING) -> object:
    lookup = _key_arg(lookup)
    if isinstance(lookup, XlError):
        return lookup
    area = _as_area(lookup_array)
    if area.height > 1 and area.width > 1:
        return NA
    mode = 1 if match_type is _MISSING else _int_arg(match_type)
    if isinstance(mode, XlError):
        return mode
    pos = _position(area, "vec", lambda: list(area.cells()), lookup,
                    0 if mode == 0 else (-1 if mode < 0 else 1))
    return NA if pos is None else float(pos + 1)


def _vh_lookup(lookup: object, table: object, index: object, range_lookup: object,
               vertical: bool) -> object:
    lookup = _key_arg(lookup)
    if isinstance(lookup, XlError):
        return lookup
    area = _as_area(table)
    idx = _int_arg(index)
    if isinstance(idx, XlError):
        return idx
    approx = True if range_lookup is _MISSING else to_bool(range_lookup)
    if isinstance(approx, XlError):
        return approx
    if idx < 1:
        return VALUE
    if idx > (area.width if vertical else area.height):
        return REF
    if vertical:
        pos = _position(area, ("col", 0), lambda: area.column(0), lookup, 1 if approx else 0)
    else:
        pos = _position(area, ("row", 0), lambda: area.row(0), lookup, 1 if approx else 0)
    if pos is None:
        return NA
    value = area.get(pos, idx - 1) if vertical else area.get(idx - 1, pos)
    return 0.0 if value is None else value


def _vlookup(lookup: object, table: object, col_index: object,
             range_lookup: object = _MISSING) -> object:
    return _vh_lookup(lookup, table, col_index, range_lookup, vertical=True)


def _hlookup(lookup: object, table: object, row_index: object,
             range_lookup: object = _MISSING) -> object:
    return _vh_lookup(lookup, table, row_index, range_lookup, vertical=False)


def _index_fn(array: object, row_num: object = _MISSING, col_num: object = _MISSING) -> object:
    area = _as_area(array)
    r = None if row_num is _MISSING else _int_arg(row_num)
    c = None if col_num is _MISSING else _int_arg(col_num)
    for x in (r, c):
        if isinstance(x, XlError):
            return x
    if (r is not None and r < 0) or (c is not None and c < 0):
        return VALUE
    if c is None and r is None:
        return area
    if c is None:
        # One index on a single row or column walks that vector; otherwise it picks a row.
        if area.height == 1:
            r, c = 1, r
        elif area.width == 1:
            c = 1
        else:
            c = 0
    rr = 0 if r is None else r
    if rr > area.height or c > area.width:
        return REF
    if rr == 0 and c == 0:
        return area
    if rr == 0:
        return Area([[v] for v in area.column(c - 1)])
    if c == 0:
        return Area([area.row(rr - 1)])
    value = area.get(rr - 1, c - 1)
    return 0.0 if value is None else value


def _xl_search(area: Area, tag: object, build: Callable[[], list], lookup: object,
               match_mode: int, search_mode: int) -> Optional[int]:
    vector = _vector(area, tag, build)
    last = search_mode == -1
    if search_mode in (2, -2) and match_mode == 0:
        pos = _approx_asc(vector, lookup) if search_mode == 2 else _approx_desc(vector, lookup)
        return pos if pos is not None and compare(vector[pos], lookup) == 0 else None
    if match_mode == 2 and isinstance(lookup, str):
        return _wild_scan(vector, lookup, last)
    if match_mode == 0 or match_mode == 2:
        return _position(area, tag, build, lookup, 0, last)
    exact = _position(area, tag, build, lookup, 0, last)
    if exact is not None:
        return exact
    # -1 keeps the largest value below lookup; 1 keeps the smallest above it.
    sign = 1 if match_mode == -1 else -1
    best = None
    for pos in _positions(len(vector), last):
        v = vector[pos]
        if v is None or isinstance(v, XlError):
            continue
        c = compare(v, lookup)
        if (c < 0) if match_mode == -1 else (c > 0):
            if best is None or sign * compare(v, vector[best]) > 0:
                best = pos
    return best


def _xlookup(lookup: object, lookup_array: object, return_array: object,
             if_not_found: object = _MISSING, match_mode: object = _MISSING,
             search_mode: object = _MISSING) -> object:
    lookup = _key_arg(lookup)
    if isinstance(lookup, XlError):
        return lookup
    la = _as_area(lookup_array)
    ra = _as_area(return_array)
    mm = 0 if match_mode is _MISSING else _int_arg(match_mode)
    sm = 1 if search_mode is _MISSING else _int_arg(search_mode)
    for x in (mm, sm):
        if isinstance(x, XlError):
            return x
    if la.width == 1:
        vertical = True
        if ra.height != la.height:
            return VALUE
    elif la.height == 1:
        vertical = False
        if ra.width != la.width:
            return VALUE
    else:
        return VALUE
    tag = ("xl-lookup", vertical)
    build = (lambda: la.column(0)) if vertical else (lambda: la.row(0))
    pos = _xl_search(la, tag, build, lookup, mm, sm)
    if pos is None:
        return NA if if_not_found is _MISSING else if_not_found
    if vertical:
        if ra.width == 1:
            value = ra.get(pos, 0)
            return 0.0 if value is None else value
        return Area([ra.row(pos)])
    if ra.height == 1:
        value = ra.get(0, pos)
        return 0.0 if value is None else value
    return Area([[v] for v in ra.column(pos)])


# ---------------------------------------------------------------- conditional aggregates

def _blank_match(v: object) -> bool:
    return v is None or v == ""


_OPERATORS = {
    "<": lambda c: c < 0,
    ">": lambda c: c > 0,
    "<=": lambda c: c <= 0,
    ">=": lambda c: c >= 0,
}
_CRITERIA = re.compile(r"(<>|>=|<=|=|<|>)?(.*)", re.DOTALL)


def _numeric_equal(num: float) -> Callable[[object], bool]:
    def match(v: object) -> bool:
        if type(v) is float:
            return v == num
        return isinstance(v, str) and parse_number(v) == num
    return match


def _criteria(crit: object) -> Callable[[object], bool]:
    """Predicate for one Excel criteria value (number, bool, blank, or operator text)."""
    if crit is None:  # Excel: an empty criteria cell counts as 0 (the *IF(S) function docs)
        return _numeric_equal(0.0)
    if isinstance(crit, bool):
        return lambda v: type(v) is bool and v == crit
    if isinstance(crit, float):
        return _numeric_equal(crit)
    op, operand = _CRITERIA.fullmatch(crit).groups()
    op = op or ""
    if operand == "":
        if op in ("", "="):
            return _blank_match
        return lambda v: v is not None and v != ""
    if operand.upper() in ("TRUE", "FALSE"):
        target = operand.upper() == "TRUE"
        if op == "<>":
            return lambda v: not (type(v) is bool and v == target)
        return lambda v: type(v) is bool and v == target
    num = parse_number(operand)
    if num is not None:
        if op in ("", "="):
            return _numeric_equal(num)
        if op == "<>":
            equal = _numeric_equal(num)
            return lambda v: not equal(v)
        test = _OPERATORS[op]
        return lambda v: type(v) is float and test(compare(v, num))
    pattern = _wild_pattern(operand)
    if op in ("", "="):
        return lambda v: isinstance(v, str) and pattern.fullmatch(v) is not None
    if op == "<>":
        return lambda v: not (isinstance(v, str) and pattern.fullmatch(v) is not None)
    test = _OPERATORS[op]
    return lambda v: isinstance(v, str) and test(compare(v, operand))


def _build_pairs(args: tuple, shape: Area) -> Union[list, XlError]:
    if not args or len(args) % 2:
        return VALUE
    pairs = []
    for i in range(0, len(args), 2):
        rng = _as_area(args[i])
        crit = scalar_of(args[i + 1])
        if isinstance(crit, XlError):
            return crit
        if rng.height != shape.height or rng.width != shape.width:
            return VALUE
        pairs.append((rng, _criteria(crit)))
    return pairs


def _matched(pairs: list) -> Iterable[tuple[int, int]]:
    base = pairs[0][0]
    for r in range(base.height):
        for c in range(base.width):
            if all(pred(rng.get(r, c)) for rng, pred in pairs):
                yield r, c


def _sum_matched(sum_area: Area, pairs: list) -> object:
    total = 0.0
    for r, c in _matched(pairs):
        if r >= sum_area.height or c >= sum_area.width:
            continue
        v = sum_area.get(r, c)
        if isinstance(v, XlError):
            return v
        if type(v) is float:
            total += v
    return total


def _sumif(range_: object, criteria: object, sum_range: object = _MISSING) -> object:
    rng = _as_area(range_)
    crit = scalar_of(criteria)
    if isinstance(crit, XlError):
        return crit
    sum_area = rng if sum_range is _MISSING else _as_area(sum_range)
    return _sum_matched(sum_area, [(rng, _criteria(crit))])


def _sumifs(sum_range: object, *pairs: object) -> object:
    sum_area = _as_area(sum_range)
    built = _build_pairs(pairs, sum_area)
    if isinstance(built, XlError):
        return built
    return _sum_matched(sum_area, built)


def _countifs(*args: object) -> object:
    if not args:
        return VALUE
    built = _build_pairs(args, _as_area(args[0]))
    if isinstance(built, XlError):
        return built
    return float(sum(1 for _ in _matched(built)))


def _rank(number: object, ref: object, order: object = _MISSING) -> object:
    """RANK / RANK.EQ: 1 + how many numbers in ref beat it; only real numbers in ref count."""
    value = scalar_of(number)
    if isinstance(value, XlError):
        return value
    value = to_number(value)
    if isinstance(value, XlError):
        return value
    ascending = False
    if order is not _MISSING and order is not None:
        flag = to_number(scalar_of(order))
        if isinstance(flag, XlError):
            return flag
        ascending = flag != 0
    numbers = [item for item in _as_area(ref).cells() if type(item) is float]
    if value not in numbers:
        return NA
    return float(1 + sum(1 for item in numbers if (item < value if ascending else item > value)))


def _ifs_extreme(target: object, pairs: tuple, pick: Callable) -> object:
    area = _as_area(target)
    built = _build_pairs(pairs, area)
    if isinstance(built, XlError):
        return built
    values = []
    for r, c in _matched(built):
        v = area.get(r, c)
        if isinstance(v, XlError):
            return v
        if type(v) is float:
            values.append(v)
    return pick(values) if values else 0.0


def _maxifs(target: object, *pairs: object) -> object:
    return _ifs_extreme(target, pairs, max)


def _minifs(target: object, *pairs: object) -> object:
    return _ifs_extreme(target, pairs, min)


# ---------------------------------------------------------------- dynamic arrays

def _sequence(rows: object, cols: object = _MISSING, start: object = _MISSING,
              step: object = _MISSING) -> object:
    r = _int_arg(rows)
    c = 1 if cols is _MISSING else _int_arg(cols)
    s = 1.0 if start is _MISSING else to_number(start)
    st = 1.0 if step is _MISSING else to_number(step)
    for x in (r, c, s, st):
        if isinstance(x, XlError):
            return x
    if r < 1 or c < 1:
        return CALC
    return Area([[s + st * (i * c + j) for j in range(c)] for i in range(r)])


def _row_key(v: object) -> tuple:
    if v is None:
        return ("blank",)
    if isinstance(v, XlError):
        return ("err", v.code)
    if isinstance(v, str):
        return ("t", v.lower())
    if isinstance(v, bool):
        return ("b", v)
    return ("n", v)


def _unique(array: object, by_col: object = _MISSING, exactly_once: object = _MISSING) -> object:
    area = _as_area(array)
    column_mode = False if by_col is _MISSING else to_bool(by_col)
    once = False if exactly_once is _MISSING else to_bool(exactly_once)
    for x in (column_mode, once):
        if isinstance(x, XlError):
            return x
    lines = (
        [area.column(c) for c in range(area.width)]
        if column_mode
        else [area.row(r) for r in range(area.height)]
    )
    keys = [tuple(_row_key(v) for v in line) for line in lines]
    counts = Counter(keys)
    seen: set = set()
    kept = []
    for line, key in zip(lines, keys):
        if key in seen or (once and counts[key] != 1):
            continue
        seen.add(key)
        kept.append(line)
    if not kept:
        return CALC
    if column_mode:
        return Area([list(row) for row in zip(*kept)])
    return Area(kept)


FUNCTIONS: dict[str, Callable] = {
    "SUBSTITUTE": _substitute,
    "TRIM": _trim,
    "NUMBERVALUE": _numbervalue,
    "CONCAT": _concat,
    "CONCATENATE": _concatenate,
    "LEFT": _left,
    "RIGHT": _right,
    "MID": _mid,
    "FIND": _find,
    "SEARCH": _search,
    "LEN": _len,
    "VALUE": _value,
    "TEXT": _text,
    "INT": _int,
    "FLOOR": _floor,
    "FLOOR.MATH": _floor_math,
    "ROUNDUP": _roundup,
    "MOD": _mod,
    "SUM": _sum,
    "MAX": _max,
    "MIN": _min,
    "SUMPRODUCT": _sumproduct,
    "ISEVEN": _iseven,
    "DAYS": _days,
    "ISERROR": _iserror,
    "ISNUMBER": _isnumber,
    "ISBLANK": _isblank,
    "ISTEXT": _istext,
    "NOT": _not,
    "AND": _and,
    "OR": _or,
    "COUNTA": _counta,
    "MATCH": _match,
    "VLOOKUP": _vlookup,
    "HLOOKUP": _hlookup,
    "INDEX": _index_fn,
    "XLOOKUP": _xlookup,
    "SUMIF": _sumif,
    "SUMIFS": _sumifs,
    "COUNTIF": _countifs,
    "RANK": _rank,
    "RANK.EQ": _rank,
    "COUNTIFS": _countifs,
    "MAXIFS": _maxifs,
    "MINIFS": _minifs,
    "SEQUENCE": _sequence,
    "UNIQUE": _unique,
}
