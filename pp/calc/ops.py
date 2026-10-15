"""Excel operators, implicit intersection and the lazy functions the compiler calls."""

from __future__ import annotations

import math

from .values import DIV0, NA, NUM, VALUE, Area, XlError, compare, scalar_of, to_bool, to_number, to_text


class RefArea(Area):
    """An Area read from cells: it remembers where it sits, for implicit intersection and COLUMN/ROW."""

    __slots__ = ("sheet", "top", "left")

    def __init__(self, rows, sheet: int, top: int, left: int) -> None:
        super().__init__(rows)
        self.sheet, self.top, self.left = sheet, top, left


def isect(value, r: int, c: int):
    """A value in a place that takes one scalar, as a non-array formula at (r, c) sees it."""
    if not isinstance(value, Area):
        return value
    if value.height == 1 and value.width == 1:
        return value.get(0, 0)
    if isinstance(value, RefArea):
        if value.width == 1 and value.top <= r < value.top + value.height:
            return value.get(r - value.top, 0)
        if value.height == 1 and value.left <= c < value.left + value.width:
            return value.get(0, c - value.left)
        return VALUE
    return value.get(0, 0)


def first(value):
    """The top-left value, as a one-cell array formula keeps it."""
    return value.get(0, 0) if isinstance(value, Area) else value


def _broadcast(a, b, fn):
    if not isinstance(a, Area) and not isinstance(b, Area):
        return fn(a, b)
    ha, wa = (a.height, a.width) if isinstance(a, Area) else (1, 1)
    hb, wb = (b.height, b.width) if isinstance(b, Area) else (1, 1)
    # Sizes that differ and are not 1 give #N/A outside the smaller array, as in Excel
    height, width = max(ha, hb), max(wa, wb)

    def at(x, h, w, i, j):
        if not isinstance(x, Area):
            return x
        if (h == 1 or i < h) and (w == 1 or j < w):
            return x.get(0 if h == 1 else i, 0 if w == 1 else j)
        return NA

    return Area([[fn(at(a, ha, wa, i, j), at(b, hb, wb, i, j)) for j in range(width)] for i in range(height)])


def _arith(fn):
    def scalar(a, b):
        if type(a) is float and type(b) is float:
            return fn(a, b)
        if isinstance(a, XlError):
            return a
        if isinstance(b, XlError):
            return b
        x, y = to_number(a), to_number(b)
        if isinstance(x, XlError):
            return x
        if isinstance(y, XlError):
            return y
        return fn(x, y)

    return lambda a, b: _broadcast(a, b, scalar)


def _div(x, y):
    return DIV0 if y == 0 else x / y


def _pow(x, y):
    try:
        result = x ** y
    except (OverflowError, ZeroDivisionError):
        return NUM if x != 0 else DIV0
    return NUM if isinstance(result, complex) or math.isinf(result) else float(result)


add = _arith(lambda x, y: x + y)
sub = _arith(lambda x, y: x - y)
mul = _arith(lambda x, y: x * y)
div = _arith(_div)
power = _arith(_pow)


def _concat(a, b):
    if isinstance(a, XlError):
        return a
    if isinstance(b, XlError):
        return b
    return to_text(a) + to_text(b)


def concat(a, b):
    return _broadcast(a, b, _concat)


def _comparer(test):
    def scalar(a, b):
        if isinstance(a, XlError):
            return a
        if isinstance(b, XlError):
            return b
        return test(compare(a, b))

    return lambda a, b: _broadcast(a, b, scalar)


eq = _comparer(lambda c: c == 0)
ne = _comparer(lambda c: c != 0)
lt = _comparer(lambda c: c < 0)
le = _comparer(lambda c: c <= 0)
gt = _comparer(lambda c: c > 0)
ge = _comparer(lambda c: c >= 0)


def _lift1(fn):
    def apply(a):
        if isinstance(a, Area):
            return Area([[fn(v) for v in a.row(i)] for i in range(a.height)])
        return fn(a)

    return apply


def _negate(a):
    if type(a) is float:
        return -a
    if isinstance(a, XlError):
        return a
    x = to_number(a)
    return x if isinstance(x, XlError) else -x


def _percent(a):
    if isinstance(a, XlError):
        return a
    x = to_number(a)
    return x if isinstance(x, XlError) else x / 100


def _plus(a):
    return 0.0 if a is None else a


neg = _lift1(_negate)
percent = _lift1(_percent)
plus = _lift1(_plus)


def result(value, r: int, c: int, array: bool):
    """A formula cell's own value: one scalar, blank references read as 0."""
    value = first(value) if array else isect(value, r, c)
    return 0.0 if value is None else value


# Lazy functions: branches are passed as zero-argument callables.

def _truth(value):
    if isinstance(value, XlError):
        return value
    return to_bool(value)


def fn_if(condition, then, otherwise):
    if isinstance(condition, Area):
        a, b = then(), otherwise()
        return _broadcast(_broadcast(condition, a, lambda x, y: (x, y)), b,
                          lambda xy, z: _pick(xy, z))
    test = _truth(condition)
    if isinstance(test, XlError):
        return test
    return then() if test else otherwise()


def _pick(xy, z):
    if isinstance(xy, XlError):
        return xy
    condition, value = xy
    test = _truth(condition)
    if isinstance(test, XlError):
        return test
    return value if test else z


def fn_iferror(value, fallback):
    if isinstance(value, Area):
        alternative = fallback()
        return _broadcast(value, alternative, lambda v, alt: alt if isinstance(v, XlError) else v)
    return fallback() if isinstance(value, XlError) else value


def fn_ifs(*pairs):
    for condition, value in pairs:
        test = _truth(scalar_of(condition()))
        if isinstance(test, XlError):
            return test
        if test:
            return value()
    return NA


def _blank_zero(value):
    """A blank cell passed to a function as a value counts as 0, so it never equals ""."""
    return 0.0 if value is None else value


def fn_switch(subject, cases, default):
    subject = _blank_zero(scalar_of(subject))
    if isinstance(subject, XlError):
        return subject
    for match, value in cases:
        candidate = _blank_zero(scalar_of(match()))
        if not isinstance(candidate, XlError) and compare(subject, candidate) == 0:
            return value()
    return default() if default is not None else NA
