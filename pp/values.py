"""Strict JSON-compatible workbook values."""

import math
import re

from ops.typed import ENGINE_ERRORS


MIN_NATIVE_INTEGER = -(2 ** 63)
MAX_NATIVE_INTEGER = 2 ** 63 - 1
SUPPORTED_ERRORS = frozenset(ENGINE_ERRORS.values())
UNSUPPORTED_XML_TEXT = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff\ufffe\uffff]")


def validate_text(value):
    if len(value) > 32_767:
        raise ValueError("text is longer than the Excel cell limit of 32767 characters")
    if UNSUPPORTED_XML_TEXT.search(value):
        raise ValueError("text contains a character OOXML cannot store")
    return value


def validate_value(value):
    if value is None or isinstance(value, bool):
        return value
    if isinstance(value, str):
        return validate_text(value)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if isinstance(value, float) and not math.isfinite(value):
            raise ValueError("numeric value must be finite")
        if isinstance(value, int) and not MIN_NATIVE_INTEGER <= value <= MAX_NATIVE_INTEGER:
            raise ValueError("integer value is outside the calculation engine range")
        return value
    if isinstance(value, dict) and set(value) == {"error"} and isinstance(value["error"], str):
        if value["error"] not in SUPPORTED_ERRORS:
            raise ValueError("unsupported workbook error value")
        return value
    raise ValueError("unsupported cell value")


def exactly_equal(first, second):
    return _tagged(first) == _tagged(second)


def _tagged(value):
    if isinstance(value, dict) and set(value) == {"error"}:
        return "error", value["error"]
    if value is None:
        return "blank", None
    if isinstance(value, bool):
        return "boolean", value
    if isinstance(value, (int, float)):
        return "number", value
    if isinstance(value, str):
        return "string", value
    return "unsupported", type(value).__name__, repr(value)
