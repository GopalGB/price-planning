"""Strict JSON-compatible workbook values."""

import math


def validate_value(value):
    if value is None or isinstance(value, (str, bool)):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if isinstance(value, float) and not math.isfinite(value):
            raise ValueError("numeric value must be finite")
        return value
    if isinstance(value, dict) and set(value) == {"error"} and isinstance(value["error"], str):
        return value
    raise ValueError("unsupported cell value")
