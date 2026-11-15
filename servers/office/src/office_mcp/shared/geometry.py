from __future__ import annotations

import re

POINT = re.compile(r"x:\s*(-?\d+(?:\.\d+)?)\s+y:\s*(-?\d+(?:\.\d+)?)")
SIZE = re.compile(r"w:\s*(-?\d+(?:\.\d+)?)\s+h:\s*(-?\d+(?:\.\d+)?)")
SIZE_NAMES = ("width", "height")


def decode_pair(
    value: object, pattern: re.Pattern[str], names: tuple[str, str]
) -> tuple[float, float] | None:
    if isinstance(value, tuple | list) and len(value) == 2:
        first, second = value
        if isinstance(first, int | float) and isinstance(second, int | float):
            return float(first), float(second)
    first_attr = getattr(value, names[0], None)
    second_attr = getattr(value, names[1], None)
    if isinstance(first_attr, int | float) and isinstance(second_attr, int | float):
        return float(first_attr), float(second_attr)
    match = pattern.search(str(value)) if value is not None else None
    return (float(match[1]), float(match[2])) if match else None


def is_empty_size(value: object) -> bool:
    size = decode_pair(value, SIZE, SIZE_NAMES)
    return size is not None and size[0] <= 0 and size[1] <= 0
