from __future__ import annotations
import re


def safe_float(value: str | None) -> float | None:
    if value is None:
        return None
    try:
        cleaned = re.sub(r"[^\d.,-]", "", str(value)).replace(",", ".")
        return float(cleaned)
    except (ValueError, TypeError):
        return None


def compute_variation(current: float | None, previous: float | None) -> float | None:
    if current is None or previous is None or previous == 0:
        return None
    return round(((current - previous) / previous) * 100, 4)
