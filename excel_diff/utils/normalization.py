from __future__ import annotations

from datetime import date, datetime
import math
import re
import unicodedata
from decimal import Decimal
from typing import Any


_NON_ALNUM_RE = re.compile(r"[^a-z0-9]+")
_WHITESPACE_RE = re.compile(r"\s+")


def strip_accents(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value)
    return "".join(char for char in normalized if not unicodedata.combining(char))


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = strip_accents(text)
    text = _WHITESPACE_RE.sub(" ", text)
    text = _NON_ALNUM_RE.sub(" ", text)
    text = _WHITESPACE_RE.sub(" ", text).strip()
    return text


def normalize_header(value: Any) -> str:
    text = normalize_text(value)
    return text.replace(" ", "_")


def format_number(value: float | int | Decimal) -> str:
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, Decimal):
        value = float(value)
    if isinstance(value, float):
        if math.isfinite(value) and value.is_integer():
            return str(int(value))
        return ("%f" % value).rstrip("0").rstrip(".")
    return str(value)


def normalize_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, (int, float, Decimal)):
        return format_number(value)
    text = str(value).strip()
    if not text:
        return ""
    return normalize_text(text)


def classify_value(value: Any) -> str:
    if value is None:
        return "empty"
    if isinstance(value, bool):
        return "bool"
    if isinstance(value, datetime):
        return "datetime"
    if isinstance(value, date):
        return "date"
    if isinstance(value, int) and not isinstance(value, bool):
        return "int"
    if isinstance(value, float):
        return "float"
    if isinstance(value, Decimal):
        return "decimal"
    text = str(value).strip()
    if not text:
        return "empty"
    return "string"


def is_blank(value: Any) -> bool:
    return normalize_value(value) == ""
