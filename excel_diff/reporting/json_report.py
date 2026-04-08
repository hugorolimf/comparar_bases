from __future__ import annotations

import json
from dataclasses import asdict, is_dataclass
from pathlib import Path

from excel_diff.models import ComparisonResult


def write_json_report(result: ComparisonResult, path: Path) -> None:
    payload = dataclass_to_dict(result)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def dataclass_to_dict(value):
    if is_dataclass(value):
        return {key: dataclass_to_dict(item) for key, item in asdict(value).items()}
    if isinstance(value, list):
        return [dataclass_to_dict(item) for item in value]
    if isinstance(value, dict):
        return {key: dataclass_to_dict(item) for key, item in value.items()}
    return value