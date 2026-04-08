from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass(slots=True)
class ColumnProfile:
    index: int
    name: str
    normalized_name: str
    dominant_type: str
    type_counts: dict[str, int]
    non_empty: int
    total: int
    null_ratio: float
    unique_ratio: float
    sample_values: list[str] = field(default_factory=list)


@dataclass(slots=True)
class SheetProfile:
    path: str
    sheet_name: str
    header_rows: list[int]
    headers: list[str]
    data_start_row: int
    sample_row_count: int
    sheet_width: int
    column_profiles: list[ColumnProfile]
    key_suggestions: list[str] = field(default_factory=list)


@dataclass(slots=True)
class ColumnMapping:
    base_column: str
    compare_column: str
    method: str
    score: float


@dataclass(slots=True)
class KeyMatchCandidate:
    index: int
    base_column: str
    compare_column: str
    method: str
    score: float
    base_unique_ratio: float
    compare_unique_ratio: float
    base_null_ratio: float
    compare_null_ratio: float
    base_type: str
    compare_type: str


@dataclass(slots=True)
class ValidationIssue:
    level: str
    code: str
    message: str


@dataclass(slots=True)
class RowDiff:
    key: str
    base_row_number: int | None
    compare_row_number: int | None
    status: str
    base_values: dict[str, Any] = field(default_factory=dict)
    compare_values: dict[str, Any] = field(default_factory=dict)
    changes: list[dict[str, Any]] = field(default_factory=list)


@dataclass(slots=True)
class ComparisonResult:
    base_profile: SheetProfile
    compare_profile: SheetProfile
    key_column: str
    resolved_compare_key: str
    column_mappings: list[ColumnMapping]
    validation_issues: list[ValidationIssue]
    matched_rows: list[RowDiff]
    only_in_base: list[RowDiff]
    only_in_compare: list[RowDiff]

    @property
    def has_errors(self) -> bool:
        return any(issue.level == "error" for issue in self.validation_issues)
