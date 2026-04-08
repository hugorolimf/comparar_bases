from __future__ import annotations

from collections import defaultdict
from pathlib import Path

from excel_diff.analysis.compatibility import map_columns, resolve_compare_key, validate_comparison
from excel_diff.analysis.schema_detector import analyze_workbook
from excel_diff.io.workbook_reader import read_sheet_data
from excel_diff.models import ComparisonResult, DiffKeyPair, RowDiff, SheetProfile
from excel_diff.utils.normalization import normalize_value


def compare_excels(
    base_path: str | Path,
    compare_path: str | Path,
    base_sheet: str | None,
    compare_sheet: str | None,
    key_column: str,
    diff_key_pairs,
    base_profile: SheetProfile | None = None,
    compare_profile: SheetProfile | None = None,
) -> ComparisonResult:
    diff_key_pairs = normalize_diff_key_pairs(diff_key_pairs)

    if base_profile is None:
        base_profile = analyze_workbook(base_path, base_sheet)
    if compare_profile is None:
        compare_profile = analyze_workbook(compare_path, compare_sheet)

    validation_issues = validate_comparison(base_profile, compare_profile, key_column)
    resolved_compare_key = resolve_compare_key(base_profile, compare_profile, key_column)
    column_mappings = map_columns(base_profile, compare_profile)

    validation_issues.extend(validate_diff_key_pairs(base_profile, compare_profile, diff_key_pairs))

    if any(issue.level == "error" for issue in validation_issues):
        return ComparisonResult(
            base_profile=base_profile,
            compare_profile=compare_profile,
            key_column=key_column,
            resolved_compare_key=resolved_compare_key,
            diff_key_pairs=diff_key_pairs,
            column_mappings=column_mappings,
            validation_issues=validation_issues,
            matched_rows=[],
            only_in_base=[],
            only_in_compare=[],
        )

    matched_rows, only_in_base, only_in_compare = build_diff_rows(
        base_profile,
        compare_profile,
        key_column,
        resolved_compare_key,
        diff_key_pairs,
        column_mappings,
    )

    return ComparisonResult(
        base_profile=base_profile,
        compare_profile=compare_profile,
        key_column=key_column,
        resolved_compare_key=resolved_compare_key,
        diff_key_pairs=diff_key_pairs,
        column_mappings=column_mappings,
        validation_issues=validation_issues,
        matched_rows=matched_rows,
        only_in_base=only_in_base,
        only_in_compare=only_in_compare,
    )


def build_diff_rows(
    base_profile: SheetProfile,
    compare_profile: SheetProfile,
    key_column: str,
    resolved_compare_key: str,
    diff_key_pairs: list[DiffKeyPair],
    column_mappings,
):
    base_rows = read_sheet_data(base_profile.path, base_profile.sheet_name, base_profile.data_start_row)
    compare_rows = read_sheet_data(compare_profile.path, compare_profile.sheet_name, compare_profile.data_start_row)

    base_by_key = group_rows_by_key(base_rows, base_profile.headers, key_column)
    compare_by_key = group_rows_by_key(compare_rows, compare_profile.headers, resolved_compare_key)

    compare_lookup = {mapping.base_column: mapping.compare_column for mapping in column_mappings}

    matched_rows: list[RowDiff] = []
    only_in_base: list[RowDiff] = []
    only_in_compare: list[RowDiff] = []

    all_keys = sorted(set(base_by_key) | set(compare_by_key))
    for key in all_keys:
        base_records = base_by_key.get(key, [])
        compare_records = compare_by_key.get(key, [])
        max_length = max(len(base_records), len(compare_records))
        for index in range(max_length):
            base_record = base_records[index] if index < len(base_records) else None
            compare_record = compare_records[index] if index < len(compare_records) else None
            if base_record and compare_record:
                diff_identifiers = build_diff_identifiers(base_record, compare_record, diff_key_pairs)
                changes = [
                    {
                        "column_base": identifier["base_column"],
                        "column_compare": identifier["compare_column"],
                        "base_value": identifier["base_value"],
                        "compare_value": identifier["compare_value"],
                    }
                    for identifier in diff_identifiers
                    if normalize_value(identifier["base_value"]) != normalize_value(identifier["compare_value"])
                ]
                matched_rows.append(
                    RowDiff(
                        key=key,
                        base_row_number=base_record["row_number"],
                        compare_row_number=compare_record["row_number"],
                        status="changed" if changes else "matched",
                        diff_identifiers=diff_identifiers,
                        base_values=base_record["values"],
                        compare_values=compare_record["values"],
                        changes=changes,
                    )
                )
            elif base_record:
                only_in_base.append(
                    RowDiff(
                        key=key,
                        base_row_number=base_record["row_number"],
                        compare_row_number=None,
                        status="only_in_base",
                        diff_identifiers=build_diff_identifiers(base_record, None, diff_key_pairs),
                        base_values=base_record["values"],
                        compare_values={},
                        changes=[],
                    )
                )
            elif compare_record:
                only_in_compare.append(
                    RowDiff(
                        key=key,
                        base_row_number=None,
                        compare_row_number=compare_record["row_number"],
                        status="only_in_compare",
                        diff_identifiers=build_diff_identifiers(None, compare_record, diff_key_pairs),
                        base_values={},
                        compare_values=compare_record["values"],
                        changes=[],
                    )
                )

    return matched_rows, only_in_base, only_in_compare


def group_rows_by_key(rows: list[tuple], headers: list[str], key_column: str) -> dict[str, list[dict]]:
    grouped: dict[str, list[dict]] = defaultdict(list)
    headers_index = {header: index for index, header in enumerate(headers)}
    key_index = headers_index[key_column]
    for row_number, row in enumerate(rows, start=1):
        values = {headers[index]: row[index] if index < len(row) else None for index in range(len(headers))}
        key_value = normalize_value(values.get(key_column))
        if not key_value:
            key_value = f"__blank__:{row_number}"
        grouped[key_value].append({"row_number": row_number, "values": values, "key_index": key_index})
    return grouped


def validate_diff_key(base_profile: SheetProfile, compare_profile: SheetProfile, diff_key_column: str, resolved_compare_diff_key: str):
    from excel_diff.models import ValidationIssue

    issues: list[ValidationIssue] = []
    base_lookup = {column.name: column for column in base_profile.column_profiles}
    compare_lookup = {column.name: column for column in compare_profile.column_profiles}

    if diff_key_column not in base_lookup:
        issues.append(ValidationIssue(level="error", code="base_diff_key_missing", message=f"A chave de diffs '{diff_key_column}' não existe na base."))
    if not resolved_compare_diff_key or resolved_compare_diff_key not in compare_lookup:
        issues.append(ValidationIssue(level="error", code="compare_diff_key_missing", message=f"Não foi possível resolver a chave de diffs '{diff_key_column}' na comparação."))

    if diff_key_column in base_lookup:
        diff_key = base_lookup[diff_key_column]
        if diff_key.null_ratio > 0.2:
            issues.append(ValidationIssue(level="warning", code="base_diff_key_nulls", message=f"A chave de diffs '{diff_key_column}' tem alto percentual de valores vazios na base."))
    if resolved_compare_diff_key and resolved_compare_diff_key in compare_lookup:
        diff_key = compare_lookup[resolved_compare_diff_key]
        if diff_key.null_ratio > 0.2:
            issues.append(ValidationIssue(level="warning", code="compare_diff_key_nulls", message=f"A chave de diffs '{resolved_compare_diff_key}' tem alto percentual de valores vazios na comparação."))

    return issues


def validate_diff_key_pairs(base_profile: SheetProfile, compare_profile: SheetProfile, diff_key_pairs: list[DiffKeyPair]):
    issues = []
    for diff_key_pair in diff_key_pairs:
        issues.extend(validate_diff_key(base_profile, compare_profile, diff_key_pair.base_column, diff_key_pair.compare_column))
    return issues


def build_diff_identifiers(base_record: dict | None, compare_record: dict | None, diff_key_pairs: list[DiffKeyPair]) -> list[dict[str, object]]:
    identifiers = []
    for diff_key_pair in diff_key_pairs:
        base_value = base_record["values"].get(diff_key_pair.base_column, "") if base_record else ""
        compare_value = compare_record["values"].get(diff_key_pair.compare_column, "") if compare_record else ""
        identifiers.append(
            {
                "base_column": diff_key_pair.base_column,
                "compare_column": diff_key_pair.compare_column,
                "base_value": base_value,
                "compare_value": compare_value,
            }
        )
    return identifiers


def normalize_diff_key_pairs(diff_key_pairs) -> list[DiffKeyPair]:
    normalized: list[DiffKeyPair] = []
    seen: set[tuple[str, str]] = set()
    for pair in diff_key_pairs or []:
        if isinstance(pair, DiffKeyPair):
            base_column = pair.base_column
            compare_column = pair.compare_column
        else:
            base_column, compare_column = pair
        key = (base_column, compare_column)
        if key not in seen:
            normalized.append(DiffKeyPair(base_column=base_column, compare_column=compare_column))
            seen.add(key)
    return normalized
