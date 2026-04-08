from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font

from excel_diff.models import ComparisonResult, DiffKeyPair, RowDiff


def write_visual_report(result: ComparisonResult, path: Path) -> None:
    workbook = Workbook()
    workbook.remove(workbook.active)

    write_sheet(workbook, "Igual", [row for row in result.matched_rows if row.status == "matched"], result, source_kind="base")
    write_sheet(workbook, "Alteração", [row for row in result.matched_rows if row.status == "changed"], result, source_kind="base")
    write_sheet(workbook, "Exclusão", result.only_in_base, result, source_kind="base")
    write_sheet(workbook, "Adição", result.only_in_compare, result, source_kind="compare")

    workbook.save(path)


def write_sheet(workbook: Workbook, title: str, rows: list[RowDiff], result: ComparisonResult, source_kind: str) -> None:
    sheet = workbook.create_sheet(title)
    headers = ["Categoria"]
    headers.extend(result.base_profile.headers)
    if title == "Alteração":
        for pair in result.diff_key_pairs:
            headers.extend([f"ibase-{pair.base_column}", f"diff-{pair.base_column}"])
    headers.extend(["Linha Base", "Linha Comparacao"])
    sheet.append(headers)
    for cell in sheet[1]:
        cell.font = Font(bold=True)

    compare_to_base = {mapping.compare_column: mapping.base_column for mapping in result.column_mappings}

    for row in rows:
        base_values = resolve_base_values(row, result, source_kind, compare_to_base)
        values = [title]
        values.extend([base_values.get(column, "") for column in result.base_profile.headers])

        if title == "Alteração":
            for pair in result.diff_key_pairs:
                identifier = find_identifier(row.diff_identifiers, pair)
                values.extend([identifier.get("base_value", ""), identifier.get("compare_value", "")])

        values.extend([
            row.base_row_number if row.base_row_number is not None else "",
            row.compare_row_number if row.compare_row_number is not None else "",
        ])

        sheet.append(values)


def resolve_base_values(
    row: RowDiff,
    result: ComparisonResult,
    source_kind: str,
    compare_to_base: dict[str, str],
) -> dict[str, object]:
    if source_kind == "compare":
        return project_compare_to_base(row.compare_values, compare_to_base, result.base_profile.headers)
    return row.base_values


def project_compare_to_base(compare_values: dict[str, object], compare_to_base: dict[str, str], base_headers: list[str]) -> dict[str, object]:
    projected = {header: "" for header in base_headers}
    for compare_column, value in compare_values.items():
        base_column = compare_to_base.get(compare_column)
        if base_column in projected:
            projected[base_column] = value
    return projected


def find_identifier(diff_identifiers: list[dict[str, object]], pair: DiffKeyPair) -> dict[str, object]:
    for identifier in diff_identifiers:
        if identifier.get("base_column") == pair.base_column and identifier.get("compare_column") == pair.compare_column:
            return identifier
    return {}