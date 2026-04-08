from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font

from excel_diff.models import ComparisonResult


def write_excel_report(result: ComparisonResult, path: Path) -> None:
    workbook = Workbook()
    summary_sheet = workbook.active
    summary_sheet.title = "Resumo"

    bold = Font(bold=True)
    summary_rows = [
        ("Arquivo base", result.base_profile.path),
        ("Aba base", result.base_profile.sheet_name),
        ("Arquivo comparação", result.compare_profile.path),
        ("Aba comparação", result.compare_profile.sheet_name),
        ("Chave base", result.key_column),
        ("Chave comparação", result.resolved_compare_key),
        ("Linhas somente base", len(result.only_in_base)),
        ("Linhas somente comparação", len(result.only_in_compare)),
        ("Linhas alteradas", len([row for row in result.matched_rows if row.status == "changed"])),
        ("Linhas iguais", len([row for row in result.matched_rows if row.status == "matched"])),
        ("Problemas de validação", len(result.validation_issues)),
    ]
    for row_index, (label, value) in enumerate(summary_rows, start=1):
        summary_sheet.cell(row=row_index, column=1, value=label).font = bold
        summary_sheet.cell(row=row_index, column=2, value=value)

    if result.validation_issues:
        start_row = len(summary_rows) + 3
        summary_sheet.cell(row=start_row, column=1, value="Validações").font = bold
        for offset, issue in enumerate(result.validation_issues, start=1):
            summary_sheet.cell(row=start_row + offset, column=1, value=issue.level)
            summary_sheet.cell(row=start_row + offset, column=2, value=issue.code)
            summary_sheet.cell(row=start_row + offset, column=3, value=issue.message)

    write_column_mappings(workbook, result)
    write_diff_sheet(workbook, "Adição", result.only_in_compare, result)
    write_diff_sheet(workbook, "Exclusão", result.only_in_base, result)
    write_diff_sheet(workbook, "Alteração", [row for row in result.matched_rows if row.status == "changed"], result)
    write_diff_sheet(workbook, "Igual", [row for row in result.matched_rows if row.status == "matched"], result)

    workbook.save(path)


def write_column_mappings(workbook: Workbook, result: ComparisonResult) -> None:
    sheet = workbook.create_sheet("Mapeamento")
    headers = ["Coluna Base", "Coluna Comparacao", "Metodo", "Score"]
    sheet.append(headers)
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    for mapping in result.column_mappings:
        sheet.append([mapping.base_column, mapping.compare_column, mapping.method, mapping.score])


def write_diff_sheet(workbook: Workbook, sheet_name: str, rows: list, result: ComparisonResult) -> None:
    sheet = workbook.create_sheet(sheet_name)
    headers = ["Categoria", "Chave", "Linha Base", "Linha Comparacao", "Status"]
    headers.extend([f"Base - {column}" for column in result.base_profile.headers])
    headers.extend([f"Comparacao - {column}" for column in result.compare_profile.headers])
    headers.append("Mudancas")
    sheet.append(headers)
    for cell in sheet[1]:
        cell.font = Font(bold=True)

    for row in rows:
        values = [
            sheet_name,
            row.key,
            row.base_row_number,
            row.compare_row_number,
            row.status,
        ]
        values.extend([row.base_values.get(column) for column in result.base_profile.headers])
        values.extend([row.compare_values.get(column) for column in result.compare_profile.headers])
        values.append(format_changes(row.changes))
        sheet.append(values)


def format_changes(changes: list[dict]) -> str:
    if not changes:
        return ""
    parts = []
    for change in changes:
        parts.append(
            f"{change.get('column_base')} => {change.get('base_value')} | {change.get('column_compare')} => {change.get('compare_value')}"
        )
    return "; ".join(parts)