from __future__ import annotations

from collections import Counter
from pathlib import Path

from excel_diff.io.workbook_reader import list_sheets, sample_sheet_rows
from excel_diff.models import ColumnProfile, SheetProfile
from excel_diff.utils.normalization import classify_value, normalize_header, normalize_text, normalize_value


MAX_SAMPLE_ROWS = 100


def analyze_workbook(path: str | Path, sheet_name: str | None = None) -> SheetProfile:
    workbook_path = Path(path)
    sheets = list_sheets(workbook_path)
    if not sheets:
        raise ValueError(f"O arquivo não possui abas: {workbook_path}")

    chosen_sheet = sheet_name or sheets[0]
    if chosen_sheet not in sheets:
        raise ValueError(f"Aba não encontrada: {chosen_sheet}")

    sample_rows = sample_sheet_rows(workbook_path, chosen_sheet, max_rows=MAX_SAMPLE_ROWS)
    header_rows, headers, data_start_row = detect_headers(sample_rows)
    column_profiles = infer_column_profiles(sample_rows, headers, data_start_row)
    key_suggestions = suggest_key_columns(column_profiles)

    return SheetProfile(
        path=str(workbook_path),
        sheet_name=chosen_sheet,
        header_rows=header_rows,
        headers=headers,
        data_start_row=data_start_row,
        sample_row_count=len(sample_rows),
        sheet_width=len(headers),
        column_profiles=column_profiles,
        key_suggestions=key_suggestions,
    )


def detect_headers(sample_rows: list[tuple]) -> tuple[list[int], list[str], int]:
    if not sample_rows:
        return [1], [], 2

    scored_rows = []
    for index, row in enumerate(sample_rows, start=1):
        scored_rows.append((index, score_header_row(row)))

    best_single = max(scored_rows, key=lambda item: item[1])
    if len(sample_rows) >= 2:
        first_row = sample_rows[0]
        second_row = sample_rows[1]
        if looks_like_title_row(first_row) and score_header_row(second_row) >= max(best_single[1] * 0.75, 4.0):
            headers = build_combined_headers(first_row, second_row)
            return [1, 2], headers, 3

    header_row = best_single[0]
    headers = build_headers_from_row(sample_rows[header_row - 1])
    return [header_row], headers, header_row + 1


def score_header_row(row: tuple) -> float:
    non_empty_values = [value for value in row if not is_blank_like(value)]
    if not non_empty_values:
        return 0.0

    text_values = [value for value in non_empty_values if classify_value(value) == "string"]
    numeric_values = [value for value in non_empty_values if classify_value(value) in {"int", "float", "decimal"}]
    unique_texts = {normalize_text(value) for value in text_values if normalize_text(value)}
    avg_length = sum(len(normalize_text(value)) for value in text_values) / max(len(text_values), 1)
    data_penalty = score_data_likeness(row)

    score = 0.0
    score += len(non_empty_values) * 1.2
    score += len(text_values) * 2.0
    score += len(unique_texts) * 1.0
    score -= len(numeric_values) * 2.5
    score -= 3.0 if len(non_empty_values) == 1 else 0.0
    score -= 1.5 if avg_length > 60 else 0.0
    score -= data_penalty
    return score


def score_data_likeness(row: tuple) -> float:
    penalty = 0.0
    for value in row:
        if is_blank_like(value):
            continue
        text = normalize_text(value)
        raw = str(value).strip()
        if "@" in raw:
            penalty += 2.0
        if len(raw) > 25:
            penalty += 1.0
        if text and sum(char.isdigit() for char in text) >= max(3, len(text) // 2):
            penalty += 1.0
        if len(raw.split()) >= 3 and len(raw) > 20:
            penalty += 0.5
    return penalty


def looks_like_title_row(row: tuple) -> bool:
    non_empty_count = len([value for value in row if not is_blank_like(value)])
    width = max(len(row), 1)
    return non_empty_count <= max(3, int(width * 0.3))


def build_headers_from_row(row: tuple) -> list[str]:
    headers = []
    seen: dict[str, int] = {}
    for index, value in enumerate(row, start=1):
        header = normalize_header(value)
        if not header:
            header = f"col_{index}"
        if header in seen:
            seen[header] += 1
            header = f"{header}_{seen[header]}"
        else:
            seen[header] = 1
        headers.append(header)
    return headers


def build_combined_headers(first_row: tuple, second_row: tuple) -> list[str]:
    headers = []
    seen: dict[str, int] = {}
    max_width = max(len(first_row), len(second_row))
    for index in range(max_width):
        top = normalize_text(first_row[index]) if index < len(first_row) and first_row[index] is not None else ""
        bottom = normalize_text(second_row[index]) if index < len(second_row) and second_row[index] is not None else ""
        parts = [part for part in [top, bottom] if part]
        header = "_".join(parts)
        header = normalize_header(header)
        if not header:
            header = f"col_{index + 1}"
        if header in seen:
            seen[header] += 1
            header = f"{header}_{seen[header]}"
        else:
            seen[header] = 1
        headers.append(header)
    return headers


def infer_column_profiles(sample_rows: list[tuple], headers: list[str], data_start_row: int) -> list[ColumnProfile]:
    if not headers:
        return []

    data_rows = sample_rows[data_start_row - 1 :]
    column_profiles: list[ColumnProfile] = []
    for column_index, header in enumerate(headers):
        type_counts: Counter[str] = Counter()
        values: list[str] = []
        non_empty = 0
        for row in data_rows:
            value = row[column_index] if column_index < len(row) else None
            value_type = classify_value(value)
            type_counts[value_type] += 1
            if value_type != "empty":
                non_empty += 1
                if len(values) < 5:
                    values.append(str(value))
        total = len(data_rows)
        unique_values = {normalize_value(row[column_index] if column_index < len(row) else None) for row in data_rows}
        unique_values.discard("")
        dominant_type = determine_dominant_type(type_counts)
        column_profiles.append(
            ColumnProfile(
                index=column_index,
                name=header,
                normalized_name=normalize_header(header),
                dominant_type=dominant_type,
                type_counts=dict(type_counts),
                non_empty=non_empty,
                total=total,
                null_ratio=(total - non_empty) / total if total else 1.0,
                unique_ratio=len(unique_values) / non_empty if non_empty else 0.0,
                sample_values=values,
            )
        )
    return column_profiles


def determine_dominant_type(type_counts: Counter[str]) -> str:
    meaningful = {key: count for key, count in type_counts.items() if key != "empty" and count > 0}
    if not meaningful:
        return "empty"
    sorted_types = sorted(meaningful.items(), key=lambda item: item[1], reverse=True)
    if len(sorted_types) == 1:
        return sorted_types[0][0]
    top_name, top_count = sorted_types[0]
    second_count = sorted_types[1][1]
    if top_count >= second_count * 2:
        return top_name
    return "mixed"


def suggest_key_columns(column_profiles: list[ColumnProfile]) -> list[str]:
    ranked = sorted(
        column_profiles,
        key=lambda column: column_key_score(column),
        reverse=True,
    )
    return [column.name for column in ranked[:5] if column.name]


def column_key_score(column: ColumnProfile) -> float:
    score = column.unique_ratio * 4.0
    score -= column.null_ratio * 3.0
    score += min(column.non_empty / max(column.total, 1), 1.0)
    normalized_name = column.normalized_name
    identifier_terms = ("ra", "cpf", "id", "codigo", "cod", "matricula", "inep", "chave", "key", "registro")
    if any(term in normalized_name for term in identifier_terms):
        score += 3.0
    if any(term in normalized_name for term in ("nome", "descricao", "endereco", "email")):
        score -= 1.0
    if column.dominant_type in {"int", "float", "decimal"}:
        score += 0.5
    if len(normalized_name) <= 4:
        score += 1.0
    return score


def is_blank_like(value) -> bool:
    return normalize_value(value) == ""
