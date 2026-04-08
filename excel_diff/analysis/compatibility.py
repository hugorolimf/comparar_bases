from __future__ import annotations

from difflib import SequenceMatcher

from excel_diff.models import ColumnMapping, KeyMatchCandidate, SheetProfile, ValidationIssue
from excel_diff.utils.normalization import normalize_header, normalize_text


def map_columns(base_profile: SheetProfile, compare_profile: SheetProfile) -> list[ColumnMapping]:
    compare_by_norm = {column.normalized_name: column.name for column in compare_profile.column_profiles}
    used_compare_columns: set[str] = set()
    mappings: list[ColumnMapping] = []

    for base_column in base_profile.column_profiles:
        exact = compare_by_norm.get(base_column.normalized_name)
        if exact and exact not in used_compare_columns:
            used_compare_columns.add(exact)
            mappings.append(ColumnMapping(base_column=base_column.name, compare_column=exact, method="exact", score=1.0))

    for base_column in base_profile.column_profiles:
        if any(mapping.base_column == base_column.name for mapping in mappings):
            continue
        exact = compare_by_norm.get(base_column.normalized_name)
        if exact and exact not in used_compare_columns:
            used_compare_columns.add(exact)
            mappings.append(ColumnMapping(base_column=base_column.name, compare_column=exact, method="exact", score=1.0))
            continue

        best_match = None
        best_score = 0.0
        for compare_column in compare_profile.column_profiles:
            if compare_column.name in used_compare_columns:
                continue
            score = similarity(base_column.normalized_name, compare_column.normalized_name)
            if score > best_score:
                best_score = score
                best_match = compare_column.name
        if best_match and best_score >= 0.72:
            used_compare_columns.add(best_match)
            mappings.append(ColumnMapping(base_column=base_column.name, compare_column=best_match, method="fuzzy", score=best_score))

    return mappings


def rank_key_candidates(base_profile: SheetProfile, compare_profile: SheetProfile) -> list[KeyMatchCandidate]:
    mappings = map_columns(base_profile, compare_profile)
    base_lookup = {column.name: column for column in base_profile.column_profiles}
    compare_lookup = {column.name: column for column in compare_profile.column_profiles}

    candidates: list[KeyMatchCandidate] = []
    for mapping in mappings:
        base_column = base_lookup.get(mapping.base_column)
        compare_column = compare_lookup.get(mapping.compare_column)
        if not base_column or not compare_column:
            continue

        score = mapping.score * 3.0
        score += (base_column.unique_ratio + compare_column.unique_ratio) * 1.5
        score += identifier_bonus(base_column.normalized_name, compare_column.normalized_name)
        score += type_bonus(base_column.dominant_type, compare_column.dominant_type)
        score -= (base_column.null_ratio + compare_column.null_ratio) * 1.2

        candidates.append(
            KeyMatchCandidate(
                index=0,
                base_column=base_column.name,
                compare_column=compare_column.name,
                method=mapping.method,
                score=score,
                base_unique_ratio=base_column.unique_ratio,
                compare_unique_ratio=compare_column.unique_ratio,
                base_null_ratio=base_column.null_ratio,
                compare_null_ratio=compare_column.null_ratio,
                base_type=base_column.dominant_type,
                compare_type=compare_column.dominant_type,
            )
        )

    candidates.sort(key=lambda candidate: candidate.score, reverse=True)
    for index, candidate in enumerate(candidates, start=1):
        candidate.index = index
    return candidates


def validate_comparison(base_profile: SheetProfile, compare_profile: SheetProfile, key_column: str) -> list[ValidationIssue]:
    issues: list[ValidationIssue] = []
    base_lookup = {column.name: column for column in base_profile.column_profiles}
    compare_lookup = {column.name: column for column in compare_profile.column_profiles}

    if key_column not in base_lookup:
        issues.append(ValidationIssue(level="error", code="base_key_missing", message=f"A chave '{key_column}' não existe na base."))

    resolved_compare_key = resolve_compare_key(base_profile, compare_profile, key_column)
    if not resolved_compare_key:
        issues.append(ValidationIssue(level="error", code="compare_key_missing", message=f"Não foi possível resolver a chave '{key_column}' na planilha de comparação."))

    shared_columns = map_columns(base_profile, compare_profile)
    if not shared_columns:
        issues.append(ValidationIssue(level="error", code="no_shared_columns", message="Não foi possível mapear colunas em comum entre os dois arquivos."))

    if key_column in base_lookup:
        base_key = base_lookup[key_column]
        if base_key.null_ratio > 0.2:
            issues.append(ValidationIssue(level="warning", code="base_key_nulls", message=f"A chave '{key_column}' tem alto percentual de valores vazios na base."))
        if base_key.unique_ratio < 0.5:
            issues.append(ValidationIssue(level="warning", code="base_key_duplicates", message=f"A chave '{key_column}' parece ter baixa unicidade na base."))

    if resolved_compare_key and resolved_compare_key in compare_lookup:
        compare_key = compare_lookup[resolved_compare_key]
        if compare_key.null_ratio > 0.2:
            issues.append(ValidationIssue(level="warning", code="compare_key_nulls", message=f"A chave '{resolved_compare_key}' tem alto percentual de valores vazios na comparação."))
        if compare_key.unique_ratio < 0.5:
            issues.append(ValidationIssue(level="warning", code="compare_key_duplicates", message=f"A chave '{resolved_compare_key}' parece ter baixa unicidade na comparação."))

    return issues


def resolve_compare_key(base_profile: SheetProfile, compare_profile: SheetProfile, key_column: str) -> str:
    base_norm = normalize_header(key_column)
    for compare_column in compare_profile.column_profiles:
        if compare_column.normalized_name == base_norm:
            return compare_column.name

    base_lookup = {column.normalized_name: column.name for column in base_profile.column_profiles}
    if base_norm not in base_lookup:
        return ""

    best_match = ""
    best_score = 0.0
    for compare_column in compare_profile.column_profiles:
        score = similarity(base_norm, compare_column.normalized_name)
        if score > best_score:
            best_score = score
            best_match = compare_column.name
    return best_match if best_score >= 0.72 else ""


def similarity(left: str, right: str) -> float:
    left_text = normalize_text(left)
    right_text = normalize_text(right)
    if not left_text or not right_text:
        return 0.0
    if left_text == right_text:
        return 1.0
    if left_text.startswith(right_text) or right_text.startswith(left_text):
        shorter = min(len(left_text), len(right_text))
        longer = max(len(left_text), len(right_text))
        if shorter <= 4 or shorter / longer >= 0.5:
            return 0.95
    if left_text in right_text or right_text in left_text:
        shorter = min(len(left_text), len(right_text))
        if shorter <= 4:
            return 0.9
    return SequenceMatcher(None, left_text, right_text).ratio()


def identifier_bonus(left_name: str, right_name: str) -> float:
    bonus = 0.0
    identifier_terms = ("ra", "cpf", "id", "codigo", "cod", "matricula", "inep", "chave", "key", "registro")
    left_has = any(term in left_name for term in identifier_terms)
    right_has = any(term in right_name for term in identifier_terms)
    if left_has and right_has:
        bonus += 2.5
    elif left_has or right_has:
        bonus += 1.0
    if len(left_name) <= 6 or len(right_name) <= 6:
        bonus += 0.5
    return bonus


def type_bonus(left_type: str, right_type: str) -> float:
    if left_type == right_type and left_type != "mixed":
        return 1.0
    numeric_types = {"int", "float", "decimal"}
    if left_type in numeric_types and right_type in numeric_types:
        return 0.6
    if {left_type, right_type} & {"string", "empty"}:
        return 0.0
    return 0.2
