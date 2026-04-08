from __future__ import annotations

from pathlib import Path

from excel_diff.analysis.schema_detector import analyze_workbook
from excel_diff.analysis.compatibility import rank_key_candidates
from excel_diff.io.workbook_reader import list_sheets
from excel_diff.comparison.diff_engine import compare_excels
from excel_diff.reporting.report_writer import write_outputs


PROJECT_ROOT = Path(__file__).resolve().parents[1]
EXCEL_FOLDER = PROJECT_ROOT / "excel"


def main() -> int:
    print("Comparador inteligente de Excel")
    print("------------------------------")
    print("Fluxo guiado: escolha os arquivos, depois as abas e por fim a chave sugerida.")

    excel_files = list_excel_files(EXCEL_FOLDER)
    if len(excel_files) < 2:
        print(f"É necessário ter pelo menos 2 arquivos Excel na pasta {EXCEL_FOLDER}.")
        return 1

    print(f"\nArquivos disponíveis em {EXCEL_FOLDER}:")
    show_file_options(excel_files)

    base_path = ask_file_choice("Escolha o Excel base", excel_files)
    compare_path = ask_file_choice("Escolha o Excel de comparação", excel_files)

    if base_path == compare_path:
        print("Os arquivos base e comparação precisam ser diferentes.")
        return 1

    base_sheets = list_sheets(base_path)
    compare_sheets = list_sheets(compare_path)

    print(f"\nArquivo base selecionado: {Path(base_path).name}")
    print(f"Arquivo de comparação selecionado: {Path(compare_path).name}")

    print("\nAbas disponíveis na base:")
    show_sheet_options(base_sheets)
    print("\nAbas disponíveis na comparação:")
    show_sheet_options(compare_sheets)

    base_preview = analyze_workbook(base_path, base_sheets[0])
    compare_preview = analyze_workbook(compare_path, compare_sheets[0])

    base_sheet = ask_sheet_choice("Escolha a aba base: ", base_sheets, default=base_preview.sheet_name)
    compare_sheet = ask_sheet_choice("Escolha a aba comparação: ", compare_sheets, default=compare_preview.sheet_name)

    base_profile = analyze_workbook(base_path, base_sheet)
    compare_profile = analyze_workbook(compare_path, compare_sheet)

    if not base_profile.column_profiles or not compare_profile.column_profiles:
        print("Não foi possível identificar colunas suficientes para comparar os arquivos.")
        return 1

    show_profile("Base", base_profile)
    show_profile("Comparação", compare_profile)

    key_candidates = rank_key_candidates(base_profile, compare_profile)
    if not key_candidates:
        print("Não foi possível sugerir uma chave com base nos dois arquivos.")
        return 1

    print("\nMelhores matches para chave de comparação:")
    show_key_candidates(key_candidates[:10])

    key_column = choose_key_column(key_candidates)
    print(f"\nChave selecionada: {key_column}")

    result = compare_excels(
        base_path,
        compare_path,
        base_sheet,
        compare_sheet,
        key_column,
        base_profile=base_profile,
        compare_profile=compare_profile,
    )
    show_validation(result)

    if result.has_errors:
        print("Comparação bloqueada devido a erros de validação.")
        return 1

    output_dir = ask_text("Diretório de saída [padrão: ./saida]: ", default=str(Path.cwd() / "saida"))
    output_name = ask_text("Nome do arquivo de saída [padrão: diff_gerado]: ", default="diff_gerado")
    excel_path, json_path = write_outputs(result, output_dir=output_dir, output_name=output_name)

    print(f"Excel gerado em: {excel_path}")
    print(f"JSON gerado em: {json_path}")
    return 0


def ask_text(prompt: str, default: str | None = None) -> str:
    value = input(prompt).strip()
    if value:
        return value
    return default or ""


def list_excel_files(folder: Path) -> list[Path]:
    if not folder.exists():
        return []
    return sorted(
        [path for path in folder.iterdir() if path.is_file() and path.suffix.lower() in {".xlsx", ".xlsm"}],
        key=lambda path: path.name.lower(),
    )


def show_file_options(options: list[Path]) -> None:
    for index, file_path in enumerate(options, start=1):
        print(f"  {index}. {file_path.name}")


def ask_file_choice(prompt: str, options: list[Path]) -> str:
    while True:
        value = input(f"{prompt} [número]: ").strip().strip('"')
        if value.isdigit():
            index = int(value) - 1
            if 0 <= index < len(options):
                selected = options[index]
                print(f"Selecionado: {selected.name}")
                return str(selected)
        print("Escolha inválida. Digite apenas o número do arquivo na lista.")


def ask_sheet_choice(prompt: str, options: list[str], default: str) -> str:
    while True:
        value = input(f"{prompt} [número, padrão: {default}]: ").strip()
        if not value:
            print(f"Selecionado: {default}")
            return default
        if value.isdigit():
            index = int(value) - 1
            if 0 <= index < len(options):
                selected = options[index]
                print(f"Selecionado: {selected}")
                return selected
        print("Escolha inválida. Digite apenas o número da aba na lista.")


def show_sheet_options(options: list[str]) -> None:
    for index, sheet_name in enumerate(options, start=1):
        print(f"  {index}. {sheet_name}")


def show_profile(label: str, profile) -> None:
    print(f"\n{label}:")
    print(f"  Aba: {profile.sheet_name}")
    print(f"  Linhas amostradas: {profile.sample_row_count}")
    print(f"  Cabeçalhos detectados: {', '.join(profile.headers[:10])}")
    print(f"  Sugestões de chave: {', '.join(profile.key_suggestions[:5])}")


def show_key_candidates(candidates) -> None:
    for candidate in candidates:
        recommendation = " (recomendado)" if candidate.index == 1 else ""
        print(
            f"  {candidate.index}. {candidate.base_column} -> {candidate.compare_column}{recommendation}"
        )
        print(
            f"     match={candidate.score:.2f} | base_unicidade={candidate.base_unique_ratio:.2f} | "
            f"comparacao_unicidade={candidate.compare_unique_ratio:.2f} | tipo_base={candidate.base_type} | tipo_comparacao={candidate.compare_type}"
        )


def choose_key_column(candidates) -> str:
    suggested = candidates[0].base_column
    while True:
        choice = input(f"Escolha a chave desejada [Enter = {suggested}]: ").strip()
        if not choice:
            return suggested
        if choice.isdigit():
            index = int(choice) - 1
            if 0 <= index < len(candidates):
                return candidates[index].base_column
        print("Escolha inválida. Digite apenas o número da chave na lista.")


def show_validation(result) -> None:
    if not result.validation_issues:
        print("Validação: ok")
        return
    print("\nValidação:")
    for issue in result.validation_issues:
        print(f"  [{issue.level}] {issue.message}")
