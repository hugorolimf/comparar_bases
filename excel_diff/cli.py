from __future__ import annotations

import sys
from pathlib import Path

from InquirerPy import inquirer
from InquirerPy.base.control import Choice
from rich import box
from rich.console import Console
from rich.panel import Panel
from rich.table import Table

from excel_diff.analysis.schema_detector import analyze_workbook
from excel_diff.analysis.compatibility import rank_key_candidates
from excel_diff.io.workbook_reader import list_sheets
from excel_diff.comparison.diff_engine import compare_excels
from excel_diff.reporting.report_writer import write_outputs


def get_project_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


PROJECT_ROOT = get_project_root()
EXCEL_FOLDER = PROJECT_ROOT / "excel"
console = Console()


def make_choice_label(prefix: int, title: str, details: str | None = None) -> str:
    if details:
        return f"{prefix}. {title} - {details}"
    return f"{prefix}. {title}"


def main() -> int:
    console.print(
        Panel.fit(
            "Fluxo guiado: escolha o primeiro arquivo, depois o segundo arquivo, as abas e por fim as chaves de pareamento e identificação.",
            title="Comparar Excel - Nectar",
            border_style="cyan",
        )
    )

    excel_files = list_excel_files(EXCEL_FOLDER)
    if len(excel_files) < 2:
        console.print(f"[red]É necessário ter pelo menos 2 arquivos Excel na pasta {EXCEL_FOLDER}.[/red]")
        return 1

    console.print(f"\n[bold]Arquivos encontrados em {EXCEL_FOLDER}:[/bold]")
    show_file_options(excel_files)

    base_path = ask_file_choice("Escolha o primeiro arquivo para iniciar a comparação", excel_files)
    compare_files = [path for path in excel_files if str(path) != base_path]
    compare_path = ask_file_choice("Escolha o segundo arquivo para comparar com o primeiro selecionado", compare_files)

    if base_path == compare_path:
        console.print("[red]O primeiro arquivo e o segundo arquivo precisam ser diferentes.[/red]")
        return 1

    base_sheets = list_sheets(base_path)
    compare_sheets = list_sheets(compare_path)
    base_file_name = Path(base_path).name
    compare_file_name = Path(compare_path).name

    console.print(f"\n[green]Arquivo selecionado para iniciar a comparação:[/green] {base_file_name}")
    console.print(f"[green]Arquivo de comparação selecionado:[/green] {compare_file_name}")

    base_preview = analyze_workbook(base_path, base_sheets[0])
    base_sheet = ask_sheet_choice(f"Escolha a aba do arquivo {base_file_name}", base_sheets, default=base_preview.sheet_name)

    base_profile = analyze_workbook(base_path, base_sheet)
    while True:
        show_profile(base_file_name, base_profile)
        if confirm_profile(f"Esta é a aba correta do arquivo {base_file_name}?"):
            break
        base_sheet = ask_sheet_choice(f"Escolha novamente a aba do arquivo {base_file_name}", base_sheets, default=base_sheet)
        base_profile = analyze_workbook(base_path, base_sheet)

    compare_preview = analyze_workbook(compare_path, compare_sheets[0])
    compare_sheet = ask_sheet_choice(f"Escolha a aba do arquivo {compare_file_name}", compare_sheets, default=compare_preview.sheet_name)
    compare_profile = analyze_workbook(compare_path, compare_sheet)

    while True:
        show_profile(compare_file_name, compare_profile)
        if confirm_profile(f"Esta é a aba correta do arquivo {compare_file_name}?"):
            break
        compare_sheet = ask_sheet_choice(f"Escolha novamente a aba do arquivo {compare_file_name}", compare_sheets, default=compare_sheet)
        compare_profile = analyze_workbook(compare_path, compare_sheet)

    if not base_profile.column_profiles or not compare_profile.column_profiles:
        console.print("[red]Não foi possível identificar colunas suficientes para analisar os dois arquivos.[/red]")
        return 1

    key_candidates = rank_key_candidates(base_profile, compare_profile)
    if not key_candidates:
        console.print("[red]Não foi possível sugerir uma chave de pareamento a partir dos dois arquivos.[/red]")
        return 1

    console.print("\n[bold]Melhores opções de chave para pareamento:[/bold]")
    show_key_candidates(key_candidates[:10])

    key_column = choose_key_column(key_candidates)
    console.print(f"\n[green]Chave de pareamento selecionada:[/green] {key_column}")

    console.print("\n[bold]Agora escolha as colunas que vão ajudar a identificar os diffs no resultado final.[/bold]")
    console.print("[dim]A ordem escolhida no primeiro arquivo será usada para montar os pares no segundo arquivo.[/dim]")
    console.print("[dim]Se a ordem estiver errada, escolha novamente antes de confirmar.[/dim]")
    show_profile_columns(base_profile)
    base_diff_columns = choose_profile_columns(base_profile, f"Escolha as colunas de identificação do arquivo {base_file_name}", allow_multiple=True)
    show_profile_columns(compare_profile)
    compare_diff_columns = choose_profile_columns(compare_profile, f"Escolha as colunas de identificação do arquivo {compare_file_name}", allow_multiple=True)

    while len(base_diff_columns) != len(compare_diff_columns):
        console.print(f"[red]A quantidade de colunas escolhidas no arquivo {base_file_name} e no arquivo {compare_file_name} precisa ser igual.[/red]")
        show_profile_columns(base_profile)
        base_diff_columns = choose_profile_columns(base_profile, f"Escolha as colunas de identificação do arquivo {base_file_name}", allow_multiple=True)
        show_profile_columns(compare_profile)
        compare_diff_columns = choose_profile_columns(compare_profile, f"Escolha as colunas de identificação do arquivo {compare_file_name}", allow_multiple=True)

    compare_diff_columns = normalize_compare_columns(base_diff_columns, compare_diff_columns)

    while True:
        diff_key_pairs = list(zip(base_diff_columns, compare_diff_columns))
        console.print("\n[bold]Pares de identificação selecionados:[/bold]")
        for base_column, compare_column in diff_key_pairs:
            console.print(f"  {base_column} -> {compare_column}")
        console.print("[dim]A ordem acima foi ajustada com base na seleção do primeiro arquivo.[/dim]")
        if confirm_profile("Confirma?"):
            break
        console.print("[yellow]Vamos selecionar as colunas novamente.[/yellow]")
        show_profile_columns(base_profile)
        base_diff_columns = choose_profile_columns(base_profile, f"Escolha as colunas de identificação do arquivo {base_file_name}", allow_multiple=True)
        show_profile_columns(compare_profile)
        compare_diff_columns = choose_profile_columns(compare_profile, f"Escolha as colunas de identificação do arquivo {compare_file_name}", allow_multiple=True)

        while len(base_diff_columns) != len(compare_diff_columns):
            console.print(f"[red]A quantidade de colunas escolhidas no arquivo {base_file_name} e no arquivo {compare_file_name} precisa ser igual.[/red]")
            show_profile_columns(base_profile)
            base_diff_columns = choose_profile_columns(base_profile, f"Escolha as colunas de identificação do arquivo {base_file_name}", allow_multiple=True)
            show_profile_columns(compare_profile)
            compare_diff_columns = choose_profile_columns(compare_profile, f"Escolha as colunas de identificação do arquivo {compare_file_name}", allow_multiple=True)

    console.print("[cyan]Processando...[/cyan]")

    result = compare_excels(
        base_path,
        compare_path,
        base_sheet,
        compare_sheet,
        key_column,
        diff_key_pairs,
        base_profile=base_profile,
        compare_profile=compare_profile,
    )
    show_validation(result)

    if result.has_errors:
        console.print("[red]A comparação foi bloqueada devido a erros de validação.[/red]")
        return 1

    output_dir = ask_text("Onde você quer salvar o resultado? [padrão: ./saida]: ", default=str(Path.cwd() / "saida"))
    output_name = ask_text("Como você quer nomear o arquivo gerado? [padrão: diff_gerado]: ", default="diff_gerado")
    excel_path, visual_path, json_path = write_outputs(result, output_dir=output_dir, output_name=output_name)

    console.print(f"[green]Excel gerado em:[/green] {excel_path}")
    console.print(f"[green]Excel visual gerado em:[/green] {visual_path}")
    console.print(f"[green]JSON gerado em:[/green] {json_path}")
    return 0


def ask_text(prompt: str, default: str | None = None) -> str:
    return inquirer.text(message=prompt, default=default or "").execute().strip()


def confirm_profile(prompt: str) -> bool:
    return bool(inquirer.confirm(message=prompt, default=True).execute())


def normalize_compare_columns(base_columns: list[str], compare_columns: list[str]) -> list[str]:
    compare_remaining = list(compare_columns)
    ordered_compare: list[str] = []

    for base_column in base_columns:
        if base_column in compare_remaining:
            ordered_compare.append(base_column)
            compare_remaining.remove(base_column)

    ordered_compare.extend(compare_remaining)
    return ordered_compare


def list_excel_files(folder: Path) -> list[Path]:
    if not folder.exists():
        return []
    return sorted(
        [path for path in folder.iterdir() if path.is_file() and path.suffix.lower() in {".xlsx", ".xlsm"}],
        key=lambda path: path.name.lower(),
    )


def show_file_options(options: list[Path]) -> None:
    table = Table(box=box.SIMPLE_HEAVY, show_header=True, header_style="bold cyan")
    table.add_column("#", style="dim", width=4)
    table.add_column("Arquivo")
    for index, file_path in enumerate(options, start=1):
        table.add_row(str(index), file_path.name)
    console.print(table)


def ask_file_choice(prompt: str, options: list[Path]) -> str:
    selected = inquirer.select(
        message=prompt,
        choices=[Choice(value=file_path, name=make_choice_label(index, file_path.name)) for index, file_path in enumerate(options, start=1)],
    ).execute()
    console.print(f"[green]Selecionado:[/green] {selected.name}")
    return str(selected)


def ask_sheet_choice(prompt: str, options: list[str], default: str) -> str:
    selected = inquirer.select(
        message=prompt,
        choices=[Choice(value=sheet_name, name=make_choice_label(index, sheet_name)) for index, sheet_name in enumerate(options, start=1)],
        default=default,
    ).execute()
    console.print(f"[green]Selecionado:[/green] {selected}")
    return selected


def show_profile(label: str, profile) -> None:
    console.print(f"\n[bold]{label}[/bold]")
    show_profile_columns(profile)


def show_profile_columns(profile) -> None:
    table = Table(box=box.SIMPLE, show_header=True, header_style="bold cyan")
    table.add_column("#", style="dim", width=4)
    table.add_column("Coluna")
    table.add_column("Tipo")
    table.add_column("Unicidade", justify="right")
    table.add_column("Vazios", justify="right")
    for index, column in enumerate(profile.column_profiles, start=1):
        table.add_row(
            str(index),
            column.name,
            column.dominant_type,
            f"{column.unique_ratio:.2f}",
            f"{column.null_ratio:.2f}",
        )
    console.print(table)


def show_key_candidates(candidates) -> None:
    table = Table(box=box.SIMPLE_HEAVY, show_header=True, header_style="bold cyan")
    table.add_column("#", style="dim", width=4)
    table.add_column("Primeiro arquivo -> Segundo arquivo")
    table.add_column("Score", justify="right")
    for candidate in candidates:
        table.add_row(
            str(candidate.index),
            f"{candidate.base_column} -> {candidate.compare_column}",
            f"{candidate.score:.2f}",
        )
    console.print(table)


def choose_key_column(candidates) -> str:
    return choose_key_columns(candidates, prompt_label="Escolha a chave que melhor liga os dois arquivos")


def choose_key_columns(candidates, prompt_label: str = "Escolha a chave desejada") -> str:
    selected = inquirer.select(
        message=prompt_label,
        choices=[
            Choice(value=candidate.index, name=f"{candidate.base_column} -> {candidate.compare_column}")
            for candidate in candidates
        ],
        default=candidates[0].index,
    ).execute()
    selected_index = selected.get("value") if isinstance(selected, dict) else selected
    selected_candidate = next((candidate for candidate in candidates if candidate.index == selected_index), candidates[0])
    return selected_candidate.base_column


def choose_profile_columns(profile, prompt_label: str = "Escolha as colunas", allow_multiple: bool = False):
    columns = profile.column_profiles
    suggested = columns[0].name if columns else ""
    if allow_multiple:
        selected = inquirer.checkbox(
            message=prompt_label,
            instruction="Use espaço para marcar ou desmarcar as opções e Enter para confirmar.",
            choices=[
                Choice(
                    value=column.name,
                    name=column.name,
                )
                for index, column in enumerate(columns, start=1)
            ],
            validate=lambda result: len(result) >= 1,
            invalid_message="Escolha pelo menos uma coluna para continuar.",
        ).execute()
        return selected

    selected = inquirer.select(
        message=prompt_label,
        choices=[
            Choice(
                value=column.name,
                name=column.name,
            )
            for index, column in enumerate(columns, start=1)
        ],
        default=suggested,
    ).execute()
    return selected


def show_validation(result) -> None:
    if not result.validation_issues:
        console.print("[green]Validação: ok[/green]")
        return
    table = Table(box=box.SIMPLE, show_header=True, header_style="bold cyan")
    table.add_column("Nível", style="bold")
    table.add_column("Mensagem")
    for issue in result.validation_issues:
        style = "red" if issue.level == "error" else "yellow"
        table.add_row(f"[{style}]{issue.level}[/{style}]", issue.message)
    console.print("\n[bold]Validação:[/bold]")
    console.print(table)
