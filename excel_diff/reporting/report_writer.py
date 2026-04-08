from __future__ import annotations

from pathlib import Path

from excel_diff.models import ComparisonResult
from excel_diff.reporting.excel_report import write_excel_report
from excel_diff.reporting.json_report import write_json_report
from excel_diff.reporting.visual_report import write_visual_report


def write_outputs(result: ComparisonResult, output_dir: str | Path | None = None, output_name: str | None = None) -> tuple[Path, Path, Path]:
    target_dir = Path(output_dir or Path.cwd() / "saida")
    target_dir.mkdir(parents=True, exist_ok=True)
    stem = output_name or f"diff_{result.base_profile.sheet_name}_vs_{result.compare_profile.sheet_name}"
    excel_path = target_dir / f"{stem}.xlsx"
    visual_path = target_dir / f"{stem}_visual.xlsx"
    json_path = target_dir / f"{stem}.json"

    write_excel_report(result, excel_path)
    write_visual_report(result, visual_path)
    write_json_report(result, json_path)
    return excel_path, visual_path, json_path
