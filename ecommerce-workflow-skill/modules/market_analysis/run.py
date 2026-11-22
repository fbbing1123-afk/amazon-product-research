from __future__ import annotations

import re
from pathlib import Path

from openpyxl import load_workbook

from shared.excel_inspect import normalize_market_source
from shared.image_conversion import make_wps_embedded_version
from shared.subprocess_runner import run_command

from .config import PROJECT_DIR, SCRIPT_PATH, SLICER_SCRIPT_PATH, STATIC_CHART_SCRIPT_PATH


def product_name_from_workflow_name(name: str) -> str:
    """从标准流程文件名中提取产品名称。"""
    product_name = re.sub(r"-?产品调研-?\d{6}.*$", "", name).strip("-_ ")
    return product_name or name


def validate_market_analysis_output(path: Path) -> None:
    """阻止未写入聚合数据或未冻结图表的中间文件继续流转。"""
    workbook = load_workbook(path, data_only=True, read_only=False)
    try:
        worksheet = next(
            (item for item in workbook.worksheets if item.title.endswith("市场分析")),
            None,
        )
        if worksheet is None:
            raise RuntimeError("市场分析阶段没有生成市场分析工作表。")

        required_cells = ("B3", "D3", "F3", "B5", "D5", "F5", "B6", "D6", "F6")
        missing_cells = [coordinate for coordinate in required_cells if worksheet[coordinate].value is None]
        if missing_cells:
            raise RuntimeError(
                "市场分析聚合数据未写入，禁止继续后续流程："
                + ", ".join(missing_cells)
            )
        images = list(getattr(worksheet, "_images", []))
        if len(images) != 3:
            raise RuntimeError(
                f"市场分析页应有 3 张冻结图表，实际检测到 {len(images)} 张，禁止继续后续流程。"
            )
        expected_anchors = {(11, 0), (23, 0), (37, 0)}
        actual_anchors = set()
        for image in images:
            marker = getattr(getattr(image, "anchor", None), "_from", None)
            if marker is not None:
                actual_anchors.add((marker.row, marker.col))
        if actual_anchors != expected_anchors:
            raise RuntimeError(
                "市场分析页冻结图表锚点不完整或重复，"
                f"期望 {sorted(expected_anchors)}，实际 {sorted(actual_anchors)}。"
            )
    finally:
        workbook.close()


def run_market_analysis(
    *,
    source: Path,
    name: str,
    output_dir: Path,
    python_exe: str,
    output_path: Path | None = None,
    source_slicers: bool = False,
    original_image_source: Path | None = None,
) -> dict[str, Path]:
    print("[market_analysis] start")
    work_dir = output_dir / "work" / name
    work_dir.mkdir(parents=True, exist_ok=True)
    market_source = normalize_market_source(source, work_dir)
    output_path = Path(output_path or output_dir / f"{name}_market_analysis.xlsx")

    generate_command: list[str | Path] = [
        python_exe,
        SCRIPT_PATH,
        market_source,
        output_path,
        "--product-name",
        product_name_from_workflow_name(name),
    ]
    if source_slicers:
        generate_command.append("--source-slicers")
    run_command(
        "market_analysis.generate",
        generate_command,
        cwd=PROJECT_DIR,
    )
    chart_image_dir = work_dir / f"{name}_market_analysis_chart_images"
    run_command(
        "market_analysis.freeze_charts",
        [
            python_exe,
            STATIC_CHART_SCRIPT_PATH,
            "--input",
            output_path,
            "--output",
            output_path,
            "--image-dir",
            chart_image_dir,
        ],
        cwd=PROJECT_DIR,
    )
    validate_market_analysis_output(output_path)
    if source_slicers:
        embedded_output = work_dir / f"{name}_slicer_images_embedded.xlsx"
        temp_output = work_dir / f"{name}_slicer_images_temp.xlsx"
        embedding = make_wps_embedded_version(
            original=Path(original_image_source or output_path),
            floating_workbook=output_path,
            output=embedded_output,
            temp_output=temp_output,
            embed_extra_sheets={"切片联动"},
        )
        embedded_output.replace(output_path)
        temp_output.unlink(missing_ok=True)
        embedded_count = embedding.get("restore_summary", {}).get("切片联动", {}).get("extra_embedded_images", 0)
        print(f"[market_analysis.source_slicers] embedded cell images: {embedded_count}")
        run_command(
            "market_analysis.source_slicers",
            [
                "powershell.exe",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                SLICER_SCRIPT_PATH,
                "-WorkbookPath",
                output_path,
            ],
            cwd=PROJECT_DIR,
        )

    return {
        "market_analysis_report": output_path,
        "current_workbook": output_path,
    }
