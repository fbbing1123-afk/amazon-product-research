from __future__ import annotations

import argparse
import os
import shutil
import sys
import time
from pathlib import Path

MODULE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = MODULE_DIR.parents[1]
MODULES_DIR = PROJECT_DIR / "modules"
if str(MODULES_DIR) not in sys.path:
    sys.path.insert(0, str(MODULES_DIR))

from shared.runtime_monitor import monitor_event
from shared.subprocess_runner import run_command


def set_monitor_paths(jsonl_path: Path | None, log_path: Path | None) -> None:
    """让续跑进程复用原工作流的监控文件。"""
    if jsonl_path:
        os.environ["ECOMMERCE_WORKFLOW_MONITOR_PATH"] = str(jsonl_path.resolve())
    if log_path:
        os.environ["ECOMMERCE_WORKFLOW_MONITOR_LOG"] = str(log_path.resolve())


def resolve_powershell_executable(explicit: str | None) -> str:
    """优先选择可用的 PowerShell 7，避免旧版运行时缺少 ZIP 类型。"""
    if explicit:
        return explicit
    for candidate in ("pwsh", "powershell"):
        resolved = shutil.which(candidate)
        if resolved:
            return resolved
    raise FileNotFoundError("未找到 pwsh 或 powershell，可用 --powershell 指定路径。")


def run_competitor_postprocessing(
    *,
    workbook: Path,
    analysis: Path,
    output: Path,
    python_exe: Path,
    powershell_exe: str,
) -> Path:
    """连续完成竞品语义回填、差评透视和原生切片器。"""
    workbook = workbook.resolve()
    analysis = analysis.resolve()
    output = output.resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    review_output = output.with_name(output.stem + "_review_analysis.xlsx")
    before_slicers = output.with_name(output.stem + "_before_slicers.xlsx")
    backfill_script = MODULE_DIR / "backfill_review_analysis.py"
    pain_view_script = MODULE_DIR / "build_pain_slicer_view.py"
    slicer_script = MODULE_DIR / "add_pain_slicers.ps1"

    started_at = time.perf_counter()
    monitor_event(
        "competitor_analysis_started",
        workbook=str(workbook),
        analysis=str(analysis),
        output=str(output),
    )

    stage_started = time.perf_counter()
    monitor_event("competitor_stage_started", stage="review_analysis_backfill")
    run_command(
        "competitor_analysis.backfill_review_analysis",
        [
            python_exe,
            backfill_script,
            "--analysis",
            analysis,
            "--workbook",
            workbook,
            "--output",
            review_output,
        ],
        cwd=PROJECT_DIR,
        watch_paths=(review_output,),
    )
    monitor_event(
        "competitor_stage_finished",
        stage="review_analysis_backfill",
        duration_seconds=round(time.perf_counter() - stage_started, 1),
        output=str(review_output),
    )

    stage_started = time.perf_counter()
    monitor_event("competitor_stage_started", stage="pain_slicer_view")
    run_command(
        "competitor_analysis.build_pain_slicer_view",
        [
            python_exe,
            pain_view_script,
            "--source",
            review_output,
            "--output",
            before_slicers,
        ],
        cwd=PROJECT_DIR,
        watch_paths=(before_slicers,),
    )
    monitor_event(
        "competitor_stage_finished",
        stage="pain_slicer_view",
        duration_seconds=round(time.perf_counter() - stage_started, 1),
        output=str(before_slicers),
    )

    stage_started = time.perf_counter()
    monitor_event("competitor_stage_started", stage="native_slicers")
    run_command(
        "competitor_analysis.add_pain_slicers",
        [powershell_exe, "-ExecutionPolicy", "Bypass", "-File", slicer_script, "-WorkbookPath", before_slicers],
        cwd=PROJECT_DIR,
        watch_paths=(before_slicers,),
    )
    monitor_event(
        "competitor_stage_finished",
        stage="native_slicers",
        duration_seconds=round(time.perf_counter() - stage_started, 1),
        output=str(before_slicers),
    )

    # 切片器已经写入后只复制文件，不再用 openpyxl 保存，避免破坏原生切片器。
    shutil.copy2(before_slicers, output)
    total_seconds = round(time.perf_counter() - started_at, 1)
    monitor_event(
        "competitor_analysis_finished",
        output=str(output),
        duration_seconds=total_seconds,
        next_step="主智能体立即执行最终结构和视觉验收",
    )
    return output


def main() -> int:
    parser = argparse.ArgumentParser(
        description="在完整流程交接点连续完成竞品分析 Excel 后处理。"
    )
    parser.add_argument("--workbook", type=Path, required=True, help="已完成评论抓取和商品卡回填的工作副本")
    parser.add_argument("--analysis", type=Path, required=True, help="已填写优缺点、总结和改进方式的 handoff JSON")
    parser.add_argument("--output", type=Path, required=True, help="最终 Excel 输出路径")
    parser.add_argument("--python", type=Path, default=Path(sys.executable))
    parser.add_argument("--powershell", help="切片器脚本使用的 PowerShell，可留空自动优先选择 pwsh")
    parser.add_argument("--monitor-jsonl", type=Path)
    parser.add_argument("--monitor-log", type=Path)
    args = parser.parse_args()
    set_monitor_paths(args.monitor_jsonl, args.monitor_log)
    result = run_competitor_postprocessing(
        workbook=args.workbook,
        analysis=args.analysis,
        output=args.output,
        python_exe=args.python,
        powershell_exe=resolve_powershell_executable(args.powershell),
    )
    print(result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
