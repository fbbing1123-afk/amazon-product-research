from __future__ import annotations

from pathlib import Path

from shared.subprocess_runner import run_command

from .config import BROWSER_STATE_DIR, DEFAULT_PREFIX, MODULE_DIR, SCRIPT_PATH


def run_keyword_trends(
    *,
    source: Path,
    name: str,
    output_dir: Path,
    python_exe: str,
    prefix: str = DEFAULT_PREFIX,
    skip_screenshot: bool = False,
    keep_run_artifacts: bool = False,
    debug_port: int | None = None,
    use_existing_browser: bool = True,
    site: str = "US",
    amazon_postal_code: str | None = None,
    ppc_bid_mode: str = "latest-week",
) -> dict[str, Path]:
    print("[keyword_trends] start")
    output_path = output_dir / f"{name}_with_keyword_trends.xlsx"
    command: list[str | Path] = [
        python_exe,
        "-u",
        SCRIPT_PATH,
        "--excel",
        source,
        "--output",
        output_path,
        "--project-dir",
        MODULE_DIR,
        "--browser-state-dir",
        BROWSER_STATE_DIR,
        "--prefix",
        prefix,
        "--site",
        site,
        "--ppc-bid-mode",
        ppc_bid_mode,
    ]
    if amazon_postal_code:
        command.extend(["--amazon-postal-code", amazon_postal_code])
    if debug_port is not None:
        command.extend(["--debug-port", str(debug_port)])
    if use_existing_browser:
        command.append("--use-existing-browser")
    if skip_screenshot:
        command.append("--skip-screenshot")
    if keep_run_artifacts:
        command.append("--keep-run-artifacts")

    run_command(
        "keyword_trends.generate",
        command,
        cwd=MODULE_DIR,
        # 趋势图片直接写入模块目录，目录变化代表某个关键词已经有进展。
        watch_paths=(MODULE_DIR,),
    )

    return {
        "keyword_trends_report": output_path,
        "current_workbook": output_path,
    }

