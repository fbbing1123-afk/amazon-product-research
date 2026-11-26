from __future__ import annotations

import json
from pathlib import Path

from shared.browser_runtime import get_debugger_address
from shared.subprocess_runner import run_command

from .config import (
    DEFAULT_DEBUGGER_ADDRESS,
    DEFAULT_EXPAND_KEYWORDS_URL,
    DEFAULT_MY_EXPORTS_URL,
    DEFAULT_SITE,
    MODULE_DIR,
    SCRIPT_PATH,
)


def run_keyword_backfill(
    *,
    source: Path,
    name: str,
    output_dir: Path,
    python_exe: str,
    export_file: Path | None = None,
    skip_browser: bool = False,
    excel_writer: str = "openpyxl",
    debugger_address: str | None = None,
    site: str = DEFAULT_SITE,
    brand_words: list[str] | None = None,
) -> dict[str, Path]:
    print("[keyword_backfill] start")
    work_dir = output_dir / "work" / name / "keyword_backfill"
    download_dir = work_dir / "downloads"
    work_dir.mkdir(parents=True, exist_ok=True)
    download_dir.mkdir(parents=True, exist_ok=True)

    output_path = output_dir / f"{name}_with_keyword_backfill.xlsx"
    config_path = work_dir / "seller_sprite_config.json"
    config = {
        "template_path": str(source),
        "output_path": str(output_path),
        "download_dir": str(download_dir),
        "storage_state": str(MODULE_DIR.parent.parent / ".workflow-state" / "sellersprite" / ".browser-state.json"),
        "expand_keywords_url": DEFAULT_EXPAND_KEYWORDS_URL,
        "my_exports_url": DEFAULT_MY_EXPORTS_URL,
        "site": site,
        "date_range_text": "",
        "excel_writer": excel_writer,
        "browser": {
            "type": "edge",
            "debugger_address": debugger_address or get_debugger_address(),
        },
    }
    config_path.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")

    # 关闭子进程输出缓冲，确保卖家精灵导出状态及时进入工作流监控。
    command: list[str | Path] = [python_exe, "-u", SCRIPT_PATH, "--config", config_path]
    if skip_browser or export_file:
        command.append("--skip-browser")
        if not export_file:
            raise ValueError("--keyword-backfill-skip-browser requires --keyword-backfill-export-file.")
    if export_file:
        command.extend(["--export-file", export_file])
    for word in brand_words or []:
        command.extend(["--brand-word", word])

    run_command("keyword_backfill.generate", command, cwd=MODULE_DIR)

    return {
        "keyword_backfill_report": output_path,
        "current_workbook": output_path,
    }
