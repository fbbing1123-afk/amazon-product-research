from __future__ import annotations

from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parent
SCRIPT_PATH = PROJECT_DIR / "excel_market_report.py"
SLICER_SCRIPT_PATH = PROJECT_DIR / "add_source_slicers.ps1"
STATIC_CHART_SCRIPT_PATH = PROJECT_DIR.parents[1] / "scripts" / "replace_market_charts_with_images.py"
