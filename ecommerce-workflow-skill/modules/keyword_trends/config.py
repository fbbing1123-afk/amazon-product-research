from __future__ import annotations

from pathlib import Path


MODULE_DIR = Path(__file__).resolve().parent
SKILL_DIR = MODULE_DIR.parents[1]
SCRIPT_PATH = MODULE_DIR / "excel-sellersprite-trends.py"
BROWSER_STATE_DIR = SKILL_DIR / ".workflow-state" / "sellersprite"
DEFAULT_PREFIX = "excel-trends"
