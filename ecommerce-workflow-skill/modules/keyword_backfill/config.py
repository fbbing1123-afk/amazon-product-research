from __future__ import annotations

from pathlib import Path


from shared.browser_runtime import get_debugger_address

MODULE_DIR = Path(__file__).resolve().parent
SCRIPT_PATH = MODULE_DIR / "seller_sprite_keyword_tool.py"
DEFAULT_EXPAND_KEYWORDS_URL = "https://www.sellersprite.com/v3/traffic/extend"
DEFAULT_MY_EXPORTS_URL = "https://www.sellersprite.com/v2/export-log"
DEFAULT_DEBUGGER_ADDRESS = get_debugger_address()
DEFAULT_SITE = "US"
