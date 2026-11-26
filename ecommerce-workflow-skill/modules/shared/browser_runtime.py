"""读取启动器确认过的项目浏览器端口，统一各阶段的默认连接地址。"""
from __future__ import annotations

import json
from pathlib import Path
from urllib.parse import urlparse


CONFIG_PATH = Path(__file__).resolve().parents[2] / '.workflow-state/sellersprite/browser-config.json'


def get_browser_port() -> int:
    if not CONFIG_PATH.exists():
        return 9223
    config = json.loads(CONFIG_PATH.read_text(encoding='utf-8-sig'))
    value = config.get('port')
    if value is None:
        address = config.get('cdp_url') or config.get('debugger_address')
        value = urlparse(address if '://' in address else 'http://' + address).port if address else 9223
    if isinstance(value, bool) or not str(value).isdigit() or not 1 <= int(value) <= 65535:
        raise ValueError(f'项目浏览器端口记录无效，请重新运行浏览器启动器：{CONFIG_PATH}')
    return int(value)


def get_debugger_address() -> str:
    return f'127.0.0.1:{get_browser_port()}'


def get_cdp_url() -> str:
    return f'http://{get_debugger_address()}'
