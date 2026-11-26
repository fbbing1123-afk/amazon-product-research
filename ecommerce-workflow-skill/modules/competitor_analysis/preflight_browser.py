from __future__ import annotations

import argparse
import json
import socket
import subprocess
import sys
from pathlib import Path
from urllib.parse import urlparse


MODULE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = MODULE_DIR.parents[1]
MODULES_DIR = PROJECT_DIR / "modules"
if str(MODULES_DIR) not in sys.path:
    sys.path.insert(0, str(MODULES_DIR))

from keyword_backfill.seller_sprite_keyword_tool import (  # noqa: E402
    ensure_sellersprite_logged_in,
)
from review_research.scrape_amazon_reviews import preflight_amazon  # noqa: E402
from shared.browser_runtime import get_debugger_address
from shared.marketplaces import get_marketplace  # noqa: E402


SELLERSPRITE_HOME = "https://www.sellersprite.com/v2/welcome"


def debugger_host_port(value: str) -> tuple[str, int]:
    text = value if "://" in value else f"http://{value}"
    parsed = urlparse(text)
    if not parsed.hostname or not parsed.port:
        raise ValueError(f"Invalid debugger address: {value}")
    return parsed.hostname, parsed.port


def debugger_ready(value: str, timeout_seconds: float = 1.0) -> bool:
    host, port = debugger_host_port(value)
    try:
        with socket.create_connection((host, port), timeout=timeout_seconds):
            return True
    except OSError:
        return False


def start_project_browser(*, debugger_address: str, site: str) -> str:
    _, port = debugger_host_port(debugger_address)
    script = PROJECT_DIR / "scripts" / "open_sellersprite_login_browser.ps1"
    subprocess.run(
        [
            "powershell",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(script),
            "-Port",
            str(port),
            "-Site",
            site,
            "-SellerSpriteUrl",
            SELLERSPRITE_HOME,
        ],
        check=True,
        cwd=PROJECT_DIR,
    )
    # 启动器可能跳过被占用的端口，必须读取成功后保存的实际地址。
    actual_address = get_debugger_address()
    if not debugger_ready(actual_address):
        raise RuntimeError(f"项目浏览器连接未就绪：{actual_address}")
    return actual_address


def find_page(context, predicate):
    for page in reversed(context.pages):
        try:
            if predicate(page):
                return page
        except Exception:
            continue
    return context.new_page()


def run_preflight(*, debugger_address: str, site: str, postal_code: str | None) -> dict[str, object]:
    from playwright.sync_api import sync_playwright

    marketplace = get_marketplace(site)
    was_ready = debugger_ready(debugger_address)
    previous_address = debugger_address
    # 即使端口有响应，也由启动器校验配置归属，避免误连其它项目浏览器。
    debugger_address = start_project_browser(debugger_address=debugger_address, site=marketplace.code)
    started_browser = not was_ready or debugger_address != previous_address

    endpoint = debugger_address if "://" in debugger_address else f"http://{debugger_address}"
    with sync_playwright() as playwright:
        browser = playwright.chromium.connect_over_cdp(endpoint)
        context = browser.contexts[0] if browser.contexts else browser.new_context()

        sellersprite_page = find_page(
            context,
            lambda item: "sellersprite.com" in (urlparse(item.url).hostname or ""),
        )
        if "sellersprite.com" not in (urlparse(sellersprite_page.url).hostname or ""):
            sellersprite_page.goto(SELLERSPRITE_HOME, wait_until="domcontentloaded", timeout=60_000)
        # 集合链接已在本地生成；这里只检查其它卖家精灵功能所需的登录态。
        ensure_sellersprite_logged_in(sellersprite_page, SELLERSPRITE_HOME)

        amazon_page = find_page(
            context,
            lambda item: (urlparse(item.url).hostname or "").lower()
            in {marketplace.amazon_domain, f"www.{marketplace.amazon_domain}"},
        )
        preflight_amazon(
            amazon_page,
            site=marketplace.code,
            postal_code=postal_code or marketplace.postal_code,
        )

        return {
            "status": "ready",
            "browser_started": started_browser,
            "debugger_address": debugger_address,
            "site": marketplace.code,
            "postal_code": postal_code or marketplace.postal_code,
            "sellersprite_url": sellersprite_page.url,
            "amazon_url": amazon_page.url,
        }


def main() -> int:
    parser = argparse.ArgumentParser(description="Prepare the project browser for competitor analysis.")
    parser.add_argument("--debugger-address", default=get_debugger_address())
    parser.add_argument("--site", default="US")
    parser.add_argument("--amazon-postal-code")
    args = parser.parse_args()
    try:
        result = run_preflight(
            debugger_address=args.debugger_address,
            site=args.site,
            postal_code=args.amazon_postal_code,
        )
    except Exception as exc:
        print(json.dumps({"status": "needs_user_action", "error": str(exc)}, ensure_ascii=False))
        return 2
    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
