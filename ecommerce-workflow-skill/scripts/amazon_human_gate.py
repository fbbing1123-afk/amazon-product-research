from __future__ import annotations

import argparse
import json
import time
from urllib.parse import urlparse


CONTINUE_TEXT = "Continue shopping"
HUMAN_CHECK_ID = "#continueCheck"


def challenge_page(browser):
    """优先选择当前 Amazon 真人确认页。"""
    pages = []
    for context in browser.contexts:
        pages.extend(context.pages)
    for page in reversed(pages):
        url = page.url.lower()
        if "/checkpoint/continue" in url:
            return page
    return None


def main() -> int:
    parser = argparse.ArgumentParser(
        description="处理 Amazon 真人确认页：人工勾选后自动点击 Continue shopping。"
    )
    parser.add_argument("--debugger-address", default="127.0.0.1:9223")
    parser.add_argument("--timeout-seconds", type=int, default=300)
    parser.add_argument("--poll-seconds", type=float, default=1.0)
    args = parser.parse_args()

    from playwright.sync_api import sync_playwright

    endpoint = (
        args.debugger_address
        if "://" in args.debugger_address
        else f"http://{args.debugger_address}"
    )
    deadline = time.monotonic() + args.timeout_seconds
    with sync_playwright() as playwright:
        browser = playwright.chromium.connect_over_cdp(endpoint, timeout=20_000)
        announced_url = ""
        missing_challenge_checks = 0
        while time.monotonic() < deadline:
            page = challenge_page(browser)
            if page is None:
                missing_challenge_checks += 1
                if missing_challenge_checks >= 5:
                    print(
                        json.dumps(
                            {
                                "status": "challenge_cleared",
                                "debugger_address": args.debugger_address,
                            },
                            ensure_ascii=False,
                        ),
                        flush=True,
                    )
                    return 0
                time.sleep(max(0.2, args.poll_seconds))
                continue
            missing_challenge_checks = 0

            try:
                # 将项目浏览器置前，但“I am human”必须由用户本人勾选。
                page.bring_to_front()
                checkbox = page.locator(HUMAN_CHECK_ID).first
                if not checkbox.is_visible(timeout=2_000):
                    time.sleep(max(0.2, args.poll_seconds))
                    continue
                if page.url != announced_url:
                    announced_url = page.url
                    print(
                        json.dumps(
                            {
                                "status": "waiting_for_user_check",
                                "message": "请在已置前的项目 Chrome 中手动勾选“I am human”；勾选后脚本会自动点击 Continue shopping。",
                                "url": page.url,
                            },
                            ensure_ascii=False,
                        ),
                        flush=True,
                    )
                if checkbox.is_checked():
                    continue_button = page.get_by_role("button", name=CONTINUE_TEXT).first
                    continue_button.wait_for(state="visible", timeout=10_000)
                    for _ in range(10):
                        if continue_button.is_enabled():
                            continue_button.click(timeout=5_000)
                            time.sleep(1.5)
                            print(
                                json.dumps(
                                    {
                                        "status": "continued",
                                        "url": page.url,
                                        "host": urlparse(page.url).netloc,
                                    },
                                    ensure_ascii=False,
                                ),
                                flush=True,
                            )
                            return 0
                        time.sleep(0.5)
                    raise RuntimeError("已勾选“I am human”，但 Continue shopping 按钮仍不可用。")
            except Exception:
                # Amazon 可能刷新挑战页令牌；重新查找最新挑战页即可继续等待。
                time.sleep(max(0.2, args.poll_seconds))
                continue
            time.sleep(max(0.2, args.poll_seconds))

    print(
        json.dumps(
            {
                "status": "needs_user_action",
                "message": "等待人工勾选超时，请在项目 Chrome 中勾选“I am human”后重新运行本脚本。",
            },
            ensure_ascii=False,
        ),
        flush=True,
    )
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
