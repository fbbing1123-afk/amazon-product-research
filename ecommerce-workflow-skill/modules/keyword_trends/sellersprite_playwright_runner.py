#!/usr/bin/env python3
"""Browser automation for Sellersprite using the Python Playwright API."""

from __future__ import annotations

import base64
import json
import random
import re
import sys
import time
import urllib.parse
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Sequence

WORKFLOW_MODULES_DIR = Path(__file__).resolve().parents[1]
if str(WORKFLOW_MODULES_DIR) not in sys.path:
    sys.path.insert(0, str(WORKFLOW_MODULES_DIR))

from shared.marketplaces import (
    amazon_domains,
    ensure_amazon_postal_code,
    ensure_sellersprite_site,
    get_marketplace,
    url_belongs_to_amazon,
)


SELLERSPRITE_URL = "https://www.sellersprite.com/v3/keyword-miner/"
MANUAL_STOP_REASONS = {"login", "captcha", "vip", "risk"}
DEFAULT_AMAZON_DELIVERY_ZIP = "10001"
TREND_EXPORT_CSS_WIDTH = 1536
TREND_EXPORT_CSS_HEIGHT = 420
TREND_EXPORT_PIXEL_RATIO = 2
TREND_EXPORT_PNG_WIDTH = TREND_EXPORT_CSS_WIDTH * TREND_EXPORT_PIXEL_RATIO
TREND_EXPORT_PNG_HEIGHT = TREND_EXPORT_CSS_HEIGHT * TREND_EXPORT_PIXEL_RATIO


@dataclass(frozen=True)
class BrowserRunResult:
    returncode: int
    stdout: str = ""
    stderr: str = ""
    command: list[str] | None = None


class ManualActionRequired(RuntimeError):
    def __init__(self, message: str, reason: str = "manual", stage: str = "browser") -> None:
        super().__init__(message)
        self.reason = reason
        self.stage = stage


def require_sync_playwright():
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as exc:
        raise RuntimeError(
            "未安装 Python Playwright。请先在运行本项目的 Python 环境里执行：\n"
            "  python -m pip install playwright\n"
            "  python -m playwright install chromium\n"
            "然后重新运行脚本。"
        ) from exc
    return sync_playwright


def safe_name(keyword: str) -> str:
    value = keyword.replace("\r", "").replace("\n", "").strip().replace(" ", "-")
    value = re.sub(r"[\\/:*?\"<>|]+", "-", value)
    value = re.sub(r"[^0-9A-Za-z._\-\u4e00-\u9fff]+", "-", value)
    value = re.sub(r"-+", "-", value).strip("-._")
    return value


def chart_path_for_keyword(project_dir: Path, prefix: str, keyword: str, index: int) -> Path:
    name = safe_name(keyword) or f"keyword-{index}"
    return project_dir / f"{prefix}-{name}-chart.png"


def build_chart_jobs(project_dir: Path, keywords: Sequence[str], prefix: str) -> list[dict[str, Any]]:
    jobs: list[dict[str, Any]] = []
    for index, keyword in enumerate(keywords, start=1):
        keyword = str(keyword).strip()
        if not keyword:
            continue
        output = chart_path_for_keyword(project_dir, prefix, keyword, index)
        jobs.append({"index": index, "keyword": keyword, "output": str(output)})
    return jobs


def connect_cdp(playwright, port: int, storage_state: Path | None = None):
    browser = playwright.chromium.connect_over_cdp(f"http://127.0.0.1:{port}")
    if browser.contexts:
        context = browser.contexts[0]
    else:
        kwargs = {"storage_state": str(storage_state)} if storage_state and storage_state.exists() else {}
        context = browser.new_context(**kwargs)
    page = next(
        (
            candidate
            for candidate in context.pages
            if "sellersprite.com" in (candidate.url or "") and sellersprite_ready(candidate)
        ),
        context.pages[0] if context.pages else context.new_page(),
    )
    return browser, context, page


def disconnect_browser(browser) -> None:
    """远程 CDP 连接由 sync_playwright 上下文释放，不能关闭项目浏览器。"""
    _ = browser


def sleep(page, ms: int) -> None:
    base = max(0, int(ms or 0))
    jitter = min(1500, max(250, round(base * 0.25)))
    page.wait_for_timeout(base + 200 + random.randrange(jitter))


def evaluate_state(page) -> dict[str, Any]:
    return page.evaluate(
        """() => {
          const isVisible = (el) => {
            if (!el) return false;
            const style = window.getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0;
          };
          const text = document.body ? document.body.innerText : '';
          const compactText = text.replace(/\\s+/g, ' ');
          const url = location.href;
          const lowerUrl = url.toLowerCase();
          const visibleDialogs = [...document.querySelectorAll('.el-dialog__wrapper, .el-message-box__wrapper, .modal, [role="dialog"]')].filter(isVisible);
          const dialogText = visibleDialogs.map(el => (el.innerText || el.textContent || '').trim()).join(' ').replace(/\\s+/g, ' ');
          const pageAndDialogText = `${compactText} ${dialogText}`;
          const hasKeywordMiner = lowerUrl.includes('/keyword-miner')
            && !!document.querySelector('input[placeholder*="输入关键词"], tbody tr, #search-anlysis-drawer-echarts-trends, #search-trends');
          const hasLoginUrl = /login|signin|sign-in/.test(lowerUrl);
          const hasVisiblePasswordInput = [...document.querySelectorAll('input[type="password"], input[placeholder*="密码"]')].some(isVisible);
          const hasVisibleLoginButton = [...document.querySelectorAll('button, .el-button, a')]
            .some(el => isVisible(el) && /^(登录|登陆|Sign in|Login)$/.test((el.innerText || el.textContent || '').trim()));
          if (hasLoginUrl || hasVisiblePasswordInput || hasVisibleLoginButton || /扫码|二维码|微信扫码|手机扫码|scan\\s*qr|qr\\s*code/i.test(pageAndDialogText)) {
            return { status: 'manual', reason: 'login', message: '检测到登录页面、扫码或登录控件', url };
          }
          if (/验证码|安全验证|滑块|机器人检测|机器人校验|captcha|robot\\s*check|geetest|arkose/i.test(pageAndDialogText)
              || [...document.querySelectorAll('iframe, [class], [id]')].some(el => /captcha|robot[-_]?check|geetest|arkose|challenge/i.test(`${el.className || ''} ${el.id || ''} ${el.src || ''}`))) {
            return { status: 'manual', reason: 'captcha', message: '检测到验证码、安全验证或机器人检测', url };
          }
          if (/VIP|会员|权限不足|无权限|试用到期/.test(pageAndDialogText)) {
            return { status: 'manual', reason: 'vip', message: '检测到 VIP、会员或权限限制提示', url };
          }
          if (/账号异常|访问过于频繁|请稍后再试|风控/.test(pageAndDialogText)) {
            return { status: 'manual', reason: 'risk', message: '检测到账号异常或访问频率限制', url };
          }
          if (/暂无数据|无数据|没有找到|未找到数据/.test(compactText) && !visibleDialogs.length) {
            return { status: 'manual', reason: 'no-data', message: '检测到当前关键词可能没有数据', url };
          }
          const visibleLoading = [...document.querySelectorAll('.el-loading-mask')].some(isVisible);
          if (visibleLoading) return { status: 'auto', reason: 'loading', message: '检测到页面仍在加载', url };
          const visibleModal = [...document.querySelectorAll('.el-dialog__wrapper, .el-message-box__wrapper, .el-notification, .el-message, .modal, [role="dialog"]')].some(isVisible);
          if (visibleModal) return { status: 'auto', reason: 'modal', message: '检测到普通弹窗或提示遮挡', url };
          if (hasKeywordMiner) return { status: 'ok', reason: 'ready', message: '页面状态正常', url };
          return { status: 'unknown', reason: 'unknown', message: '无法识别当前页面状态', url };
        }"""
    )


def guard_manual_state(page, stage: str) -> dict[str, Any]:
    state = evaluate_state(page)
    if state.get("status") == "manual" and state.get("reason") == "login":
        if recover_sellersprite_login_shell(page):
            state = evaluate_state(page)
    if state.get("status") == "manual" and state.get("reason") in MANUAL_STOP_REASONS:
        raise ManualActionRequired(state.get("message") or state.get("reason") or "需要人工处理", state.get("reason", "manual"), stage)
    return state


def wait_loading_gone(page, timeout: int = 30000) -> None:
    page.wait_for_function(
        """() => {
          const isVisible = (el) => {
            if (!el) return false;
            const style = window.getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0;
          };
          return ![...document.querySelectorAll('.el-loading-mask')].some(isVisible);
        }""",
        timeout=timeout,
    )


def safe_auto_handle_page_issue(page) -> dict[str, str]:
    try:
        state = evaluate_state(page)
    except Exception:
        return {"action": "none", "reason": "unknown"}
    if state.get("status") == "manual" and state.get("reason") in MANUAL_STOP_REASONS:
        return {"action": "manual", "reason": state.get("reason", "manual")}
    if state.get("status") == "auto" and state.get("reason") == "modal":
        page.reload(wait_until="domcontentloaded", timeout=45000)
        try:
            wait_loading_gone(page, 15000)
        except Exception:
            pass
        sleep(page, 1200)
        return {"action": "reload", "reason": "modal"}
    if state.get("status") == "auto" and state.get("reason") == "loading":
        try:
            wait_loading_gone(page, 15000)
        except Exception:
            pass
        sleep(page, 600)
        return {"action": "wait", "reason": "loading"}
    return {"action": "none", "reason": str(state.get("reason") or "")}


def cleanup_browser_tabs(project_dir: Path, port: int = 9222, keep_sellersprite_keyword_pages: int = 1, close_amazon_tabs: bool = True) -> int:
    sync_playwright = require_sync_playwright()
    closed_count = 0
    with sync_playwright() as p:
        browser, context, _page = connect_cdp(p, port)
        try:
            sellersprite_kept = 0
            for tab in list(context.pages):
                url = (tab.url or "").strip().lower()
                should_close = False
                if not url or url == "about:blank" or url.startswith("chrome://new-tab") or url.startswith("chrome://welcome") or "google.com/chrome" in url:
                    should_close = True
                elif "sellersprite.com/v3/keyword-miner" in url:
                    sellersprite_kept += 1
                    should_close = sellersprite_kept > keep_sellersprite_keyword_pages
                elif close_amazon_tabs and url_belongs_to_amazon(url, amazon_domains()):
                    should_close = True
                if should_close and len(context.pages) - closed_count > 1:
                    try:
                        tab.close(run_before_unload=False)
                        closed_count += 1
                    except Exception:
                        pass
            return closed_count
        finally:
            disconnect_browser(browser)


def save_storage_state(context, state_file: Path) -> None:
    state_file.parent.mkdir(parents=True, exist_ok=True)
    context.storage_state(path=str(state_file))


def restore_storage_state(context, page, state_file: Path) -> None:
    if not state_file.exists():
        return
    try:
        state = json.loads(state_file.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return

    cookies = state.get("cookies", [])
    if isinstance(cookies, list) and cookies:
        try:
            context.add_cookies(cookies)
        except Exception:
            pass

    origins = state.get("origins", [])
    if not isinstance(origins, list):
        return
    for origin in origins:
        if not isinstance(origin, dict):
            continue
        origin_url = str(origin.get("origin") or "")
        local_storage = origin.get("localStorage", [])
        if not origin_url or not isinstance(local_storage, list):
            continue
        if "sellersprite.com" not in origin_url:
            continue
        try:
            page.goto(origin_url, wait_until="domcontentloaded", timeout=30000)
            page.evaluate(
                """(items) => {
                  for (const item of items) {
                    if (item && typeof item.name === 'string') {
                      localStorage.setItem(item.name, String(item.value ?? ''));
                    }
                  }
                }""",
                local_storage,
            )
        except Exception:
            pass


def sellersprite_ready(page) -> bool:
    try:
        return bool(page.evaluate(
            """() => {
              const isVisible = (el) => {
                if (!el) return false;
                const style = window.getComputedStyle(el);
                const rect = el.getBoundingClientRect();
                return style.display !== 'none'
                  && style.visibility !== 'hidden'
                  && style.opacity !== '0'
                  && rect.width > 0
                  && rect.height > 0;
              };
              const compact = (el) => (el.innerText || el.textContent || '').replace(/\\s+/g, ' ').trim();
              const visitorRe = /\\u672a\\u767b\\u5f55|\\u6e38\\u5ba2|\\u7acb\\u5373\\u767b\\u5f55|\\u767b\\u5f55\\s*\\/\\s*\\u6ce8\\u518c/;
              const accountRe = /XWTKFB\\d+|\\u9ad8\\u7ea7\\u4f1a\\u5458|VIP|\\u5269\\u4f59\\d+\\u5929|\\u5957\\u9910\\u8fc7\\u671f[:\\uff1a]?\\s*\\d{4}-\\d{2}-\\d{2}/i;
              const accountNodes = [...document.querySelectorAll(
                '.user-info, .u-header__user, .u-header__user-avatar_nickname, .user-name, .right'
              )].filter(isVisible);
              const hasVisibleAccount = accountNodes.some(el => {
                const text = compact(el);
                if (!text || visitorRe.test(text)) return false;
                return accountRe.test(text)
                  || el.classList.contains('u-header__user-avatar_nickname')
                  || (el.classList.contains('user-name') && text.length >= 3);
              });
              if (hasVisibleAccount) return true;

              const hasVisibleVisitor = accountNodes.some(el => visitorRe.test(compact(el)));
              if (hasVisibleVisitor) return false;

              const visibleText = [...document.querySelectorAll('h1, h2, h3, input, textarea, button, a')]
                .filter(isVisible)
                .map(compact)
                .join(' ');
              const lowerUrl = location.href.toLowerCase();
              const hasKeywordInput = [...document.querySelectorAll(
                'input[placeholder*="keyword" i], input[placeholder*="\\u5173\\u952e\\u8bcd"], input[placeholder*="\\u8f93\\u5165"]'
              )].some(isVisible);
              const onUsefulPage = lowerUrl.includes('/v2/welcome')
                || (lowerUrl.includes('/v2/export-log') && /\\u6211\\u7684\\u5bfc\\u51fa/.test(visibleText))
                || (lowerUrl.includes('/v3/keyword-miner') && hasKeywordInput)
                || /\\u5356\\u5bb6\\u7cbe\\u7075\\u540e\\u53f0|\\u5173\\u952e\\u8bcd\\u6316\\u6398|Keyword Miner/i.test(visibleText);
              return onUsefulPage;
            }"""
        ))
    except Exception:
        return False


def sellersprite_has_real_login_challenge(page) -> bool:
    try:
        return bool(page.evaluate(
            """() => {
              const isVisible = (el) => {
                if (!el) return false;
                const style = window.getComputedStyle(el);
                const rect = el.getBoundingClientRect();
                return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0;
              };
              const text = document.body ? document.body.innerText || '' : '';
              const hasCredentialInput = [...document.querySelectorAll(
                'input[type="password"], input[placeholder*="password" i], input[placeholder*="\\u5bc6\\u7801"]'
              )].some(isVisible);
              return hasCredentialInput || /captcha|robot\\s*check|geetest|arkose|qr\\s*code|scan\\s*qr|\\u626b\\u7801|\\u4e8c\\u7ef4\\u7801/i.test(text);
            }"""
        ))
    except Exception:
        return False


def recover_sellersprite_login_shell(page) -> bool:
    if sellersprite_ready(page):
        return True
    if sellersprite_has_real_login_challenge(page):
        return False

    for probe_url in ("https://www.sellersprite.com/v2/welcome", SELLERSPRITE_URL):
        try:
            page.goto(probe_url, wait_until="domcontentloaded", timeout=45000)
            sleep(page, 1500)
            try:
                wait_loading_gone(page, 8000)
            except Exception:
                pass
            if sellersprite_ready(page):
                if probe_url != SELLERSPRITE_URL:
                    page.goto(SELLERSPRITE_URL, wait_until="domcontentloaded", timeout=45000)
                    sleep(page, 1200)
                if sellersprite_ready(page):
                    return True
        except Exception:
            pass

    if sellersprite_has_real_login_challenge(page):
        return False
    try:
        visitor_re = re.compile(r"未登录|游客")
        for selector in (".user-info", ".right", ".user-name"):
            trigger = page.locator(selector).filter(has_text=visitor_re)
            if trigger.count() and trigger.first.is_visible(timeout=1_000):
                trigger.first.click(timeout=10_000)
                sleep(page, 900)
                break

        clicked = False
        login_re = re.compile(r"^(立即登录|登录\s*/\s*注册|登录|登陆|Sign in|Login)$", re.I)
        for selector in ("a.text-link", "a", "button", ".el-button"):
            candidates = page.locator(selector).filter(has_text=login_re)
            for index in range(min(candidates.count(), 8)):
                candidate = candidates.nth(index)
                if candidate.is_visible(timeout=500):
                    candidate.click(timeout=10_000)
                    clicked = True
                    break
            if clicked:
                break
        if clicked:
            deadline = time.monotonic() + 15
            while time.monotonic() < deadline and not sellersprite_ready(page):
                sleep(page, 500)
            if sellersprite_ready(page):
                if urllib.parse.urlparse(page.url or '').path.rstrip('/') == urllib.parse.urlparse(SELLERSPRITE_URL).path.rstrip('/'):
                    return True
                page.goto(SELLERSPRITE_URL, wait_until="domcontentloaded", timeout=45000)
                deadline = time.monotonic() + 10
                while time.monotonic() < deadline and not sellersprite_ready(page):
                    sleep(page, 500)
                return sellersprite_ready(page)
            page.goto(SELLERSPRITE_URL, wait_until="domcontentloaded", timeout=45000)
            deadline = time.monotonic() + 10
            while time.monotonic() < deadline and not sellersprite_ready(page):
                sleep(page, 500)
            return sellersprite_ready(page)
    except Exception:
        pass
    return False


def is_logged_in(page) -> bool:
    if sellersprite_ready(page):
        return True
    state = evaluate_state(page)
    if state.get("reason") == "login":
        return recover_sellersprite_login_shell(page)
    text = page.evaluate("() => document.body ? document.body.innerText.slice(0, 1000) : ''")
    return not bool(re.search(r"未登录|游客|立即登录|登录|Login|Sign in", str(text), re.I))


def navigate_sellersprite_tool(page, target_url: str, link_text: str) -> None:
    """Use SellerSprite's visible navigation before falling back to a direct URL."""
    target_path = urllib.parse.urlparse(target_url).path.rstrip("/")
    current_path = urllib.parse.urlparse(page.url or "").path.rstrip("/")
    if current_path == target_path:
        return

    try:
        link = page.get_by_role("link", name=link_text, exact=True)
        if link.count() and link.first.is_visible(timeout=2_000):
            link.first.click(timeout=10_000)
            sleep(page, 1_200)
            if urllib.parse.urlparse(page.url or "").path.rstrip("/") == target_path:
                return
    except Exception:
        pass

    page.goto(target_url, wait_until="domcontentloaded", timeout=45_000)


def ensure_sellersprite_login(project_dir: Path, port: int = 9222, state_file_name: str = ".browser-state.json") -> bool:
    sync_playwright = require_sync_playwright()
    state_file = project_dir / state_file_name
    with sync_playwright() as p:
        browser, context, page = connect_cdp(p, port, state_file)
        try:
            restore_storage_state(context, page, state_file)
            page.goto(SELLERSPRITE_URL, wait_until="domcontentloaded", timeout=45000)
            sleep(page, 1500)
            try:
                wait_loading_gone(page, 10000)
            except Exception:
                pass
            if is_logged_in(page):
                print("登录状态有效，无需重新登录。")
                save_storage_state(context, state_file)
                return True

            print("=" * 50)
            print("请在 Chrome 浏览器中手动登录卖家精灵。")
            print("登录完成后，回到终端按回车继续。")
            print("=" * 50)
            try:
                input()
            except EOFError:
                print("终端不可交互，请在 Chrome 中完成登录后重新运行。")
                return False

            page.goto(SELLERSPRITE_URL, wait_until="domcontentloaded", timeout=45000)
            sleep(page, 1500)
            if not is_logged_in(page):
                print("仍未检测到有效登录状态。")
                return False
            save_storage_state(context, state_file)
            print(f"登录状态已保存到 {state_file}")
            return True
        finally:
            disconnect_browser(browser)


def close_drawer_if_open(page) -> str:
    closed = page.evaluate(
        """() => {
          const isVisible = (el) => {
            if (!el) return false;
            const style = window.getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0;
          };
          const drawer = document.querySelector('#search-anlysis-drawer-echarts-trends') || document.querySelector('.el-drawer__wrapper');
          if (!drawer || !isVisible(drawer)) return 'already-closed';
          const closeBtn = drawer.querySelector('.fa.fa-times.close')
            || drawer.querySelector('.close-wrap .close')
            || drawer.querySelector('.el-drawer__close-btn')
            || drawer.querySelector('.el-drawer__header button')
            || drawer.querySelector('[aria-label="Close"]')
            || drawer.querySelector('[aria-label="close"]')
            || drawer.querySelector('.el-icon-close')
            || drawer.querySelector('[class*="close"]');
          if (!closeBtn) return 'no-close-button';
          closeBtn.click();
          return 'clicked';
        }"""
    )
    if closed == "clicked":
        try:
            page.wait_for_function(
                """() => {
                  const el = document.querySelector('#search-trends');
                  if (!el) return true;
                  const style = window.getComputedStyle(el);
                  const rect = el.getBoundingClientRect();
                  return style.display === 'none' || style.visibility === 'hidden' || rect.width === 0 || rect.height === 0;
                }""",
                timeout=10000,
            )
        except Exception:
            pass
        sleep(page, 400)
    return str(closed)


def first_visible_row_text(page) -> str:
    return str(page.evaluate(
        """() => {
          const isVisible = (el) => {
            if (!el) return false;
            const style = window.getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0;
          };
          const row = [...document.querySelectorAll('tbody tr')].find(isVisible);
          return row ? (row.textContent || '').replace(/\\s+/g, ' ').trim() : '';
        }"""
    ) or "")


def wait_first_row_ready(page, keyword: str, previous_row_text: str = "", timeout_ms: int = 45000) -> dict[str, Any]:
    started = time.monotonic()
    last_state: dict[str, Any] = {}
    while (time.monotonic() - started) * 1000 < timeout_ms:
        try:
            wait_loading_gone(page, 8000)
        except Exception:
            pass
        last_state = page.evaluate(
            """({ keyword, previousRowText, startedAt }) => {
              const isVisible = (el) => {
                if (!el) return false;
                const style = window.getComputedStyle(el);
                const rect = el.getBoundingClientRect();
                return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0;
              };
              const rows = [...document.querySelectorAll('tbody tr')].filter(isVisible);
              const row = rows[0];
              if (!row) return { ready: false, reason: 'no-row' };
              const rowText = (row.textContent || '').replace(/\\s+/g, ' ').trim();
              const hasKeyword = !keyword || rowText.toLowerCase().includes(String(keyword).toLowerCase());
              const rowChanged = !!previousRowText && rowText !== previousRowText;
              const numericCells = [...row.querySelectorAll('td span, td a, td div')]
                .filter(el => isVisible(el))
                .map(el => (el.textContent || '').trim())
                .filter(text => /^\\d+(?:,\\d{3})*$/.test(text));
              const waitedLongEnoughForRefresh = Date.now() - startedAt > 3000;
              const hasReadableResult = rowText.length > 0 || numericCells.length > 0;
              const ready = hasReadableResult && (hasKeyword || rowChanged || waitedLongEnoughForRefresh);
              return { ready, reason: numericCells.length ? 'ready-with-number' : 'ready-without-monthly-number', rowText, rowChanged, hasKeyword, waitedLongEnoughForRefresh, numericCells };
            }""",
            {"keyword": keyword, "previousRowText": previous_row_text, "startedAt": int(time.time() * 1000)},
        )
        if last_state.get("ready"):
            return last_state
        sleep(page, 1000)
    raise RuntimeError(f"等待搜索结果首行月搜索量超时: {last_state.get('reason', 'unknown')}")


def search_keyword(page, keyword: str) -> None:
    guard_manual_state(page, "search-before")
    previous = first_visible_row_text(page)
    input_locator = page.locator('input[placeholder*="输入关键词"]').first
    input_locator.wait_for(state="visible", timeout=10000)
    input_locator.click(timeout=10000)
    input_locator.fill("", timeout=10000)
    input_locator.fill(keyword, timeout=10000)
    page.evaluate(
        """(keyword) => {
          const input = document.querySelector('input[placeholder*="输入关键词"]');
          if (!input) return;
          const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set;
          if (setter) setter.call(input, keyword);
          else input.value = keyword;
          input.dispatchEvent(new Event('input', { bubbles: true }));
          input.dispatchEvent(new Event('change', { bubbles: true }));
        }""",
        keyword,
    )
    sleep(page, 1200)
    page.evaluate(
        """() => {
          const isVisible = (el) => {
            if (!el) return false;
            const style = window.getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0;
          };
          const input = document.querySelector('input[placeholder*="输入关键词"]');
          const inputRect = input && input.getBoundingClientRect();
          const buttons = [...document.querySelectorAll('button, .el-button')].filter(isVisible);
          const queryButtons = buttons.filter(b => (b.innerText || b.textContent || '').trim().includes('立即查询'));
          const btn = (inputRect && queryButtons.find(b => {
            const rect = b.getBoundingClientRect();
            return Math.abs(rect.y - inputRect.y) < 80 && rect.x > inputRect.x;
          })) || queryButtons[0] || buttons.find(b => /查询|搜索|Search/i.test((b.innerText || b.textContent || '').trim()));
          if (!btn) throw new Error('找不到“立即查询”按钮');
          btn.click();
        }"""
    )
    sleep(page, 1200)
    wait_first_row_ready(page, keyword, previous)


def open_first_row_monthly_drawer(page) -> None:
    page.evaluate(
        """() => {
          const isVisible = (el) => {
            if (!el) return false;
            const style = window.getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0;
          };
          const normalize = (text) => (text || '').replace(/\\s+/g, ' ').trim();
          const tables = [...document.querySelectorAll('table')].filter(isVisible);
          const table = tables.find(t => [...t.querySelectorAll('tbody tr')].some(isVisible)) || document;
          const row = [...table.querySelectorAll('tbody tr')].find(isVisible);
          if (!row) throw new Error('找不到表格第一行');
          const headerRows = [...table.querySelectorAll('thead tr')].filter(isVisible);
          const headerCells = headerRows.length ? [...headerRows[0].querySelectorAll('th,td')].map(cell => normalize(cell.innerText || cell.textContent)) : [];
          const monthlyIndex = headerCells.findIndex(text => /月搜索量|月搜索|搜索量/.test(text) && !/周|Google|谷歌|购买|购买率|排名|ABA|竞价|PPC/i.test(text));
          const trendIndex = headerCells.findIndex(text => /月搜索趋势|搜索趋势|趋势/.test(text) && !/Google|谷歌|购买|购买率|排名|ABA|竞价|PPC/i.test(text));
          const actionIndex = headerCells.findIndex(text => /操作|分析/.test(text));
          const cells = [...row.querySelectorAll('td,th')].filter(isVisible);
          let target = null;
          const monthlyValuePattern = /^\\d+(?:,\\d{3})*$/;
          if (monthlyIndex >= 0 && cells[monthlyIndex]) {
            target = [...cells[monthlyIndex].querySelectorAll('a, span, div, button')]
              .find(el => isVisible(el) && monthlyValuePattern.test(normalize(el.textContent))) || null;
          }
          if (!target && monthlyIndex < 0 && cells[6]) {
            const monthlyLikeControls = [cells[6], ...cells[6].querySelectorAll('a, span, div, button')].filter(isVisible);
            target = monthlyLikeControls.find(el => {
              const text = normalize(el.innerText || el.textContent);
              const styleText = `${el.getAttribute('style') || ''} ${window.getComputedStyle(el).cursor || ''} ${el.className || ''}`;
              return /border-bottom/i.test(styleText) && monthlyValuePattern.test(text);
            }) || monthlyLikeControls.find(el => {
              const text = normalize(el.innerText || el.textContent);
              const styleText = `${el.getAttribute('style') || ''} ${window.getComputedStyle(el).cursor || ''} ${el.className || ''}`;
              return /cursor\\s*:\\s*pointer|pointer|clickable/i.test(styleText)
                && (monthlyValuePattern.test(text) || /^[-\\s]+$/.test(text) || /^\\d+(?:,\\d{3})*\\s+\\d+(?:,\\d{3})*$/.test(text));
            }) || null;
          }
          if (!target) {
            target = [...row.querySelectorAll('td span, td a, td div')].find(el => isVisible(el) && monthlyValuePattern.test(normalize(el.textContent)));
          }
          if (!target && actionIndex >= 0 && cells[actionIndex]) {
            const controls = [...cells[actionIndex].querySelectorAll('i, svg, a, button, span, div')].filter(isVisible);
            target = controls.find(el => /chart|trend|echarts|analysis|anlysis|趋势|分析|柱|bar|histogram/i.test(`${el.className || ''} ${el.getAttribute('class') || ''} ${el.getAttribute('title') || ''} ${el.getAttribute('aria-label') || ''} ${normalize(el.textContent)}`))
              || controls.find(el => ['I', 'SVG', 'A', 'BUTTON'].includes(el.tagName))
              || controls[0]
              || cells[actionIndex];
          }
          if (!target) {
            const controls = [...row.querySelectorAll('i, svg, a, button, span, div')].filter(isVisible);
            target = controls.find(el => /chart|trend|echarts|analysis|anlysis|趋势|分析|柱|bar|histogram/i.test(`${el.className || ''} ${el.getAttribute('class') || ''} ${el.getAttribute('title') || ''} ${el.getAttribute('aria-label') || ''} ${normalize(el.textContent)}`));
          }
          if (!target && trendIndex >= 0 && cells[trendIndex]) {
            target = [...cells[trendIndex].querySelectorAll('canvas, svg, a, span, div, button')].find(isVisible) || cells[trendIndex];
          }
          if (!target) throw new Error('未找到可打开第一行趋势分析抽屉的控件');
          (target.closest('a, button') || target).click();
        }"""
    )
    page.wait_for_function(
        """() => {
          const el = document.querySelector('#search-trends');
          if (!el) return false;
          const style = window.getComputedStyle(el);
          const rect = el.getBoundingClientRect();
          return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0;
        }""",
        timeout=6000,
    )
    page.wait_for_function(
        """() => {
          const el = document.getElementById('search-trends');
          return !!(el && el.__vue__ && el.__vue__.chart);
        }""",
        timeout=6000,
    )
    sleep(page, 700)


def scroll_drawer_top(page) -> None:
    page.evaluate(
        """() => {
          const drawer = document.querySelector('#search-anlysis-drawer-echarts-trends') || document.querySelector('.el-drawer__wrapper');
          const body = drawer && (drawer.querySelector('.el-drawer__body') || drawer.querySelector('[class*="body"]'));
          if (body) body.scrollTop = 0;
          window.scrollTo(0, 0);
        }"""
    )
    sleep(page, 300)


def click_search_trend_tab(page) -> None:
    page.evaluate(
        """() => {
          const isVisible = (el) => {
            if (!el) return false;
            const style = window.getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0;
          };
          const normalize = (text) => (text || '').replace(/\\s+/g, ' ').trim();
          const drawer = document.querySelector('#search-anlysis-drawer-echarts-trends') || document.querySelector('.el-drawer__wrapper') || document;
          const candidates = [...drawer.querySelectorAll('button, .el-button, label, span, div, a, li')]
            .filter(isVisible)
            .filter(el => {
              const text = normalize(el.innerText || el.textContent);
              return text === '\\u641c\\u7d22\\u8d8b\\u52bf' || text.includes('\\u641c\\u7d22\\u8d8b\\u52bf');
            });
          const target = candidates.find(el => {
            const rect = el.getBoundingClientRect();
            return normalize(el.innerText || el.textContent) === '\\u641c\\u7d22\\u8d8b\\u52bf' && rect.width < 220 && rect.height < 90;
          }) || candidates.find(el => {
            const rect = el.getBoundingClientRect();
            return rect.width < 220 && rect.height < 90;
          }) || candidates[0];
          if (target) target.click();
        }"""
    )
    page.wait_for_function(
        """() => {
          const el = document.querySelector('#search-trends');
          if (!el) return false;
          const style = window.getComputedStyle(el);
          const rect = el.getBoundingClientRect();
          return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0;
        }""",
        timeout=10000,
    )
    sleep(page, 500)


def configure_month_google_trend(page) -> None:
    page.evaluate(
        """() => {
          const isVisible = (el) => {
            if (!el) return false;
            const style = window.getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0;
          };
          const monthButton = [...document.querySelectorAll('button, .el-button, label, span')]
            .find(el => isVisible(el) && (el.innerText || el.textContent || '').trim() === '按月');
          if (!monthButton) throw new Error('找不到“按月”按钮');
          monthButton.click();
        }"""
    )
    page.wait_for_function(
        """() => {
          const el = document.getElementById('search-trends');
          const chart = el && el.__vue__ && el.__vue__.chart;
          if (!chart) return false;
          const opt = chart.getOption();
          const data = opt && opt.xAxis && opt.xAxis[0] && opt.xAxis[0].data;
          return Array.isArray(data) && data.length > 0 && data.length <= 90 && /^\\d{4}-\\d{2}$/.test(String(data[0]));
        }""",
        timeout=30000,
    )
    google_ready = page.evaluate(
        """() => {
          const isVisible = (el) => {
            if (!el) return false;
            const style = window.getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0;
          };
          const el = document.getElementById('search-trends');
          const chart = el && el.__vue__ && el.__vue__.chart;
          if (!chart) throw new Error('找不到 ECharts 实例');
          const hasGoogleSeries = () => {
            const opt = chart.getOption();
            const series = (opt && opt.series) || [];
            return series.some(s => String(s.name || '').includes('Google') && Array.isArray(s.data) && s.data.some(v => v !== null && v !== undefined && v !== ''));
          };
          if (hasGoogleSeries()) return true;
          const label = [...document.querySelectorAll('label.el-checkbox, label, .el-checkbox')]
            .find(node => isVisible(node) && (node.innerText || node.textContent || '').includes('对比Google趋势'));
          if (!label) return false;
          const input = label.querySelector('input[type="checkbox"]');
          const checked = !!(input && input.checked) || /is-checked/.test(label.className || '') || !!label.querySelector('.is-checked');
          if (checked) label.click();
          label.click();
          return false;
        }"""
    )
    if not google_ready:
        try:
            page.wait_for_function(
                """() => {
                  const el = document.getElementById('search-trends');
                  const chart = el && el.__vue__ && el.__vue__.chart;
                  if (!chart) return false;
                  const opt = chart.getOption();
                  const series = (opt && opt.series) || [];
                  return series.some(s => String(s.name || '').includes('Google') && Array.isArray(s.data) && s.data.some(v => v !== null && v !== undefined && v !== ''));
                }""",
                timeout=6000,
            )
        except Exception:
            pass
    sleep(page, 1200)


def click_ppc_bid_tab(page) -> None:
    page.evaluate(
        """() => {
          const isVisible = (el) => {
            if (!el) return false;
            const style = window.getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0;
          };
          const normalize = (text) => (text || '').replace(/\\s+/g, ' ').trim();
          const drawer = document.querySelector('#search-anlysis-drawer-echarts-trends') || document.querySelector('.el-drawer__wrapper') || document;
          const candidates = [...drawer.querySelectorAll('button, .el-button, label, span, div, a, li')]
            .filter(isVisible)
            .filter(el => {
              const text = normalize(el.innerText || el.textContent);
              return text === 'PPC\\u7ade\\u4ef7' || text.includes('PPC\\u7ade\\u4ef7');
            });
          const target = candidates.find(el => {
            const rect = el.getBoundingClientRect();
            return rect.width < 220 && rect.height < 90;
          }) || candidates[0];
          if (!target) throw new Error('找不到 PPC 竞价页签');
          target.click();
        }"""
    )
    page.wait_for_function(
        """() => {
          const el = document.getElementById('ppc-trends');
          if (!el) return false;
          const style = window.getComputedStyle(el);
          const rect = el.getBoundingClientRect();
          const chart = (el.__vue__ && el.__vue__.chart) || (window.echarts && window.echarts.getInstanceByDom && window.echarts.getInstanceByDom(el));
          return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0 && !!chart;
        }""",
        timeout=15000,
    )
    sleep(page, 1000)


def click_ppc_monthly_history(page) -> bool:
    clicked = bool(page.evaluate(
        """() => {
          const isVisible = (el) => {
            if (!el) return false;
            const style = window.getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0;
          };
          const normalize = (text) => String(text || '').replace(/\\s+/g, ' ').trim();
          const drawer = document.querySelector('#search-anlysis-drawer-echarts-trends') || document.querySelector('.el-drawer__wrapper') || document;
          const candidates = [...drawer.querySelectorAll('button, .el-button, a, span, div, li')]
            .filter(isVisible)
            .filter(el => {
              const text = normalize(el.innerText || el.textContent);
              return /查看.*历史.*月|历史.*月数据|月度.*历史|Monthly|month history|historical month/i.test(text);
            });
          const target = candidates.find(el => {
            const rect = el.getBoundingClientRect();
            return rect.width < 260 && rect.height < 100;
          }) || candidates[0];
          if (!target) return false;
          target.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
          target.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
          target.click();
          return true;
        }"""
    ))
    if clicked:
        sleep(page, 1800)
    return clicked


def merge_ppc_data(primary: dict[str, Any], secondary: dict[str, Any]) -> dict[str, Any]:
    def dedupe(rows: Sequence[dict], keys: Sequence[str]) -> list[dict]:
        seen: set[tuple] = set()
        merged: list[dict] = []
        for row in rows:
            marker = tuple(row.get(key) for key in keys)
            if marker in seen:
                continue
            seen.add(marker)
            merged.append(dict(row))
        return merged

    weekly = dedupe([*primary.get("ppcWeekly", []), *secondary.get("ppcWeekly", [])], ("week", "weekEnd", "bid"))
    monthly = dedupe([*primary.get("ppcMonthly", []), *secondary.get("ppcMonthly", [])], ("month", "bid"))
    merged = dict(primary)
    merged["ppcWeekly"] = weekly
    merged["ppcMonthly"] = monthly
    merged["ppcStatus"] = "success" if weekly or monthly else primary.get("ppcStatus") or secondary.get("ppcStatus") or "not_available"
    merged["ppcWeeklySource"] = primary.get("ppcWeeklySource") or secondary.get("ppcWeeklySource") or ("table-or-echarts" if weekly else "")
    merged["ppcMonthlySource"] = primary.get("ppcMonthlySource") or secondary.get("ppcMonthlySource") or ("table-or-echarts" if monthly else "")
    if secondary.get("ppcError") and not merged.get("ppcError"):
        merged["ppcError"] = secondary.get("ppcError", "")
    return merged


def set_zoom_to_2023(page) -> dict[str, Any]:
    return page.evaluate(
        """() => {
          const el = document.getElementById('search-trends');
          if (!el) throw new Error('找不到 #search-trends 元素');
          const chart = el.__vue__ && el.__vue__.chart;
          if (!chart) throw new Error('找不到 ECharts 实例');
          const opt = chart.getOption();
          const data = opt && opt.xAxis && opt.xAxis[0] && opt.xAxis[0].data;
          if (!Array.isArray(data) || !data.length) throw new Error('找不到按月 xAxis 数据');
          const idx2023 = data.findIndex(d => d === '2023-01');
          const startIndex = idx2023 >= 0 ? idx2023 : 0;
          chart.dispatchAction({ type: 'dataZoom', startValue: startIndex, endValue: data.length - 1 });
          return { visibleStart: data[startIndex], visibleEnd: data[data.length - 1], totalMonths: data.length, usedFallbackStart: idx2023 < 0 };
        }"""
    )


def chart_png_dimensions(base64_png: str) -> tuple[int, int]:
    encoded = base64_png.split(",", 1)[1] if base64_png.startswith("data:") and "," in base64_png else base64_png
    missing_padding = (-len(encoded)) % 4
    if missing_padding:
        encoded += "=" * missing_padding
    data = base64.b64decode(encoded, validate=True)
    if len(data) < 24 or not data.startswith(b"\x89PNG\r\n\x1a\n"):
        raise ValueError("ECharts 导出结果不是有效 PNG")
    return int.from_bytes(data[16:20], "big"), int.from_bytes(data[20:24], "big")


def normalize_chart_export_size(page) -> dict[str, int]:
    size = page.evaluate(
        """({ width, height }) => {
          const el = document.getElementById('search-trends');
          if (!el) throw new Error('找不到 #search-trends 元素');
          const chart = el.__vue__ && el.__vue__.chart;
          if (!chart) throw new Error('找不到 ECharts 实例');
          chart.resize({ width, height, silent: true });
          const renderer = chart.getZr && chart.getZr();
          if (renderer && renderer.flush) renderer.flush();
          return { width: chart.getWidth(), height: chart.getHeight() };
        }""",
        {"width": TREND_EXPORT_CSS_WIDTH, "height": TREND_EXPORT_CSS_HEIGHT},
    )
    page.wait_for_function(
        """({ width, height }) => {
          const el = document.getElementById('search-trends');
          const chart = el && el.__vue__ && el.__vue__.chart;
          return !!chart && chart.getWidth() === width && chart.getHeight() === height;
        }""",
        arg={"width": TREND_EXPORT_CSS_WIDTH, "height": TREND_EXPORT_CSS_HEIGHT},
        timeout=5000,
    )
    sleep(page, 500)
    return {"width": int(size["width"]), "height": int(size["height"])}


def export_chart_base64(page) -> str:
    expected = (TREND_EXPORT_PNG_WIDTH, TREND_EXPORT_PNG_HEIGHT)
    for attempt in range(2):
        normalize_chart_export_size(page)
        sleep(page, 1000)
        base64_png = str(page.evaluate(
            """(pixelRatio) => {
              const el = document.getElementById('search-trends');
              if (!el) throw new Error('找不到 #search-trends 元素');
              const chart = el.__vue__ && el.__vue__.chart;
              if (!chart) throw new Error('找不到 ECharts 实例');
              return chart.getDataURL({ type: 'png', pixelRatio, backgroundColor: '#fff' }).split(',')[1];
            }""",
            TREND_EXPORT_PIXEL_RATIO,
        ))
        actual = chart_png_dimensions(base64_png)
        if actual == expected:
            return base64_png
        if attempt == 0:
            continue
        raise RuntimeError(f"搜索趋势图尺寸异常: expected={expected}, actual={actual}")
    raise RuntimeError("搜索趋势图导出失败")


def extract_search_trend_monthly_data(page) -> list[dict[str, Any]]:
    return page.evaluate(
        """() => {
          const normalizeNumber = (value) => {
            if (value === null || value === undefined || value === '') return null;
            if (typeof value === 'number') return Number.isFinite(value) ? value : null;
            const match = String(value).match(/-?\\d+(?:,\\d{3})*(?:\\.\\d+)?/);
            return match ? Number(match[0].replace(/,/g, '')) : null;
          };
          const normalizeSeriesValue = (raw) => {
            if (Array.isArray(raw)) return normalizeNumber(raw[1] !== undefined ? raw[1] : raw[0]);
            if (raw && typeof raw === 'object') {
              const value = raw.value;
              if (Array.isArray(value)) return normalizeNumber(value[1] !== undefined ? value[1] : value[0]);
              return normalizeNumber(value);
            }
            return normalizeNumber(raw);
          };
          const el = document.getElementById('search-trends');
          if (!el) throw new Error('找不到 #search-trends 元素');
          const chart = el.__vue__ && el.__vue__.chart;
          if (!chart) throw new Error('找不到 ECharts 实例');
          const opt = chart.getOption();
          const months = opt && opt.xAxis && opt.xAxis[0] && opt.xAxis[0].data;
          const series = (opt && opt.series) || [];
          if (!Array.isArray(months) || !months.length) return [];
          const searchSeries = series.find(s => /搜索量/.test(String(s.name || '')) && !/Google|谷歌|购买|购买率|排名|ABA/i.test(String(s.name || '')))
            || series.find(s => /搜索/.test(String(s.name || '')) && !/Google|谷歌/i.test(String(s.name || '')));
          if (!searchSeries || !Array.isArray(searchSeries.data)) return [];
          return months.map((month, index) => ({ month: String(month), searchVolume: normalizeSeriesValue(searchSeries.data[index]) }))
            .filter(point => /^\\d{4}-\\d{2}$/.test(point.month) && point.searchVolume !== null)
            .map(point => ({ month: point.month, searchVolume: Math.round(point.searchVolume) }));
        }"""
    )


def extract_ppc_data(page) -> dict[str, Any]:
    return page.evaluate(
        """() => {
          const blockedPattern = /VIP|会员|权限不足|无权限|试用到期|验证码|安全验证|访问过于频繁|风控/;
          const normalizeNumber = (value) => {
            if (value === null || value === undefined || value === '') return null;
            const match = String(value).match(/-?\\d+(?:,\\d{3})*(?:\\.\\d+)?/);
            return match ? Number(match[0].replace(/,/g, '')) : null;
          };
          const normalizeSeriesValue = (raw) => {
            if (Array.isArray(raw)) return normalizeNumber(raw[1] !== undefined ? raw[1] : raw[0]);
            if (raw && typeof raw === 'object') {
              const value = raw.value;
              if (Array.isArray(value)) return normalizeNumber(value[1] !== undefined ? value[1] : value[0]);
              return normalizeNumber(value);
            }
            return normalizeNumber(raw);
          };
          const parseMonth = (text) => {
            const match = String(text || '').match(/(20\\d{2})[-\\/年.](\\d{1,2})/);
            return match ? `${match[1]}-${String(Number(match[2])).padStart(2, '0')}` : null;
          };
          const parseWeekRange = (text) => {
            const raw = String(text || '');
            const weeklyLabel = raw.match(/(20\\d{2})\\s*(?:\\u7b2c)?\\d{1,2}\\s*\\u5468\\s*[\\(\\uff08]\\s*(\\d{1,2})[\\/.-](\\d{1,2})\\s*[~\\-\\u2013\\u2014\\u81f3]\\s*(\\d{1,2})[\\/.-](\\d{1,2})\\s*[\\)\\uff09]/);
            if (weeklyLabel) {
              const year = Number(weeklyLabel[1]);
              const startMonth = Number(weeklyLabel[2]);
              const startDay = Number(weeklyLabel[3]);
              const endMonth = Number(weeklyLabel[4]);
              const endDay = Number(weeklyLabel[5]);
              const endYear = endMonth < startMonth ? year + 1 : year;
              return {
                week: `${year}-${String(startMonth).padStart(2, '0')}-${String(startDay).padStart(2, '0')}`,
                weekEnd: `${endYear}-${String(endMonth).padStart(2, '0')}-${String(endDay).padStart(2, '0')}`,
              };
            }
            const matches = [...raw.matchAll(/(20\\d{2})[-\\/年.](\\d{1,2})[-\\/月.](\\d{1,2})/g)];
            const formatParts = (year, month, day) => `${year}-${String(Number(month)).padStart(2, '0')}-${String(Number(day)).padStart(2, '0')}`;
            if (matches.length) {
              const result = { week: formatParts(matches[0][1], matches[0][2], matches[0][3]) };
              if (matches.length > 1) result.weekEnd = formatParts(matches[1][1], matches[1][2], matches[1][3]);
              return result;
            }
            return null;
          };
          const isVisible = (el) => {
            if (!el) return false;
            const style = window.getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0;
          };
          const drawer = document.querySelector('#search-anlysis-drawer-echarts-trends') || document.querySelector('.el-drawer__wrapper') || document;
          const visibleTextNodes = [...drawer.querySelectorAll('.el-alert, .el-message, .empty, .no-data, [class*=empty], [class*=permission], [class*=vip], [class*=lock]')]
            .filter(isVisible)
            .map(el => (el.innerText || el.textContent || '').replace(/\\s+/g, ' '));
          if (visibleTextNodes.some(text => blockedPattern.test(text))) {
            return { ppcStatus: 'blocked', ppcError: 'PPC 区域出现权限、验证或风控提示', ppcWeekly: [], ppcMonthly: [] };
          }
          const tables = [...drawer.querySelectorAll('table')].filter(isVisible);
          const extractTable = (mode) => {
            for (const table of tables) {
              const rows = [...table.querySelectorAll('tr')].filter(isVisible);
              if (rows.length < 2) continue;
              const headers = [...rows[0].querySelectorAll('th,td')].map(cell => (cell.innerText || cell.textContent || '').trim());
              const timeIndex = headers.findIndex(text => mode === 'weekly' ? /周|日期|时间|week|date/i.test(text) : /月份|月|month/i.test(text));
              let bidIndex = headers.findIndex(text => /最大值|最高|上限|max/i.test(text));
              if (bidIndex < 0) bidIndex = headers.findIndex(text => /PPC|竞价|建议竞价|广告竞价|bid/i.test(text));
              if (timeIndex < 0 || bidIndex < 0) continue;
              const data = [];
              for (const row of rows.slice(1)) {
                const cells = [...row.querySelectorAll('td,th')].map(cell => (cell.innerText || cell.textContent || '').trim());
                const bid = normalizeNumber(cells[bidIndex]);
                if (bid === null) continue;
                if (mode === 'weekly') {
                  const week = parseWeekRange(cells[timeIndex]);
                  if (week) data.push({ ...week, bid });
                } else {
                  const month = parseMonth(cells[timeIndex]);
                  if (month) data.push({ month, bid });
                }
              }
              if (data.length) return data;
            }
            return [];
          };
          const vueCharts = [];
          const visit = (node) => {
            if (!node || vueCharts.length > 30) return;
            if (isVisible(node) && node.__vue__ && node.__vue__.chart && !vueCharts.includes(node)) vueCharts.push(node);
            for (const child of node.children || []) visit(child);
          };
          visit(drawer);
          if (!vueCharts.length && window.echarts && typeof window.echarts.getInstanceByDom === 'function') {
            for (const el of [...drawer.querySelectorAll('div')].filter(isVisible)) {
              try {
                const chart = window.echarts.getInstanceByDom(el);
                if (chart && !vueCharts.includes(el)) {
                  el.__codexChart = chart;
                  vueCharts.push(el);
                }
              } catch {}
            }
          }
          const extractChart = (mode) => {
            for (const el of vueCharts) {
              const chart = (el.__vue__ && el.__vue__.chart) || el.__codexChart;
              const opt = chart.getOption && chart.getOption();
              const labels = opt && opt.xAxis && opt.xAxis[0] && opt.xAxis[0].data;
              const series = (opt && opt.series) || [];
              if (!Array.isArray(labels)) continue;
              const bidSeries = series.find(s => /最大值|最高|上限|max/i.test(String(s.name || '')) && Array.isArray(s.data))
                || series.find(s => /PPC|竞价|建议竞价|广告竞价|bid/i.test(String(s.name || '')) && Array.isArray(s.data));
              if (!bidSeries) continue;
              const data = labels.map((label, index) => {
                const bid = normalizeSeriesValue(bidSeries.data[index]);
                if (bid === null) return null;
                if (mode === 'weekly') {
                  const week = parseWeekRange(label);
                  return week ? { ...week, bid } : null;
                }
                const month = /^\\d{4}-\\d{2}$/.test(String(label)) ? String(label) : parseMonth(label);
                return month ? { month, bid } : null;
              }).filter(Boolean);
              if (data.length) return data;
            }
            return [];
          };
          const ppcWeekly = extractTable('weekly').concat(extractChart('weekly'));
          const ppcMonthly = extractTable('monthly').concat(extractChart('monthly'));
          return {
            ppcStatus: (ppcWeekly.length || ppcMonthly.length) ? 'success' : 'not_available',
            ppcError: '',
            ppcSource: 'python-playwright',
            ppcWeeklySource: ppcWeekly.length ? 'table-or-echarts' : '',
            ppcMonthlySource: ppcMonthly.length ? 'table-or-echarts' : '',
            ppcWeekly,
            ppcMonthly,
          };
        }"""
    )


def extract_count_from_amazon_text(text: str) -> dict[str, Any]:
    compact = re.sub(r"\s+", " ", text or "")
    if re.search(r"captcha|robot check|enter the characters|sorry, we just need to make sure|automated access", compact, re.I):
        return {"status": "blocked", "error": "Amazon 出现验证码或机器人校验"}
    patterns = [
        r"(\d+\s*-\s*\d+\s+of\s+)?(over\s+)?([\d,]+)\s+results?\s+for",
        r"\d+\s*-\s*\d+\s+von\s+([\d.,]+)\s+Ergebnissen?",
        r"von\s+(?:mehr\s+als\s+|über\s+)?([\d.,]+)\s+Ergebnissen?",
        r"(?:mehr\s+als|über)\s+([\d.,]+)\s+Ergebnisse?",
        r"([\d.,]+)\s+Ergebnisse?\s+für",
        r"([\d,]+)\s+results?\s+for",
        r"over\s+([\d,]+)\s+results?",
        r"([\d,]+)\s+results?",
        r"of\s+(over\s+)?([\d,]+)\s+results?\s+for",
        r"of\s+(over\s+)?([\d,]+)\s+results?",
        r"results?\s*[:：]\s*(over\s+)?([\d,]+)",
        r"(over\s+)?([\d,]+)\s+results?\s+for",
        r"显示[:：]\s*\d+\s*-\s*\d+\s*条[，,]\s*共\s*([\d,]+)\s*条",
        r"共\s*超过\s*([\d,]+)\s*条",
        r"共\s*([\d,]+)\s*条\s*[\"“]",
        r"共\s*([\d,]+)\s*条\b",
        r"共\s*([\d,]+)\s*条\s*结果",
        r"共\s*约\s*([\d,]+)\s*条",
        r"超过\s*([\d,]+)\s*条\s*结果",
        r"共\s*超过\s*([\d,]+)\s*个\s*结果",
        r"共\s*([\d,]+)\s*个\s*结果",
    ]
    for pattern in patterns:
        match = re.search(pattern, compact, re.I)
        if not match:
            continue
        count_text = next((group for group in reversed(match.groups()) if group and re.search(r"\d", group)), match.group(0))
        count = int(re.sub(r"[^\d]", "", count_text))
        return {
            "status": "success",
            "text": match.group(0).strip(),
            "count": count,
            "isApproximate": bool(re.search(r"over|超过|约", match.group(0), re.I)),
        }
    return {"status": "failed", "error": "未找到 Amazon 搜索结果数量文本"}


def amazon_debug_excerpt(text: str) -> str:
    compact = re.sub(r"\s+", " ", text or "").strip()
    if not compact:
        return ""
    patterns = [
        r".{0,120}(results?|RESULTS?|search results?|共.{0,12}结果|验证码|captcha|robot check).{0,220}",
        r".{0,120}(No results|没有找到|Try checking your spelling|Enter the characters).{0,220}",
    ]
    snippets: list[str] = []
    for pattern in patterns:
        for match in re.finditer(pattern, compact, re.I):
            snippet = match.group(0).strip()
            if snippet and snippet not in snippets:
                snippets.append(snippet)
            if len(snippets) >= 3:
                break
        if len(snippets) >= 3:
            break
    return " || ".join(snippets)[:1000] or compact[:1000]


def ensure_amazon_delivery_zip(amazon_page, zip_code: str = DEFAULT_AMAZON_DELIVERY_ZIP) -> dict[str, Any]:
    zip_code = str(zip_code or "").strip()
    if not zip_code:
        return {"status": "skipped", "reason": "empty-zip"}
    try:
        current = amazon_page.evaluate(
            """() => {
              const text = (document.querySelector('#glow-ingress-line2')?.textContent || document.querySelector('#nav-global-location-slot')?.textContent || '').replace(/\\s+/g, ' ').trim();
              return { text, hasTargetZip: text.includes('__ZIP__') };
            }""".replace("__ZIP__", zip_code)
        )
        if current.get("hasTargetZip"):
            return {"status": "already-set", "text": current.get("text", "")}

        amazon_page.evaluate(
            """() => {
              const isVisible = (el) => {
                if (!el) return false;
                const style = window.getComputedStyle(el);
                const rect = el.getBoundingClientRect();
                return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0;
              };
              const target = document.querySelector('#nav-global-location-popover-link')
                || document.querySelector('#glow-ingress-block')
                || document.querySelector('#nav-global-location-slot')
                || [...document.querySelectorAll('a, button, div')].find(el => isVisible(el) && /Deliver to|配送至|送至|邮编|郵編|zip/i.test(el.innerText || el.textContent || ''));
              if (!target) throw new Error('找不到 Amazon 配送位置入口');
              target.click();
            }"""
        )
        amazon_page.wait_for_selector(
            "#GLUXZipUpdateInput, #GLUXChangePostalCodeLink, input[name='zipCode'], input[placeholder*='zip'], input[aria-label*='zip'], input[aria-label*='postal']",
            state="visible",
            timeout=12000,
        )
        try:
            change_link = amazon_page.locator("#GLUXChangePostalCodeLink").first
            if change_link.is_visible(timeout=1000):
                change_link.click(force=True, timeout=5000)
                amazon_page.wait_for_timeout(1000)
        except Exception:
            pass
        amazon_page.wait_for_selector(
            "#GLUXZipUpdateInput, input[name='zipCode'], input[placeholder*='zip'], input[aria-label*='zip'], input[aria-label*='postal']",
            state="visible",
            timeout=12000,
        )
        zip_locator = amazon_page.locator(
            "#GLUXZipUpdateInput, input[name='zipCode'], input[placeholder*='zip'], input[aria-label*='zip'], input[aria-label*='postal']"
        ).first
        zip_locator.fill(zip_code, timeout=10000)

        amazon_page.evaluate(
            """() => {
              const isVisible = (el) => {
                if (!el) return false;
                const style = window.getComputedStyle(el);
                const rect = el.getBoundingClientRect();
                return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0;
              };
              const buttons = [...document.querySelectorAll('input[type="submit"], input[type="button"], button, .a-button-input, .a-button-text, span')]
                .filter(isVisible);
              const apply = buttons.find(el => /^(Apply|应用|套用|确定)$/i.test((el.value || el.innerText || el.textContent || '').replace(/\\s+/g, ' ').trim()))
                || [...document.querySelectorAll('#GLUXZipUpdate input[type="submit"], #GLUXZipUpdate .a-button-input')].find(isVisible)
                || document.querySelector('#GLUXZipUpdate input[type="submit"], #GLUXZipUpdate .a-button-input, span#GLUXZipUpdate input');
              if (!apply) throw new Error('找不到 Amazon 邮编 Apply 按钮');
              apply.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
              apply.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
              apply.click();
            }"""
        )
        amazon_page.wait_for_timeout(2500)

        try:
            amazon_page.evaluate(
                """() => {
                  const isVisible = (el) => {
                    if (!el) return false;
                    const style = window.getComputedStyle(el);
                    const rect = el.getBoundingClientRect();
                    return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0;
                  };
                  const direct = document.querySelector('button[name="glowDoneButton"], input[name="glowDoneButton"]');
                  const buttons = [...document.querySelectorAll('button, input[type="submit"], input[type="button"], .a-button-input, .a-button-text, span')]
                    .filter(isVisible);
                  const done = (direct && isVisible(direct) ? direct : null)
                    || buttons.find(el => /^(Done|完成|确定|确认)$/i.test((el.value || el.innerText || el.textContent || '').replace(/\\s+/g, ' ').trim()))
                    || (document.querySelector('#GLUXConfirmClose') && isVisible(document.querySelector('#GLUXConfirmClose')) ? document.querySelector('#GLUXConfirmClose') : null)
                    || document.querySelector('.a-popover-footer .a-button-input');
                  if (done) {
                    done.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
                    done.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
                    done.click();
                  }
                }"""
            )
            try:
                amazon_page.wait_for_function(
                    """(zipCode) => {
                      const text = (document.querySelector('#glow-ingress-line2')?.textContent || document.querySelector('#nav-global-location-slot')?.textContent || '').replace(/\\s+/g, ' ').trim();
                      return text.includes(zipCode);
                    }""",
                    zip_code,
                    timeout=10000,
                )
            except Exception:
                amazon_page.wait_for_timeout(2000)
        except Exception:
            pass

        updated = amazon_page.evaluate(
            """() => {
              const text = (document.querySelector('#glow-ingress-line2')?.textContent || document.querySelector('#nav-global-location-slot')?.textContent || '').replace(/\\s+/g, ' ').trim();
              return { text, hasTargetZip: text.includes('__ZIP__') };
            }""".replace("__ZIP__", zip_code)
        )
        if updated.get("hasTargetZip"):
            return {"status": "success", "text": updated.get("text", "")}
        return {
            "status": "failed",
            "error": f"Amazon location text did not include target postal code {zip_code}",
            "text": updated.get("text", ""),
        }
    except Exception as exc:
        return {"status": "failed", "error": str(exc)}


def ensure_amazon_delivery_zip(amazon_page, zip_code: str = DEFAULT_AMAZON_DELIVERY_ZIP) -> dict[str, Any]:
    zip_code = str(zip_code or "").strip()
    if not zip_code:
        return {"status": "skipped", "reason": "empty-zip"}

    def read_location() -> dict[str, Any]:
        return amazon_page.evaluate(
            """(zipCode) => {
              const text = (document.querySelector('#glow-ingress-line2')?.textContent || document.querySelector('#nav-global-location-slot')?.textContent || '').replace(/\\s+/g, ' ').trim();
              return { text, hasTargetZip: text.includes(zipCode) };
            }""",
            zip_code,
        )

    try:
        current = read_location()
        if current.get("hasTargetZip"):
            return {"status": "already-set", "text": current.get("text", "")}

        amazon_page.locator("#nav-global-location-popover-link, #glow-ingress-block, #nav-global-location-slot").first.click(
            force=True,
            timeout=10000,
        )
        amazon_page.wait_for_selector("#GLUXZipUpdateInput, #GLUXChangePostalCodeLink", state="attached", timeout=12000)
        amazon_page.wait_for_timeout(700)

        try:
            change_link = amazon_page.locator("#GLUXChangePostalCodeLink").first
            if change_link.is_visible(timeout=1000):
                change_link.click(force=True, timeout=5000)
                amazon_page.wait_for_timeout(700)
        except Exception:
            pass

        zip_input = amazon_page.locator("#GLUXZipUpdateInput").first
        zip_input.wait_for(state="visible", timeout=12000)
        zip_input.fill(zip_code, timeout=10000)
        amazon_page.wait_for_timeout(300)

        amazon_page.evaluate(
            """() => {
              const isVisible = (el) => {
                if (!el) return false;
                const style = window.getComputedStyle(el);
                const rect = el.getBoundingClientRect();
                return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0;
              };
              const candidates = [
                document.querySelector('#GLUXZipUpdate-announce'),
                ...document.querySelectorAll('#GLUXZipUpdate input, #GLUXZipUpdate button, #GLUXZipUpdate .a-button-input, #GLUXZipUpdate .a-button-text')
              ].filter(Boolean);
              const apply = candidates.find(el => isVisible(el) && /Apply|应用|套用|确定/i.test((el.value || el.innerText || el.textContent || '').replace(/\\s+/g, ' ').trim()))
                || candidates.find(isVisible)
                || candidates[0];
              if (!apply) throw new Error('Amazon zip Apply button not found');
              apply.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
              apply.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
              apply.click();
            }"""
        )
        amazon_page.wait_for_timeout(2500)

        amazon_page.evaluate(
            """() => {
              const isVisible = (el) => {
                if (!el) return false;
                const style = window.getComputedStyle(el);
                const rect = el.getBoundingClientRect();
                return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0;
              };
              const candidates = [
                ...document.querySelectorAll('button[name="glowDoneButton"], input[name="glowDoneButton"], #GLUXConfirmClose, #GLUXConfirmClose-announce'),
                ...document.querySelectorAll('button, input[type="submit"], input[type="button"], .a-button-input, .a-button-text')
              ];
              const done = candidates.find(el => isVisible(el) && /^(Done|Continue|完成|确定|确认)$/i.test((el.value || el.innerText || el.textContent || '').replace(/\\s+/g, ' ').trim()))
                || candidates.find(el => isVisible(el) && el.getAttribute('name') === 'glowDoneButton');
              if (done) {
                done.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
                done.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
                done.click();
              }
            }"""
        )

        try:
            amazon_page.wait_for_function(
                """(zipCode) => {
                  const text = (document.querySelector('#glow-ingress-line2')?.textContent || document.querySelector('#nav-global-location-slot')?.textContent || '').replace(/\\s+/g, ' ').trim();
                  return text.includes(zipCode);
                }""",
                zip_code,
                timeout=12000,
            )
        except Exception:
            amazon_page.wait_for_timeout(2000)

        updated = read_location()
        if updated.get("hasTargetZip"):
            return {"status": "success", "text": updated.get("text", "")}
        return {
            "status": "failed",
            "error": f"Amazon location text did not include target postal code {zip_code}",
            "text": updated.get("text", ""),
        }
    except Exception as exc:
        return {"status": "failed", "error": str(exc)}


def ensure_amazon_delivery_zip(amazon_page, zip_code: str = DEFAULT_AMAZON_DELIVERY_ZIP) -> dict[str, Any]:
    zip_code = str(zip_code or "").strip()
    if not zip_code:
        return {"status": "skipped", "reason": "empty-zip"}
    try:
        for selector in ("#sp-cc-accept", "#sp-cc-rejectall-link"):
            try:
                button = amazon_page.locator(selector).first
                if button.is_visible(timeout=800):
                    button.click(timeout=3000)
                    amazon_page.wait_for_timeout(1000)
                    break
            except Exception:
                continue
        return amazon_page.evaluate(
            """async (zipCode) => {
              const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));
              const isVisible = (el) => {
                if (!el) return false;
                const style = window.getComputedStyle(el);
                const rect = el.getBoundingClientRect();
                return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0;
              };
              const normalize = (text) => String(text || '').replace(/\\s+/g, ' ').trim();
              const compactPostal = (text) => normalize(text)
                .replace(/[\\u200c-\\u200f\\ufeff]/g, '')
                .replace(/[^a-z0-9]/gi, '')
                .toUpperCase();
              const postalNeedles = (code) => {
                const full = compactPostal(code);
                const sector = full.replace(/[A-Z]{2}$/, '');
                return [full, sector].filter((value, index, list) => value.length >= 4 && list.indexOf(value) === index);
              };
              const matchesPostalCode = (text) => {
                const haystack = compactPostal(text);
                return postalNeedles(zipCode).some((needle) => haystack.includes(needle));
              };
              const locationText = () => normalize((document.querySelector('#glow-ingress-line2')?.textContent || document.querySelector('#nav-global-location-slot')?.textContent || ''));
              const clickElement = (el) => {
                if (!el) return false;
                el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
                el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
                el.click();
                return true;
              };
              const dismissCookieBanner = async () => {
                const cookieButton = [
                  document.querySelector('#sp-cc-accept'),
                  document.querySelector('#sp-cc-rejectall-link'),
                  ...document.querySelectorAll('input[name="accept"], input[aria-label="Accept"], input[value="Accept"], button, input[type="submit"], .a-button-input, .a-button-text')
                ].find(el => isVisible(el) && /^(Accept|Decline|Reject)$/i.test(normalize(el.value || el.innerText || el.textContent || el.getAttribute('aria-label'))));
                if (clickElement(cookieButton)) {
                  await sleep(800);
                }
              };
              if (matchesPostalCode(locationText())) {
                return { status: 'already-set', text: locationText() };
              }
              await dismissCookieBanner();

              const entry = document.querySelector('#nav-global-location-popover-link')
                || document.querySelector('#glow-ingress-block')
                || document.querySelector('#nav-global-location-slot');
              if (!clickElement(entry)) return { status: 'failed', error: 'Amazon location entry not found', text: locationText() };
              await sleep(1200);

              const change = document.querySelector('#GLUXChangePostalCodeLink');
              if (isVisible(change)) {
                clickElement(change);
                await sleep(700);
              }

              const input = document.querySelector('#GLUXZipUpdateInput');
              if (!isVisible(input)) {
                return { status: 'failed', error: 'Amazon zip input not visible', text: locationText() };
              }
              const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set;
              if (setter) setter.call(input, zipCode);
              else input.value = zipCode;
              input.dispatchEvent(new Event('input', { bubbles: true }));
              input.dispatchEvent(new Event('change', { bubbles: true }));
              await sleep(300);

              const applyCandidates = [
                document.querySelector('#GLUXZipUpdate-announce'),
                ...document.querySelectorAll('#GLUXZipUpdate input, #GLUXZipUpdate button, #GLUXZipUpdate .a-button-input, #GLUXZipUpdate .a-button-text')
              ].filter(Boolean);
              const apply = applyCandidates.find(el => isVisible(el) && /Apply|应用|套用|确定/i.test(normalize(el.value || el.innerText || el.textContent)))
                || applyCandidates.find(isVisible)
                || applyCandidates[0];
              if (!clickElement(apply)) return { status: 'failed', error: 'Amazon zip Apply button not found', text: locationText() };
              await sleep(2500);

              const doneCandidates = [
                ...document.querySelectorAll('button[name="glowDoneButton"], input[name="glowDoneButton"], #GLUXConfirmClose, #GLUXConfirmClose-announce'),
                ...document.querySelectorAll('button, input[type="submit"], input[type="button"], .a-button-input, .a-button-text')
              ];
              const done = doneCandidates.find(el => isVisible(el) && /^(Done|Continue|完成|确定|确认)$/i.test(normalize(el.value || el.innerText || el.textContent)))
                || doneCandidates.find(el => isVisible(el) && el.getAttribute('name') === 'glowDoneButton');
              clickElement(done);

              for (let i = 0; i < 24; i++) {
                await sleep(500);
                if (matchesPostalCode(locationText())) {
                  return { status: 'success', text: locationText() };
                }
              }
              return {
                status: 'failed',
                error: `Amazon location text did not include target postal code ${zipCode}`,
                text: locationText()
              };
            }""",
            zip_code,
        )
    except Exception as exc:
        return {"status": "failed", "error": str(exc)}


def check_amazon_login_state(amazon_page) -> dict[str, Any]:
    try:
        return amazon_page.evaluate(
            """() => {
              const normalize = (text) => String(text || '').replace(/\\s+/g, ' ').trim();
              const text = normalize(document.body ? document.body.innerText : '');
              const url = location.href;
              const accountText = normalize(
                document.querySelector('#nav-link-accountList')?.innerText
                || document.querySelector('#nav-link-accountList-nav-line-1')?.innerText
                || ''
              );
              const hasPasswordInput = !!document.querySelector('input[type="password"]');
              const hasSigninUrl = /signin|sign-in|login|ap\\/signin/i.test(url);
              const hasSigninText = /Hello,\\s*sign in|Sign in securely|Sign in for the best experience|登录|登陆|请登录|您好[，,]?\\s*请登录/i.test(`${accountText} ${text}`);
              const hasAccountLists = /Account\\s*&\\s*Lists|账户与列表|账户及列表/i.test(text);
              const looksLoggedIn = /Hello,\\s*(?!sign in)[^\\n]+|您好[，,]?\\s*(?!请登录)[^\\n]+/i.test(accountText)
                || /Your Account|您的账户|账户服务/i.test(accountText);
              if (hasPasswordInput || hasSigninUrl || hasSigninText || (hasAccountLists && /Hello,\\s*sign in/i.test(text))) {
                return { status: 'logged-out', accountText, url, reason: 'signin-indicator' };
              }
              if (looksLoggedIn) {
                return { status: 'logged-in', accountText, url };
              }
              return { status: 'unknown', accountText, url };
            }"""
        )
    except Exception as exc:
        return {"status": "failed", "error": str(exc)}


def require_amazon_login(amazon_page) -> dict[str, Any]:
    state = check_amazon_login_state(amazon_page)
    if state.get("status") == "logged-out":
        raise ManualActionRequired(
            "Amazon 登录态已失效，请在项目 Chrome 浏览器中手动登录 Amazon 后重新运行，避免未登录状态导致站内结果数偏差。",
            reason="login",
            stage="amazon-login",
        )
    return state


def check_amazon_login_state(amazon_page) -> dict[str, Any]:
    try:
        return amazon_page.evaluate(
            """() => {
              const normalize = (text) => String(text || '').replace(/\\s+/g, ' ').trim();
              const isVisible = (el) => {
                if (!el) return false;
                const style = window.getComputedStyle(el);
                const rect = el.getBoundingClientRect();
                return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0;
              };
              const text = normalize(document.body ? document.body.innerText : '');
              const url = location.href;
              const accountText = normalize(
                document.querySelector('#nav-link-accountList-nav-line-1')?.innerText
                || document.querySelector('#nav-link-accountList')?.innerText
                || ''
              );
              const hasPasswordInput = [...document.querySelectorAll('input[type="password"], #ap_password')].some(isVisible);
              const hasEmailInput = [...document.querySelectorAll('input[type="email"], #ap_email')].some(isVisible);
              const hasSigninUrl = /signin|sign-in|login|ap\\/signin/i.test(url);
              const hasSigninText = /Hello,\\s*sign in|Sign in securely|Sign in for the best experience|登录|登陆|请登录|您好[，,]?\\s*请登录/i.test(`${accountText} ${text}`);
              const looksLoggedIn = /Hello,\\s*(?!sign in)[^\\n]+|您好[，,]?\\s*(?!请登录)[^\\n]+/i.test(accountText)
                || /Your Account|您的账户|账户服务/i.test(`${accountText} ${text}`);
              if (looksLoggedIn) {
                return { status: 'logged-in', accountText, url };
              }
              if (hasPasswordInput || hasEmailInput) {
                return { status: 'logged-out', accountText, url, reason: 'credential-input' };
              }
              if (hasSigninUrl || hasSigninText) {
                return { status: 'unknown', accountText, url, reason: 'signin-shell' };
              }
              return { status: 'unknown', accountText, url };
            }"""
        )
    except Exception as exc:
        return {"status": "failed", "error": str(exc)}


def require_amazon_login(amazon_page, site: str = "US") -> dict[str, Any]:
    marketplace = get_marketplace(site)
    state = check_amazon_login_state(amazon_page)
    if state.get("status") == "logged-in":
        return state

    original_url = str(state.get("url") or amazon_page.url or "")
    if state.get("reason") == "signin-shell" or state.get("status") == "unknown":
        for probe_url in (marketplace.amazon_account_url, marketplace.amazon_base_url + "/"):
            try:
                amazon_page.goto(probe_url, wait_until="domcontentloaded", timeout=45000)
                amazon_page.wait_for_timeout(2500)
            except Exception:
                continue
            probe_state = check_amazon_login_state(amazon_page)
            if probe_state.get("status") == "logged-in":
                if original_url and url_belongs_to_amazon(original_url, (marketplace.amazon_domain,)) and original_url != amazon_page.url:
                    try:
                        amazon_page.goto(original_url, wait_until="domcontentloaded", timeout=45000)
                        amazon_page.wait_for_timeout(1500)
                    except Exception:
                        pass
                probe_state["validatedVia"] = probe_url
                return probe_state
            if probe_state.get("status") == "logged-out":
                state = probe_state
                break

    if state.get("status") == "logged-out":
        raise ManualActionRequired(
            "Amazon login appears expired. Please sign in to Amazon in the project Chrome browser, then rerun the workflow.",
            reason="login",
            stage="amazon-login",
        )
    return state


def extract_amazon_result(
    page,
    keyword: str,
    site: str = "US",
    amazon_postal_code: str | None = None,
) -> dict[str, Any]:
    amazon_page = None
    marketplace = get_marketplace(site)
    delivery_zip = str(amazon_postal_code or marketplace.postal_code or DEFAULT_AMAZON_DELIVERY_ZIP)
    try:
        amazon_page = page.context.new_page()
        url = marketplace.amazon_search_url(keyword)
        amazon_page.goto(url, wait_until="domcontentloaded", timeout=45000)
        login_state = require_amazon_login(amazon_page, marketplace.code)
        zip_result = ensure_amazon_postal_code(amazon_page, delivery_zip)
        if zip_result.get("status") in {"success", "already-set"}:
            amazon_page.goto(url, wait_until="domcontentloaded", timeout=45000)
            require_amazon_login(amazon_page, marketplace.code)
            if not (amazon_page.url or "").startswith(url):
                amazon_page.goto(url, wait_until="domcontentloaded", timeout=45000)
        else:
            return {
                "amazonResultStatus": "failed",
                "amazonResultSource": "direct-url",
                "amazonResultUrl": amazon_page.url,
                "amazonSite": marketplace.code,
                "amazonDomain": marketplace.amazon_domain,
                "amazonDeliveryZipStatus": zip_result.get("status"),
                "amazonDeliveryZip": delivery_zip,
                "amazonDeliveryZipError": zip_result.get("error") or "Amazon delivery postal code was not confirmed.",
                "amazonDeliveryLocationText": zip_result.get("text") or "",
                "amazonLoginStatus": login_state.get("status"),
            }
        amazon_page.wait_for_timeout(6000)
        body_text = amazon_page.evaluate("() => document.body ? document.body.innerText || '' : ''")
        parsed = extract_count_from_amazon_text(str(body_text))
        fields = {
            "amazonResultStatus": parsed.get("status"),
            "amazonResultSource": "direct-url",
            "amazonResultUrl": amazon_page.url,
            "amazonSite": marketplace.code,
            "amazonDomain": marketplace.amazon_domain,
            "amazonDeliveryZipStatus": zip_result.get("status"),
            "amazonDeliveryZip": delivery_zip,
            "amazonLoginStatus": login_state.get("status"),
        }
        if login_state.get("accountText"):
            fields["amazonAccountText"] = login_state.get("accountText")
        if zip_result.get("text"):
            fields["amazonDeliveryLocationText"] = zip_result.get("text")
        if zip_result.get("error"):
            fields["amazonDeliveryZipError"] = zip_result.get("error")
        if parsed.get("status") == "success":
            fields["amazonResultText"] = parsed.get("text")
            fields["amazonResultCount"] = parsed.get("count")
            fields["amazonResultIsApproximate"] = bool(parsed.get("isApproximate"))
        else:
            fields["amazonResultError"] = parsed.get("error") or "Amazon 结果数提取失败"
            fields["amazonResultDebugText"] = amazon_debug_excerpt(str(body_text))
        return fields
    except ManualActionRequired:
        raise
    except Exception as exc:
        return {"amazonResultStatus": "failed", "amazonResultError": str(exc), "amazonResultSource": "direct-url"}
    finally:
        if amazon_page is not None:
            try:
                amazon_page.close(run_before_unload=False)
            except Exception:
                pass
        try:
            page.bring_to_front()
        except Exception:
            pass
        sleep(page, 700)


def collect_trend_result(page) -> dict[str, Any]:
    open_first_row_monthly_drawer(page)
    click_search_trend_tab(page)
    scroll_drawer_top(page)
    configure_month_google_trend(page)
    zoom = set_zoom_to_2023(page)
    trend_monthly = extract_search_trend_monthly_data(page)
    base64_png = export_chart_base64(page)
    try:
        click_ppc_bid_tab(page)
        ppc = extract_ppc_data(page)
        if click_ppc_monthly_history(page):
            ppc = merge_ppc_data(ppc, extract_ppc_data(page))
    except Exception as exc:
        ppc = {"ppcStatus": "not_available", "ppcError": str(exc), "ppcMonthly": [], "ppcWeekly": []}
    return {"trendStatus": "success", "base64": base64_png, "zoom": zoom, "trendMonthly": trend_monthly, **ppc}


def write_chart_png(base64_png: str, output: Path) -> int:
    if "," in base64_png and base64_png.startswith("data:"):
        base64_png = base64_png.split(",", 1)[1]
    missing_padding = (-len(base64_png)) % 4
    if missing_padding:
        base64_png += "=" * missing_padding
    data = base64.b64decode(base64_png, validate=True)
    if not data.startswith(b"\x89PNG\r\n\x1a\n"):
        raise ValueError("解码结果不是 PNG 文件")
    output.write_bytes(data)
    return output.stat().st_size


def process_keyword(
    page,
    job: dict[str, Any],
    site: str = "US",
    amazon_postal_code: str | None = None,
) -> dict[str, Any]:
    keyword = str(job["keyword"])
    guard_manual_state(page, "before-keyword")
    close_drawer_if_open(page)
    safe_auto_handle_page_issue(page)
    search_keyword(page, keyword)
    trend_result: dict[str, Any] = {
        "trendStatus": "not_available",
        "trendError": "",
        "trendMonthly": [],
        "ppcMonthly": [],
        "ppcWeekly": [],
        "ppcStatus": "not_available",
        "ppcError": "",
        "base64": "",
    }
    try:
        try:
            trend_result = collect_trend_result(page)
        except Exception:
            close_drawer_if_open(page)
            search_keyword(page, keyword)
            trend_result = collect_trend_result(page)
    except Exception as exc:
        trend_result["trendError"] = str(exc)
        safe_auto_handle_page_issue(page)
        close_drawer_if_open(page)
    finally:
        close_drawer_if_open(page)
    amazon_result = extract_amazon_result(page, keyword, site=site, amazon_postal_code=amazon_postal_code)
    close_drawer_if_open(page)
    return {**job, "status": "success", **trend_result, **amazon_result}


def run_sellersprite_workflow(
    project_dir: Path,
    keywords: Sequence[str],
    prefix: str,
    data_output: Path,
    port: int = 9222,
    site: str = "US",
    amazon_postal_code: str | None = None,
) -> BrowserRunResult:
    sync_playwright = require_sync_playwright()
    marketplace = get_marketplace(site)
    amazon_postal_code = amazon_postal_code or marketplace.postal_code
    jobs = build_chart_jobs(project_dir, keywords, prefix)
    payload: dict[str, Any] = {"halted": False, "haltReason": "", "total": len(jobs), "successCount": 0, "results": []}
    logs: list[str] = []

    with sync_playwright() as p:
        browser, _context, page = connect_cdp(p, port)
        try:
            navigate_sellersprite_tool(page, SELLERSPRITE_URL, "关键词挖掘")
            sleep(page, 1200)
            ensure_sellersprite_site(page, marketplace.code)
            halted = False
            halt_reason = ""
            for job in jobs:
                if halted:
                    payload["results"].append({**job, "status": "failed", "stage": "halted", "error": f"前序关键词触发人工处理状态: {halt_reason}", "halted": True})
                    continue
                keyword = str(job["keyword"])
                try:
                    result = process_keyword(page, job, site=marketplace.code, amazon_postal_code=amazon_postal_code)
                    output = Path(str(job["output"]))
                    base64_png = str(result.get("base64") or "")
                    if base64_png:
                        size = write_chart_png(base64_png, output)
                        logs.append(json.dumps({"status": "success", "keyword": keyword, "output": str(output), "bytes": size}, ensure_ascii=False))
                    else:
                        logs.append(json.dumps({"status": "success", "keyword": keyword, "output": str(output), "bytes": 0, "noImage": True}, ensure_ascii=False))
                    result.pop("base64", None)
                    payload["results"].append(result)
                    print(f"完成: {keyword} -> {output}")
                except ManualActionRequired as exc:
                    halted = True
                    halt_reason = exc.reason
                    payload["results"].append({**job, "status": "failed", "stage": exc.stage, "error": str(exc), "halted": True, "reason": exc.reason})
                    print(f"停止: {keyword} ({exc.stage}) {exc}")
                except Exception as exc:
                    payload["results"].append({**job, "status": "failed", "stage": "keyword", "error": str(exc), "halted": False})
                    print(f"失败: {keyword} {exc}")
                    safe_auto_handle_page_issue(page)
                    close_drawer_if_open(page)
            payload["halted"] = halted
            payload["haltReason"] = halt_reason
            payload["successCount"] = sum(1 for item in payload["results"] if item.get("status") == "success")
        finally:
            disconnect_browser(browser)

    data_output.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return BrowserRunResult(returncode=1 if payload["halted"] else 0, stdout="\n".join(logs), stderr="", command=["python-playwright"])


def retry_failed_amazon_results(
    data_path: Path,
    port: int = 9222,
    site: str = "US",
    amazon_postal_code: str | None = None,
) -> int:
    if not data_path.exists():
        return 0
    try:
        payload = json.loads(data_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return 0
    results = payload.get("results", []) if isinstance(payload, dict) else []
    failed = [
        item for item in results
        if isinstance(item, dict)
        and item.get("amazonResultStatus") == "failed"
        and not str(item.get("amazonResultError", "")).startswith("blocked")
    ]
    if not failed:
        return 0

    fixed = 0
    attempted = 0
    marketplace = get_marketplace(site)
    amazon_postal_code = amazon_postal_code or marketplace.postal_code
    sync_playwright = require_sync_playwright()
    with sync_playwright() as p:
        browser, _context, page = connect_cdp(p, port)
        try:
            for item in failed:
                keyword = str(item.get("keyword", ""))
                print(f"  重试: {keyword} ...", end=" ")
                attempted += 1
                result = extract_amazon_result(page, keyword, site=marketplace.code, amazon_postal_code=amazon_postal_code)
                if result.get("amazonResultStatus") == "success":
                    item.update(result)
                    item.pop("amazonResultError", None)
                    item.pop("amazonResultDebugText", None)
                    fixed += 1
                    print(f"OK {int(item.get('amazonResultCount', 0)):,}")
                elif result.get("amazonResultStatus") == "blocked":
                    item["amazonResultError"] = f"blocked: {result.get('amazonResultError', '')}"
                    print("BLOCKED")
                else:
                    item.update(result)
                    print("FAIL")
        finally:
            disconnect_browser(browser)

    if attempted:
        payload["results"] = results
        data_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return fixed
