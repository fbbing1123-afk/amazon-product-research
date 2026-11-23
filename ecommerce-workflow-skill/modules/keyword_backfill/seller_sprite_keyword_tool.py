from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
import tempfile
import time
from copy import copy
from urllib.parse import unquote, urlparse
from urllib.request import Request, urlopen
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Font, PatternFill
from openpyxl.styles import Alignment
from openpyxl.worksheet.worksheet import Worksheet

MODULES_DIR = Path(__file__).resolve().parents[1]
if str(MODULES_DIR) not in sys.path:
    sys.path.insert(0, str(MODULES_DIR))

from shared.worksheet_rows import insert_rows_preserving_merges
from shared.marketplaces import ensure_sellersprite_site, get_marketplace, normalize_site_code


ASIN_RE = re.compile(r"^[A-Z0-9]{10}$", re.IGNORECASE)
ASIN_IN_TEXT_RE = re.compile(r"\b[A-Z0-9]{10}\b", re.IGNORECASE)
EXPORT_RE = re.compile(
    r"ExpandKeywords-(?P<site>[A-Z]{2})-(?P<first_asin>[A-Z0-9]{10})-batch\((?P<count>\d+)\)-",
    re.IGNORECASE,
)
KEYWORD_LINK_FONT = Font(name="宋体", sz=11, color="0000FF", underline="single")
CDP_CONNECT_TIMEOUT_MS = 20_000
PAGE_PROBE_TIMEOUT_MS = 2_000


@dataclass
class TemplateInfo:
    market_sheet: Worksheet
    product_name: str
    asins: list[str]
    keyword_header_row: int
    keyword_col: int


@dataclass
class KeywordRow:
    raw: list[Any]
    keyword: str
    translation: str
    related_products: int
    monthly_searches: int


def load_config(path: Path) -> dict[str, Any]:
    if not path.exists():
        raise FileNotFoundError(f"找不到配置文件：{path}")
    with path.open("r", encoding="utf-8-sig") as f:
        config = json.load(f)
    return config


def restore_storage_state(context: Any, page: Any, state_file: Path | None) -> None:
    if not state_file or not state_file.exists():
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
        if "sellersprite.com" not in origin_url or not isinstance(local_storage, list):
            continue
        try:
            page.goto(origin_url, wait_until="domcontentloaded", timeout=30_000)
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


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def parse_number(value: Any) -> int:
    if value is None:
        return 0
    if isinstance(value, (int, float)):
        return int(value)
    text = str(value).strip()
    if not text or text in {"-", "--"}:
        return 0
    text = re.sub(r"[,，\s]", "", text)
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    if not match:
        return 0
    return int(float(match.group(0)))


def sanitize_sheet_name(name: str) -> str:
    cleaned = re.sub(r"[:\\/?*\[\]]", " ", name).strip()
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned[:31] or "关键词数据导出"


def find_market_sheet(wb: Workbook) -> Worksheet:
    matches = [ws for ws in wb.worksheets if "市场调研" in ws.title]
    if not matches:
        raise ValueError("模板中没有找到名称包含“市场调研”的工作表")
    return matches[0]


def extract_product_name(sheet_name: str) -> str:
    product_name = sheet_name.replace("市场调研", "").strip()
    product_name = re.sub(r"[-_\s]+$", "", product_name)
    return product_name or "产品"


def find_competitor_asins(ws: Worksheet) -> list[str]:
    for row in ws.iter_rows():
        for cell in row:
            if is_competitor_asin_label(cell.value):
                asins: list[str] = []
                col = cell.column + 1
                while col <= ws.max_column:
                    value = ws.cell(row=cell.row, column=col).value
                    if not normalize_text(value):
                        break
                    asins.extend(extract_asins_from_text(value))
                    col += 1
                if not asins:
                    raise ValueError("找到了“竞对ASIN：”，但右侧没有读取到有效 ASIN")
                return unique_asins(asins)

    asins = find_asins_under_keyword_trend(ws)
    if asins:
        return asins
    raise ValueError("模板中没有找到“竞对ASIN：”单元格，也没有在“关键词搜索量及趋势”下一行读取到 ASIN")


def is_competitor_asin_label(value: Any) -> bool:
    text = normalize_text(value).upper().replace(" ", "").replace("：", ":")
    if not text or ASIN_RE.match(text):
        return False
    return "ASIN" in text and text.endswith(":")


def find_asins_under_keyword_trend(ws: Worksheet) -> list[str]:
    for row in range(1, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            if "关键词搜索量及趋势" not in normalize_text(ws.cell(row=row, column=col).value):
                continue
            scan_rows = range(row + 1, min(row + 4, ws.max_row + 1))
            for scan_row in scan_rows:
                row_text = " ".join(
                    normalize_text(ws.cell(row=scan_row, column=scan_col).value)
                    for scan_col in range(1, ws.max_column + 1)
                )
                asins = unique_asins(extract_asins_from_text(row_text))
                if asins:
                    return asins
    return []


def unique_asins(values: Iterable[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for value in values:
        asin = normalize_text(value).upper()
        if not ASIN_RE.match(asin) or asin in seen:
            continue
        seen.add(asin)
        result.append(asin)
    return result


def extract_asins_from_text(value: Any) -> list[str]:
    text = normalize_text(value).upper()
    return [match.group(0).upper() for match in ASIN_IN_TEXT_RE.finditer(text)]


def find_template_keyword_position(ws: Worksheet) -> tuple[int, int]:
    for row in ws.iter_rows():
        for cell in row:
            if normalize_text(cell.value) == "关键词":
                return cell.row, cell.column
    raise ValueError("市场调研页没有找到“关键词”表头")


def inspect_template(wb: Workbook) -> TemplateInfo:
    ws = find_market_sheet(wb)
    product_name = extract_product_name(ws.title)
    asins = find_competitor_asins(ws)
    keyword_header_row, keyword_col = find_template_keyword_position(ws)
    return TemplateInfo(ws, product_name, asins, keyword_header_row, keyword_col)


def import_export_workbook(export_path: Path) -> tuple[list[str], list[list[Any]]]:
    wb = load_workbook(export_path, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise ValueError(f"导出表为空：{export_path}")

    header_row_idx = find_export_header_row(rows)
    headers = [normalize_text(v) for v in rows[header_row_idx]]
    data_rows = [list(row[: len(headers)]) for row in rows[header_row_idx + 1 :] if any(v is not None for v in row)]
    return headers, data_rows


def find_export_header_row(rows: list[tuple[Any, ...]]) -> int:
    for idx, row in enumerate(rows[:30]):
        normalized = [normalize_text(v) for v in row]
        if "关键词" in normalized and find_header_index(normalized, ("相关产品数", "相关产品")) is not None and "月搜索量" in normalized:
            return idx
    raise ValueError("导出表前 30 行没有找到表头：关键词、相关产品/相关产品数、月搜索量")


def find_header_index(headers: list[str], candidates: tuple[str, ...]) -> int | None:
    for candidate in candidates:
        if candidate in headers:
            return headers.index(candidate)
    return None


def build_keyword_rows(
    headers: list[str],
    data_rows: list[list[Any]],
    brand_blacklist: Iterable[str],
    allowed_related_products: set[int] | None = None,
) -> list[KeywordRow]:
    keyword_idx = headers.index("关键词")
    related_idx = find_header_index(headers, ("相关产品数", "相关产品"))
    if related_idx is None:
        raise ValueError("导出表没有找到“相关产品”或“相关产品数”列")
    search_idx = headers.index("月搜索量")
    translation_idx = find_header_index(headers, ("关键词翻译", "翻译"))
    blacklist = [normalize_brand_text(b) for b in brand_blacklist if b and b.strip()]

    seen: set[str] = set()
    result: list[KeywordRow] = []
    for row in data_rows:
        keyword = normalize_text(row[keyword_idx] if keyword_idx < len(row) else "")
        translation = normalize_text(row[translation_idx] if translation_idx is not None and translation_idx < len(row) else "")
        # 卖家精灵在部分 Amazon Choice 词后追加 AC 标记，它不是搜索词正文。
        keyword = re.sub(r"\s+AC\s*$", "", keyword, flags=re.IGNORECASE).strip()
        keyword_key = keyword.lower()
        related_products = parse_number(row[related_idx] if related_idx < len(row) else None)
        monthly_searches = parse_number(row[search_idx] if search_idx < len(row) else None)
        if allowed_related_products is not None and related_products not in allowed_related_products:
            continue
        if not keyword or keyword_key in seen:
            continue
        seen.add(keyword_key)
        if monthly_searches <= 0:
            continue
        if is_brand_keyword(keyword_key, blacklist):
            continue
        result.append(
            KeywordRow(
                raw=row,
                keyword=keyword,
                translation=translation,
                related_products=related_products,
                monthly_searches=monthly_searches,
            )
        )
    return result


def normalize_brand_text(value: Any) -> str:
    text = normalize_text(value).lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def is_brand_keyword(keyword_key: str, blacklist: Iterable[str]) -> bool:
    keyword_words = normalize_brand_text(keyword_key)
    if not keyword_words:
        return False
    padded_keyword = f" {keyword_words} "
    tokens = set(keyword_words.split())
    for brand in blacklist:
        if not brand:
            continue
        if " " in brand:
            if f" {brand} " in padded_keyword:
                return True
        elif brand in tokens:
            return True
    return False


def sort_export_rows(headers: list[str], data_rows: list[list[Any]]) -> list[list[Any]]:
    related_idx = find_header_index(headers, ("相关产品数", "相关产品"))
    if related_idx is None:
        raise ValueError("导出表没有找到“相关产品”或“相关产品数”列")
    search_idx = headers.index("月搜索量")
    keyword_idx = headers.index("关键词")
    return sorted(
        data_rows,
        key=lambda row: (
            -parse_number(row[related_idx] if related_idx < len(row) else None),
            -parse_number(row[search_idx] if search_idx < len(row) else None),
            normalize_text(row[keyword_idx] if keyword_idx < len(row) else "").lower(),
        ),
    )


def top_related_product_values(headers: list[str], data_rows: list[list[Any]], count: int = 2) -> set[int]:
    related_idx = find_header_index(headers, ("相关产品数", "相关产品"))
    if related_idx is None:
        raise ValueError("导出表没有找到“相关产品”或“相关产品数”列")
    values = sorted(
        {parse_number(row[related_idx] if related_idx < len(row) else None) for row in data_rows},
        reverse=True,
    )
    return set(values[:count])


def move_monthly_search_next_to_related(headers: list[str], data_rows: list[list[Any]]) -> None:
    related_idx = find_header_index(headers, ("相关产品数", "相关产品"))
    if related_idx is None:
        raise ValueError("导出表没有找到“相关产品”或“相关产品数”列")
    search_idx = headers.index("月搜索量")
    target_idx = related_idx + 1
    if search_idx == target_idx:
        return
    if search_idx < target_idx:
        target_idx -= 1

    header = headers.pop(search_idx)
    headers.insert(target_idx, header)
    for row in data_rows:
        value = row.pop(search_idx) if search_idx < len(row) else None
        row.insert(target_idx, value)


def select_keywords(rows: list[KeywordRow]) -> list[KeywordRow]:
    ordered = sorted(rows, key=lambda r: (-r.related_products, -r.monthly_searches, r.keyword.lower()))
    top_related = sorted({r.related_products for r in ordered}, reverse=True)[:2]
    groups = {value: [r for r in ordered if r.related_products == value] for value in top_related}

    selected: list[KeywordRow] = []
    if len(top_related) >= 1:
        selected.extend(groups[top_related[0]][:5])
    if len(top_related) >= 2:
        selected.extend(groups[top_related[1]][:3])
    return selected


def keyword_display_text(item: KeywordRow) -> str:
    if item.translation:
        return f"{item.keyword}\n({item.translation})"
    return item.keyword


def amazon_search_url(keyword: str, site: str = "US") -> str:
    return get_marketplace(site).amazon_search_url(keyword)


def apply_keyword_link_style(cell) -> None:
    cell.font = copy(KEYWORD_LINK_FONT)


def refresh_export_sheet(
    wb: Workbook,
    sheet_name: str,
    headers: list[str],
    rows: list[list[Any]],
    selected: list[KeywordRow],
) -> None:
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    ws.append(headers)
    selected_row_ids = {id(item.raw) for item in selected}
    selected_fill = PatternFill(fill_type="solid", fgColor="FFFF00")
    for row in rows:
        ws.append(row)
        if id(row) in selected_row_ids:
            for cell in ws[ws.max_row]:
                cell.fill = selected_fill
    if ws.max_row > 1 and ws.max_column > 0:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
    for col_cells in ws.columns:
        max_len = max(len(normalize_text(cell.value)) for cell in col_cells)
        ws.column_dimensions[col_cells[0].column_letter].width = min(max(max_len + 2, 10), 40)


def write_selected_keywords(info: TemplateInfo, selected: list[KeywordRow], site: str = "US") -> None:
    ws = info.market_sheet
    start_row = info.keyword_header_row + 1
    col = info.keyword_col
    next_section_row = find_next_section_row(ws, start_row)
    available_rows = next_section_row - start_row
    if len(selected) > available_rows:
        rows_to_insert = len(selected) - available_rows
        template_row = max(start_row, next_section_row - 1)
        insert_rows_preserving_merges(ws, next_section_row, rows_to_insert, template_row)
        copy_row_format(ws, template_row, next_section_row, rows_to_insert)

    clear_until = start_row + max(len(selected), available_rows) - 1
    for row in range(start_row, clear_until + 1):
        ws.cell(row=row, column=col).value = None
    for offset, item in enumerate(selected):
        cell = ws.cell(row=start_row + offset, column=col)
        cell.value = keyword_display_text(item)
        cell.hyperlink = amazon_search_url(item.keyword, site)
        apply_keyword_link_style(cell)
        cell.alignment = copy(cell.alignment)
        cell.alignment = Alignment(
            horizontal=cell.alignment.horizontal,
            vertical=cell.alignment.vertical,
            text_rotation=cell.alignment.text_rotation,
            wrap_text=True,
            shrink_to_fit=cell.alignment.shrink_to_fit,
            indent=cell.alignment.indent,
        )


def find_next_section_row(ws: Worksheet, start_row: int) -> int:
    for row in range(start_row, ws.max_row + 1):
        value = normalize_text(ws.cell(row=row, column=1).value)
        if re.match(r"^\d+\.", value):
            return row
    return ws.max_row + 1


def copy_row_format(ws: Worksheet, source_row: int, target_start_row: int, row_count: int) -> None:
    max_col = max(ws.max_column, 1)
    source_height = ws.row_dimensions[source_row].height
    for row_offset in range(row_count):
        target_row = target_start_row + row_offset
        ws.row_dimensions[target_row].height = source_height
        for col in range(1, max_col + 1):
            source = ws.cell(row=source_row, column=col)
            target = ws.cell(row=target_row, column=col)
            if source.has_style:
                target._style = copy(source._style)
            if source.number_format:
                target.number_format = source.number_format
            if source.alignment:
                target.alignment = copy(source.alignment)
            if source.protection:
                target.protection = copy(source.protection)
            if not isinstance(target, MergedCell):
                target.value = None


def target_export_pattern(asins: list[str], site: str) -> re.Pattern[str]:
    first_asin = re.escape(asins[0])
    count = len(asins)
    return re.compile(rf"ExpandKeywords-{re.escape(site)}-{first_asin}-batch\({count}\)-", re.IGNORECASE)


def wait_for_download_complete(download_dir: Path, pattern: re.Pattern[str], start_time: float, timeout: int = 300) -> Path:
    deadline = time.time() + timeout
    while time.time() < deadline:
        candidates = []
        for path in download_dir.iterdir():
            if path.suffix.lower() not in {".xlsx", ".xls"}:
                continue
            if pattern.search(path.name) and path.stat().st_mtime >= start_time:
                candidates.append(path)
        partials = list(download_dir.glob("*.crdownload")) + list(download_dir.glob("*.tmp"))
        if candidates and not partials:
            return max(candidates, key=lambda p: p.stat().st_mtime)
        time.sleep(2)
    raise TimeoutError("等待导出文件下载超时")


def run_browser_flow(config: dict[str, Any], info: TemplateInfo) -> Path:
    try:
        from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
        from playwright.sync_api import sync_playwright
    except ImportError as exc:
        raise RuntimeError(
            "当前 Python 环境没有安装 playwright，无法执行网页自动化。"
            "可以先手动下载导出文件，再使用 --skip-browser --export-file 处理 Excel。"
        ) from exc

    expand_url = normalize_text(config.get("expand_keywords_url"))
    exports_url = normalize_text(config.get("my_exports_url"))
    if not expand_url or "请粘贴" in expand_url or not exports_url or "请粘贴" in exports_url:
        raise ValueError("请先在配置文件里填写 expand_keywords_url 和 my_exports_url")

    download_dir = Path(config["download_dir"]).expanduser().resolve()
    download_dir.mkdir(parents=True, exist_ok=True)

    debugger_address = normalize_text(config.get("browser", {}).get("debugger_address")) or "http://127.0.0.1:9222"
    if not debugger_address.startswith(("http://", "https://", "ws://", "wss://")):
        debugger_address = f"http://{debugger_address}"

    site = normalize_site_code(config.get("site", "US"))
    pattern = target_export_pattern(info.asins, site)
    asin_text = "\n".join(info.asins)

    with sync_playwright() as p:
        try:
            # CDP 连接必须有硬上限，避免浏览器渲染器无响应时卡住整个工作流。
            browser = p.chromium.connect_over_cdp(
                debugger_address,
                timeout=CDP_CONNECT_TIMEOUT_MS,
            )
        except Exception as exc:
            raise RuntimeError(
                "无法连接浏览器调试端口。请先关闭 Edge/Chrome 后，用 "
                "--remote-debugging-port=9222 打开已登录卖家精灵的浏览器。"
            ) from exc

        print("SellerSprite CDP connected", flush=True)

        context = browser.contexts[0] if browser.contexts else browser.new_context(accept_downloads=True)
        page = None
        for candidate in context.pages:
            try:
                if "sellersprite.com" in (candidate.url or "") and sellersprite_ready(candidate):
                    page = candidate
                    break
            except Exception:
                continue
        using_existing_logged_page = page is not None
        if page is None:
            page = context.new_page()
        storage_state = normalize_text(config.get("storage_state"))
        if not using_existing_logged_page:
            restore_storage_state(context, page, Path(storage_state).expanduser() if storage_state else None)
            # 恢复本地存储后先等页面导航结束，避免马上读取被销毁的执行上下文。
            try:
                page.wait_for_load_state("domcontentloaded", timeout=15_000)
            except Exception:
                pass
        print("SellerSprite checking login state", flush=True)
        ensure_sellersprite_logged_in(page, expand_url)
        # 先完成登录态恢复，再进入工具页，规避登录按钮触发刷新时的上下文失效。
        print("SellerSprite opening keyword tool", flush=True)
        navigate_sellersprite_tool(page, expand_url, "拓展流量词")
        print(f"SellerSprite switching site to {site}", flush=True)
        site_state = ensure_sellersprite_site(page, site)
        print(f"SellerSprite site: {site_state['site']} ({site_state['status']})")
        if dismiss_sellersprite_promotions(page):
            print("SellerSprite promotional dialog dismissed", flush=True)
        page.locator("input, textarea").first.wait_for(timeout=60_000)

        fields = page.locator("input, textarea")
        input_box = fields.nth(2) if fields.count() >= 3 else page.locator("input[placeholder*=ASIN], textarea[placeholder*=ASIN]").first
        print(f"SellerSprite filling {len(info.asins)} ASINs", flush=True)
        input_box.click()
        input_box.fill(asin_text, timeout=30_000)
        page.keyboard.press("Tab")
        page.wait_for_timeout(1000)

        print("SellerSprite submitting keyword query", flush=True)
        page.locator("button.el-button--primary", has_text="立即查询").last.click(timeout=10_000)
        current_variant = page.locator("button", has_text="用当前变体拓词").last
        print("SellerSprite selecting current variants", flush=True)
        current_variant.click(timeout=60_000)
        result_count = wait_for_keyword_results_ready(page, timeout=120)
        print(f"SellerSprite keyword results ready: {result_count}", flush=True)

        page.get_by_text("月搜索量").first.wait_for(timeout=90_000)
        export_button = page.locator("button", has_text=re.compile(r"^\s*导出\s*$")).last
        wait_for_enabled_locator(export_button, timeout=90)
        export_button.scroll_into_view_if_needed()
        export_button.click(timeout=30_000)
        page.wait_for_timeout(3000)

        body_text = page.locator("body").inner_text(timeout=15_000)
        if "免费用户" in body_text and "升级套餐" in body_text:
            raise PermissionError("当前卖家精灵账号是免费用户，页面提示升级套餐后才可导出。")

        navigate_sellersprite_exports(page, exports_url)
        download_url = wait_for_export_ready_playwright(page, pattern)
        downloaded = download_export_file(download_url, download_dir)
        if storage_state:
            try:
                Path(storage_state).expanduser().parent.mkdir(parents=True, exist_ok=True)
                context.storage_state(path=str(Path(storage_state).expanduser()))
            except Exception:
                pass
        # 远程 CDP 浏览器由 sync_playwright 上下文断开客户端连接，不能调用 close() 关闭项目浏览器。
        return downloaded


def sellersprite_ready(page: Any) -> bool:
    try:
        return bounded_page_probe(page, (
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
                '.user-info, .u-header__user, .u-header__user-avatar_nickname, .user-name, .right, button.profile-trigger'
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
              // 页面即使仍停留在关键词工具 URL，也可能已经退回游客态；游客态不能继续切站或提交查询。
              if (/\u672a\u767b\u5f55|\u6e38\u5ba2|\u7acb\u5373\u767b\u5f55|\u767b\u5f55\\s*\\/\\s*\u6ce8\u518c/.test(visibleText)) {
                return false;
              }
              const lowerUrl = location.href.toLowerCase();
              const hasKeywordInput = [...document.querySelectorAll(
                'input[placeholder*="keyword" i], input[placeholder*="\\u5173\\u952e\\u8bcd"], input[placeholder*="\\u8f93\\u5165"], textarea[placeholder*="ASIN" i]'
              )].some(isVisible);
              const onUsefulPage = lowerUrl.includes('/v2/welcome')
                || (lowerUrl.includes('/v2/export-log') && /\\u6211\\u7684\\u5bfc\\u51fa/.test(visibleText))
                || (lowerUrl.includes('/v3/keyword-miner') && hasKeywordInput)
                || (lowerUrl.includes('/v3/traffic/extend') && /\\u62d3\\u5c55\\u6d41\\u91cf\\u8bcd/.test(visibleText))
                || /\\u5356\\u5bb6\\u7cbe\\u7075\\u540e\\u53f0|\\u5173\\u952e\\u8bcd\\u6316\\u6398|Keyword Miner/i.test(visibleText);
              return onUsefulPage;
            }"""
        ))
    except Exception:
        return False


def dismiss_sellersprite_promotions(page: Any) -> bool:
    """关闭会遮挡 ASIN 输入框的卖家精灵营销弹窗。"""
    selectors = (
        ".kol-asin-signal-mask .close-btn",
        ".el-dialog__wrapper:visible .el-dialog__headerbtn",
        ".el-dialog__wrapper:visible .el-dialog__close",
        ".el-message-box__wrapper:visible .el-message-box__headerbtn",
    )
    dismissed = False
    for selector in selectors:
        candidates = page.locator(selector)
        for index in range(candidates.count()):
            candidate = candidates.nth(index)
            try:
                if not candidate.is_visible(timeout=500):
                    continue
                candidate.click(timeout=5_000)
                dismissed = True
                try:
                    page.locator(".kol-asin-signal-mask").first.wait_for(state="hidden", timeout=5_000)
                except Exception:
                    page.wait_for_timeout(800)
            except Exception:
                continue
    return dismissed


def sellersprite_has_real_login_challenge(page: Any) -> bool:
    try:
        if sellersprite_saved_credentials_available(page):
            return False
        return bounded_page_probe(page, (
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


def bounded_page_probe(page: Any, predicate: str, timeout_ms: int = PAGE_PROBE_TIMEOUT_MS) -> bool:
    """用 Playwright 自带超时执行页面探测，避免无界 page.evaluate 阻塞。"""
    try:
        page.wait_for_function(
            predicate,
            timeout=max(250, int(timeout_ms)),
            polling=100,
        )
        return True
    except Exception:
        return False


def click_sellersprite_login_shell(page: Any) -> bool:
    try:
        visitor_re = re.compile(r"未登录|游客")
        for selector in (
            "button.profile-trigger",
            ".user-info",
            ".right",
            ".user-name",
            "[class*='profile-trigger']",
        ):
            trigger = page.locator(selector).filter(has_text=visitor_re)
            if trigger.count() and trigger.first.is_visible(timeout=1_000):
                trigger.first.click(timeout=10_000)
                page.wait_for_timeout(900)
                break

        if not page.locator("text=立即登录").count() and not page.locator("text=登录").count():
            trigger = page.get_by_role("button", name=visitor_re)
            if trigger.count() and trigger.first.is_visible(timeout=1_000):
                trigger.first.click(timeout=10_000)
                page.wait_for_timeout(900)

        login_re = re.compile(r"^(立即登录|登录\s*/?\s*注册|登录|登陆|Sign in|Login)$", re.I)
        for selector in (
            "a.profile-menu-item",
            "a[href*='/w/user/login']",
            "a.text-link",
            "a",
            "button",
            ".el-button",
        ):
            candidates = page.locator(selector).filter(has_text=login_re)
            for index in range(candidates.count()):
                candidate = candidates.nth(index)
                if candidate.is_visible(timeout=500):
                    candidate.click(timeout=10_000)
                    return True
        return False
    except Exception:
        return False


def sellersprite_saved_credentials_available(page: Any) -> bool:
    """判断登录页是否已经由浏览器自动填入账号和密码。"""
    try:
        user_selectors = (
            'input[type="email"]',
            'input[name="email"]',
            'input[name="username"]',
            'input[placeholder*="手机号"]',
            'input[placeholder*="邮箱"]',
        )
        password_selector = 'input[type="password"]'

        def has_visible_value(selector: str) -> bool:
            locator = page.locator(selector)
            for index in range(locator.count()):
                field = locator.nth(index)
                if field.is_visible(timeout=300) and str(field.input_value(timeout=500) or '').strip():
                    return True
            return False

        return any(has_visible_value(selector) for selector in user_selectors) and has_visible_value(password_selector)
    except Exception:
        return False


def submit_sellersprite_saved_login(page: Any) -> bool:
    """点击已填充登录页中的提交按钮，并等待页面离开登录页。"""
    if "/w/user/login" not in (page.url or ""):
        return False
    if not sellersprite_saved_credentials_available(page):
        return False

    login_re = re.compile(r"^\s*立即登录\s*$", re.I)
    for selector in ("button#login_btn", "button[type='submit']", "button"):
        candidates = page.locator(selector).filter(has_text=login_re)
        for index in range(candidates.count()):
            candidate = candidates.nth(index)
            if not candidate.is_visible(timeout=500):
                continue
            candidate.click(timeout=10_000, no_wait_after=True)
            deadline = time.monotonic() + 20
            while time.monotonic() < deadline:
                if "/w/user/login" not in (page.url or ""):
                    return True
                page.wait_for_timeout(500)
            return True
    return False


def wait_for_sellersprite_ready(page: Any, timeout_ms: int = 15_000) -> bool:
    deadline = time.monotonic() + timeout_ms / 1000
    while time.monotonic() < deadline:
        if sellersprite_ready(page):
            return True
        page.wait_for_timeout(500)
    return sellersprite_ready(page)


def recover_sellersprite_login_shell(page: Any, return_url: str) -> bool:
    if sellersprite_ready(page):
        return True
    if submit_sellersprite_saved_login(page):
        try:
            page.goto(return_url, wait_until="domcontentloaded", timeout=60_000)
        except Exception:
            pass
        return wait_for_sellersprite_ready(page, timeout_ms=15_000)
    if sellersprite_has_real_login_challenge(page):
        return False

    for probe_url in ("https://www.sellersprite.com/v2/welcome", return_url):
        try:
            page.goto(probe_url, wait_until="domcontentloaded", timeout=60_000)
            page.wait_for_timeout(1500)
            if sellersprite_ready(page):
                if probe_url != return_url:
                    page.goto(return_url, wait_until="domcontentloaded", timeout=60_000)
                    page.wait_for_timeout(1200)
                if sellersprite_ready(page):
                    return True
        except Exception:
            pass

    if sellersprite_has_real_login_challenge(page):
        return False
    if click_sellersprite_login_shell(page):
        page.wait_for_timeout(1_200)
        submit_sellersprite_saved_login(page)
        try:
            page.goto(return_url, wait_until="domcontentloaded", timeout=60_000)
        except Exception:
            pass
        return wait_for_sellersprite_ready(page, timeout_ms=15_000)
    return False


def ensure_sellersprite_logged_in(page: Any, return_url: str) -> None:
    if recover_sellersprite_login_shell(page, return_url):
        return

    raise RuntimeError(
        "卖家精灵当前未能自动恢复登录态。请先在已打开的 9223 浏览器里登录卖家精灵，"
        "确认页面不再显示“未登录/游客”后再重新运行。"
    )


def navigate_sellersprite_tool(page: Any, target_url: str, link_text: str) -> None:
    """Prefer SellerSprite's visible in-app links so a live tab keeps its session state."""
    target_path = urlparse(target_url).path.rstrip("/")
    current_path = urlparse(page.url or "").path.rstrip("/")
    if current_path == target_path:
        return

    try:
        link = page.get_by_role("link", name=link_text, exact=True)
        if link.count() and link.first.is_visible(timeout=2_000):
            link.first.click(timeout=10_000)
            try:
                page.wait_for_load_state("domcontentloaded", timeout=15_000)
            except Exception:
                pass
            page.wait_for_timeout(1_200)
            if urlparse(page.url or "").path.rstrip("/") == target_path:
                return
    except Exception:
        pass

    page.goto(target_url, wait_until="domcontentloaded", timeout=60_000)


def navigate_sellersprite_exports(page: Any, exports_url: str) -> None:
    """Open the export history inside the authenticated SellerSprite tab."""
    target_path = urlparse(exports_url).path.rstrip("/")
    if urlparse(page.url or "").path.rstrip("/") == target_path:
        return

    try:
        view_exports = page.get_by_text("前往查看", exact=True)
        if view_exports.count() and view_exports.first.is_visible(timeout=1_000):
            view_exports.first.click(timeout=10_000)
            page.wait_for_timeout(1_200)
            if urlparse(page.url or "").path.rstrip("/") == target_path:
                return
    except Exception:
        pass

    try:
        dismiss = page.get_by_text("等会儿看", exact=True)
        if dismiss.count() and dismiss.first.is_visible(timeout=1_000):
            dismiss.first.click(timeout=5_000)
        user_menu = page.locator(".user-name, .user-info").first
        if user_menu.count() and user_menu.is_visible(timeout=2_000):
            user_menu.click(timeout=5_000)
            page.wait_for_timeout(400)
        export_link = page.locator("a[href*='/v2/export-log']:visible").first
        if export_link.count() and export_link.is_visible(timeout=2_000):
            export_link.click(timeout=10_000)
            page.wait_for_timeout(1_200)
            if urlparse(page.url or "").path.rstrip("/") == target_path:
                return
    except Exception:
        pass

    page.goto(exports_url, wait_until="domcontentloaded", timeout=60_000)


def wait_for_enabled_locator(locator: Any, timeout: int = 90) -> None:
    deadline = time.time() + timeout
    last_state = ""
    while time.time() < deadline:
        try:
            if locator.count() and locator.is_visible(timeout=1000):
                disabled = locator.get_attribute("disabled", timeout=1000)
                aria_disabled = locator.get_attribute("aria-disabled", timeout=1000)
                classes = locator.get_attribute("class", timeout=1000) or ""
                last_state = f"disabled={disabled!r}, aria-disabled={aria_disabled!r}, class={classes!r}"
                if disabled is None and aria_disabled != "true" and "is-disabled" not in classes:
                    return
        except Exception as exc:
            last_state = str(exc).splitlines()[0]
        time.sleep(2)
    raise TimeoutError(f"等待导出按钮可点击超时: {last_state}")


def wait_for_keyword_results_ready(page: Any, timeout: int = 120) -> int:
    """等待拓词结果真正加载完成，避免加载遮罩拦截导出按钮。"""
    deadline = time.time() + timeout
    last_state = ""
    while time.time() < deadline:
        try:
            body_text = (page.locator("body").inner_text(timeout=5_000) or "").replace("\n", " ")
            match = re.search(r"搜索结果数\s*[:：]\s*([0-9,]+)", body_text)
            result_count = int(match.group(1).replace(",", "")) if match else 0
            visible_masks = page.evaluate(
                """() => [...document.querySelectorAll('.el-loading-mask')]
                  .filter(node => node.offsetWidth > 0 || node.offsetHeight > 0).length"""
            )
            last_state = f"result_count={result_count}, visible_masks={visible_masks}"
            if result_count > 0 and visible_masks == 0:
                return result_count
            if "处理请求出现错误" in body_text and visible_masks == 0:
                last_state += ", page_error=处理请求出现错误"
        except Exception as exc:
            last_state = str(exc).splitlines()[0]
        time.sleep(1)
    raise TimeoutError(f"等待卖家精灵关键词结果加载超时: {last_state}")


def wait_for_export_ready_playwright(page: Any, pattern: re.Pattern[str], timeout: int = 600) -> str:
    deadline = time.time() + timeout
    poll_count = 0
    last_status = ""
    while time.time() < deadline:
        poll_count += 1
        page.wait_for_timeout(5000)
        href = find_matching_download_url(page, pattern)
        if href:
            print(f"导出任务已完成，发现下载链接（第 {poll_count} 次检查）", flush=True)
            return href

        rows = page.locator("tr, .el-table__row")
        matched_status = "未找到目标任务"
        for index in range(rows.count()):
            row = rows.nth(index)
            try:
                row_text = row.inner_text(timeout=1000)
            except Exception:
                continue
            if not pattern.search(row_text):
                continue
            matched_status = re.sub(r"\s+", " ", row_text).strip()[:180]
            if "已完成" in row_text:
                page.reload(wait_until="domcontentloaded")
                break

            refresh = row.locator(".lhy-refresh").first
            if refresh.count():
                refresh.click(timeout=3000)
            break
        else:
            page.reload(wait_until="domcontentloaded")
        if matched_status != last_status or poll_count % 3 == 0:
            remaining = max(0, int(deadline - time.time()))
            print(f"等待卖家精灵导出：{matched_status}；剩余上限约 {remaining}s", flush=True)
            last_status = matched_status
        time.sleep(5)
    raise TimeoutError("我的导出页面没有等到目标任务变成“已完成”")


def find_matching_download_url(page: Any, pattern: re.Pattern[str]) -> str | None:
    urls: list[str] = []
    for locator, attr in ((page.locator("input[data-item-url]"), "data-item-url"), (page.locator("a.download-excel[href]"), "href")):
        for index in range(locator.count()):
            href = first_attribute(locator.nth(index), attr)
            if href and pattern.search(unquote(href)):
                urls.append(href)
    return urls[0] if urls else None


def first_attribute(locator: Any, name: str) -> str | None:
    try:
        if locator.count() == 0:
            return None
        return locator.first.get_attribute(name, timeout=1000)
    except Exception:
        return None


def download_export_file(url: str, download_dir: Path) -> Path:
    download_dir.mkdir(parents=True, exist_ok=True)
    filename = Path(unquote(urlparse(url).path)).name
    if not filename:
        raise ValueError(f"无法从下载链接解析文件名：{url}")
    output = download_dir / filename
    request = Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urlopen(request, timeout=120) as response:
        output.write_bytes(response.read())
    return output


def process_export(config: dict[str, Any], export_path: Path, brand_blacklist: Iterable[str] | None = None) -> Path:
    template_path = Path(config["template_path"]).expanduser()
    output_path = Path(config["output_path"]).expanduser()
    if not template_path.exists():
        raise FileNotFoundError(f"模板文件不存在：{template_path}")
    if not export_path.exists():
        raise FileNotFoundError(f"导出文件不存在：{export_path}")

    wb = load_workbook(template_path, rich_text=True)
    info = inspect_template(wb)
    headers, data_rows = import_export_workbook(export_path)
    data_rows = sort_export_rows(headers, data_rows)
    candidate_related_values = top_related_product_values(headers, data_rows)
    keyword_rows = build_keyword_rows(
        headers,
        data_rows,
        brand_blacklist if brand_blacklist is not None else config.get("brand_blacklist", []),
        candidate_related_values,
    )
    selected = select_keywords(keyword_rows)
    move_monthly_search_next_to_related(headers, data_rows)

    sheet_name = "关键词导出"
    writer = normalize_text(config.get("excel_writer", "excel_com")).lower()
    if writer in {"openpyxl", "legacy"}:
        refresh_export_sheet(wb, sheet_name, headers, data_rows, selected)
        site = normalize_site_code(config.get("site", "US"))
        write_selected_keywords(info, selected, site=site)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        if output_path.resolve() == template_path.resolve():
            backup = template_path.with_suffix(".bak.xlsx")
            shutil.copy2(template_path, backup)
        wb.save(output_path)
    else:
        write_workbook_with_excel_com(
            template_path,
            output_path,
            info,
            sheet_name,
            headers,
            data_rows,
            selected,
            site=normalize_site_code(config.get("site", "US")),
        )

    print(f"产品名称: {info.product_name}")
    print(f"竞对 ASIN: {', '.join(info.asins)}")
    print(f"导出表: {export_path}")
    print(f"写入 sheet: {sheet_name}")
    print(f"入选关键词数量: {len(selected)}")
    print(f"输出文件: {output_path}")
    return output_path


def write_workbook_with_excel_com(
    template_path: Path,
    output_path: Path,
    info: TemplateInfo,
    sheet_name: str,
    headers: list[str],
    data_rows: list[list[Any]],
    selected: list[KeywordRow],
    site: str = "US",
) -> None:
    marketplace = get_marketplace(site)
    payload = {
        "template_path": str(template_path.resolve()),
        "output_path": str(output_path.resolve()),
        "market_sheet": info.market_sheet.title,
        "market_max_row": info.market_sheet.max_row,
        "market_max_col": info.market_sheet.max_column,
        "keyword_header_row": info.keyword_header_row,
        "keyword_col": info.keyword_col,
        "export_sheet": sheet_name,
        "headers": [json_safe_value(value) for value in headers],
        "rows": [[json_safe_value(value) for value in row] for row in data_rows],
        "selected_keywords": [keyword_display_text(item) for item in selected],
        "selected_keyword_terms": [item.keyword for item in selected],
        "amazon_base_url": marketplace.amazon_base_url,
    }
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="seller-sprite-excel-") as tmp_dir:
        tmp_path = Path(tmp_dir)
        payload_path = tmp_path / "payload.json"
        script_path = tmp_path / "write_excel.ps1"
        payload_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
        script_path.write_text(EXCEL_COM_WRITER_SCRIPT, encoding="utf-8-sig")
        result = subprocess.run(
            [
                "powershell.exe",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(script_path),
                str(payload_path),
            ],
            capture_output=True,
            text=True,
            encoding="utf-8",
        )
        if result.returncode != 0:
            raise RuntimeError(
                "Excel COM 写回失败。请确认输出文件没有被 Excel 打开，并且本机安装了 Excel。\n"
                f"STDOUT:\n{result.stdout}\nSTDERR:\n{result.stderr}"
            )


def json_safe_value(value: Any) -> Any:
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    return str(value)


EXCEL_COM_WRITER_SCRIPT = r'''
param([string]$PayloadPath)
$ErrorActionPreference = 'Stop'
$payload = Get-Content -LiteralPath $PayloadPath -Raw -Encoding UTF8 | ConvertFrom-Json

$xlOpenXMLWorkbook = 51
$xlPasteFormats = -4122
$xlYes = 1
$yellow = 65535

function Release-ComObject($obj) {
    if ($null -ne $obj) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj)
    }
}

function Get-RowMergeSpans($ws, [int]$row, [int]$lastCol) {
    $spans = @()
    $seen = @{}
    for ($col = 1; $col -le $lastCol; $col++) {
        $cell = $ws.Cells.Item($row, $col)
        if ($cell.MergeCells) {
            $area = $cell.MergeArea
            $address = [string]$area.Address($false, $false)
            if (-not $seen.ContainsKey($address) -and [int]$area.Row -eq $row -and [int]$area.Rows.Count -eq 1) {
                $seen[$address] = $true
                $spans += [pscustomobject]@{
                    FirstCol = [int]$area.Column
                    ColCount = [int]$area.Columns.Count
                }
            }
        }
    }
    return $spans
}

function Apply-RowMergeSpans($ws, [int]$row, $spans) {
    foreach ($span in @($spans)) {
        $firstCol = [int]$span.FirstCol
        $lastCol = $firstCol + [int]$span.ColCount - 1
        if ($lastCol -gt $firstCol) {
            $range = $ws.Range($ws.Cells.Item($row, $firstCol), $ws.Cells.Item($row, $lastCol))
            if (-not $range.MergeCells) {
                [void]$range.Merge()
            }
        }
    }
}

$excel = $null
$wb = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false

    $templatePath = [string]$payload.template_path
    $outputPath = [string]$payload.output_path
    $samePath = [System.String]::Equals(
        [System.IO.Path]::GetFullPath($templatePath),
        [System.IO.Path]::GetFullPath($outputPath),
        [System.StringComparison]::OrdinalIgnoreCase
    )
    if ($samePath) {
        $wb = $excel.Workbooks.Open($templatePath)
    } else {
        if (Test-Path -LiteralPath $outputPath) {
            Remove-Item -LiteralPath $outputPath -Force
        }
        $wb = $excel.Workbooks.Open($templatePath)
        $wb.SaveAs($outputPath, $xlOpenXMLWorkbook)
    }

    $ws = $wb.Worksheets.Item([string]$payload.market_sheet)
    $startRow = [int]$payload.keyword_header_row + 1
    $keywordCol = [int]$payload.keyword_col
    $lastRow = [Math]::Max([int]$payload.market_max_row, $startRow)
    $lastCol = [Math]::Max([int]$payload.market_max_col, $keywordCol)
    $nextSectionRow = $lastRow + 1
    for ($row = $startRow; $row -le $lastRow; $row++) {
        $value = [string]$ws.Cells.Item($row, 1).Text
        if ($value -match '^\d+\.') {
            $nextSectionRow = $row
            break
        }
    }

    $selectedKeywords = @($payload.selected_keywords)
    $selectedKeywordTerms = @($payload.selected_keyword_terms)
    $availableRows = $nextSectionRow - $startRow
    $originalColumnWidths = @{}
    for ($col = 1; $col -le $lastCol; $col++) {
        $originalColumnWidths[$col] = $ws.Columns.Item($col).ColumnWidth
    }
    $originalRowHeights = @{}
    for ($row = 1; $row -le $lastRow; $row++) {
        $originalRowHeights[$row] = $ws.Rows.Item($row).RowHeight
    }
    if ($selectedKeywords.Count -gt $availableRows) {
        $insertCount = $selectedKeywords.Count - $availableRows
        $templateRow = [Math]::Max($startRow, $nextSectionRow - 1)
        $templateRowHeight = $ws.Rows.Item($templateRow).RowHeight
        $templateMergeSpans = Get-RowMergeSpans $ws $templateRow $lastCol
        $targetRows = "$nextSectionRow`:$($nextSectionRow + $insertCount - 1)"
        $ws.Rows.Item($targetRows).Insert()
        $ws.Rows.Item($templateRow).Copy()
        $ws.Rows.Item($targetRows).PasteSpecial($xlPasteFormats)
        $excel.CutCopyMode = $false
        for ($row = $nextSectionRow; $row -lt $nextSectionRow + $insertCount; $row++) {
            $ws.Rows.Item($row).RowHeight = $templateRowHeight
            Apply-RowMergeSpans $ws $row $templateMergeSpans
        }
        for ($row = $nextSectionRow; $row -le $lastRow; $row++) {
            $ws.Rows.Item($row + $insertCount).RowHeight = $originalRowHeights[$row]
        }
        for ($col = 1; $col -le $lastCol; $col++) {
            $ws.Columns.Item($col).ColumnWidth = $originalColumnWidths[$col]
        }
    }

    $clearRows = [Math]::Max($selectedKeywords.Count, $availableRows)
    for ($offset = 0; $offset -lt $clearRows; $offset++) {
        $ws.Cells.Item($startRow + $offset, $keywordCol).ClearContents()
    }
    for ($index = 0; $index -lt $selectedKeywords.Count; $index++) {
        $keywordCell = $ws.Cells.Item($startRow + $index, $keywordCol)
        $keywordCell.Value2 = [string]$selectedKeywords[$index]
        $keywordCell.WrapText = $true
        $keywordCell.Font.Name = '宋体'
        $keywordCell.Font.Size = 11
        $keywordCell.Font.Color = 16711680
        $keywordCell.Font.Underline = 2
        if ($index -lt $selectedKeywordTerms.Count) {
            $keyword = [string]$selectedKeywordTerms[$index]
            $address = [string]$payload.amazon_base_url + '/s?k=' + [System.Uri]::EscapeDataString($keyword)
            [void]$ws.Hyperlinks.Add($keywordCell, $address)
        }
    }

    foreach ($sheet in @($wb.Worksheets)) {
        if ($sheet.Name -eq [string]$payload.export_sheet) {
            $sheet.Delete()
            break
        }
    }
    $exportSheet = $wb.Worksheets.Add([System.Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count))
    $exportSheet.Name = [string]$payload.export_sheet

    $headers = @($payload.headers)
    $rows = @($payload.rows)
    $rowCount = $rows.Count + 1
    $colCount = $headers.Count
    if ($colCount -gt 0) {
        $matrix = New-Object 'object[,]' $rowCount, $colCount
        for ($col = 0; $col -lt $colCount; $col++) {
            $matrix[0, [int]$col] = $headers[$col]
        }
        for ($dataIndex = 0; $dataIndex -lt $rows.Count; $dataIndex++) {
            $rowValues = @($rows[$dataIndex])
            for ($col = 0; $col -lt $colCount; $col++) {
                if ($col -lt $rowValues.Count) {
                    $matrix[[int]($dataIndex + 1), [int]$col] = $rowValues[$col]
                }
            }
        }
        $target = $exportSheet.Range($exportSheet.Cells.Item(1, 1), $exportSheet.Cells.Item($rowCount, $colCount))
        $target.Value2 = $matrix
        $exportSheet.Rows.Item(1).Font.Bold = $true
        if ($rowCount -gt 1) {
            [void]$target.AutoFilter(1)
        }

        $keywordCol = 1
        for ($col = 1; $col -le $colCount; $col++) {
            if ([string]$exportSheet.Cells.Item(1, $col).Value2 -eq '关键词') {
                $keywordCol = $col
                break
            }
        }
        $selectedSet = @{}
        foreach ($keyword in @($payload.selected_keyword_terms)) {
            $selectedSet[[string]$keyword] = $true
        }
        for ($row = 2; $row -le $rowCount; $row++) {
            $keyword = [string]$exportSheet.Cells.Item($row, $keywordCol).Value2
            if ($selectedSet.ContainsKey($keyword)) {
                $exportSheet.Range($exportSheet.Cells.Item($row, 1), $exportSheet.Cells.Item($row, $colCount)).Interior.Color = $yellow
            }
        }
    }

    $wb.Save()
}
finally {
    if ($null -ne $wb) {
        $wb.Close($true)
    }
    if ($null -ne $excel) {
        $excel.Quit()
    }
    Release-ComObject $wb
    Release-ComObject $excel
}
'''


def main() -> int:
    parser = argparse.ArgumentParser(description="卖家精灵拓展流量词自动化与模板写回")
    parser.add_argument("--config", default="seller_sprite_config.json", help="配置文件路径")
    parser.add_argument("--export-file", help="已下载的卖家精灵导出 Excel")
    parser.add_argument("--skip-browser", action="store_true", help="跳过网页自动化，只处理 --export-file")
    parser.add_argument("--brand-word", action="append", default=None, help="本次临时剔除的品牌词，可重复传入")
    args = parser.parse_args()

    config = load_config(Path(args.config))
    if args.skip_browser:
        if not args.export_file:
            raise ValueError("--skip-browser 模式必须提供 --export-file")
        export_path = Path(args.export_file).expanduser()
    else:
        wb = load_workbook(Path(config["template_path"]), rich_text=True)
        info = inspect_template(wb)
        export_path = run_browser_flow(config, info)

    process_export(config, export_path, args.brand_word)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"错误: {exc}", file=sys.stderr)
        raise SystemExit(1)
