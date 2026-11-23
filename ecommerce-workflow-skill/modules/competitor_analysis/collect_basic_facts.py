"""批量读取竞品三项基础字段：数据源优先，缺失时读取美国站卖家精灵插件。"""

from __future__ import annotations

import argparse
import json
import math
import re
import sys
import time
from datetime import date, datetime, timezone
from pathlib import Path
from urllib.parse import urlparse

from openpyxl import load_workbook


FIELDS = {"monthly_sales": "月销量", "price": "价格($)", "listed_at": "上架时间"}
PANEL = "#seller-sprite-extension-quick-view-listing"
SCHEMA = "competitor-basic-facts.v1"
MISSING = {"", "/", "-", "--", "n/a", "none", "null", "暂无", "无数据"}


def asin_from_input(value: str) -> str:
    """接受一个 ASIN 或美国站商品链接，拒绝其它站点和非商品地址。"""
    value = value.strip()
    if re.fullmatch(r"[A-Z0-9]{10}", value, re.I):
        return value.upper()
    url = urlparse(value)
    if url.scheme not in {"http", "https"} or url.hostname not in {"amazon.com", "www.amazon.com"}:
        raise ValueError(f"请提供美国站商品链接或十位 ASIN：{value}")
    match = re.search(r"/(?:dp|gp/product)/([A-Z0-9]{10})(?:/|$)", url.path, re.I)
    if not match:
        raise ValueError(f"商品链接中没有找到 ASIN：{value}")
    return match[1].upper()


def normalize_value(field: str, value):
    """零销量是有效值；范围销量、无效价格和缺失标记不能当作精确数值。"""
    if value is None or isinstance(value, bool) or str(value).strip().lower() in MISSING:
        return None
    if field == "listed_at":
        if isinstance(value, (date, datetime)):
            return value.strftime("%Y-%m-%d")
        match = re.fullmatch(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:\s+00:00:00)?", str(value).strip())
        if match:
            try:
                return date(*map(int, match.groups())).isoformat()
            except ValueError:
                pass
        return None
    text = str(value).strip().replace(",", "")
    if field == "price":
        text = re.sub(r"^(?:US\$|\$)\s*", "", text)
    if not re.fullmatch(r"\d+(?:\.\d+)?", text):
        return None
    number = float(text)
    if not math.isfinite(number):
        return None
    if field == "monthly_sales":
        return int(number) if number.is_integer() and number >= 0 else None
    return number if number > 0 else None


def parse_plugin_panel(text: str, expected_asin: str, page_url: str) -> dict:
    """只解析当前商品插件区域，避免误取推荐商品、均价或子体销量。"""
    if asin_from_input(page_url) != expected_asin:
        raise ValueError("页面已跳转到其它 ASIN，拒绝串商品取数")
    identifiers = set(re.findall(r"\bASIN\s*[:：]\s*([A-Z0-9]{10})\b", text, re.I))
    if {value.upper() for value in identifiers} != {expected_asin}:
        raise ValueError("插件 ASIN 缺失或与目标不一致，可能尚未加载完成")
    patterns = {
        "monthly_sales": r"近\s*30\s*天销量\s*[（(]父体[）)]\s*[:：]\s*([^\s]+)",
        "price": r"(?:^|\n)\s*价格\s*[:：]\s*((?:US)?\$?\s*[\d,.]+)",
        "listed_at": r"上架时间\s*[:：]\s*(\d{4}[-/]\d{1,2}[-/]\d{1,2})",
    }
    labels = {"monthly_sales": "近30天销量(父体)", "price": "价格", "listed_at": "上架时间"}
    values = {}
    for field, pattern in patterns.items():
        match = re.search(pattern, text)
        raw = match[1].strip() if match else None
        values[field] = {"value": normalize_value(field, raw), "label": labels[field], "raw": raw}
    return {"asin": expected_asin, "url": page_url, "fields": values,
            "observed_at": datetime.now(timezone.utc).isoformat()}


def read_source(path: Path, asins: list[str], sheet_name: str | None = None) -> dict:
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        sheets = [sheet for sheet in wb if (sheet.title == sheet_name if sheet_name else sheet.title.endswith("数据源"))]
        if len(sheets) != 1:
            raise ValueError("数据源工作表不唯一或不存在，请用 --source-sheet 指定")
        sheet = sheets[0]
        rows = sheet.iter_rows(values_only=True)
        header = [str(value or "").strip() for value in next(rows)]
        if header.count("ASIN") != 1:
            raise ValueError("数据源必须包含唯一 ASIN 列")
        for label in FIELDS.values():
            if header.count(label) > 1:
                raise ValueError(f"数据源字段重复：{label}")
        result = {asin: {field: [] for field in FIELDS} for asin in asins}
        for row_number, row in enumerate(rows, 2):
            record = dict(zip(header, row))
            asin = str(record.get("ASIN") or "").strip().upper()
            if asin not in result:
                continue
            for field, label in FIELDS.items():
                value = normalize_value(field, record.get(label))
                if value is not None:
                    result[asin][field].append({"value": value, "source": "workbook", "label": label,
                                               "sheet": sheet.title, "row": row_number})
        return result
    finally:
        wb.close()


def resolve_product(asin: str, candidates: dict, fetch_page=None) -> dict:
    result = {"asin": asin, "fields": {}, "errors": []}
    for field, label in FIELDS.items():
        values = candidates.get(field, [])
        distinct = {entry["value"] for entry in values}
        if len(distinct) > 1:
            result["fields"][field] = {"value": None, "source": "conflict", "label": label}
            result["errors"].append(f"数据源同一 ASIN 的{label}存在冲突，需核对")
        elif values:
            result["fields"][field] = dict(values[0])
        else:
            result["fields"][field] = {"value": None, "source": "missing", "label": label}
    needed = [field for field in FIELDS if result["fields"][field]["source"] == "missing"]
    if needed and fetch_page:
        try:
            observation = fetch_page(asin, needed)
            result["page"] = {key: observation[key] for key in ("url", "observed_at")}
            for field in needed:
                item = observation["fields"][field]
                result["fields"][field] = {**item, "source": "sellersprite_plugin" if item["value"] is not None else "missing"}
        except Exception as error:
            result["errors"].append(f"网页补取失败：{error}")
    result["missing_fields"] = [FIELDS[field] for field in FIELDS if result["fields"][field]["value"] is None]
    result["status"] = "complete" if not result["missing_fields"] else "incomplete"
    return result


class PluginCollector:
    """按需连接已登录项目浏览器；仅关闭本脚本创建的页面。"""

    def __init__(self, endpoint: str, timeout: float = 25):
        self.endpoint = endpoint
        self.timeout = timeout
        self.runtime = None
        self.page = None

    def __enter__(self):
        return self

    def __exit__(self, *args):
        try:
            if self.page is not None:
                self.page.close()
        finally:
            if self.runtime is not None:
                self.runtime.stop()

    def __call__(self, asin: str, needed: list[str]):
        if self.page is None:
            from playwright.sync_api import sync_playwright

            if self.runtime is None:
                self.runtime = sync_playwright().start()
            browser = self.runtime.chromium.connect_over_cdp(self.endpoint, timeout=20_000)
            if not browser.contexts:
                raise RuntimeError("项目浏览器没有已登录上下文，请先打开项目浏览器")
            self.page = browser.contexts[0].new_page()
        page = self.page
        page.goto(f"https://www.amazon.com/dp/{asin}", wait_until="domcontentloaded", timeout=40_000)
        deadline = time.monotonic() + self.timeout
        last_error = "卖家精灵商品插件尚未出现，请检查插件是否启用及登录"
        latest = None
        previous = None
        while time.monotonic() < deadline:
            if "/ap/signin" in page.url or "/errors/validateCaptcha" in page.url:
                raise RuntimeError("页面要求登录或验证码，请人工处理后重试")
            panel = page.locator(PANEL)
            if panel.count() == 1 and panel.is_visible():
                try:
                    latest = parse_plugin_panel(panel.inner_text(timeout=3000), asin, page.url)
                    signature = [latest["fields"][field]["value"] for field in needed]
                    # 两次连续读取一致且所需字段齐全后返回，避免插件刚出现时仍是占位值。
                    if all(value is not None for value in signature) and signature == previous:
                        return latest
                    previous = signature
                except ValueError as error:
                    last_error = str(error)
                    latest = None
                    previous = None
            page.wait_for_timeout(500)
        if latest is not None:
            return latest
        raise RuntimeError(last_error)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("items", nargs="+", help="一个或多个 ASIN / 美国站商品链接")
    parser.add_argument("--workbook", type=Path, required=True)
    parser.add_argument("--source-sheet")
    parser.add_argument("--output", type=Path, required=True, help="取数结果 JSON，不修改 Excel")
    parser.add_argument("--cdp", default="http://127.0.0.1:9223")
    parser.add_argument("--timeout", type=float, default=25)
    parser.add_argument("--source-only", action="store_true", help="仅检查 Excel，不打开网页")
    args = parser.parse_args()
    if not 1 <= args.timeout <= 60:
        parser.error("--timeout 必须在 1～60 秒之间")
    if args.output.resolve() == args.workbook.resolve() or args.output.suffix.lower() != ".json":
        parser.error("--output 必须是独立的 .json 文件，不能覆盖工作簿")
    asins = list(dict.fromkeys(asin_from_input(item) for item in args.items))
    source = read_source(args.workbook, asins, args.source_sheet)
    report = {"schema_version": SCHEMA, "site": "US", "currency": "USD", "products": []}
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with PluginCollector(args.cdp, args.timeout) as collector:
        for asin in asins:
            print(f"正在处理 {asin}", flush=True)
            result = resolve_product(asin, source[asin], None if args.source_only else collector)
            report["products"].append(result)
            # 每完成一个商品即保存检查点；部分失败不丢失其它商品的取数结果。
            temporary = args.output.with_suffix(".json.tmp")
            temporary.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
            temporary.replace(args.output)
            print(json.dumps(result, ensure_ascii=False), flush=True)
    return 0 if all(item["status"] == "complete" for item in report["products"]) else 2


if __name__ == "__main__":
    sys.exit(main())
