"""批量读取竞品基础字段和卖点：数据源优先，缺失时读取美国站商品网页。"""

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
ALL_FIELDS = {**FIELDS, "selling_points": "产品卖点"}
AVAILABLE_FIELDS = {**ALL_FIELDS, "alexa_questions": "Alexa推荐问题"}
PANEL = "#seller-sprite-extension-quick-view-listing"
SCHEMA = "competitor-basic-facts.v1"
MISSING = {"", "/", "-", "--", "n/a", "none", "null", "暂无", "无数据"}


sys.path.insert(0, str((Path(__file__).resolve().parent / "..").resolve()))
from shared.browser_runtime import get_cdp_url
from shared.amazon_access import AmazonAccessRequired, ensure_amazon_access, goto_amazon, pending_amazon_access_page

def normalize_selling_points(value) -> str | None:
    """保留原始卖点内容，只清理空项和重复条目，不强行凑足或截断五条。"""
    if value is None or isinstance(value, (bool, int, float)):
        return None
    items = value if isinstance(value, (list, tuple)) else str(value).splitlines()
    cleaned = []
    for item in items:
        if not isinstance(item, str):
            continue
        text = re.sub(r"\s+", " ", item).strip()
        if text.lower() in MISSING or text in {"›", "See more product details", "About this item"}:
            continue
        if text and text not in cleaned:
            cleaned.append(text)
    return "\n".join(cleaned) or None


def read_page_selling_points(page, asin: str) -> dict:
    """只读取当前商品 About this item 的条目，不混入 A+、推荐商品或评论。"""
    if asin_from_input(page.url) != asin:
        raise ValueError("页面已跳转到其它 ASIN，拒绝串商品读取卖点")
    root = page.locator("#feature-bullets")
    value = None
    if root.count() == 1 and root.is_visible():
        # 此区域包含五点列表；展开折叠项后读取页面原文，部分商品实际多于五条。
        expand = page.get_by_role("button", name="See more About this item", exact=True)
        if expand.count() == 1 and expand.is_visible() and expand.get_attribute("aria-expanded") != "true":
            expand.click(timeout=3000)
        if re.search(r"About this item", root.inner_text(timeout=3000), re.I):
            value = normalize_selling_points(root.locator("li").all_inner_texts())
    if asin_from_input(page.url) != asin:
        raise ValueError("读取卖点时商品发生变化，请重试")
    return {"value": value, "raw": value, "label": "About this item", "source": "amazon_product_page"}


def normalize_alexa_questions(value) -> str | None:
    """只保留问题按钮文案；这些是推荐问题，不能解释为真实买家提问频次。"""
    if not isinstance(value, (str, list, tuple)):
        return None
    items = value.splitlines() if isinstance(value, str) else value
    questions, seen = [], set()
    for item in items:
        if not isinstance(item, str):
            continue
        question = re.sub(r"\s+", " ", item).strip()
        if not question.endswith(("?", "？")) or question.casefold() in seen:
            continue
        questions.append(question)
        seen.add(question.casefold())
    return "\n".join(questions) or None


def read_page_alexa_questions(page, asin: str) -> dict:
    """读取商品页 Looking for specific info 区域的现有推荐按钮，不向 Alexa 发送提问。"""
    if asin_from_input(page.url) != asin:
        raise ValueError("页面已跳转到其它 ASIN，拒绝串商品读取推荐问题")
    root = page.locator("#dpx-rex-nice-widget-container")
    value = None
    if root.count() == 1 and root.is_visible():
        if re.search(r"Looking for specific info", root.inner_text(timeout=3000), re.I):
            # Amazon用 input+span 渲染问题按钮，读取推荐列表文案而非仅匹配button标签。
            pills = root.locator("#dpx-rex-nile-inline-default-pills-container li")
            value = normalize_alexa_questions(pills.all_inner_texts())
    if asin_from_input(page.url) != asin:
        raise ValueError("读取推荐问题时商品发生变化，请重试")
    return {"value": value, "raw": value, "label": "Looking for specific info / Alexa推荐问题",
            "source": "amazon_alexa_suggestions"}


def aggregate_question_evidence(products: list[dict]) -> list[dict]:
    """按原问题归并展示过它的商品；近义归并留给有上下文的分析阶段。"""
    groups = {}
    for product in products:
        value = product.get("fields", {}).get("alexa_questions", {}).get("value")
        for question in (normalize_alexa_questions(value) or "").splitlines():
            key = question.casefold()
            entry = groups.setdefault(key, {"question": question, "asins": []})
            if product["asin"] not in entry["asins"]:
                entry["asins"].append(product["asin"])
    return [{**entry, "product_page_count": len(entry["asins"])} for entry in groups.values()]


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
    if field == "selling_points":
        return normalize_selling_points(value)
    if field == "alexa_questions":
        return normalize_alexa_questions(value)
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


def read_source(path: Path, asins: list[str], sheet_name: str | None = None, fields=None) -> dict:
    fields = FIELDS if fields is None else fields
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        sheets = [sheet for sheet in wb if (sheet.title.strip() == sheet_name.strip() if sheet_name else (sheet.title.strip().endswith("数据源") and not any(word in sheet.title for word in ("评价", "评论"))))]
        if len(sheets) != 1:
            raise ValueError("数据源工作表不唯一或不存在，请用 --source-sheet 指定")
        sheet = sheets[0]
        rows = sheet.iter_rows(values_only=True)
        header = [str(value or "").strip() for value in next(rows)]
        if header.count("ASIN") != 1:
            raise ValueError("数据源必须包含唯一 ASIN 列")
        for label in fields.values():
            if header.count(label) > 1:
                raise ValueError(f"数据源字段重复：{label}")
        result = {asin: {field: [] for field in fields} for asin in asins}
        for row_number, row in enumerate(rows, 2):
            record = dict(zip(header, row))
            asin = str(record.get("ASIN") or "").strip().upper()
            if asin not in result:
                continue
            for field, label in fields.items():
                value = normalize_value(field, record.get(label))
                if value is not None:
                    result[asin][field].append({"value": value, "source": "workbook", "label": label,
                                               "sheet": sheet.title, "row": row_number})
        return result
    finally:
        wb.close()


def resolve_product(asin: str, candidates: dict, fetch_page=None, fields=None) -> dict:
    fields = FIELDS if fields is None else fields
    result = {"asin": asin, "fields": {}, "errors": []}
    for field, label in fields.items():
        values = candidates.get(field, [])
        distinct = {entry["value"] for entry in values}
        if len(distinct) > 1:
            result["fields"][field] = {"value": None, "source": "conflict", "label": label}
            result["errors"].append(f"数据源同一 ASIN 的{label}存在冲突，需核对")
        elif values:
            result["fields"][field] = dict(values[0])
        else:
            result["fields"][field] = {"value": None, "source": "missing", "label": label}
    needed = [field for field in fields if result["fields"][field]["source"] == "missing"]
    if needed and fetch_page:
        try:
            observation = fetch_page(asin, needed)
            result["page"] = {key: observation[key] for key in ("url", "observed_at")}
            if observation.get("warning"):
                result["errors"].append(observation["warning"])
            for field in needed:
                item = observation["fields"][field]
                result["fields"][field] = {**item, "source": item.get("source", "sellersprite_plugin") if item["value"] is not None else "missing"}
        except Exception as error:
            result["errors"].append(f"网页补取失败：{error}")
            if isinstance(error, AmazonAccessRequired):
                result['access_paused'] = True
    result["missing_fields"] = [fields[field] for field in fields if result["fields"][field]["value"] is None]
    result["status"] = "complete" if not result["missing_fields"] else "incomplete"
    if result.get('access_paused'):
        result['status'] = 'needs_user_action'
    return result


class PluginCollector:
    """按需连接已登录项目浏览器；仅关闭本脚本创建的页面。"""

    def __init__(self, endpoint: str, timeout: float = 25, postal_code='10001'):
        self.endpoint = endpoint
        self.timeout = timeout
        self.postal_code = postal_code
        self.runtime = None
        self.page = None
        self.access_paused = False

    def __enter__(self):
        return self

    def __exit__(self, *args):
        try:
            if self.page is not None and not self.access_paused:
                self.page.close()
        finally:
            if self.runtime is not None:
                self.runtime.stop()

    def __call__(self, asin: str, needed: list[str]):
        if self.access_paused:
            raise AmazonAccessRequired('已有访问受阻页面，暂停新的竞品取数')
        if self.page is None:
            from playwright.sync_api import sync_playwright

            if self.runtime is None:
                self.runtime = sync_playwright().start()
            browser = self.runtime.chromium.connect_over_cdp(self.endpoint, timeout=20_000)
            if not browser.contexts:
                raise RuntimeError("项目浏览器没有已登录上下文，请先打开项目浏览器")
            context = browser.contexts[0]
            pending = pending_amazon_access_page(context)
            if pending is not None:
                try:
                    ensure_amazon_access(pending)
                except AmazonAccessRequired:
                    self.access_paused = True
                    raise
            self.page = context.new_page()
        page = self.page
        target_url = f'https://www.amazon.com/dp/{asin}'
        try:
            goto_amazon(page, target_url, timeout=40000)
            location = page.locator('#glow-ingress-line2').inner_text(timeout=5000)
            if self.postal_code not in location:
                raise AmazonAccessRequired(f'竞品页配送邮编不符：{location}')
        except AmazonAccessRequired:
            self.access_paused = True
            raise
        deadline = time.monotonic() + self.timeout
        last_error = "请检查所选商品字段/推荐问题是否展示，以及所需插件是否启用并登录"
        latest = None
        previous = None
        while time.monotonic() < deadline:
            try:
                access = ensure_amazon_access(page, expected_url=target_url)
                if access['status'] == 'cleared':
                    location = page.locator('#glow-ingress-line2').inner_text(timeout=5000)
                    if self.postal_code not in location:
                        raise AmazonAccessRequired(f'确认恢复后配送邮编不符：{location}')
            except AmazonAccessRequired:
                self.access_paused = True
                raise
            try:
                if asin_from_input(page.url) != asin:
                    raise ValueError("页面已跳转到其它 ASIN")
                latest = {"asin": asin, "url": page.url, "observed_at": datetime.now(timezone.utc).isoformat(),
                          "fields": {field: {"value": None, "label": AVAILABLE_FIELDS[field], "raw": None} for field in needed}}
                if any(field in FIELDS for field in needed):
                    panel = page.locator(PANEL)
                    if panel.count() == 1 and panel.is_visible():
                        try:
                            plugin = parse_plugin_panel(panel.inner_text(timeout=3000), asin, page.url)
                            latest["fields"].update(plugin["fields"])
                        except ValueError as error:
                            last_error = str(error)
                # 卖点直接来自 Amazon，不依赖卖家精灵插件是否就绪。
                if "selling_points" in needed:
                    latest["fields"]["selling_points"] = read_page_selling_points(page, asin)
                if "alexa_questions" in needed:
                    latest["fields"]["alexa_questions"] = read_page_alexa_questions(page, asin)
                signature = [latest["fields"][field]["value"] for field in needed]
                # 两次连续读取一致且所需字段齐全后返回，避免页面刚出现时仍是占位值。
                if all(value is not None for value in signature) and signature == previous:
                    return latest
                previous = signature
            except ValueError as error:
                last_error = str(error)
                latest = None
                previous = None
            page.wait_for_timeout(500)
        if latest is not None:
            latest["warning"] = "等待网页字段超时；缺失项未补齐。" + last_error
            return latest
        raise RuntimeError(last_error)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("items", nargs="+", help="一个或多个 ASIN / 美国站商品链接")
    parser.add_argument("--workbook", type=Path, required=True)
    parser.add_argument("--source-sheet")
    parser.add_argument("--output", type=Path, required=True, help="取数结果 JSON，不修改 Excel")
    parser.add_argument("--cdp", default=get_cdp_url())
    parser.add_argument("--timeout", type=float, default=25)
    parser.add_argument('--postal-code', default='10001')
    parser.add_argument("--source-only", action="store_true", help="仅检查 Excel，不打开网页")
    parser.add_argument("--fields", nargs="+", choices=list(AVAILABLE_FIELDS), default=list(ALL_FIELDS),
                        help="默认基础三项及卖点；可单选 selling_points 或 alexa_questions")
    args = parser.parse_args()
    if not 1 <= args.timeout <= 60:
        parser.error("--timeout 必须在 1～60 秒之间")
    if args.output.resolve() == args.workbook.resolve() or args.output.suffix.lower() != ".json":
        parser.error("--output 必须是独立的 .json 文件，不能覆盖工作簿")
    asins = list(dict.fromkeys(asin_from_input(item) for item in args.items))
    fields = {field: AVAILABLE_FIELDS[field] for field in args.fields}
    source = read_source(args.workbook, asins, args.source_sheet, fields)
    report = {"schema_version": SCHEMA, "site": "US", "currency": "USD", "requested_fields": list(fields), "products": []}
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with PluginCollector(args.cdp, args.timeout, args.postal_code) as collector:
        for asin in asins:
            print(f"正在处理 {asin}", flush=True)
            result = resolve_product(asin, source[asin], None if args.source_only else collector, fields)
            report["products"].append(result)
            if "alexa_questions" in fields:
                report["question_evidence"] = aggregate_question_evidence(report["products"])
            # 每完成一个商品即保存检查点；部分失败不丢失其它商品的取数结果。
            temporary = args.output.with_suffix(".json.tmp")
            temporary.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
            temporary.replace(args.output)
            print(json.dumps(result, ensure_ascii=False), flush=True)
            if result.get('access_paused'):
                break
    return 0 if all(item["status"] == "complete" for item in report["products"]) else 2


if __name__ == "__main__":
    sys.exit(main())
