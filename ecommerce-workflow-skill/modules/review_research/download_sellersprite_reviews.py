"""通过项目 Chrome 的卖家精灵插件下载评论，保存原文件和可追溯的评论 JSON。"""

from __future__ import annotations

import argparse
from collections import Counter
from datetime import date, datetime, timezone
import hashlib
import json
import math
from pathlib import Path
import re
import tempfile
import time
from urllib.parse import urlparse

from openpyxl import load_workbook


PANEL = "#seller-sprite-extension-quick-view-listing"
REQUIRED_HEADERS = {"ASIN", "标题", "内容", "VP评论", "Vine Voice评论", "星级", "评论链接", "评论时间"}


class NeedsUserAction(RuntimeError):
    """需要保留页面，交给用户处理登录或真人验证。"""


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def asin_from_input(value: str) -> str:
    """只接受明确的 ASIN 或美国站商品链接，避免误开其它站点。"""
    value = value.strip()
    if re.fullmatch(r"[A-Za-z0-9]{10}", value):
        return value.upper()
    url = urlparse(value)
    match = re.search(r"/(?:dp|gp/product)/([A-Za-z0-9]{10})(?:/|$)", url.path)
    if url.scheme == "https" and url.hostname in {"amazon.com", "www.amazon.com"} and match:
        return match[1].upper()
    raise ValueError(f"请输入 ASIN 或美国站 HTTPS 商品链接：{value}")


def text(value) -> str:
    return str(value).strip() if value is not None else ""


def source_flag(value) -> bool:
    """导出表用 Y 表示有标记；空白仅表示导出文件未标记。"""
    marker = text(value).upper()
    if marker not in {"", "Y", "N"}:
        raise ValueError(f"未知的 VP/Vine 标记：{marker}")
    return marker == "Y"


def helpful_count(value) -> int:
    """赞同数空白按零处理；拒绝负数、缩写和非整数，避免错误排序。"""
    if value is None or text(value) == "":
        return 0
    if isinstance(value, bool):
        raise ValueError("赞同数不能是布尔值")
    if isinstance(value, float) and not math.isfinite(value):
        raise ValueError("赞同数必须是有限的非负整数")
    if isinstance(value, (int, float)):
        if value >= 0 and value == int(value):
            return int(value)
    elif re.fullmatch(r"(?:\d+|\d{1,3}(?:,\d{3})+)", text(value)):
        return int(text(value).replace(",", ""))
    raise ValueError(f"赞同数必须是非负整数：{value}")


def json_value(value):
    return value.isoformat() if isinstance(value, (date, datetime)) else value


def read_export(path: Path, asin: str) -> dict:
    """只读校验下载表；保留所有原始字段，不擅自筛选评论或补造子体 ASIN。"""
    asin = asin_from_input(asin)
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        # 导出页签包含下载目标和条数；不能仅凭调用参数给一份无关文件贴上目标 ASIN。
        candidates = [sheet for sheet in workbook if re.fullmatch(
            rf"{re.escape(asin)}-Reviews\(\d+\)", sheet.title, re.I)]
        if len(candidates) != 1:
            raise ValueError(f"未找到唯一的 {asin}-Reviews(条数) 页，拒绝导入错商品文件")
        sheet = candidates[0]
        rows = sheet.iter_rows(values_only=True)
        headers = [text(value) for value in next(rows, ())]
        if not REQUIRED_HEADERS.issubset(headers) or len(headers) != len(set(headers)):
            raise ValueError("评论导出表头缺失或重复，请核对插件导出格式")
        reviews = []
        for row_number, row in enumerate(rows, 2):
            if not any(value is not None and text(value) for value in row):
                continue
            raw = {key: json_value(value) for key, value in zip(headers, row)}
            rating_value = raw.get("星级")
            try:
                rating = float(rating_value)
            except (ValueError, TypeError):
                raise ValueError(f"第 {row_number} 行星级无效：{rating_value}") from None
            if rating not in {1, 2, 3, 4, 5} or isinstance(rating_value, bool):
                raise ValueError(f"第 {row_number} 行星级必须是 1～5")
            url = text(raw.get("评论链接"))
            parsed = urlparse(url)
            match = re.search(r"/(?:customer-reviews|srp/-)/([A-Z0-9]+)(?:/|$)", parsed.path)
            # 美国站的插件导出会包含“其它国家/地区”评论，保留真实原链接与所属国家。
            review_domains = {"amazon.com", "amazon.de", "amazon.co.uk", "amazon.ca", "amazon.fr", "amazon.in"}
            hostname = (parsed.hostname or "").removeprefix("www.")
            if parsed.scheme != "https" or hostname not in review_domains or not match:
                raise ValueError(f"第 {row_number} 行缺少有效的Amazon评论链接")
            row_asin = text(raw.get("ASIN")).upper()
            if row_asin and not re.fullmatch(r"[A-Z0-9]{10}", row_asin):
                raise ValueError(f"第 {row_number} 行原始 ASIN 无效")
            reviews.append({
                "asin": asin, "source_row_asin": row_asin or None,
                "review_id": match[1], "rating": int(rating),
                "title": text(raw.get("标题")), "body": text(raw.get("内容")),
                "title_translated": text(raw.get("标题(翻译)")),
                "body_translated": text(raw.get("内容(翻译)")),
                "verified": source_flag(raw.get("VP评论")),
                "vine": source_flag(raw.get("Vine Voice评论")),
                "helpful_count": helpful_count(raw.get("赞同数")) if "赞同数" in headers else None,
                "date_text": text(raw.get("评论时间")), "variant": text(raw.get("型号")),
                "country": text(raw.get("所属国家")), "review_url": url,
                "source_sheet": sheet.title, "source_row": row_number, "raw": raw,
            })
        declared_count = int(re.search(r"\((\d+)\)$", sheet.title)[1])
        if declared_count != len(reviews):
            raise ValueError(f"页签声明 {declared_count} 条，实际 {len(reviews)} 条，数量不一致")
        unique_count = len({item["review_id"] for item in reviews})
        if unique_count != len(reviews):
            raise ValueError("导出文件存在重复评论 ID，需核对后再用于分析")
        warnings = []
        missing_asin = sum(item["source_row_asin"] is None for item in reviews)
        other_asin = sum(item["source_row_asin"] not in {None, asin} for item in reviews)
        # 用户已确认 ASIN 空白可以接受；保留空值及计数，不将其作为错误或阻断条件。
        if other_asin:
            warnings.append(f"{other_asin} 条评论带其它子体 ASIN，保留原值；后续分析须确认变体范围")
        empty_body = sum(not item["body"] for item in reviews)
        if empty_body:
            warnings.append(f"{empty_body} 条评论正文为空，保留原记录供用户核查，不补造或删除")
        if not reviews:
            warnings.append("导出文件为空，不能据此断言该商品没有评论")
        translation = {
            "columns_present": {"标题(翻译)", "内容(翻译)"}.issubset(headers),
            "missing_title_rows": [r["source_row"] for r in reviews if r["title"] and not r["title_translated"]],
            "missing_body_rows": [r["source_row"] for r in reviews if r["body"] and not r["body_translated"]],
        }
        translation["complete"] = (translation["columns_present"] and not translation["missing_title_rows"]
                                   and not translation["missing_body_rows"])
        if not translation["complete"]:
            warnings.append("翻译列缺失或部分翻译为空，原文件已保留，尚不能作为翻译完整的数据源")
        return {
            "schema": "sellersprite-reviews.v1", "asin": asin, "site": "US",
            "source_sheet": sheet.title,
            "source": "sellersprite_plugin_export", "source_file": str(path.resolve()),
            "source_sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
            "read_at": now_iso(), "review_count": len(reviews),
            "headers": headers,
            "blank_source_asin_count": missing_asin,
            "rating_counts": dict(sorted(Counter(str(r["rating"]) for r in reviews).items())),
            "verified_count": sum(r["verified"] for r in reviews),
            "vine_count": sum(r["vine"] for r in reviews),
            "trusted_count": sum(r["verified"] or r["vine"] for r in reviews),
            "all_amazon_reviews_confirmed": False, "translation": translation,
            "warnings": warnings, "reviews": reviews,
        }
    finally:
        workbook.close()


def write_json(path: Path, payload: dict) -> None:
    """逐商品写检查点，避免后续商品失败导致已完成记录丢失。"""
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8", newline="\n")
    temporary.replace(path)


def check_product(page, asin: str) -> None:
    """核对网页和插件 ASIN，不处理登录、验证码或付费提示。"""
    try:
        matched = asin_from_input(page.url) == asin
    except ValueError:
        matched = False
    if not matched:
        raise ValueError("页面跳转或 ASIN 不一致，请人工核对商品或登录/验证码页面")
    panel = page.locator(PANEL)
    panel.wait_for(state="visible", timeout=45_000)
    found = re.search(r"ASIN\s*[:：]\s*([A-Z0-9]{10})\b", panel.inner_text(), re.I)
    if not found or found[1].upper() != asin:
        raise ValueError("卖家精灵插件 ASIN 与目标不一致")


def continue_after_user_check(page) -> bool:
    """只在真人框已经勾选后点击继续；本函数不会主动勾选验证框。"""
    if "/checkpoint/" not in urlparse(page.url).path:
        return False
    checkbox = page.locator("#continueCheck")
    if checkbox.count() != 1 or not checkbox.is_checked():
        return False
    button = page.get_by_role("button", name="Continue shopping", exact=True)
    if button.count() == 1 and button.is_enabled():
        button.click(timeout=5000)
        return True
    return False


def export_status(payload: dict, success: str) -> str:
    """文件下载完成与翻译数据源就绪分开，空文件不冒充无评论。"""
    if not payload["review_count"]:
        return "empty_export"
    return success if payload["translation"]["complete"] else "incomplete_translation"


class NamedDownload:
    """使用 Chrome 原始文件名落盘，避免 Playwright 的随机编号出现在下载列表。"""

    def __init__(self, page, destination: Path):
        self.page = page
        self.destination = destination.resolve()
        self.session = None
        self.guid = None
        self.filename = None
        self.state = None
        self.file_path = None
        self.error = None

    def __enter__(self):
        frame_session = self.page.context.new_cdp_session(self.page)
        try:
            self.frame_id = frame_session.send("Page.getFrameTree")["frameTree"]["frame"]["id"]
        finally:
            frame_session.detach()
        self.session = self.page.context.browser.new_browser_cdp_session()
        self.session.on("Browser.downloadWillBegin", self.on_begin)
        self.session.on("Browser.downloadProgress", self.on_progress)
        try:
            self.session.send("Browser.setDownloadBehavior", {
                "behavior": "allow", "downloadPath": str(self.destination), "eventsEnabled": True,
            })
        except Exception:
            self.session.detach()
            raise
        return self

    def on_begin(self, event: dict) -> None:
        # 只接收本商品页发起的下载；其它标签页的下载事件不能冒充本次结果。
        if event.get("frameId") != self.frame_id or self.guid is not None:
            return
        filename = event.get("suggestedFilename", "")
        self.guid = event.get("guid")
        if (not filename or Path(filename).name != filename or "\\" in filename
                or ":" in filename or not filename.lower().endswith(".xlsx")):
            self.error = "收到的文件名不是预期的评论 Excel"
            return
        self.filename = filename

    def on_progress(self, event: dict) -> None:
        if self.guid is None or event.get("guid") != self.guid:
            return
        self.state = event.get("state")
        if event.get("filePath"):
            self.file_path = event["filePath"]

    def completed_path(self) -> Path:
        if self.error:
            raise ValueError(self.error)
        if self.state != "completed" or not self.filename:
            raise RuntimeError("浏览器下载尚未完成")
        # 新建的独立目录没有同名旧文件；有实际路径时仍核对目录，避免读错其它下载。
        path = Path(self.file_path).resolve() if self.file_path else self.destination / self.filename
        if path.parent != self.destination or not path.name.lower().endswith(".xlsx"):
            raise ValueError("下载文件路径不在当前商品输出目录")
        if not path.is_file():
            raise RuntimeError("浏览器报告完成，但评论文件尚未落盘")
        return path

    def __exit__(self, *args):
        if self.session is not None:
            try:
                # 退出后恢复 Chrome 默认下载方式，后续手动下载不继续落入项目目录。
                self.session.send("Browser.setDownloadBehavior", {"behavior": "default", "eventsEnabled": False})
            finally:
                self.session.detach()


def download_one(page, asin: str, destination: Path, timeout: float) -> dict:
    """按实测页面流程点一次按钮；以真实下载事件和文件校验为完成依据。"""
    started = time.monotonic()
    page.goto(f"https://www.amazon.com/dp/{asin}", wait_until="domcontentloaded", timeout=45_000)
    blocked_paths = ("/checkpoint/", "/ap/signin", "/errors/validateCaptcha")
    if any(part in urlparse(page.url).path for part in blocked_paths):
        page.bring_to_front()
        print(f"[{asin}] 等待人工处理登录/真人验证，请在项目浏览器中完成后继续；最多等待 5 分钟", flush=True)
        manual_deadline = time.monotonic() + 300
        while any(part in urlparse(page.url).path for part in blocked_paths):
            if time.monotonic() >= manual_deadline:
                raise NeedsUserAction("等待人工处理超时，验证页已保留，请完成后仅重试本商品")
            continue_after_user_check(page)
            page.wait_for_timeout(1000)
    check_product(page, asin)
    panel = page.locator(PANEL)
    button = panel.get_by_text("评论下载", exact=True)
    button.wait_for(state="visible", timeout=15_000)
    if button.count() != 1:
        raise ValueError("评论下载入口不唯一，停止以免误点")
    # 先悬停展开提示层，再勾选翻译；不能先点下载后才改变选项。
    button.hover(timeout=10_000)
    translate = page.get_by_role("checkbox", name="翻译标题和内容", exact=True)
    label = page.get_by_text("翻译标题和内容", exact=True)
    label.wait_for(state="visible", timeout=10_000)
    # 插件隐藏原生 input，用可见文字标签控制复选框；读取 input 状态确认选中。
    if not translate.is_checked():
        label.click(timeout=10_000)
    if not translate.is_checked():
        raise RuntimeError("翻译标题和内容未勾选，停止下载")
    print(f"[{asin}] 已勾选翻译标题和内容", flush=True)
    last_progress = ""
    with NamedDownload(page, destination) as download:
        download_started = time.monotonic()
        button.click(timeout=10_000)
        deadline = time.monotonic() + timeout
        while download.state != "completed" and time.monotonic() < deadline:
            if download.error:
                raise ValueError(download.error)
            if download.state == "canceled":
                raise RuntimeError("浏览器已取消评论下载，未生成完整数据源")
            if page.is_closed():
                raise RuntimeError("下载页面已关闭")
            progress = re.search(r"下载中\s*\.{0,3}\s*\d+%", panel.inner_text(timeout=3000))
            if progress and progress[0] != last_progress:
                last_progress = progress[0]
                print(f"[{asin}] {last_progress}", flush=True)
            page.wait_for_timeout(500)
        if download.state != "completed":
            raise TimeoutError(f"{timeout:g} 秒内未收到下载文件；最后进度：{last_progress or '未出现'}。请检查插件登录或提示；不自动重复点击")
        path = download.completed_path()
        check_product(page, asin)
        payload = read_export(path, asin)
        payload.update({"page_url": page.url, "downloaded_at": now_iso(), "translation_requested": True,
                        "download_seconds": round(time.monotonic() - download_started, 2),
                        "elapsed_seconds": round(time.monotonic() - started, 2)})
        write_json(destination / "reviews.json", payload)
        return {"asin": asin, "status": export_status(payload, "downloaded"),
                "source_file": str(path.resolve()), "reviews_file": str((destination / "reviews.json").resolve()),
                "review_count": payload["review_count"], "elapsed_seconds": payload["elapsed_seconds"],
                "download_seconds": payload["download_seconds"], "translation": payload["translation"],
                "warnings": payload["warnings"]}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("items", nargs="+", help="一个或多个 ASIN / 美国站商品链接")
    parser.add_argument("--output-dir", required=True, type=Path, help="本地输出父目录；每次创建独立运行目录")
    parser.add_argument("--cdp", default="http://127.0.0.1:9223")
    parser.add_argument("--download-timeout", type=float, default=180)
    parser.add_argument("--import-file", type=Path, help="只读校验已下载文件，仅支持一个 ASIN，不连接浏览器")
    args = parser.parse_args()
    try:
        asins = list(dict.fromkeys(asin_from_input(item) for item in args.items))
    except ValueError as error:
        parser.error(str(error))
    if not 0 < args.download_timeout < float("inf"):
        parser.error("下载超时必须为有限正数")
    if args.import_file and len(asins) != 1:
        parser.error("--import-file 只能对应一个 ASIN")
    args.output_dir.mkdir(parents=True, exist_ok=True)
    run_dir = Path(tempfile.mkdtemp(prefix="sellersprite-reviews-", dir=args.output_dir)).resolve()
    manifest = {"schema": "sellersprite-review-downloads.v1", "started_at": now_iso(),
                "requested_asins": asins, "products": []}
    manifest_path = run_dir / "manifest.json"
    write_json(manifest_path, manifest)
    print(f"输出目录：{run_dir}", flush=True)
    if args.import_file:
        payload = read_export(args.import_file, asins[0])
        write_json(run_dir / "reviews.json", payload)
        manifest["products"].append({"asin": asins[0], "status": export_status(payload, "imported"),
                                     "source_file": payload["source_file"], "reviews_file": str(run_dir / "reviews.json"),
                                     "review_count": payload["review_count"], "translation": payload["translation"]})
    else:
        from playwright.sync_api import sync_playwright

        # 使用项目已有本地连接，不依赖 ChatGPT 扩展，不读取凭据，也不关闭用户浏览器。
        with sync_playwright() as runtime:
            browser = runtime.chromium.connect_over_cdp(args.cdp, timeout=20_000)
            # CDP 连接本身会让 Playwright 接管下载；先立即恢复，避免导航失败时遗留随机命名。
            setup_session = browser.new_browser_cdp_session()
            try:
                setup_session.send("Browser.setDownloadBehavior", {"behavior": "default", "eventsEnabled": False})
            finally:
                setup_session.detach()
            if not browser.contexts:
                raise RuntimeError("未找到项目浏览器上下文，请先启动项目专用浏览器")
            context = browser.contexts[0]
            for asin in asins:
                destination = run_dir / asin
                destination.mkdir()
                page = context.new_page()
                keep_page = False
                try:
                    result = download_one(page, asin, destination, args.download_timeout)
                except Exception as error:
                    keep_page = isinstance(error, NeedsUserAction)
                    result = {"asin": asin, "status": "needs_user_action" if keep_page else "failed", "error": str(error)}
                    # 失败截图和原下载文件只留本地，便于定点重试。
                    try:
                        page.screenshot(path=str(destination / "failure.png"))
                    except Exception:
                        pass
                finally:
                    if not keep_page and not page.is_closed():
                        page.close()
                manifest["products"].append(result)
                write_json(manifest_path, manifest)
                print(json.dumps(result, ensure_ascii=False), flush=True)
                if keep_page:
                    completed = {item["asin"] for item in manifest["products"]}
                    manifest["products"].extend({"asin": pending, "status": "not_attempted"}
                                                for pending in asins if pending not in completed)
                    break
    manifest["finished_at"] = now_iso()
    write_json(manifest_path, manifest)
    print(f"结果清单：{manifest_path}", flush=True)
    return 0 if all(item["status"] in {"downloaded", "imported"} for item in manifest["products"]) else 1


if __name__ == "__main__":
    raise SystemExit(main())
