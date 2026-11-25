from __future__ import annotations

import argparse
import base64
import json
import math
import re
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from io import BytesIO
from typing import Any

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.drawing.image import Image
from PIL import Image as PILImage

WORKFLOW_MODULES_DIR = Path(__file__).resolve().parents[2]
if str(WORKFLOW_MODULES_DIR) not in sys.path:
    sys.path.insert(0, str(WORKFLOW_MODULES_DIR))

from shared.marketplaces import ensure_amazon_postal_code, get_marketplace, normalize_amazon_url


CATEGORY_SECTION_MARKER = "所属平台分类及常用分类节点"
CATEGORY_HEADER_MARKER = "一级类目"
NEXT_SECTION_RE = re.compile(r"^\s*\d+\.")
CATEGORY_PERFORMANCE_COL = 6
CATEGORY_SALES_COL = 9
CATEGORY_LEVEL_START_COL = 1
CATEGORY_LEVEL_FIXED_COLS = 5
CATEGORY_EXTRA_LEVEL_START_COL = 10
OVERVIEW_SELECTOR = ".fullPanoramaBox"
DEFAULT_IMAGE_WIDTH = 360
DEFAULT_IMAGE_HEIGHT = 138
DEFAULT_IMAGE_SCALE = 2.0
DEFAULT_ATTEMPTS = 1
DEFAULT_TIMEOUT_MS = 30000
FIXED_CAPTURE_WAIT_MS = 20000
SCRIPT_DIR = Path(__file__).resolve().parent
NODE_CAPTURE_SCRIPT = SCRIPT_DIR / "chart-runtime" / "capture_sellersprite_overview.mjs"


def require_sync_playwright():
    try:
        from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
        from playwright.sync_api import sync_playwright
    except ImportError as exc:
        raise RuntimeError(
            "未安装 Python Playwright。请在当前运行环境里执行：\n"
            "  python -m pip install playwright\n"
            "并确保使用带 softtime/sorftime 插件的浏览器调试端口。"
        ) from exc
    return sync_playwright, PlaywrightTimeoutError


def find_research_sheet(workbook, requested: str | None = None):
    if requested:
        return workbook[requested]
    for ws in workbook.worksheets:
        if "市场调研" in ws.title:
            return ws
    return workbook.worksheets[0]


def find_row_containing(ws, text: str) -> int | None:
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and text in str(cell.value):
                return cell.row
    return None


def find_next_section_row(ws, start_row: int) -> int:
    for row in range(start_row, ws.max_row + 1):
        value = ws.cell(row, 1).value
        if value and NEXT_SECTION_RE.match(str(value)):
            return row
    return ws.max_row + 1


def cell_hyperlink_target(cell) -> str | None:
    if cell.hyperlink:
        return cell.hyperlink.target or cell.hyperlink.location
    return None


def merged_parent_cell(ws, row: int, col: int):
    cell = ws.cell(row, col)
    if not isinstance(cell, MergedCell):
        return cell
    for merged_range in ws.merged_cells.ranges:
        if merged_range.min_row <= row <= merged_range.max_row and merged_range.min_col <= col <= merged_range.max_col:
            return ws.cell(merged_range.min_row, merged_range.min_col)
    return cell


def category_leaf_cell(ws, row: int):
    candidate_cols = list(range(CATEGORY_LEVEL_START_COL, CATEGORY_LEVEL_FIXED_COLS + 1))
    candidate_cols.extend(range(CATEGORY_EXTRA_LEVEL_START_COL, ws.max_column + 1))
    for col in reversed(candidate_cols):
        cell = ws.cell(row, col)
        if cell.value not in (None, ""):
            return cell
    return None


def iter_category_rows(ws, site: str = "US"):
    section_row = find_row_containing(ws, CATEGORY_SECTION_MARKER)
    header_row = find_row_containing(ws, CATEGORY_HEADER_MARKER)
    if section_row is None or header_row is None:
        return
    first_data_row = header_row + 1
    next_section_row = find_next_section_row(ws, first_data_row)
    for row in range(first_data_row, next_section_row):
        leaf = category_leaf_cell(ws, row)
        if leaf is None or leaf.value in (None, ""):
            continue
        link = cell_hyperlink_target(leaf)
        if not link:
            continue
        yield {
            "row": row,
            "label": str(leaf.value).strip(),
            "url": normalize_amazon_category_url(link, site),
        }


def normalize_amazon_category_url(url: str, site: str = "US") -> str:
    return normalize_amazon_url(str(url).strip(), site)


def parse_monthly_sales(text: str) -> int | None:
    compact = re.sub(r"\s+", " ", text or "")
    match = re.search(r"月总销量\s*[:：]?\s*([0-9][0-9,]*)\b", compact)
    if not match:
        return None
    return int(match.group(1).replace(",", ""))


def clear_existing_category_images(ws, rows: set[int]) -> None:
    kept = []
    for image in getattr(ws, "_images", []):
        anchor = getattr(image, "anchor", None)
        anchor_row = None
        anchor_col = None
        try:
            anchor_row = anchor._from.row + 1
            anchor_col = anchor._from.col + 1
        except Exception:
            pass
        if anchor_row in rows and CATEGORY_PERFORMANCE_COL <= (anchor_col or 0) <= 8:
            continue
        kept.append(image)
    ws._images = kept


def insert_overview_image(ws, row: int, image_path: Path, image_width: int, image_height: int) -> None:
    anchor_cell = merged_parent_cell(ws, row, CATEGORY_PERFORMANCE_COL)
    # 清除原模板中失效的 WPS 图片公式，避免截图没有显示时出现 #REF!。
    anchor_cell.value = None
    # 保留截图原始像素，只缩小 Excel 显示框，避免实际重采样造成模糊。
    img = Image(str(image_path))
    img.width = image_width
    img.height = image_height
    ws.add_image(img, anchor_cell.coordinate)


def image_display_size(image_path: Path, image_scale: float) -> tuple[int, int]:
    """按类目表现单元格适配显示框；原始高清 PNG 文件保持不变。"""
    img = Image(str(image_path))
    # 标准类目概览统一使用模板显示尺寸，避免不同页面截图造成大小混用。
    aspect_ratio = img.width / max(1, img.height)
    if 2.9 <= aspect_ratio <= 3.4:
        return DEFAULT_IMAGE_WIDTH, 115
    scale = min(
        DEFAULT_IMAGE_WIDTH / max(1, img.width),
        DEFAULT_IMAGE_HEIGHT / max(1, img.height),
    )
    return (
        max(1, round(img.width * scale)),
        max(1, round(img.height * scale)),
    )


def is_real_overview(text: str) -> bool:
    return "月总销量" in text and ("Listing垄断系数" in text or "品牌垄断系数" in text)


def capture_overview_element(page, context, output_path: Path, timeout_ms: int = 10000) -> None:
    """使用 CDP 截取完整的 Softtime 面板，避免字体加载阻塞 Playwright 截图。"""
    overview = page.locator(OVERVIEW_SELECTOR).first
    if not overview.count():
        raise RuntimeError("Softtime market overview has no valid screenshot element.")
    overview.scroll_into_view_if_needed(timeout=max(1000, timeout_ms))
    box = overview.bounding_box(timeout=max(1000, timeout_ms))
    if not box or box["width"] <= 0 or box["height"] <= 0:
        raise RuntimeError("Softtime market overview has no valid screenshot bounds.")
    cdp = context.new_cdp_session(page)
    device_scale = float(page.evaluate("() => window.devicePixelRatio || 1"))
    try:
        result = cdp.send(
            "Page.captureScreenshot",
            {
                "format": "png",
                "fromSurface": False,
                "captureBeyondViewport": False,
            },
        )
        full_screenshot = base64.b64decode(result["data"])
        with PILImage.open(BytesIO(full_screenshot)) as source:
            # CSS 像素要换算为设备像素，否则会裁切左侧区域。
            left = max(0, math.floor(box["x"] * device_scale))
            top = max(0, math.floor(box["y"] * device_scale))
            right = min(
                source.width,
                left + math.ceil(box["width"] * device_scale),
            )
            bottom = min(
                source.height,
                top + math.ceil(box["height"] * device_scale),
            )
            if right <= left or bottom <= top:
                raise RuntimeError(f"Softtime screenshot bounds exceed viewport: {source.width}x{source.height}")
            source.crop((left, top, right, bottom)).save(output_path, format="PNG")
    finally:
        cdp.detach()


def prepare_capture_window(page, context):
    """截图期间固定 Chrome 外层窗口，避免动态窗口宽度造成右侧内容缺失。"""
    cdp = context.new_cdp_session(page)
    original = cdp.send("Browser.getWindowForTarget")
    bounds = original.get("bounds", {})
    cdp.send(
        "Browser.setWindowBounds",
        {
            "windowId": original["windowId"],
            "bounds": {"left": 0, "top": 0, "width": 1920, "height": 1080, "windowState": "normal"},
        },
    )
    page.bring_to_front()
    page.wait_for_timeout(750)
    page.set_viewport_size({"width": 1920, "height": 1080})

    def restore():
        try:
            cdp.send(
                "Browser.setWindowBounds",
                {
                    "windowId": original["windowId"],
                    "bounds": {
                        "left": bounds.get("left", 0),
                        "top": bounds.get("top", 0),
                        "width": bounds.get("width", 1920),
                        "height": bounds.get("height", 1080),
                        "windowState": bounds.get("windowState", "normal"),
                    },
                },
            )
        finally:
            cdp.detach()

    return restore


def capture_one(
    page,
    context,
    url: str,
    output_path: Path,
    timeout_ms: int,
    allow_partial: bool = False,
    amazon_postal_code: str | None = None,
) -> dict[str, Any]:
    page.goto(url, wait_until="domcontentloaded", timeout=timeout_ms)
    postal_result = ensure_amazon_postal_code(page, amazon_postal_code or "")
    if postal_result.get("status") not in {"skipped", "success", "already-set"}:
        raise RuntimeError(
            f"Amazon delivery postal code was not confirmed: "
            f"{postal_result.get('status')} {postal_result.get('error') or ''} "
            f"{postal_result.get('text') or ''}".strip()
        )
    if postal_result.get("status") == "success":
        page.goto(url, wait_until="domcontentloaded", timeout=timeout_ms)
    page.wait_for_selector(OVERVIEW_SELECTOR, timeout=timeout_ms)
    overview = page.locator(OVERVIEW_SELECTOR)
    overview.scroll_into_view_if_needed(timeout=timeout_ms)
    readiness_started = time.monotonic()
    # 插件区域出现后固定等待满 20 秒，避免结构先出现但 Canvas 尚未完成绘制。
    page.wait_for_timeout(FIXED_CAPTURE_WAIT_MS)
    complete = page.evaluate(
        """selector => {
          const el = document.querySelector(selector);
          if (!el) return false;
          const rect = el.getBoundingClientRect();
          const text = (el.innerText || '').replace(/\\s+/g, ' ');
          const canvases = [...el.querySelectorAll('canvas')];
          const required = [
            /月总销量/,
            /月总销售额/,
            /Listing.*垄断/,
            /卖家.*垄断/,
            /品牌.*垄断|品牌销量垄断/,
            /平均售价/,
            /平均评价数量/,
            /平均星级/,
            /A\\+视频占比/,
          ];
          return rect.width >= 300
            && rect.height >= 120
            && required.every(pattern => pattern.test(text))
            && canvases.length >= 12
            && canvases.every(canvas => canvas.width > 0 && canvas.height > 0);
        }""",
        OVERVIEW_SELECTOR,
    )
    if not complete and not allow_partial:
        raise RuntimeError("Softtime market overview is incomplete after the fixed 20-second wait.")
    text = page.locator(OVERVIEW_SELECTOR).inner_text(timeout=3000)
    # 保持单一路径，避免超时后两个截图任务同时写入同一个文件。
    capture_overview_element(page, context, output_path)
    return {
        "url": page.url,
        "monthly_sales": parse_monthly_sales(text),
        "complete": complete,
        "waited_ms": int((time.monotonic() - readiness_started) * 1000),
        "text": re.sub(r"\s+", " ", text).strip(),
        "amazon_delivery_zip_status": postal_result.get("status"),
        "amazon_delivery_location_text": postal_result.get("text"),
    }


def capture_one_with_node(
    *,
    node_exe: str,
    url: str,
    output_path: Path,
    debugger_address: str,
    timeout_ms: int,
    allow_partial: bool = False,
    amazon_postal_code: str | None = None,
) -> dict[str, Any]:
    completed = subprocess.run(
        [
            node_exe,
            str(NODE_CAPTURE_SCRIPT),
            "--url",
            url,
            "--output",
            str(output_path),
            "--debugger-address",
            debugger_address,
            "--timeout-ms",
            str(timeout_ms),
            "--allow-partial",
            "1" if allow_partial else "0",
            "--postal-code",
            amazon_postal_code or "",
        ],
        cwd=str(SCRIPT_DIR),
        text=True,
        encoding="utf-8",
        errors="replace",
        capture_output=True,
    )
    if completed.returncode != 0:
        message = (completed.stderr or completed.stdout or "").strip().splitlines()
        # 优先取最后一条真正的 Error，避免把普通阶段日志误当成失败原因。
        meaningful = next(
            (
                line.strip()
                for line in reversed(message)
                if line.strip().startswith(("Error:", "TypeError:", "RangeError:"))
            ),
            message[-1].strip()
            if message
            else f"Node capture failed with exit code {completed.returncode}",
        )
        raise RuntimeError(meaningful)
    lines = [line.strip() for line in completed.stdout.splitlines() if line.strip()]
    if not lines:
        raise RuntimeError("Node capture did not return JSON.")
    return json.loads(lines[-1])


def capture_category_overviews(
    *,
    workbook_path: Path,
    output_path: Path,
    debugger_address: str,
    image_dir: Path,
    sheet_name: str | None = None,
    limit: int | None = None,
    timeout_ms: int = DEFAULT_TIMEOUT_MS,
    image_width: int | None = None,
    image_height: int | None = None,
    image_scale: float = DEFAULT_IMAGE_SCALE,
    attempts: int = DEFAULT_ATTEMPTS,
    node_exe: str | None = None,
    site: str = "US",
    amazon_postal_code: str | None = None,
    start_row: int | None = None,
    capture_runtime: str = "node",
) -> list[dict[str, Any]]:
    if capture_runtime != "node":
        raise ValueError("类目截图只允许使用 Node Playwright 截图器。")
    marketplace = get_marketplace(site)
    amazon_postal_code = amazon_postal_code or marketplace.postal_code
    # 类目截图只增删图片，必须保留竞品对比维度等已有富文本。
    wb = load_workbook(workbook_path, rich_text=True)
    ws = find_research_sheet(wb, sheet_name)
    rows = list(iter_category_rows(ws, marketplace.code))
    if start_row is not None:
        rows = [item for item in rows if item["row"] >= start_row]
    if limit is not None:
        rows = rows[: max(0, limit)]
    clear_existing_category_images(ws, {item["row"] for item in rows})
    image_dir.mkdir(parents=True, exist_ok=True)

    results: list[dict[str, Any]] = []
    node_exe = node_exe or "node"
    for index, item in enumerate(rows, start=1):
        image_path = image_dir / f"category_overview_row_{item['row']}.png"
        data = {}
        status = "error"
        error = None
        attempt_used = 0
        for attempt in range(1, max(1, attempts) + 1):
            attempt_used = attempt
            try:
                data = capture_one_with_node(
                    node_exe=node_exe,
                    url=item["url"],
                    output_path=image_path,
                    debugger_address=debugger_address,
                    timeout_ms=timeout_ms,
                    allow_partial=attempt >= max(1, attempts),
                    amazon_postal_code=amazon_postal_code,
                )
                display_width, display_height = (
                    (image_width, image_height)
                    if image_width is not None and image_height is not None
                    else image_display_size(image_path, image_scale)
                )
                insert_overview_image(
                    ws,
                    item["row"],
                    image_path,
                    display_width,
                    display_height,
                )
                if data["monthly_sales"] is not None:
                    from product_research.scripts.fill_custom_template import set_category_monthly_sales
                    set_category_monthly_sales(ws.cell(item["row"], CATEGORY_SALES_COL), data["monthly_sales"])
                status = "ok"
                error = None
                break
            except Exception as exc:
                error = str(exc).splitlines()[0]
        results.append(
            {
                "index": index,
                "row": item["row"],
                "label": item["label"],
                "url": item["url"],
                "status": status,
                "error": error,
                "attempts": attempt_used,
                "complete": data.get("complete") if status == "ok" else False,
                "waited_ms": data.get("waited_ms") if status == "ok" else None,
                "image": str(image_path) if status == "ok" else None,
                "display_size": [display_width, display_height]
                if status == "ok"
                else None,
                "monthly_sales": data.get("monthly_sales"),
                "resolved_url": data.get("url"),
            }
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return results


def main() -> int:
    parser = argparse.ArgumentParser(description="Capture softtime/sorftime category overview screenshots into market research rows.")
    parser.add_argument("--input", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--image-dir", type=Path)
    parser.add_argument("--debugger-address", default="127.0.0.1:9223")
    parser.add_argument("--sheet")
    parser.add_argument("--limit", type=int)
    parser.add_argument("--timeout-ms", type=int, default=DEFAULT_TIMEOUT_MS)
    parser.add_argument("--image-width", type=int)
    parser.add_argument("--image-height", type=int)
    parser.add_argument("--image-scale", type=float, default=DEFAULT_IMAGE_SCALE)
    parser.add_argument("--attempts", type=int, default=DEFAULT_ATTEMPTS)
    parser.add_argument("--node", default="node")
    parser.add_argument("--site", default="US")
    parser.add_argument("--amazon-postal-code")
    parser.add_argument("--start-row", type=int)
    parser.add_argument(
        "--capture-runtime",
        choices=["node"],
        default="node",
        help="兼容参数；类目截图固定使用 Node Playwright。",
    )
    args = parser.parse_args()

    image_dir = args.image_dir
    if image_dir is None:
        image_dir = Path(tempfile.mkdtemp(prefix="category-overviews-"))
    results = capture_category_overviews(
        workbook_path=args.input,
        output_path=args.output,
        debugger_address=args.debugger_address,
        image_dir=image_dir,
        sheet_name=args.sheet,
        limit=args.limit,
        timeout_ms=args.timeout_ms,
        image_width=args.image_width,
        image_height=args.image_height,
        image_scale=args.image_scale,
        attempts=args.attempts,
        node_exe=args.node,
        site=args.site,
        amazon_postal_code=args.amazon_postal_code,
        start_row=args.start_row,
        capture_runtime=args.capture_runtime,
    )
    print(json.dumps(results, ensure_ascii=False, indent=2))
    print(args.output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
