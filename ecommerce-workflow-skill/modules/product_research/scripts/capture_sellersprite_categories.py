"""抓取卖家精灵插件的类目概览，按实际类目行回填并保留逐项检查点。"""
from __future__ import annotations

import argparse
import base64
import json
import re
import sys
import time
from pathlib import Path

from openpyxl import load_workbook
from PIL import Image as PILImage, ImageStat
from playwright.sync_api import sync_playwright

MODULES = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(MODULES))
from product_research.scripts.capture_category_overviews import iter_category_rows, clear_existing_category_images, insert_overview_image
from product_research.scripts.insert_sellersprite_screenshots import region_pixels


PANEL = "#seller-sprite-best-sellers-box"


def read_sales(text):
    match = re.search(r"(?:^|\n)月销量\s*\n\s*([\d,]+)(?:\n|$)", text)
    return int(match[1].replace(",", "")) if match else None


def capture(page, url, image, postal_code):
    if page.url != url:
        page.goto(url, wait_until="domcontentloaded", timeout=45000)
    if any(part in page.url for part in ("/checkpoint/", "/ap/signin", "validateCaptcha")):
        raise RuntimeError("Amazon需要人工验证或登录，页面已保留")
    location = page.locator("#glow-ingress-line2").inner_text(timeout=5000)
    if postal_code not in location:
        raise RuntimeError(f"配送邮编不符：{location}")
    panel = page.locator(PANEL)
    panel.wait_for(state="visible", timeout=30000)
    expand = panel.get_by_text("展开", exact=True)
    if expand.count() and expand.is_visible():
        expand.click()
    # 保留项目固定等待20秒的策略，等待插件绘图完成。
    page.wait_for_timeout(20000)
    refreshed = False
    deadline = time.monotonic() + 120
    while True:
        text = panel.inner_text(timeout=5000)
        if "需要登录" in text:
            raise RuntimeError("卖家精灵插件未登录")
        sales = read_sales(text)
        if sales is not None and panel.locator("canvas").count() > 0:
            break
        if time.monotonic() > deadline:
            raise TimeoutError("类目样本抓取未在120秒内返回有效月销量")
        refresh = panel.get_by_text("立即刷新", exact=True)
        if not refreshed and refresh.count() and refresh.is_visible():
            refresh.click()
            refreshed = True
        print(f"[类目] 等待样本：{'抓取中' if '抓取中' in text else '等待插件数据'}", flush=True)
        page.wait_for_timeout(10000)
    tip = panel.get_by_text("知道了", exact=True)
    if tip.count() and tip.is_visible():
        tip.click()
    # 插件持续重排会使元素截图一直等待稳定，改用当前可见元素边界裁剪浏览器图像。
    panel.evaluate("element => element.scrollIntoView({block: 'start', inline: 'nearest'})")
    box = panel.bounding_box(timeout=5000)
    if not box or min(box["width"], box["height"]) < 20:
        raise ValueError("插件图表没有有效边界")
    scroll = page.evaluate("({x: scrollX, y: scrollY})")
    session = page.context.new_cdp_session(page)
    try:
        result = session.send("Page.captureScreenshot", {
            "format": "png", "fromSurface": True, "captureBeyondViewport": True,
            "clip": {"x": box["x"] + scroll["x"], "y": box["y"] + scroll["y"],
                     "width": box["width"], "height": box["height"], "scale": 1},
        })
        image.write_bytes(base64.b64decode(result["data"]))
    finally:
        session.detach()
    with PILImage.open(image) as preview:
        if max(ImageStat.Stat(preview.convert("RGB")).stddev) < 3:
            raise ValueError("插件截图接近空白")
        size = preview.size
    return {"monthly_sales": sales, "image": str(image), "pixels": list(size), "text": text, "url": page.url}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--image-dir", type=Path, required=True)
    parser.add_argument("--postal-code", default="10001")
    parser.add_argument("--start", type=int, default=0)
    parser.add_argument("--limit", type=int, default=5)
    args = parser.parse_args()
    if args.input.resolve() == args.output.resolve():
        raise ValueError("请输出到独立检查点")
    if not 1 <= args.limit <= 5:
        raise ValueError("单批处理1到5个类目")
    args.image_dir.mkdir(parents=True, exist_ok=True)
    manifest_path = args.image_dir / "manifest.json"
    results = json.loads(manifest_path.read_text(encoding="utf-8")) if manifest_path.exists() else {}
    wb = load_workbook(args.input, rich_text=True)
    ws = next(ws for ws in wb if ws.title.endswith("市场调研"))
    items = list(iter_category_rows(ws))[args.start:args.start + args.limit]
    failed = False
    with sync_playwright() as runtime:
        browser = runtime.chromium.connect_over_cdp("http://127.0.0.1:9223", timeout=20000)
        context = browser.contexts[0]
        for item in items:
            url = item["url"]
            started = time.monotonic()
            print(f"[类目] 开始 {item['label']}，目标第{item['row']}行", flush=True)
            try:
                prior = results.get(url, {})
                if prior.get("status") == "ok" and Path(prior["image"]).exists():
                    result = prior
                else:
                    page = next((page for page in context.pages if page.url == url), None)
                    if page is None:
                        page = context.new_page()
                    node = re.search(r"/(\d+)(?:/|$)", url)[1]
                    result = capture(page, url, args.image_dir / f"{node}.png", args.postal_code)
                    result["status"] = "ok"
                row = item["row"]
                region = next(area for area in ws.merged_cells.ranges if f"F{row}" in area)
                width, height = region_pixels(ws, region)
                # 显示框同时服从单元格和质量门禁上限，截图原始像素不变。
                scale = min(min(width - 6, 420) / result["pixels"][0],
                            min(height - 6, 170) / result["pixels"][1])
                clear_existing_category_images(ws, {row})
                insert_overview_image(ws, row, Path(result["image"]), int(result["pixels"][0] * scale), int(result["pixels"][1] * scale))
                ws.cell(row, 9).value = result["monthly_sales"]
                print(f"[类目] 完成 {item['label']}：月销量 {result['monthly_sales']}", flush=True)
            except Exception as error:
                failed = True
                result = {"status": "failed", "error": str(error)}
                print(f"[类目] 失败 {item['label']}：{error}", flush=True)
            result.update({"row": item["row"], "label": item["label"], "elapsed_seconds": round(time.monotonic() - started, 1)})
            results[url] = result
            manifest_path.write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
            args.output.parent.mkdir(parents=True, exist_ok=True)
            temporary = args.output.with_suffix(".part.xlsx")
            wb.save(temporary)
            temporary.replace(args.output)
            # 保存会关闭已加载图片的字节流；下一项从检查点重新打开以保留全部图片。
            wb.close()
            wb = load_workbook(args.output, rich_text=True)
            ws = next(ws for ws in wb if ws.title.endswith("市场调研"))
    wb.close()
    return int(failed)


if __name__ == "__main__":
    raise SystemExit(main())
