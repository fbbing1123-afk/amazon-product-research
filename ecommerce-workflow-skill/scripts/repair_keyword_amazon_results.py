#!/usr/bin/env python3
"""Fill only missing Amazon result counts in an existing workflow workbook."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font

MODULES_DIR = Path(__file__).resolve().parents[1] / "modules"
sys.path.insert(0, str(MODULES_DIR))

from keyword_trends.sellersprite_playwright_runner import extract_amazon_result


sys.path.insert(0, str((Path(__file__).resolve().parent / "../modules").resolve()))
from shared.browser_runtime import get_browser_port

def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--workbook", required=True)
    parser.add_argument("--port", type=int, default=get_browser_port())
    parser.add_argument("--site", default="US")
    parser.add_argument("--postal-code", default="10001")
    args = parser.parse_args()

    from playwright.sync_api import sync_playwright

    path = Path(args.workbook)
    wb = load_workbook(path)
    ws = next(sheet for sheet in wb.worksheets if "市场调研" in sheet.title)
    result_column = next(cell.column for cell in ws[4] if str(cell.value or "").strip() == "站内结果呈现")
    keyword_rows = []
    for row in range(5, ws.max_row + 1):
        value = str(ws.cell(row, 1).value or "").strip()
        if value.startswith("2."):
            break
        if value and str(ws.cell(row, result_column).value or "").strip() in {"", "/"}:
            keyword_rows.append(row)

    with sync_playwright() as playwright:
        browser = playwright.chromium.connect_over_cdp(f"http://127.0.0.1:{args.port}")
        page = next(
            page
            for context in browser.contexts
            for page in context.pages
            if "sellersprite.com" in (page.url or "") and "/v3/" in (page.url or "")
        )
        for row in keyword_rows:
            keyword = str(ws.cell(row, 1).value).split("\n", 1)[0].strip()
            result = extract_amazon_result(page, keyword, site=args.site, amazon_postal_code=args.postal_code)
            if result.get("amazonResultStatus") != "success":
                print(f"FAILED {keyword}: {result.get('amazonResultError') or result.get('amazonDeliveryZipError')}")
                continue
            ws.cell(row, result_column).value = int(result["amazonResultCount"])
            keyword_cell = ws.cell(row, 1)
            keyword_cell.hyperlink = result.get("amazonResultUrl")
            keyword_cell.font = Font(name="宋体", sz=11, color="0000FF", underline="single")
            print(f"OK {keyword}: {int(result['amazonResultCount']):,}")
        if keyword_rows:
            wb.save(path)
        # 远程 CDP 连接由 sync_playwright 上下文释放，保留项目浏览器供后续任务使用。

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
