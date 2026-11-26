"""复用已完成的趋势报告，只刷新最新周/月竞价并按旺季搜索量排序。"""

from __future__ import annotations

import argparse
import importlib.util
import json
import sys
import time
from pathlib import Path

MODULE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(MODULE_DIR))
sys.path.insert(0, str(MODULE_DIR.parent))

from openpyxl import load_workbook
from keyword_rows import latest_monthly_bid, latest_weekly_bid, update_keyword_rows
import sellersprite_playwright_runner as runner
from shared.browser_runtime import get_browser_port
from shared.marketplaces import ensure_sellersprite_site
from shared.quality_gates import validate_keyword_research


def load_trend_module():
    spec = importlib.util.spec_from_file_location("keyword_trends_refresh", MODULE_DIR / "excel-sellersprite-trends.py")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def collect_latest_bid(page, mode: str) -> dict:
    """先切换到指定周期，再读取该周期最新一期的数据。"""
    monthly = mode == "latest-month"
    label = "月度" if monthly else "周度"
    field = "ppcMonthly" if monthly else "ppcWeekly"
    runner.open_first_row_monthly_drawer(page)
    runner.click_ppc_bid_tab(page)
    if not runner.set_ppc_history_mode(page, monthly=monthly):
        raise RuntimeError(f"无法切换到{label} PPC 数据。")
    deadline = time.monotonic() + 15
    while True:
        result = runner.extract_ppc_data(page)
        if result.get("ppcStatus") == "blocked":
            raise runner.ManualActionRequired(result.get("ppcError", "竞价读取受限"), stage="latest-ppc")
        points = result.get(field, [])
        if points:
            period, bid = latest_monthly_bid(points) if monthly else latest_weekly_bid(points)
            if period is not None:
                return {"period": period, "bid": bid, field: points}
        if time.monotonic() >= deadline:
            raise RuntimeError(f"未读取到带有效日期的{label} PPC 数据，未使用旧值或其它周期代替。")
        page.wait_for_timeout(500)


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--cache", type=Path, required=True)
    parser.add_argument("--site", default="US")
    parser.add_argument("--debug-port", type=int, default=get_browser_port())
    parser.add_argument("--ppc-bid-mode", choices=("latest-week", "latest-month"), default="latest-week")
    parser.add_argument("--resume", action="store_true", help="显式复用本次缓存，仅补未成功的关键词")
    args = parser.parse_args(argv)
    if args.source.resolve() == args.output.resolve():
        raise ValueError("输出必须使用独立路径，不能覆盖原始工作簿。")
    module = load_trend_module()
    workbook = load_workbook(args.source, rich_text=True)
    try:
        sheet = module.find_market_research_sheet(workbook)
        header = module.find_keyword_header(sheet)
        headers = module.find_header_cells_in_row(sheet, header.row, ["搜索趋势", "旺季月搜索量", "旺季竞价"])
        items = module.extract_keywords(sheet, header, headers["搜索趋势"].column)
        sheet_name = sheet.title
        peak_col, bid_col = headers["旺季月搜索量"].column, headers["旺季竞价"].column
    finally:
        workbook.close()
    if not items:
        raise ValueError("没有可处理的关键词。")
    payload = {"site": args.site, "mode": args.ppc_bid_mode, "results": {}}
    if args.resume and args.cache.exists():
        payload = json.loads(args.cache.read_text(encoding="utf-8"))
        if payload.get("site") != args.site:
            raise ValueError("缓存站点不符，不能跨站点复用。")
        if payload.get("mode") != args.ppc_bid_mode:
            raise ValueError("缓存取数周期不符，不能把月度竞价用于周度任务。")
    args.cache.parent.mkdir(parents=True, exist_ok=True)
    with runner.require_sync_playwright()() as playwright:
        browser, _context, page = runner.connect_cdp(playwright, args.debug_port)
        try:
            runner.navigate_sellersprite_tool(page, runner.SELLERSPRITE_URL, "关键词挖掘")
            runner.sleep(page, 1200)
            if not runner.is_logged_in(page):
                raise runner.ManualActionRequired("请在项目浏览器中登录卖家精灵。", reason="login")
            ensure_sellersprite_site(page, args.site)
            for index, item in enumerate(items, 1):
                if payload["results"].get(item.keyword, {}).get("status") == "success":
                    print(f"[{index}/{len(items)}] 复用已成功竞价：{item.keyword}", flush=True)
                    continue
                print(f"[{index}/{len(items)}] 读取 {args.ppc_bid_mode} 竞价：{item.keyword}", flush=True)
                runner.close_drawer_if_open(page)
                runner.search_keyword(page, item.keyword)
                result = collect_latest_bid(page, args.ppc_bid_mode)
                result.update(status="success", keyword=item.keyword)
                payload["results"][item.keyword] = result
                args.cache.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
                page.screenshot(path=str(args.cache.parent / f"{args.ppc_bid_mode}-ppc-{index:02d}.png"))
                print(f"[{index}/{len(items)}] {item.keyword}：{result['period']}，竞价 {result['bid'] if result['bid'] is not None else '/'}", flush=True)
                runner.close_drawer_if_open(page)
        finally:
            runner.disconnect_browser(browser)
    result = update_keyword_rows(
        args.source, args.output, sheet_name=sheet_name,
        rows=[module.keyword_item_row(item) for item in items], peak_column=peak_col, bid_column=bid_col,
        bids={module.keyword_item_row(item): payload["results"][item.keyword]["bid"] for item in items},
    )
    gate = validate_keyword_research(args.output)
    print(json.dumps({"output": str(args.output), "sort": result, "quality_gate": gate.as_dict()}, ensure_ascii=False), flush=True)
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
