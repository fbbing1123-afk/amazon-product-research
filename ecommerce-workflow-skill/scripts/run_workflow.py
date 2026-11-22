from __future__ import annotations

import argparse
import os
import re
import sys
from contextlib import ExitStack
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse


SKILL_DIR = Path(__file__).resolve().parents[1]
VENV_PYTHON = SKILL_DIR / ".venv" / "Scripts" / "python.exe"


def _same_executable(left: Path | str, right: Path | str) -> bool:
    try:
        left_path = os.path.normcase(os.path.realpath(str(left)))
        right_path = os.path.normcase(os.path.realpath(str(right)))
    except OSError:
        return False
    return left_path == right_path


def reexec_with_project_venv() -> None:
    if os.environ.get("ECOMMERCE_WORKFLOW_SKIP_VENV_REEXEC") == "1":
        return
    if not VENV_PYTHON.exists():
        return
    if _same_executable(sys.executable, VENV_PYTHON):
        return

    print(f"[workflow] switching to project Python: {VENV_PYTHON}", flush=True)
    env = os.environ.copy()
    env["ECOMMERCE_WORKFLOW_SKIP_VENV_REEXEC"] = "1"
    env.setdefault("PLAYWRIGHT_BROWSERS_PATH", "0")
    os.execve(str(VENV_PYTHON), [str(VENV_PYTHON), *sys.argv], env)


reexec_with_project_venv()
os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", "0")

MODULES_DIR = SKILL_DIR / "modules"
if str(MODULES_DIR) not in sys.path:
    sys.path.insert(0, str(MODULES_DIR))

from keyword_backfill.run import run_keyword_backfill
from keyword_trends.run import run_keyword_trends
from market_analysis.run import run_market_analysis
from product_research.run import run_category_research, run_product_research_charts
from review_research.run import run_review_research
from shared.excel_inspect import find_source_sheet, inspect_image_risk
from shared.image_conversion import convert_wps_dispimg_to_floating, make_wps_embedded_version
from shared.marketplaces import MARKETPLACES, get_marketplace, normalize_site_code
from shared.paths import ensure_dir, resolve_existing_file
from shared.power import workflow_power_hold
from shared.quality_gates import run_step_gate, validate_zip_package
from shared.runtime_monitor import initialize_monitor, monitor_event


DEFAULT_STEPS = ("keyword_research", "category_research", "market_analysis", "review_research")
LEGACY_FULL_WORKFLOW_STEPS = ("keyword_backfill", "keyword_trends", "category_research", "market_analysis")
AVAILABLE_STEPS = (*DEFAULT_STEPS, "keyword_backfill", "keyword_trends", "product_research")
DEFAULT_FINAL_OUTPUT_DIR_NAME = "亚马逊产品调研流程输出表"
DEFAULT_WORK_CACHE_DIR_NAME = "亚马逊产品调研流程缓存"
DEFAULT_CATEGORY_OVERVIEW_LIMIT = None
MARKET_RESEARCH_SUFFIX = "市场调研"
STEP_OUTPUT_LABELS = {
    "keyword_research": "关键词调研",
    "keyword_backfill": "关键词导出",
    "keyword_trends": "关键词趋势",
    "category_research": "类目调研",
    "product_research": "产品调研",
    "market_analysis": "市场分析",
    "review_research": "评论调研",
}

STEP_ORDER = {
    "keyword_research": 1,
    "keyword_backfill": 1,
    "keyword_trends": 2,
    "category_research": 3,
    "product_research": 3,
    "market_analysis": 4,
    "review_research": 5,
}

ASIN_PATTERN = re.compile(r"(?<![A-Z0-9])([A-Z0-9]{10})(?![A-Z0-9])", re.I)


def parse_steps(value: str) -> list[str]:
    steps = [item.strip() for item in value.split(",") if item.strip()]
    unknown = [step for step in steps if step not in AVAILABLE_STEPS]
    if unknown:
        raise argparse.ArgumentTypeError(f"Unknown step(s): {', '.join(unknown)}")
    if not steps:
        raise argparse.ArgumentTypeError("At least one step is required.")
    return steps


def validate_step_selection(steps: list[str]) -> None:
    """校验步骤组合，避免别名重复执行或按逆序覆盖前一步结果。"""
    if len(steps) != len(set(steps)):
        raise ValueError("步骤列表中不能重复同一个步骤。")
    if "keyword_research" in steps and {
        "keyword_backfill",
        "keyword_trends",
    }.intersection(steps):
        raise ValueError(
            "keyword_research 已包含 keyword_backfill 和 keyword_trends，不能与旧步骤混合重复运行。"
        )
    if "category_research" in steps and "product_research" in steps:
        raise ValueError("category_research 和 product_research 是同一个兼容入口，不能同时运行。")
    order = [STEP_ORDER[step] for step in steps]
    if order != sorted(order):
        raise ValueError(
            "步骤必须按关键词、类目、市场分析、评论的顺序传入；单独运行任一步骤不受影响。"
        )


def inspect_workbook_capabilities(workbook_path: Path) -> dict[str, bool]:
    """只读检查工作簿是否具备各模块的最小输入结构。"""
    from openpyxl import load_workbook

    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        source_sheet = find_source_sheet(workbook)
        has_market_sheet = any("市场调研" in name for name in workbook.sheetnames)
        has_keyword_header = False
        has_asin_header = False
        for worksheet in workbook.worksheets:
            for row in worksheet.iter_rows(
                min_row=1,
                max_row=min(12, worksheet.max_row),
                values_only=True,
            ):
                normalized = {
                    re.sub(r"[\s_\-]+", "", str(value or "")).lower()
                    for value in row
                }
                has_keyword_header = has_keyword_header or "关键词" in normalized
                has_asin_header = has_asin_header or bool(
                    {"asin", "竞品asin"}.intersection(normalized)
                )
        return {
            "has_source_sheet": source_sheet is not None,
            "has_market_sheet": has_market_sheet,
            "has_keyword_header": has_keyword_header,
            "has_asin_header": has_asin_header,
        }
    finally:
        workbook.close()


def validate_step_input(step: str, workbook_path: Path, args) -> None:
    """在启动浏览器或子进程前验证当前步骤的工作簿输入。"""
    capabilities = inspect_workbook_capabilities(workbook_path)
    if step in {"keyword_research", "keyword_backfill", "keyword_trends"}:
        if not capabilities["has_market_sheet"]:
            raise ValueError(
                f"{step} 需要包含“市场调研”的工作表：{workbook_path}"
            )
        if step == "keyword_trends" and not capabilities["has_keyword_header"]:
            raise ValueError("keyword_trends 需要在市场调研页找到“关键词”表头。")
    elif step in {"category_research", "product_research"}:
        if not capabilities["has_source_sheet"]:
            raise ValueError(
                f"{step} 需要数据源工作表（产品名-数据源、数据源、US 或 Sheet1）：{workbook_path}"
            )
        template_path = (
            resolve_existing_file(args.research_template, "research template")
            if args.research_template
            else workbook_path
        )
        template_capabilities = inspect_workbook_capabilities(template_path)
        if not template_capabilities["has_market_sheet"]:
            raise ValueError(
                f"{step} 需要研究模板包含“市场调研”工作表：{template_path}"
            )
    elif step == "market_analysis" and not capabilities["has_source_sheet"]:
        raise ValueError(
            f"market_analysis 需要数据源工作表（产品名-数据源、数据源、US 或 Sheet1）：{workbook_path}"
        )
    elif step == "review_research":
        if not args.review_asin and not capabilities["has_asin_header"]:
            raise ValueError(
                "review_research 没有找到 ASIN 表头；请在工作簿补充 ASIN，或传入 --review-asin。"
            )
    monitor_event(
        "step_preflight_passed",
        step=step,
        workbook=str(workbook_path),
        capabilities=", ".join(
            key for key, available in capabilities.items() if available
        ),
    )


def extract_competitor_asins(workbook_path: Path) -> list[str]:
    """从竞品分析页顶部和详情链接中按展示顺序提取竞品 ASIN。"""
    from openpyxl import load_workbook

    workbook = load_workbook(workbook_path, read_only=False, data_only=True)
    try:
        worksheet = next(
            (sheet for sheet in workbook.worksheets if sheet.title.endswith("竞品分析")),
            None,
        )
        if worksheet is None:
            return []

        found: list[str] = []
        seen: set[str] = set()

        def add_asins(value: object) -> None:
            for match in ASIN_PATTERN.finditer(str(value or "").upper()):
                asin = match.group(1)
                if asin not in seen:
                    seen.add(asin)
                    found.append(asin)

        # 模板顶部标题通常已经按竞品列顺序列出 ASIN。
        for row in worksheet.iter_rows(min_row=1, max_row=min(5, worksheet.max_row)):
            for cell in row:
                if cell.hyperlink and cell.hyperlink.target:
                    add_asins(cell.hyperlink.target)
                add_asins(cell.value)
        return found
    finally:
        workbook.close()


def infer_site_from_workbook(workbook_path: Path) -> str:
    """从数据源中的 Amazon 域名推断站点；多站点混用时拒绝猜测。"""
    from openpyxl import load_workbook

    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        source_sheet = next(
            (
                sheet
                for sheet in workbook.worksheets
                if sheet.title.endswith("数据源")
            ),
            None,
        )
        if source_sheet is None:
            source_sheet = next(
                (
                    workbook[name]
                    for name in ("数据源", "US", "DE", "UK", "Sheet1")
                    if name in workbook.sheetnames
                ),
                None,
            )
        if source_sheet is None:
            source_sheet = workbook.worksheets[0]

        domain_to_site = {
            marketplace.amazon_domain: site
            for site, marketplace in MARKETPLACES.items()
        }
        detected_sites: set[str] = set()
        for row in source_sheet.iter_rows(values_only=True):
            for value in row:
                text = str(value or "")
                if "amazon." not in text.lower():
                    continue
                host = (urlparse(text).hostname or "").lower()
                host = host.removeprefix("www.")
                if host in domain_to_site:
                    detected_sites.add(domain_to_site[host])

        if len(detected_sites) > 1:
            sites = ", ".join(sorted(detected_sites))
            raise ValueError(
                f"数据源检测到多个 Amazon 站点（{sites}），请通过 --site 明确指定，不能自动混合调研。"
            )
        if detected_sites:
            return next(iter(detected_sites))
        return "US"
    finally:
        workbook.close()


def inspect_workflow_sheets(workbook_path: Path) -> dict[str, bool]:
    """检查脚本阶段和竞品分析阶段的关键工作表是否已经生成。"""
    from openpyxl import load_workbook

    expected = {
        "评论抓取摘要": False,
        "商品评论": False,
        "差评痛点": False,
        "竞品分析": False,
        "差评透视": False,
    }
    workbook = load_workbook(workbook_path, read_only=True, data_only=False)
    try:
        sheet_names = set(workbook.sheetnames)
        for sheet_name in expected:
            expected[sheet_name] = sheet_name in sheet_names or any(
                name.endswith(sheet_name) for name in sheet_names
            )
        return expected
    finally:
        workbook.close()


def sanitize_filename_stem(value: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', " ", value)
    cleaned = re.sub(r"\s+", " ", cleaned).strip().strip(".")
    cleaned = re.sub(r"[-_\s]+$", "", cleaned)
    return cleaned or "亚马逊产品调研"


def extract_date_suffix(source: Path) -> str:
    match = re.search(r"(\d{6})(?!\d)", source.stem)
    if match:
        return match.group(1)
    return datetime.now().strftime("%y%m%d")


def infer_product_name_from_workbook(source: Path) -> str | None:
    try:
        from openpyxl import load_workbook
    except Exception:
        return None

    try:
        wb = load_workbook(source, read_only=True, data_only=True)
    except Exception:
        return None

    try:
        for sheet_name in wb.sheetnames:
            if sheet_name.endswith(MARKET_RESEARCH_SUFFIX):
                product_name = sheet_name[: -len(MARKET_RESEARCH_SUFFIX)].strip()
                product_name = re.sub(r"[-_\s]+$", "", product_name)
                if product_name and product_name not in {"产品名称", "产品名"}:
                    return product_name
    finally:
        wb.close()
    return None


def default_output_name(source: Path) -> str:
    product_name = infer_product_name_from_workbook(source)
    if not product_name:
        product_name = re.sub(r"-?产品调研-?\d{6}$", "", source.stem).strip("-_ ")
    product_name = sanitize_filename_stem(product_name)
    return f"{product_name}-产品调研-{extract_date_suffix(source)}"


def default_final_output_path(name: str) -> Path:
    desktop = Path.home() / "Desktop"
    output_dir = desktop / DEFAULT_FINAL_OUTPUT_DIR_NAME
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir / f"{name}.xlsx"


def default_work_cache_dir() -> Path:
    desktop = Path.home() / "Desktop"
    return desktop / DEFAULT_WORK_CACHE_DIR_NAME


def default_partial_output_path(name: str, steps: list[str]) -> Path:
    desktop = Path.home() / "Desktop"
    product_name = re.sub(r"-?产品调研-?\d{6}$", "", name).strip("-_ ")
    product_name = sanitize_filename_stem(product_name)
    label = "、".join(STEP_OUTPUT_LABELS.get(step, step) for step in steps)
    return desktop / f"{product_name}-{label}.xlsx"


def run_keyword_research_step(*, current_workbook: Path, name: str, output_dir: Path, args, workflow_site: str, amazon_postal_code: str) -> dict[str, Path]:
    results: dict[str, Path] = {}
    export_file = (
        resolve_existing_file(args.keyword_backfill_export_file, "keyword backfill export")
        if args.keyword_backfill_export_file
        else None
    )
    results.update(
        run_keyword_backfill(
            source=current_workbook,
            name=name,
            output_dir=output_dir,
            python_exe=args.python_exe,
            export_file=export_file,
            skip_browser=args.keyword_backfill_skip_browser,
            excel_writer=args.keyword_backfill_excel_writer,
            debugger_address=args.keyword_backfill_debugger_address,
            site=workflow_site,
            brand_words=args.keyword_backfill_brand_word,
        )
    )
    results.update(
        run_keyword_trends(
            source=results["current_workbook"],
            name=name,
            output_dir=output_dir,
            python_exe=args.python_exe,
            prefix=args.keyword_trends_prefix,
            ppc_bid_mode=args.keyword_ppc_bid_mode,
            skip_screenshot=args.keyword_trends_skip_screenshot,
            keep_run_artifacts=args.keyword_trends_keep_artifacts,
            debug_port=args.keyword_trends_debug_port,
            use_existing_browser=args.keyword_trends_use_existing_browser,
            site=workflow_site,
            amazon_postal_code=amazon_postal_code,
        )
    )
    return results


def _main(power_hold_stack: ExitStack) -> int:
    parser = argparse.ArgumentParser(description="Run modular ecommerce Excel workflow.")
    parser.add_argument("--source", required=True, help="Amazon/SellerSprite source .xlsx")
    parser.add_argument("--research-template", help="Product research template .xlsx")
    parser.add_argument("--name", help="Output file prefix/name. If omitted, infer from the workbook market-research sheet.")
    parser.add_argument("--output-dir", help="Internal cache directory. Defaults to a desktop cache folder.")
    parser.add_argument("--final-output", help="Final combined workbook path")
    parser.add_argument("--steps", type=parse_steps, default=list(DEFAULT_STEPS), help="Comma-separated steps in execution order")
    parser.add_argument(
        "--continue-full-workflow",
        action="store_true",
        help="完整流程在登录或局部失败后续跑时使用；评论阶段结束后继续交接竞品分析，不将本次续跑误记为最终完成",
    )
    parser.add_argument("--python", dest="python_exe", default=str(VENV_PYTHON if VENV_PYTHON.exists() else sys.executable), help="Python executable")
    parser.add_argument("--node", dest="node_exe", default=os.environ.get("NODE_EXE", "node"), help="Node executable")
    parser.add_argument(
        "--site",
        default="AUTO",
        help="SellerSprite/Amazon marketplace: AUTO, US, DE, UK, CA or FR; AUTO reads the source Amazon domain",
    )
    parser.add_argument("--amazon-postal-code", help="Override the marketplace default Amazon delivery postal code")
    parser.add_argument("--keyword-backfill-export-file", help="Existing SellerSprite expand-keywords export .xlsx")
    parser.add_argument("--keyword-backfill-skip-browser", action="store_true", help="Skip SellerSprite browser automation and process --keyword-backfill-export-file")
    parser.add_argument("--keyword-backfill-excel-writer", default="openpyxl", choices=("openpyxl", "excel_com"), help="Workbook writer for keyword backfill")
    parser.add_argument("--keyword-backfill-debugger-address", default="127.0.0.1:9223", help="Edge/Chrome CDP debugger address for keyword backfill")
    parser.add_argument("--keyword-backfill-site", default=None, help="Compatibility alias for --site")
    parser.add_argument("--keyword-backfill-brand-word", action="append", default=None, help="Temporary brand word to exclude; repeat as needed")
    parser.add_argument("--keyword-trends-prefix", default="excel-trends", help="Keyword trend screenshot file prefix")
    parser.add_argument("--keyword-ppc-bid-mode", choices=("latest-month", "peak-period"), default="latest-month", help="竞价默认取最新月，可选保留旧旺季区间逻辑")
    parser.add_argument("--keyword-trends-debug-port", type=int, default=9223, help="Chrome/Edge DevTools port for keyword trend screenshots")
    parser.add_argument(
        "--keyword-trends-use-existing-browser",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="默认复用已打开的项目 DevTools 浏览器；使用 --no-keyword-trends-use-existing-browser 才启动独立浏览器",
    )
    parser.add_argument("--keyword-trends-skip-screenshot", action="store_true", help="Use existing keyword trend PNG/JSON artifacts instead of opening Sellersprite")
    parser.add_argument("--keyword-trends-keep-artifacts", action="store_true", help="Keep keyword trend PNG/JSON run artifacts")
    parser.add_argument("--category-overviews", action=argparse.BooleanOptionalAction, default=True, help="Capture softtime/sorftime market overview screenshots for category rows")
    parser.add_argument("--category-overviews-debugger-address", default="127.0.0.1:9223", help="Chrome/Edge CDP debugger address for category overview screenshots")
    parser.add_argument(
        "--category-overviews-limit",
        type=int,
        default=DEFAULT_CATEGORY_OVERVIEW_LIMIT,
        help="Maximum deduplicated category overview captures; omit to capture every category",
    )
    parser.add_argument("--market-analysis-charts", action=argparse.BooleanOptionalAction, default=True, help="Generate and insert the 9 market-research chart images before market_analysis")
    parser.add_argument("--market-analysis-source-slicers", action=argparse.BooleanOptionalAction, default=False, help="Optionally add source-data market-layer slicers; disabled by default")
    parser.add_argument("--review-debugger-address", default="127.0.0.1:9223", help="Chrome/Edge CDP debugger address for Amazon review research")
    parser.add_argument("--review-asin", action="append", default=[], help="Explicit ASIN for review research; repeat as needed. Overrides workbook ASINs.")
    parser.add_argument("--review-asin-limit", type=int, help="Limit workbook ASINs for review research")
    parser.add_argument("--review-max-pages", type=int, default=10, help="Maximum 10-review batches per ASIN; kept as pages for compatibility")
    parser.add_argument("--review-max-reviews", type=int, help="Compatibility override for both positive and critical review limits")
    parser.add_argument("--review-max-positive-reviews", type=int, default=50, help="Maximum positive reviews per ASIN")
    parser.add_argument("--review-max-critical-reviews", type=int, default=100, help="Maximum critical reviews per ASIN")
    parser.add_argument("--review-sort-by", choices=("recent", "helpful"), default="helpful")
    parser.add_argument("--review-reviewer-type", choices=("default", "verified", "all"), default="default")
    parser.add_argument("--review-star-filter", choices=("both", "all_stars", "positive", "critical", "five_star", "four_star", "three_star", "two_star", "one_star"), default="both")
    parser.add_argument("--preserve-wps-images", action=argparse.BooleanOptionalAction, default=True, help="Convert WPS DISPIMG cell images to floating images before openpyxl processing")
    parser.add_argument("--wps-embedded-output", help="Optional extra output path that restores WPS DISPIMG cell images after the workflow")
    parser.add_argument("--wps-embed-extra-sheet", action="append", default=[], help="For --wps-embedded-output, also embed floating images on this sheet. Can be repeated.")
    parser.add_argument("--force", action="store_true", help="Continue despite non-fatal warnings")
    args = parser.parse_args()
    validate_step_selection(args.steps)

    source = resolve_existing_file(args.source, "source")
    name = args.name or default_output_name(source)
    output_dir = ensure_dir(args.output_dir) if args.output_dir else ensure_dir(default_work_cache_dir())
    requested_site = args.keyword_backfill_site or args.site
    if str(requested_site or "").strip().upper() in {"AUTO", "自动", "自动识别"}:
        workflow_site = infer_site_from_workbook(source)
        print(f"[workflow] site auto-detected from source: {workflow_site}")
    else:
        workflow_site = normalize_site_code(requested_site)
    marketplace = get_marketplace(workflow_site)
    amazon_postal_code = args.amazon_postal_code or marketplace.postal_code

    print(f"[workflow] source: {source}")
    print(f"[workflow] name: {name}")
    print(f"[workflow] output_dir: {output_dir}")
    print(f"[workflow] steps: {', '.join(args.steps)}")
    print(f"[workflow] python: {args.python_exe}")
    print(f"[workflow] node: {args.node_exe}")
    print(f"[workflow] site: {marketplace.code} ({marketplace.amazon_domain}), amazon_postal_code: {amazon_postal_code}")

    monitor_jsonl = output_dir / f"{name}_workflow_monitor.jsonl"
    monitor_log = output_dir / f"{name}_workflow_monitor.log"
    initialize_monitor(monitor_jsonl, monitor_log)
    power_hold_stack.enter_context(workflow_power_hold())
    monitor_event(
        "workflow_started",
        source=str(source),
        name=name,
        site=marketplace.code,
        postal_code=amazon_postal_code,
        steps=", ".join(args.steps),
    )
    print(f"[workflow] monitor jsonl: {monitor_jsonl}")
    print(f"[workflow] monitor log: {monitor_log}")

    results: dict[str, Path] = {}
    current_workbook = source
    requested_steps = tuple(args.steps)
    # 续跑可能只传入 review_research，但它仍属于之前已启动的完整流程。
    # 单独运行评论模块不带此标记，因此不会被强行串入竞品分析。
    is_full_workflow = requested_steps == DEFAULT_STEPS
    continues_full_workflow = bool(args.continue_full_workflow)
    if continues_full_workflow and requested_steps != ("review_research",):
        raise ValueError("--continue-full-workflow 只允许与单独的 review_research 续跑组合使用。")
    needs_competitor_handoff = is_full_workflow or continues_full_workflow
    if args.final_output:
        final_output = Path(args.final_output).expanduser().resolve()
    elif is_full_workflow:
        final_output = default_final_output_path(name)
    else:
        final_output = default_partial_output_path(name, args.steps)
    original_workbook = source

    category_research_only = tuple(args.steps) in {("category_research",), ("product_research",)}
    if args.preserve_wps_images and not category_research_only:
        risk = inspect_image_risk(current_workbook)
        if risk.has_dispimg:
            preprocessed = output_dir / f"{name}_wps_images_floating.xlsx"
            print(f"[workflow] converting WPS DISPIMG images to floating images: {preprocessed}")
            convert_wps_dispimg_to_floating(current_workbook, preprocessed)
            current_workbook = preprocessed
            results["wps_images_floating_source"] = preprocessed

    for step in args.steps:
        monitor_event("step_dispatched", step=step)
        validate_step_input(step, current_workbook, args)
        if step == "keyword_research":
            results.update(
                run_keyword_research_step(
                    current_workbook=current_workbook,
                    name=name,
                    output_dir=output_dir,
                    args=args,
                    workflow_site=workflow_site,
                    amazon_postal_code=amazon_postal_code,
                )
            )
            current_workbook = results["current_workbook"]

        elif step == "keyword_backfill":
            export_file = (
                resolve_existing_file(args.keyword_backfill_export_file, "keyword backfill export")
                if args.keyword_backfill_export_file
                else None
            )
            results.update(
                run_keyword_backfill(
                    source=current_workbook,
                    name=name,
                    output_dir=output_dir,
                    python_exe=args.python_exe,
                    export_file=export_file,
                    skip_browser=args.keyword_backfill_skip_browser,
                    excel_writer=args.keyword_backfill_excel_writer,
                    debugger_address=args.keyword_backfill_debugger_address,
                    site=workflow_site,
                    brand_words=args.keyword_backfill_brand_word,
                )
            )
            current_workbook = results["current_workbook"]

        elif step == "keyword_trends":
            results.update(
                run_keyword_trends(
                    source=current_workbook,
                    name=name,
                    output_dir=output_dir,
                    python_exe=args.python_exe,
                    prefix=args.keyword_trends_prefix,
                    ppc_bid_mode=args.keyword_ppc_bid_mode,
                    skip_screenshot=args.keyword_trends_skip_screenshot,
                    keep_run_artifacts=args.keyword_trends_keep_artifacts,
                    debug_port=args.keyword_trends_debug_port,
                    use_existing_browser=args.keyword_trends_use_existing_browser,
                    site=workflow_site,
                    amazon_postal_code=amazon_postal_code,
                )
            )
            current_workbook = results["current_workbook"]

        elif step in {"category_research", "product_research"}:
            template = resolve_existing_file(args.research_template, "research template") if args.research_template else current_workbook
            results.update(
                run_category_research(
                    source=current_workbook,
                    template=template,
                    name=name,
                    output_dir=output_dir,
                    python_exe=args.python_exe,
                    node_exe=args.node_exe,
                    force=args.force,
                    capture_category_overviews=args.category_overviews,
                    category_overview_debugger_address=args.category_overviews_debugger_address,
                    category_overview_limit=args.category_overviews_limit,
                    site=workflow_site,
                    amazon_postal_code=amazon_postal_code,
                )
            )
            current_workbook = results["product_research_report"]
            results["current_workbook"] = current_workbook

        elif step == "market_analysis":
            if args.market_analysis_charts:
                results.update(
                    run_product_research_charts(
                        source=current_workbook,
                        name=name,
                        output_dir=output_dir,
                        python_exe=args.python_exe,
                        node_exe=args.node_exe,
                    )
                )
                current_workbook = results["current_workbook"]
            results.update(
                run_market_analysis(
                    source=current_workbook,
                    name=name,
                    output_dir=output_dir,
                    python_exe=args.python_exe,
                    output_path=final_output,
                    source_slicers=args.market_analysis_source_slicers,
                    original_image_source=original_workbook,
                )
            )
            current_workbook = results["current_workbook"]

        elif step == "review_research":
            review_asins = list(args.review_asin)
            if not review_asins:
                review_asins = extract_competitor_asins(current_workbook)
                if not review_asins:
                    if needs_competitor_handoff:
                        raise ValueError(
                            "完整流程需要从竞品分析页读取目标 ASIN，但当前工作簿没有找到竞品 ASIN。"
                        )
                else:
                    print("[workflow] competitor review ASINs: " + ", ".join(review_asins))
            results.update(
                run_review_research(
                    source=current_workbook,
                    name=name,
                    output_dir=output_dir,
                    python_exe=args.python_exe,
                    site=workflow_site,
                    amazon_postal_code=amazon_postal_code,
                    debugger_address=args.review_debugger_address,
                    asins=review_asins,
                    asin_limit=args.review_asin_limit,
                    max_pages=args.review_max_pages,
                    max_reviews=args.review_max_reviews,
                    max_positive_reviews=args.review_max_positive_reviews,
                    max_critical_reviews=args.review_max_critical_reviews,
                    sort_by=args.review_sort_by,
                    reviewer_type=args.review_reviewer_type,
                    star_filter=args.review_star_filter,
                )
            )
            current_workbook = results["current_workbook"]

        if step in {
            "keyword_research",
            "category_research",
            "product_research",
            "market_analysis",
            "review_research",
        }:
            gate_report = run_step_gate(step, current_workbook)
            monitor_event(
                "step_quality_gate_passed",
                step=step,
                gate=gate_report.gate,
                workbook=str(current_workbook),
                metrics=str(gate_report.metrics),
            )
            print(
                f"[quality_gate] {step} passed: {gate_report.metrics}",
                flush=True,
            )

    # 最终交付必须来自最后一个成功阶段，不能停留在中途的市场分析版本。
    if current_workbook != final_output:
        final_output.parent.mkdir(parents=True, exist_ok=True)
        final_output.write_bytes(current_workbook.read_bytes())
        current_workbook = final_output
        results["current_workbook"] = current_workbook

    results["final_workbook"] = current_workbook
    results["workflow_monitor_jsonl"] = monitor_jsonl
    results["workflow_monitor_log"] = monitor_log

    if args.wps_embedded_output:
        embedded_output = Path(args.wps_embedded_output).expanduser().resolve()
        temp_output = embedded_output.with_name(embedded_output.stem + "_temp.xlsx")
        print(f"[workflow] creating WPS embedded-cell-image output: {embedded_output}")
        make_wps_embedded_version(
            original=original_workbook,
            floating_workbook=current_workbook,
            output=embedded_output,
            temp_output=temp_output,
            embed_extra_sheets=set(args.wps_embed_extra_sheet),
        )
        zip_report = validate_zip_package(embedded_output)
        monitor_event(
            "wps_package_quality_gate_passed",
            workbook=str(embedded_output),
            gate=zip_report.gate,
        )
        results["wps_embedded_workbook"] = embedded_output

    if needs_competitor_handoff:
        results["agent_handoff_workbook"] = current_workbook
        detected_sheets = inspect_workflow_sheets(current_workbook)
        missing_sheets = [name for name, present in detected_sheets.items() if not present]
        monitor_event(
            "scripted_stages_finished",
            final_workbook=str(current_workbook),
            detected_sheets=", ".join(
                name for name, present in detected_sheets.items() if present
            ),
            missing_sheets=", ".join(missing_sheets),
        )
        monitor_event(
            "workflow_handoff_required",
            final_workbook=str(current_workbook),
            required_stage="competitor_analysis",
            required_outputs="竞品分析、差评透视、3 个橙色切片器",
            reason=(
                "完整流程脚本阶段不会在 Python 进程内调用专用竞品分析子智能体"
                if is_full_workflow
                else "完整流程续跑已完成评论阶段，必须立即接续竞品分析"
            ),
            continuation="主智能体完成 handoff JSON 后立即调用 modules\\competitor_analysis\\run.py；不得结束完整流程或重新启动四个脚本阶段",
        )
        print("[workflow] scripted stages completed")
        print(
            "[workflow] handoff required: complete competitor analysis, review translation, "
            "pain-point grouping, pain pivot, and slicers"
        )
        print(
            "[workflow] continuation required: after filling the handoff JSON, "
            "run modules\\competitor_analysis\\run.py immediately and reuse this monitor."
        )
    else:
        monitor_event("workflow_finished", final_workbook=str(current_workbook))
        print("[workflow] completed")
    for key, path in results.items():
        print(f"[workflow] {key}: {path}")
    return 0


def main() -> int:
    """运行完整工作流，并确保退出时恢复 Windows 电源状态。"""
    with ExitStack() as power_hold_stack:
        return _main(power_hold_stack)


if __name__ == "__main__":
    raise SystemExit(main())
