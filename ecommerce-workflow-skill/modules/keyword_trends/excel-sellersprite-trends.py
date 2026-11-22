#!/usr/bin/env python3
"""Read keywords from an Excel market research sheet and insert Sellersprite trend screenshots."""

from __future__ import annotations

import argparse
import calendar
import json
import os
import re
import shutil
import socket
import subprocess
import sys
import time
import urllib.request
from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path
from statistics import mean
from typing import Iterable, Sequence

from openpyxl import load_workbook
from openpyxl.cell.cell import Cell
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# 让脚本从任意工作目录启动时都能找到同目录的浏览器驱动和 WPS 工具。
SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import sellersprite_playwright_runner as browser_runner
from wps_cell_images import convert_floating_images_to_wps_cell_images
from keyword_rows import latest_monthly_bid, latest_weekly_bid, update_keyword_rows

WORKFLOW_MODULES_DIR = Path(__file__).resolve().parents[1]
if str(WORKFLOW_MODULES_DIR) not in sys.path:
    sys.path.insert(0, str(WORKFLOW_MODULES_DIR))

from shared.marketplaces import amazon_domains, get_marketplace, normalize_site_code

DEFAULT_PROJECT_DIR = Path(__file__).resolve().parent
DEFAULT_EXCEL = DEFAULT_PROJECT_DIR / "input.xlsx"
DEFAULT_OUTPUT = DEFAULT_PROJECT_DIR / "output_with_trends.xlsx"
DEFAULT_TEMP_KEYWORDS = ".excel-keywords.txt"
DEFAULT_DATA_OUTPUT = ".excel-trends-data.json"
DEFAULT_PREFIX = "excel-trends"
TREND_FIELD_TITLES = ("趋势旺季", "淡季月搜索量", "旺季月搜索量", "旺季竞价")
AMAZON_RESULT_FIELD_TITLES = ("站内搜索结果", "站内结果呈现")
KEYWORD_LINK_FONT = Font(name="宋体", sz=11, color="0000FF", underline="single")


def find_browser_executable() -> Path:
    """查找当前电脑可用于 CDP 自动化的 Chrome 或 Edge。"""
    candidates = (
        Path(os.environ.get("PROGRAMFILES", "")) / "Google/Chrome/Application/chrome.exe",
        Path(os.environ.get("PROGRAMFILES(X86)", "")) / "Google/Chrome/Application/chrome.exe",
        Path(os.environ.get("LOCALAPPDATA", "")) / "Google/Chrome/Application/chrome.exe",
        Path(os.environ.get("PROGRAMFILES", "")) / "Microsoft/Edge/Application/msedge.exe",
        Path(os.environ.get("PROGRAMFILES(X86)", "")) / "Microsoft/Edge/Application/msedge.exe",
    )
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    for command_name in ("chrome", "chrome.exe", "msedge", "msedge.exe"):
        resolved = shutil.which(command_name)
        if resolved:
            return Path(resolved)
    return candidates[0]


CHROME_PATH = find_browser_executable()
CHROME_USER_DATA_DIR = ".chrome-profile"
BROWSER_STATE_FILE = ".browser-state.json"
SELLERSPRITE_URL = "https://www.sellersprite.com/v3/keyword-miner/"
DEBUG_PORT = 9222
ACTIVE_DEBUG_PORT = DEBUG_PORT
NO_TREND_MAX_MIN_RATIO = 1.5
NO_TREND_RANGE_AVG_RATIO = 0.35
MIN_VALID_SEASONALITY_SEARCH_VOLUME = 2
PEAK_AVG_MULTIPLIER = 1.20
PEAK_MAX_RATIO = 0.65
BURST_MAX_RATIO = 0.90
BURST_AVG_MULTIPLIER = 1.35
GIT_BASH_CANDIDATES = (
    Path(r"E:\Git\bin\bash.exe"),
    Path(r"E:\Git\usr\bin\bash.exe"),
    Path(r"C:\Program Files\Git\bin\bash.exe"),
    Path(r"C:\Program Files\Git\usr\bin\bash.exe"),
)


@dataclass(frozen=True)
class KeywordItem:
    keyword: str
    keyword_cell: str
    image_cell: str


@dataclass(frozen=True)
class ScreenshotRunResult:
    returncode: int
    stdout: str
    stderr: str
    command: list[str]


@dataclass(frozen=True)
class MonthlyTrendPoint:
    month: str
    search_volume: int


@dataclass(frozen=True)
class SeasonalityResult:
    peak_label: str
    low_search_volume: int | None
    peak_search_volume: int | None
    peak_months: list[int]
    burst_months: list[int]
    peak_bid: float | None = None



def cell_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def find_market_research_sheet(workbook):
    for sheet_name in workbook.sheetnames:
        if "市场调研" in sheet_name:
            return workbook[sheet_name]
    raise ValueError("找不到名称包含“市场调研”的工作表")


def find_keyword_header(worksheet) -> Cell:
    for row in worksheet.iter_rows():
        for cell in row:
            if cell_text(cell.value) == "关键词":
                return cell
    raise ValueError(f"工作表 {worksheet.title} 中找不到文本等于“关键词”的标题单元格")


def find_header_cells_in_row(worksheet, row: int, titles: Sequence[str]) -> dict[str, Cell]:
    remaining = set(titles)
    found: dict[str, Cell] = {}
    for cell in worksheet[row]:
        text = cell_text(cell.value)
        if text in remaining:
            found[text] = cell
            remaining.remove(text)
    if remaining:
        missing = "、".join(titles for titles in titles if titles in remaining)
        raise ValueError(f"工作表 {worksheet.title} 第 {row} 行缺少标题: {missing}")
    return found


def find_optional_header_cell_in_row(worksheet, row: int, titles: str | Sequence[str]) -> Cell | None:
    if isinstance(titles, str):
        targets = {titles}
    else:
        targets = set(titles)
    for cell in worksheet[row]:
        if cell_text(cell.value) in targets:
            return cell
    return None


def keyword_item_row(item: KeywordItem) -> int:
    match = re.search(r"\d+", item.keyword_cell)
    if not match:
        raise ValueError(f"无法从关键词单元格解析行号: {item.keyword_cell}")
    return int(match.group(0))


def is_section_title(text: str) -> bool:
    normalized = text.strip()
    if re.match(r"^\d+[\.．、]\s*\S+", normalized):
        return True
    section_markers = (
        "所属平台分类",
        "数据分析概览",
        "品牌分析",
        "上架时间分析",
        "Review数量分析",
        "Rating分析",
        "价格分析",
        "总结",
    )
    return any(marker in normalized for marker in section_markers)


def normalize_keyword_for_query(text: str) -> str:
    """Use only the query keyword, not translated notes in parentheses."""
    normalized = cell_text(text)
    if not normalized:
        return ""
    first_line = normalized.splitlines()[0].strip()
    keyword = first_line or normalized
    keyword = re.sub(r"\s*[（(][^（）()]*[）)]\s*$", "", keyword).strip()
    return keyword


def extract_keywords(worksheet, header_cell: Cell, image_column: int | None = None) -> list[KeywordItem]:
    items: list[KeywordItem] = []
    column = header_cell.column
    image_column = image_column or column + 1
    row = header_cell.row + 1

    while row <= worksheet.max_row:
        text = cell_text(worksheet.cell(row=row, column=column).value)
        if not text:
            break
        if is_section_title(text):
            break
        keyword = normalize_keyword_for_query(text)
        if not keyword:
            row += 1
            continue
        keyword_cell = f"{get_column_letter(column)}{row}"
        image_cell = f"{get_column_letter(image_column)}{row}"
        items.append(KeywordItem(keyword=keyword, keyword_cell=keyword_cell, image_cell=image_cell))
        row += 1

    return items


def safe_name(keyword: str) -> str:
    value = keyword.replace("\r", "").replace("\n", "").strip().replace(" ", "-")
    value = re.sub(r"[\\/:*?\"<>|]+", "-", value)
    value = re.sub(r"[^0-9A-Za-z._\-一-鿿]+", "-", value)
    value = re.sub(r"-+", "-", value).strip("-._")
    return value


def chart_path_for_keyword(project_dir: Path, prefix: str, keyword: str, index: int | None = None) -> Path:
    name = safe_name(keyword)
    if not name and index is not None:
        name = f"keyword-{index}"
    if not name:
        name = "keyword"
    return project_dir / f"{prefix}-{name}-chart.png"


def recent_complete_year(today: date | None = None) -> int:
    today = today or date.today()
    return today.year - 1


def parse_number(value: object) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    match = re.search(r"-?\d+(?:,\d{3})*(?:\.\d+)?", text)
    if not match:
        return None
    try:
        return float(match.group(0).replace(",", ""))
    except ValueError:
        return None


def parse_monthly_search_points(raw_points: Sequence[dict]) -> list[MonthlyTrendPoint]:
    points: list[MonthlyTrendPoint] = []
    for raw in raw_points:
        month = str(raw.get("month", "")).strip()
        if not re.match(r"^\d{4}-\d{2}$", month):
            continue
        value = parse_number(raw.get("searchVolume", raw.get("search_volume")))
        if value is None:
            continue
        points.append(MonthlyTrendPoint(month=month, search_volume=int(round(value))))
    return points


def filter_points_by_year(points: Sequence[MonthlyTrendPoint], year: int) -> list[MonthlyTrendPoint]:
    prefix = f"{year}-"
    return [point for point in points if point.month.startswith(prefix)]


def month_number(month: str) -> int | None:
    match = re.match(r"^\d{4}-(\d{2})$", month)
    if not match:
        return None
    value = int(match.group(1))
    if 1 <= value <= 12:
        return value
    return None


def group_contiguous_months(months: Sequence[int]) -> list[list[int]]:
    groups: list[list[int]] = []
    for month in sorted(set(months)):
        if not groups or month != groups[-1][-1] + 1:
            groups.append([month])
        else:
            groups[-1].append(month)
    return groups


def format_month_group(group: Sequence[int], burst: bool = False) -> str:
    if len(group) == 1:
        label = f"{group[0]}月"
    else:
        label = f"{group[0]}-{group[-1]}月"
    if burst:
        label += "爆发"
    return label


def format_cross_year_month_group(group: Sequence[int]) -> str:
    return f"{group[0]}月-次年{group[-1]}月"


def group_peak_months_for_display(months: Sequence[int]) -> list[list[int]]:
    groups = group_contiguous_months(months)
    if len(groups) >= 2 and groups[0][0] == 1 and groups[-1][-1] == 12:
        return [groups[-1] + groups[0]] + groups[1:-1]
    return groups


def is_cross_year_month_group(group: Sequence[int]) -> bool:
    return bool(group) and any(month == 12 for month in group) and any(month == 1 for month in group)


def format_peak_label_from_months(peak_months: Sequence[int], burst_months: Sequence[int]) -> str:
    peaks = sorted(set(peak_months))
    bursts = set(burst_months)
    if not peaks:
        return "无明显趋势"
    labels: list[str] = []
    for group in group_peak_months_for_display(peaks):
        group_bursts = [month for month in group if month in bursts]
        if is_cross_year_month_group(group):
            labels.append(format_cross_year_month_group(group))
        elif len(group) == 1 and group_bursts:
            labels.append(format_month_group(group, burst=True))
        else:
            labels.append(format_month_group(group))
    return "，".join(labels)


def is_cross_year_peak_months(peak_months: Sequence[int]) -> bool:
    return any(is_cross_year_month_group(group) for group in group_peak_months_for_display(peak_months))


def resolve_peak_month_keys_for_base_year(peak_months: Sequence[int], base_year: int) -> list[str]:
    keys: list[str] = []
    for group in group_peak_months_for_display(peak_months):
        if is_cross_year_month_group(group):
            for month in group:
                year = base_year + 1 if month < group[0] else base_year
                keys.append(f"{year}-{month:02d}")
        else:
            for month in group:
                keys.append(f"{base_year}-{month:02d}")
    return keys


def resolved_season_points_for_base_year(points: Sequence[MonthlyTrendPoint], peak_months: Sequence[int], base_year: int) -> list[MonthlyTrendPoint]:
    point_values = {point.month: point for point in points}
    current_keys = resolve_peak_month_keys_for_base_year(peak_months, base_year)
    if any(point_values[key].search_volume > 0 for key in current_keys if key in point_values):
        return [point_values[key] for key in current_keys if key in point_values]

    for candidate_year in range(base_year - 1, 1999, -1):
        candidate_keys = resolve_peak_month_keys_for_base_year(peak_months, candidate_year)
        candidate_points = [point_values[key] for key in candidate_keys if key in point_values]
        if any(point.search_volume > 0 for point in candidate_points):
            return candidate_points
    return [point_values[key] for key in current_keys if key in point_values]


def valid_seasonality_search_volume(value: int) -> bool:
    return value >= MIN_VALID_SEASONALITY_SEARCH_VOLUME


def select_seasonality_points(
    points: Sequence[MonthlyTrendPoint],
    preferred_year: int,
) -> tuple[list[MonthlyTrendPoint], int, str]:
    """Use a full preferred year when available, otherwise start at keyword entry."""
    ordered = sorted(points, key=lambda point: point.month)
    preferred_points = filter_points_by_year(ordered, preferred_year)
    preferred_valid_months = {
        point.month
        for point in preferred_points
        if valid_seasonality_search_volume(point.search_volume)
    }
    if len(preferred_valid_months) == 12:
        return preferred_points, preferred_year, "complete-year"

    first_valid = next(
        (point for point in ordered if valid_seasonality_search_volume(point.search_volume)),
        None,
    )
    if first_valid is None:
        return [], preferred_year, "no-valid-data"

    selected = [point for point in ordered if point.month >= first_valid.month]
    latest_year = int(selected[-1].month[:4])
    return selected, latest_year, "since-keyword-entry"


def analyze_seasonality(raw_points: Sequence[dict] | Sequence[MonthlyTrendPoint]) -> SeasonalityResult:
    if not raw_points:
        return SeasonalityResult("无明显趋势", None, None, [], [])
    if isinstance(raw_points[0], MonthlyTrendPoint):
        points = list(raw_points)  # type: ignore[arg-type]
    else:
        points = parse_monthly_search_points(raw_points)  # type: ignore[arg-type]
    values = [point.search_volume for point in points]
    analysis_values = [value for value in values if valid_seasonality_search_volume(value)]
    if not values or not analysis_values:
        return SeasonalityResult("无明显趋势", None, None, [], [])

    low = min(analysis_values)
    high = max(analysis_values)
    avg = mean(analysis_values)
    min_for_ratio = max(low, 1)
    if high / min_for_ratio < NO_TREND_MAX_MIN_RATIO or (high - low) / max(avg, 1) < NO_TREND_RANGE_AVG_RATIO:
        return SeasonalityResult("无明显趋势", low, high, [], [])

    peak_months: list[int] = []
    burst_months: list[int] = []
    for point in points:
        month = month_number(point.month)
        if month is None:
            continue
        value = point.search_volume
        if value >= avg * PEAK_AVG_MULTIPLIER and value >= high * PEAK_MAX_RATIO:
            peak_months.append(month)
        if value >= high * BURST_MAX_RATIO and value >= avg * BURST_AVG_MULTIPLIER:
            burst_months.append(month)

    groups = group_contiguous_months(peak_months)
    labels: list[str] = []
    for group in groups:
        group_bursts = [month for month in group if month in burst_months]
        if len(group) == 1 and group_bursts:
            labels.append(format_month_group(group, burst=True))
        else:
            labels.append(format_month_group(group))
    label = "，".join(labels) if labels else "无明显趋势"
    return SeasonalityResult(label, low, high, sorted(set(peak_months)), sorted(set(burst_months)))


def max_bid_for_peak_months(ppc_points: Sequence[dict], year: int, peak_months: Sequence[int]) -> float | None:
    peak_keys = set(resolve_peak_month_keys_for_base_year(peak_months, year))
    if not peak_keys:
        return None
    values: list[float] = []
    for point in ppc_points:
        month = str(point.get("month", "")).strip()
        if month not in peak_keys:
            continue
        bid = parse_number(point.get("bid"))
        if valid_ppc_bid(bid):
            values.append(bid)
    return max(values) if values else None


def parse_iso_date(value: object) -> date | None:
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    match = re.search(r"(20\d{2})[-/.年](\d{1,2})[-/.月](\d{1,2})", text)
    if not match:
        return None
    year, month, day = (int(part) for part in match.groups())
    try:
        return date(year, month, day)
    except ValueError:
        return None


def month_date_range(year: int, month: int) -> tuple[date, date] | None:
    if month < 1 or month > 12:
        return None
    last_day = calendar.monthrange(year, month)[1]
    return date(year, month, 1), date(year, month, last_day)


def date_ranges_overlap(start_a: date, end_a: date, start_b: date, end_b: date) -> bool:
    return start_a <= end_b and start_b <= end_a


def valid_ppc_bid(bid: float | None) -> bool:
    return bid is not None and bid > 0


def max_recent_weekly_bid(ppc_points: Sequence[dict], days: int = 30) -> float | None:
    parsed: list[tuple[date, date, float]] = []
    for point in ppc_points:
        week_start = parse_iso_date(point.get("week", point.get("date", point.get("time"))))
        if week_start is None:
            continue
        week_end = parse_iso_date(point.get("weekEnd", point.get("week_end"))) or week_start
        bid = parse_number(point.get("bid"))
        if valid_ppc_bid(bid):
            parsed.append((week_start, week_end, bid))
    if not parsed:
        return None
    latest_date = max(week_end for _week_start, week_end, _bid in parsed)
    window_start = latest_date - timedelta(days=days)
    values = [
        bid
        for week_start, week_end, bid in parsed
        if date_ranges_overlap(week_start, week_end, window_start, latest_date)
    ]
    return max(values) if values else None


def max_recent_monthly_bid(ppc_points: Sequence[dict]) -> float | None:
    parsed: list[tuple[str, float]] = []
    for point in ppc_points:
        month = str(point.get("month", "")).strip()
        if not re.match(r"^\d{4}-\d{2}$", month):
            continue
        bid = parse_number(point.get("bid"))
        if valid_ppc_bid(bid):
            parsed.append((month, bid))
    if not parsed:
        return None
    latest_month = max(month for month, _bid in parsed)
    values = [bid for month, bid in parsed if month == latest_month]
    return max(values) if values else None


def max_recent_bid_for_no_trend(ppc_weekly: Sequence[dict], ppc_monthly: Sequence[dict]) -> float | None:
    weekly_bid = max_recent_weekly_bid(ppc_weekly, days=30)
    if weekly_bid is not None:
        return weekly_bid
    return max_recent_monthly_bid(ppc_monthly)


def max_weekly_bid_for_peak_months(ppc_points: Sequence[dict], year: int, peak_months: Sequence[int]) -> float | None:
    ranges: list[tuple[date, date]] = []
    for key in resolve_peak_month_keys_for_base_year(peak_months, year):
        key_year, key_month = (int(part) for part in key.split("-"))
        month_range = month_date_range(key_year, key_month)
        if month_range:
            ranges.append(month_range)
    if not ranges:
        return None
    values: list[float] = []
    for point in ppc_points:
        week_start = parse_iso_date(point.get("week", point.get("date", point.get("time"))))
        if week_start is None:
            continue
        week_end = parse_iso_date(point.get("weekEnd", point.get("week_end")))
        bid = parse_number(point.get("bid"))
        if not valid_ppc_bid(bid):
            continue
        if week_end is not None:
            if any(date_ranges_overlap(week_start, week_end, start, end) for start, end in ranges):
                values.append(bid)
        elif any(start <= week_start <= end for start, end in ranges):
            values.append(bid)
    return max(values) if values else None


def max_bid_for_peak_period(
    ppc_weekly: Sequence[dict],
    ppc_monthly: Sequence[dict],
    year: int,
    peak_months: Sequence[int],
) -> float | None:
    if not peak_months:
        return max_recent_bid_for_no_trend(ppc_weekly, ppc_monthly)
    weekly_bid = max_weekly_bid_for_peak_months(ppc_weekly, year, peak_months)
    if weekly_bid is not None:
        return weekly_bid
    return max_bid_for_peak_months(ppc_monthly, year, peak_months)


def currency_number_format(site: str) -> str:
    marketplace = get_marketplace(site)
    return f'"{marketplace.currency_symbol}"#,##0.00'


def write_trend_fields(
    worksheet,
    item: KeywordItem,
    header_cells: dict[str, Cell],
    result: SeasonalityResult,
    site: str = "US",
    missing_value: object | None = None,
) -> None:
    row = keyword_item_row(item)
    values = {
        "趋势旺季": result.peak_label,
        "淡季月搜索量": result.low_search_volume,
        "旺季月搜索量": result.peak_search_volume,
        "旺季竞价": result.peak_bid,
    }
    for title, value in values.items():
        if title not in header_cells:
            continue
        cell = worksheet.cell(row=row, column=header_cells[title].column)
        cell.value = missing_value if value is None else value
        if title == "旺季竞价" and value is not None:
            cell.number_format = currency_number_format(site)


def write_missing_trend_fields(
    worksheet,
    item: KeywordItem,
    header_cells: dict[str, Cell],
    missing_value: object = "/",
) -> None:
    row = keyword_item_row(item)
    for title in TREND_FIELD_TITLES:
        if title in header_cells:
            worksheet.cell(row=row, column=header_cells[title].column).value = missing_value


def sanitize_trend_data_output(path: Path) -> None:
    if not path.exists():
        return
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return
    if not isinstance(payload, dict) or not isinstance(payload.get("results"), list):
        return
    cleaned_results = []
    changed = False
    for item in payload["results"]:
        if isinstance(item, dict) and "base64" in item:
            item = {key: value for key, value in item.items() if key != "base64"}
            changed = True
        cleaned_results.append(item)
    if changed:
        payload["results"] = cleaned_results
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def load_trend_data_json(path: Path) -> dict[str, dict]:
    if not path.exists():
        return {}
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        print(f"趋势数据 JSON 解析失败: {exc}", file=sys.stderr)
        return {}
    results = payload.get("results", []) if isinstance(payload, dict) else []
    loaded: dict[str, dict] = {}
    for item in results:
        if not isinstance(item, dict):
            continue
        keyword = cell_text(item.get("keyword"))
        if not keyword or item.get("status") == "failed":
            continue
        if (
            item.get("trendMonthly")
            or item.get("ppcMonthly")
            or item.get("ppcWeekly")
            or item.get("amazonResultCount") is not None
            or item.get("amazonResultText")
        ):
            loaded[keyword] = item
    return loaded


def apply_trend_data_fields(
    worksheet,
    items: Sequence[KeywordItem],
    header_cells: dict[str, Cell],
    trend_data: dict[str, dict],
    year: int,
    site: str = "US",
    ppc_bid_mode: str = "latest-week",
) -> int:
    if ppc_bid_mode not in {"latest-week", "latest-month", "peak-period"}:
        raise ValueError(f"不支持的竞价取数方式：{ppc_bid_mode}")
    written = 0
    for item in items:
        data = trend_data.get(item.keyword)
        if not data:
            continue
        all_points = parse_monthly_search_points(data.get("trendMonthly", []))
        seasonality_points, ppc_year, selection_mode = select_seasonality_points(all_points, year)
        if not seasonality_points:
            write_missing_trend_fields(worksheet, item, header_cells)
            # 最新周/月竞价独立于搜索量历史，缺少趋势时也可回填。
            if ppc_bid_mode in {"latest-week", "latest-month"} and "旺季竞价" in header_cells:
                _period, bid = latest_weekly_bid(data.get("ppcWeekly", [])) if ppc_bid_mode == "latest-week" else latest_monthly_bid(data.get("ppcMonthly", []))
                target = worksheet.cell(keyword_item_row(item), header_cells["旺季竞价"].column)
                target.value = bid if bid is not None else "/"
                if bid is not None:
                    target.number_format = currency_number_format(site)
            written += 1
            print(f"已回填趋势字段占位: {item.keyword} -> /")
            continue
        result = analyze_seasonality(seasonality_points)
        if ppc_bid_mode == "latest-week":
            _period, peak_bid = latest_weekly_bid(data.get("ppcWeekly", []))
        elif ppc_bid_mode == "latest-month":
            _period, peak_bid = latest_monthly_bid(data.get("ppcMonthly", []))
        else:
            peak_bid = max_bid_for_peak_period(data.get("ppcWeekly", []), data.get("ppcMonthly", []), ppc_year, result.peak_months)
        result = SeasonalityResult(
            peak_label=result.peak_label,
            low_search_volume=result.low_search_volume,
            peak_search_volume=result.peak_search_volume,
            peak_months=result.peak_months,
            burst_months=result.burst_months,
            peak_bid=peak_bid,
        )
        write_trend_fields(worksheet, item, header_cells, result, site=site, missing_value="/")
        written += 1
        print(f"已回填趋势字段: {item.keyword} -> {result.peak_label} ({selection_mode})")
    return written


def write_amazon_result_fields(
    worksheet,
    items: Sequence[KeywordItem],
    header_cell: Cell | None,
    trend_data: dict[str, dict],
    site: str = "US",
) -> int:
    if header_cell is None:
        return 0
    print(f"站内搜索结果标题单元格: {header_cell.coordinate} ({cell_text(header_cell.value)})")
    written = 0
    for item in items:
        row = keyword_item_row(item)
        target = worksheet.cell(row=row, column=header_cell.column)
        data = trend_data.get(item.keyword)
        if not data:
            print(f"跳过站内搜索结果: {item.keyword} -> 缺少结构化数据")
            target.value = "/"
            continue
        count = parse_number(data.get("amazonResultCount"))
        if count is None:
            status = cell_text(data.get("amazonResultStatus"))
            error = cell_text(data.get("amazonResultError"))
            reason = error or status or "缺少结果数"
            print(f"跳过站内搜索结果: {item.keyword} -> {reason}")
            target.value = "/"
            continue
        value = int(round(count))
        target.value = value
        keyword_cell = worksheet[item.keyword_cell]
        amazon_url = cell_text(data.get("amazonResultUrl")) or get_marketplace(site).amazon_search_url(item.keyword)
        keyword_cell.hyperlink = amazon_url
        keyword_cell.font = KEYWORD_LINK_FONT
        written += 1
        print(f"已回填站内搜索结果: {item.keyword} -> {target.coordinate} = {value}")
    return written



def build_screenshot_environment() -> dict[str, str]:
    env = os.environ.copy()
    env.setdefault("PLAYWRIGHT_SESSION", "default")
    return env

def run_screenshot_script(
    project_dir: Path,
    keywords: Sequence[str],
    prefix: str,
    site: str = "US",
    amazon_postal_code: str | None = None,
    temp_name: str = DEFAULT_TEMP_KEYWORDS,
    data_output: Path | None = None,
) -> ScreenshotRunResult:
    """Run Sellersprite screenshot/data extraction through the Python Playwright API."""
    data_output = data_output or (project_dir / DEFAULT_DATA_OUTPUT)
    try:
        print("?? Python Playwright API ?????????????...")
        completed = browser_runner.run_sellersprite_workflow(
            project_dir=project_dir,
            keywords=keywords,
            prefix=prefix,
            data_output=data_output,
            port=ACTIVE_DEBUG_PORT,
            site=site,
            amazon_postal_code=amazon_postal_code,
        )
        if completed.stdout:
            sys.stdout.write(completed.stdout + "\n")
        if completed.stderr:
            sys.stderr.write(completed.stderr + "\n")
        sanitize_trend_data_output(data_output)
        return ScreenshotRunResult(
            returncode=completed.returncode,
            stdout=completed.stdout,
            stderr=completed.stderr,
            command=completed.command or ["python-playwright"],
        )
    except RuntimeError as exc:
        message = str(exc)
        print(message, file=sys.stderr)
        return ScreenshotRunResult(
            returncode=1,
            stdout="",
            stderr=message,
            command=["python-playwright"],
        )


def cleanup_run_artifacts(
    project_dir: Path,
    chart_paths: Sequence[Path],
) -> list[str]:
    """Remove disposable run artifacts while keeping login state and trend diagnostics."""
    removed: list[str] = []

    def remove_file(path: Path) -> None:
        if path.exists() and path.is_file():
            path.unlink()
            removed.append(str(path))

    def remove_dir(path: Path) -> None:
        if path.exists() and path.is_dir():
            shutil.rmtree(path, ignore_errors=True)
            if not path.exists():
                removed.append(str(path))

    for chart_path in chart_paths:
        remove_file(chart_path)


    remove_file(project_dir / ".amazon-retry-temp.js")
    remove_dir(project_dir / ".playwright-cli")
    remove_dir(project_dir / "__pycache__")
    remove_dir(project_dir / "tests" / "__pycache__")

    playwright_dir = project_dir / ".playwright"
    if playwright_dir.exists() and playwright_dir.is_dir() and not any(playwright_dir.iterdir()):
        playwright_dir.rmdir()
        removed.append(str(playwright_dir))

    return removed


def resolve_project_path(project_dir: Path, path: Path) -> Path:
    return path if path.is_absolute() else project_dir / path


def image_anchor_cell(image) -> str:
    anchor = image.anchor
    if isinstance(anchor, str):
        return anchor
    marker = getattr(anchor, "_from", None)
    if marker is None:
        return ""
    column = marker.col + 1
    row = marker.row + 1
    return f"{get_column_letter(column)}{row}"


def remove_images_at_cells(worksheet, cells: Iterable[str]) -> int:
    targets = set(cells)
    before = len(worksheet._images)
    worksheet._images = [image for image in worksheet._images if image_anchor_cell(image) not in targets]
    return before - len(worksheet._images)


def png_size(path: Path) -> tuple[int, int]:
    data = path.read_bytes()
    if not data.startswith(b"\x89PNG\r\n\x1a\n"):
        return (1, 1)
    return (int.from_bytes(data[16:20], "big"), int.from_bytes(data[20:24], "big"))


def make_workbook_image(path: Path):
    try:
        return OpenpyxlImage(str(path))
    except ImportError:
        data = path.read_bytes()

        image = object.__new__(OpenpyxlImage)
        image.ref = str(path)
        image.width, image.height = png_size(path)
        image.format = path.suffix.lstrip(".").lower() or "png"
        image._data = lambda data=data: data
        return image


def column_width_to_pixels(width: float | None) -> int:
    if width is None:
        width = 8.43
    return max(1, int(width * 7 + 5))


def row_height_to_pixels(height: float | None) -> int:
    if height is None:
        height = 15
    return max(1, int(height * 96 / 72))


def image_target_region(worksheet, cell_ref: str) -> tuple[int, int, int, int]:
    cell = worksheet[cell_ref]
    for merged_range in worksheet.merged_cells.ranges:
        if cell.coordinate in merged_range:
            return merged_range.min_col, merged_range.max_col, merged_range.min_row, merged_range.max_row
    return cell.column, cell.column, cell.row, cell.row


def fit_image_to_anchor_region(worksheet, image, cell_ref: str, padding: int = 8) -> None:
    min_col, max_col, min_row, max_row = image_target_region(worksheet, cell_ref)
    max_width = sum(
        column_width_to_pixels(worksheet.column_dimensions[get_column_letter(col)].width)
        for col in range(min_col, max_col + 1)
    )
    max_height = sum(
        row_height_to_pixels(worksheet.row_dimensions[row].height)
        for row in range(min_row, max_row + 1)
    )
    max_width = max(1, max_width - padding)
    max_height = max(1, max_height - padding)
    original_width = max(1, int(getattr(image, "width", 1) or 1))
    original_height = max(1, int(getattr(image, "height", 1) or 1))
    scale = min(max_width / original_width, max_height / original_height)
    image.width = max(1, int(original_width * scale))
    image.height = max(1, int(original_height * scale))


def insert_images(worksheet, image_pairs: Sequence[tuple[KeywordItem, Path]]) -> list[str]:
    target_cells = [item.image_cell for item, _path in image_pairs]
    removed = remove_images_at_cells(worksheet, target_cells)
    if removed:
        print(f"已删除旧图数量: {removed}")

    inserted: list[str] = []
    for item, path in image_pairs:
        image = make_workbook_image(path)
        fit_image_to_anchor_region(worksheet, image, item.image_cell)
        image.anchor = item.image_cell
        worksheet.add_image(image)
        inserted.append(item.keyword)
        print(f"已插入: {item.keyword} -> {item.image_cell}")
    return inserted


def reanchor_section3_chart_images(worksheet, trend_cells: Sequence[str]) -> int:
    """关键词区域扩行后，把第 3 小节的图表重新锚定到当前标题行。"""
    label_columns = {
        "品牌分析": ("B", "D", "F"),
        "上架时间分析": ("B", "D", "F"),
        "Review数量分析": ("B",),
        "Rating分析": ("B",),
        "价格分析": ("B",),
    }
    target_cells: list[str] = []
    for label, columns in label_columns.items():
        label_row = None
        for row in worksheet.iter_rows():
            if any(label in str(cell.value or "") for cell in row):
                label_row = row[0].row
                break
        if label_row is None:
            return 0
        target_cells.extend(f"{column}{label_row}" for column in columns)

    trend_cell_set = set(trend_cells)
    chart_images = []
    for image in list(getattr(worksheet, "_images", [])):
        anchor_cell = image_anchor_cell(image)
        if anchor_cell in trend_cell_set:
            continue
        width = max(1, int(getattr(image, "width", 1) or 1))
        height = max(1, int(getattr(image, "height", 1) or 1))
        # 类目表现图和趋势图都是宽图，第 3 小节图表是窄图或横向统计图。
        if width / height >= 2.5:
            continue
        chart_images.append(image)

    if len(chart_images) != len(target_cells):
        return 0
    moved = 0
    for image, target_cell in zip(chart_images, target_cells):
        if image_anchor_cell(image) != target_cell:
            image.anchor = target_cell
            moved += 1
    return moved


def reanchor_category_overview_images(worksheet, trend_cells: Sequence[str]) -> int:
    """关键词区域扩行后，把类目表现截图重新锚定到当前类目行。"""
    category_rows: list[int] = []
    in_category_section = False
    for row in worksheet.iter_rows():
        row_number = row[0].row
        row_text = " ".join(str(cell.value or "") for cell in row)
        if "所属平台分类" in row_text and "常用分类节点" in row_text:
            in_category_section = True
            continue
        if in_category_section and "数据分析概览" in row_text:
            break
        if in_category_section and row_number > 1:
            has_category_value = any(
                worksheet.cell(row_number, column).value not in (None, "")
                for column in range(1, 6)
            )
            if has_category_value and not any(
                marker in row_text for marker in ("一级类目", "二级", "三级", "四级", "五级")
            ):
                category_rows.append(row_number)

    # 类目数量由当前工作簿实际内容决定，不能固定假设为五行。
    target_rows = category_rows
    if not target_rows:
        return 0
    trend_cell_set = set(trend_cells)
    category_images = []
    for image in list(getattr(worksheet, "_images", [])):
        anchor_cell = image_anchor_cell(image)
        if anchor_cell in trend_cell_set:
            continue
        width = max(1, int(getattr(image, "width", 1) or 1))
        height = max(1, int(getattr(image, "height", 1) or 1))
        ratio = width / height
        if 2.5 <= ratio < 3.2:
            category_images.append(image)

    if len(category_images) != len(target_rows):
        return 0
    moved = 0
    for image, row_number in zip(category_images, target_rows):
        target_cell = f"F{row_number}"
        if image_anchor_cell(image) != target_cell:
            image.anchor = target_cell
            moved += 1
    return moved


def is_port_open(port: int = DEBUG_PORT) -> bool:
    """Check if a TCP port is listening on localhost."""
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        sock.settimeout(1)
        return sock.connect_ex(("127.0.0.1", port)) == 0
    finally:
        sock.close()



def is_devtools_port(port: int) -> bool:
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{port}/json/version/", timeout=2) as response:
            if response.status != 200:
                return False
            payload = json.loads(response.read().decode("utf-8", errors="replace"))
    except Exception:
        return False
    return bool(payload.get("webSocketDebuggerUrl"))


def find_available_debug_port(start: int = DEBUG_PORT, attempts: int = 20) -> int:
    for candidate in range(start, start + attempts):
        if not is_port_open(candidate):
            return candidate
        if is_devtools_port(candidate):
            return candidate
    raise RuntimeError(f"????? Chrome ????: {start}-{start + attempts - 1}")

def launch_chrome(port: int = DEBUG_PORT, user_data_dir: Path | None = None) -> None:
    """Launch Chrome with remote debugging enabled."""
    if user_data_dir is None:
        user_data_dir = DEFAULT_PROJECT_DIR / CHROME_USER_DATA_DIR
    user_data_dir = Path(user_data_dir)
    user_data_dir.mkdir(parents=True, exist_ok=True)

    subprocess.Popen(
        [
            str(CHROME_PATH),
            f"--remote-debugging-port={port}",
            f"--user-data-dir={str(user_data_dir)}",
            SELLERSPRITE_URL,
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )

    for _ in range(40):
        if is_devtools_port(port):
            return
        time.sleep(0.5)
    raise RuntimeError(f"Chrome 启动超时，端口 {port} 没有响应")


def classify_browser_tab_url(url: str) -> str:
    """Classify a browser tab URL for safe tab cleanup."""
    normalized = (url or "").strip().lower()
    if not normalized or normalized == "about:blank":
        return "disposable"
    if normalized.startswith("chrome://new-tab") or normalized.startswith("chrome://welcome"):
        return "disposable"
    if "google.com/chrome" in normalized:
        return "disposable"
    if any(domain in normalized for domain in amazon_domains()):
        return "amazon_temporary"
    if "sellersprite.com" in normalized:
        if "/v3/keyword-miner" in normalized:
            return "sellersprite_keyword_miner"
        if any(token in normalized for token in ("login", "signin", "passport", "account")):
            return "manual_action"
        return "sellersprite_other"
    return "unknown"


def cleanup_browser_tabs(
    session: str = "default",
    keep_sellersprite_keyword_pages: int = 1,
    close_amazon_tabs: bool = True,
) -> int:
    """Close disposable browser tabs and return the number of closed tabs."""
    try:
        return browser_runner.cleanup_browser_tabs(
            project_dir=DEFAULT_PROJECT_DIR,
            port=ACTIVE_DEBUG_PORT,
            keep_sellersprite_keyword_pages=keep_sellersprite_keyword_pages,
            close_amazon_tabs=close_amazon_tabs,
        )
    except RuntimeError as exc:
        print(str(exc), file=sys.stderr)
        return 0


def setup_browser_session(
    project_dir: Path,
    session: str = "default",
    port: int = DEBUG_PORT,
    browser_state_dir: Path | None = None,
    use_existing_browser: bool = False,
) -> bool:
    """Ensure Chrome is running and Sellersprite is logged in using Python Playwright."""
    global ACTIVE_DEBUG_PORT
    state_dir = browser_state_dir or project_dir
    user_data_dir = state_dir / CHROME_USER_DATA_DIR
    env = build_screenshot_environment()

    if not use_existing_browser:
        subprocess.run(
            ["taskkill", "//F", "//IM", CHROME_PATH.name],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            env=env,
        )
        time.sleep(1)

    print("?? Chrome ???...")
    if use_existing_browser:
        if not is_devtools_port(port):
            print(f"Chrome ???? {port} ??????? DevTools", file=sys.stderr)
            return False
    else:
        port = find_available_debug_port(port)
    ACTIVE_DEBUG_PORT = port
    print(f"Chrome ????: {port}")
    if not use_existing_browser:
        launch_chrome(port, user_data_dir)
        time.sleep(2)

    try:
        closed_tabs = cleanup_browser_tabs(
            session=session,
            keep_sellersprite_keyword_pages=1,
            close_amazon_tabs=True,
        )
        if closed_tabs:
            print(f"????????????????: {closed_tabs} ?")
        return browser_runner.ensure_sellersprite_login(
            project_dir=state_dir,
            port=port,
            state_file_name=BROWSER_STATE_FILE,
        )
    except RuntimeError as exc:
        print(str(exc), file=sys.stderr)
        return False


def retry_failed_amazon_extractions(
    data_path: Path,
    session: str = "default",
    site: str = "US",
    amazon_postal_code: str | None = None,
) -> int:
    """Retry Amazon result extraction through the Python Playwright API."""
    if not data_path.exists():
        return 0
    try:
        payload = json.loads(data_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return 0
    results = payload.get("results", []) if isinstance(payload, dict) else []
    has_retry_candidate = any(
        isinstance(item, dict)
        and item.get("amazonResultStatus") == "failed"
        and not str(item.get("amazonResultError", "")).startswith("blocked")
        for item in results
    )
    if not has_retry_candidate:
        return 0
    try:
        fixed = browser_runner.retry_failed_amazon_results(
            data_path,
            port=ACTIVE_DEBUG_PORT,
            site=site,
            amazon_postal_code=amazon_postal_code,
        )
    except RuntimeError as exc:
        print(str(exc), file=sys.stderr)
        return 0
    if fixed:
        print(f"Amazon ??????? {fixed} ????? {data_path}")
    closed_tabs = cleanup_browser_tabs(
        session=session,
        keep_sellersprite_keyword_pages=1,
        close_amazon_tabs=True,
    )
    if closed_tabs:
        print(f"????????????????: {closed_tabs} ?")
    return fixed


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="把卖家精灵关键词趋势截图插入 Excel 市场调研页")
    parser.add_argument("--excel", type=Path, default=DEFAULT_EXCEL, help="源 Excel 文件路径")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT, help="输出 Excel 副本路径")
    parser.add_argument("--project-dir", type=Path, default=DEFAULT_PROJECT_DIR, help="包含截图脚本的项目目录")
    parser.add_argument("--browser-state-dir", type=Path, default=None, help="共享浏览器登录态目录，默认使用 project-dir")
    parser.add_argument("--debug-port", type=int, default=DEBUG_PORT, help="Chrome/Edge DevTools 调试端口")
    parser.add_argument("--use-existing-browser", action="store_true", help="直接连接已经打开的调试浏览器，不杀进程、不新开 Chrome")
    parser.add_argument("--prefix", default=DEFAULT_PREFIX, help="截图输出前缀")
    parser.add_argument("--dry-run", action="store_true", help="只识别关键词，不执行截图或写 Excel")
    parser.add_argument("--skip-screenshot", action="store_true", help="不调用截图脚本，直接使用项目目录中已有 PNG 插入 Excel")
    parser.add_argument("--keep-run-artifacts", action="store_true", help="保留本次运行的 PNG 和调试缓存")
    parser.add_argument("--ppc-bid-mode", choices=("latest-week", "latest-month", "peak-period"), default="latest-week", help="默认取最新周 PPC；可选最新月或旧旺季区间逻辑")
    parser.add_argument("--site", default="US", help="Marketplace site code, for example US, DE, UK, CA or FR")
    parser.add_argument("--amazon-postal-code", help="Override the marketplace default Amazon postal code")
    parser.add_argument(
        "--keyword",
        action="append",
        dest="selected_keywords",
        help="只处理指定关键词；可重复传入，用于补跑缺失项。",
    )
    parser.add_argument(
        "--embed-wps-cell-images",
        action="store_true",
        help="把本次插入的趋势截图转换成 WPS DISPIMG 嵌入单元格图片。完整工作流默认不要开启，避免后续 openpyxl 保存时丢图。",
    )
    args = parser.parse_args(argv)
    site = normalize_site_code(args.site)
    marketplace = get_marketplace(site)
    amazon_postal_code = args.amazon_postal_code or marketplace.postal_code


    if not args.excel.exists():
        print(f"Excel 文件不存在: {args.excel}", file=sys.stderr)
        return 1

    workbook = load_workbook(args.excel, rich_text=True)
    worksheet = find_market_research_sheet(workbook)
    header = find_keyword_header(worksheet)
    header_cells: dict[str, Cell] = {}
    try:
        header_cells = find_header_cells_in_row(worksheet, header.row, ["搜索趋势", *TREND_FIELD_TITLES])
    except ValueError as exc:
        print(f"提示: {exc}；将使用关键词右侧列插图，跳过趋势字段回填。")
    amazon_result_header = find_optional_header_cell_in_row(worksheet, header.row, AMAZON_RESULT_FIELD_TITLES)
    if amazon_result_header is None:
        titles = "、".join(f"“{title}”" for title in AMAZON_RESULT_FIELD_TITLES)
        print(f"提示: 第 {header.row} 行未找到{titles}标题，跳过站内搜索结果回填。")
    image_column = header_cells.get("搜索趋势").column if "搜索趋势" in header_cells else None
    items = extract_keywords(worksheet, header, image_column=image_column)
    if args.selected_keywords:
        selected = {keyword.strip().casefold() for keyword in args.selected_keywords if keyword.strip()}
        items = [item for item in items if item.keyword.casefold() in selected]

    print(f"目标工作表: {worksheet.title}")
    print(f"关键词标题单元格: {header.coordinate}")
    print(f"识别到关键词数量: {len(items)}")
    for item in items:
        print(f"  {item.keyword_cell} -> {item.image_cell}: {item.keyword}")

    if not items:
        print("没有关键词可处理，脚本结束。")
        return 0

    if args.dry_run:
        print("dry-run 模式：不执行截图，不写入 Excel。")
        return 0

    data_output = args.project_dir / DEFAULT_DATA_OUTPUT
    browser_state_dir = args.browser_state_dir or args.project_dir
    generated_chart_paths: list[Path] = []

    if args.skip_screenshot:
        print("skip-screenshot 模式：不调用截图脚本，直接检查已有 PNG。")
    else:
        if not setup_browser_session(
            args.project_dir,
            port=args.debug_port,
            browser_state_dir=browser_state_dir,
            use_existing_browser=args.use_existing_browser,
        ):
            print("浏览器会话设置失败，请手动启动 Chrome 并登录卖家精灵后重试。", file=sys.stderr)
            return 1
        if data_output.exists():
            data_output.unlink()
        result = run_screenshot_script(
            project_dir=args.project_dir,
            keywords=[item.keyword for item in items],
            prefix=args.prefix,
            site=site,
            amazon_postal_code=amazon_postal_code,
            data_output=data_output,
        )
        if result.returncode != 0:
            print(f"截图脚本返回非零状态 {result.returncode}，将继续检查已生成的 PNG。")

    available = []
    missing = []
    for index, item in enumerate(items, start=1):
        chart_path = chart_path_for_keyword(args.project_dir, args.prefix, item.keyword, index=index)
        if chart_path.exists():
            available.append((item, chart_path))
            generated_chart_paths.append(chart_path)
            print(f"已生成: {item.keyword} -> {chart_path}")
        else:
            missing.append((item, chart_path))
            print(f"跳过: {item.keyword}，未找到图片 {chart_path}")

    if not available:
        print("没有可插入图片，脚本结束，不更新 Excel 副本。")
        return 0

    inserted = insert_images(worksheet, available)
    category_moved = reanchor_category_overview_images(
        worksheet,
        [item.image_cell for item, _path in available],
    )
    if category_moved:
        print(f"已重新对齐类目表现截图锚点: {category_moved} 张")
    chart_moved = reanchor_section3_chart_images(
        worksheet,
        [item.image_cell for item, _path in available],
    )
    if chart_moved:
        print(f"已重新对齐第 3 小节图表锚点: {chart_moved} 张")
    if args.skip_screenshot:
        print("skip-screenshot 模式：跳过 Amazon 失败项浏览器重试。")
    else:
        retry_failed_amazon_extractions(
            data_output,
            site=site,
            amazon_postal_code=amazon_postal_code,
        )
    trend_data = load_trend_data_json(data_output)
    trend_headers = {title: cell for title, cell in header_cells.items() if title in TREND_FIELD_TITLES}
    year = recent_complete_year()
    print("趋势旺季判断模式: 固定规则")
    if trend_data and len(trend_headers) == len(TREND_FIELD_TITLES):
        written = apply_trend_data_fields(
            worksheet,
            items,
            trend_headers,
            trend_data,
            year,
            site=site,
            ppc_bid_mode=args.ppc_bid_mode,
        )
        print(f"成功回填趋势字段数量: {written}")
    elif trend_data:
        print("已读取趋势数据，但模板缺少趋势字段标题，跳过字段回填。")
    if trend_data and amazon_result_header is not None:
        amazon_written = write_amazon_result_fields(
            worksheet,
            items,
            amazon_result_header,
            trend_data,
            site=site,
        )
        print(f"成功回填站内搜索结果数量: {amazon_written}")
    args.output.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(args.output)
    if args.embed_wps_cell_images:
        wps_temp_output = args.output.with_name(args.output.stem + "_floating_temp.xlsx")
        wps_result = convert_floating_images_to_wps_cell_images(
            floating_workbook=args.output,
            output=args.output,
            sheet_name=worksheet.title,
            target_cells=[item.image_cell for item, _path in available],
            temp_output=wps_temp_output,
            source_workbook=args.excel,
        )
        if wps_temp_output.exists():
            wps_temp_output.unlink()
        print(f"WPS 嵌入单元格趋势截图数量: {wps_result['embedded']}")
    else:
        print("趋势截图已保持为浮动图片，方便后续流程继续保存而不丢图。")
    if "旺季月搜索量" in trend_headers:
        # 回填与图片嵌入完成后再整行排序，包括本次未补跑的已有关键词。
        all_items = extract_keywords(worksheet, header, image_column=image_column)
        sort_result = update_keyword_rows(
            args.output, args.output, sheet_name=worksheet.title,
            rows=[keyword_item_row(item) for item in all_items],
            peak_column=trend_headers["旺季月搜索量"].column,
        )
        print(f"已按旺季月搜索量降序排列：{sort_result['rows']} 个关键词")
    print(f"已保存 Excel 副本: {args.output}")
    print(f"成功插入数量: {len(inserted)}")
    if missing:
        print("跳过关键词:")
        for item, chart_path in missing:
            print(f"  {item.keyword}: 缺少图片 {chart_path}")
    if args.keep_run_artifacts:
        print("已按参数保留运行临时文件。")
    else:
        removed_artifacts = cleanup_run_artifacts(
            project_dir=args.project_dir,
            chart_paths=generated_chart_paths,
        )
        if removed_artifacts:
            print(f"已清理运行临时文件数量: {len(removed_artifacts)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

