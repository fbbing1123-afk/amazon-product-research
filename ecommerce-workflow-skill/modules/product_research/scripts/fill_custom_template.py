from __future__ import annotations

import argparse
import json
import math
import re
from copy import copy
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from openpyxl.cell.cell import MergedCell
from openpyxl import load_workbook
from openpyxl.styles import Border, Font, PatternFill, Side
from openpyxl.worksheet.cell_range import CellRange


CATEGORY_SECTION_MARKER = "所属平台分类及常用分类节点"
CATEGORY_HEADER_MARKER = "一级类目"
NEXT_SECTION_RE = re.compile(r"^\s*\d+\.")
CATEGORY_LEVEL_START_COL = 1
CATEGORY_LEVEL_FIXED_COLS = 5
CATEGORY_SALES_COL = 9
CATEGORY_TABLE_END_COL = 9
CATEGORY_EXTRA_LEVEL_START_COL = 10
MONTHLY_SALES_HEADER = "月销量"
CATEGORY_TEXT_FONT = Font(name="宋体", sz=11, color="000000")
CATEGORY_LINK_FONT = Font(name="宋体", sz=11, color="0000FF", underline="single")
CATEGORY_BORDER_SIDE = Side(style="thin", color="000000")
CATEGORY_PRIMARY_SHARE_THRESHOLD = 0.5
CATEGORY_PRIMARY_FILL = PatternFill(fill_type="solid", fgColor="D9EAF7")
CATEGORY_TRANSFER_FILL = PatternFill(fill_type="solid", fgColor="FCE4D6")
CATEGORY_TRANSLATION_CACHE_PATH = Path(__file__).resolve().parents[3] / "resources" / "category_translations.json"
AMAZON_BESTSELLER_ROOT_SLUGS = {
    "Home & Kitchen": "home-garden",
    "Toys & Games": "toys-and-games",
    "Electronics": "electronics",
    "Health & Household": "hpc",
    "Beauty & Personal Care": "beauty",
    "Baby Products": "baby-products",
    "Office Products": "office-products",
}
CATEGORY_PATH_TRANSLATION_HEADERS = (
    "类目路径翻译",
    "类目路径中文",
    "类目路径中文翻译",
    "中文类目路径",
    "分类路径翻译",
    "分类路径中文",
    "类目路径(中文)",
    "类目路径（中文）",
)
CATEGORY_BIG_TRANSLATION_HEADERS = (
    "大类目翻译",
    "大类目中文",
    "大类目中文翻译",
    "大类目(中文)",
    "大类目（中文）",
)
CATEGORY_LEAF_TRANSLATION_HEADERS = (
    "最细类目翻译",
    "最细类目中文",
    "小类目翻译",
    "小类目中文",
    "类目翻译",
    "类目中文",
    "中文类目",
)


@dataclass
class CategorySummary:
    big_category: str
    path_text: str
    levels: list[str]
    translated_levels: list[str]
    final_link: str | None
    product_count: int = 0
    monthly_sales: float = 0
    product_share: float = 0


def clean(value):
    if value is None:
        return None
    if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
        return None
    return value


def extract_year(value):
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value.year
    text = str(value).strip()
    if not text:
        return None
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"):
        try:
            return datetime.strptime(text[:10], fmt).year
        except ValueError:
            pass
    if len(text) >= 4 and text[:4].isdigit():
        return int(text[:4])
    return None


def read_rows_by_header(ws):
    headers = [ws.cell(1, col).value for col in range(1, ws.max_column + 1)]
    rows = []
    for row in range(2, ws.max_row + 1):
        item = {}
        for col, header in enumerate(headers, start=1):
            if header:
                item[str(header)] = clean(ws.cell(row, col).value)
        if any(v not in (None, "") for v in item.values()):
            rows.append(item)
    return headers, rows


def header_columns(ws) -> dict[str, int]:
    columns = {}
    for col in range(1, ws.max_column + 1):
        header = ws.cell(1, col).value
        if header not in (None, ""):
            columns[str(header)] = col
    return columns


def first_header_column(columns: dict[str, int], candidates: tuple[str, ...]) -> int | None:
    for candidate in candidates:
        if candidate in columns:
            return columns[candidate]
    return None


def parse_number(value) -> float:
    if value in (None, ""):
        return 0
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return 0
        return float(value)
    text = str(value).strip()
    if not text:
        return 0
    text = text.replace(",", "")
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    return float(match.group(0)) if match else 0


def cell_hyperlink_target(cell) -> str | None:
    if cell.hyperlink:
        return cell.hyperlink.target or cell.hyperlink.location
    return None


def extract_node_id(url: str | None) -> str | None:
    if not url:
        return None
    parsed = urlparse(url)
    node = parse_qs(parsed.query).get("node")
    if node and node[0]:
        return node[0]
    match = re.search(r"(?:[?&]node=|/node/)(\d+)", url)
    return match.group(1) if match else None


def build_leaf_category_url(
    big_category_url: str | None,
    path_url: str | None,
    big_category: str | None = None,
) -> str | None:
    node_id = extract_node_id(path_url)
    if not node_id:
        return path_url or big_category_url
    if not big_category_url:
        root_slug = AMAZON_BESTSELLER_ROOT_SLUGS.get(normalize_text(big_category))
        if root_slug:
            return f"https://www.amazon.com/gp/bestsellers/{root_slug}/{node_id}"
        return path_url
    return f"{big_category_url.rstrip('/')}/{node_id}"


def category_path_key(path_text: str) -> str:
    return re.sub(r"\s+", " ", path_text).strip().casefold()


def big_category_matches_path_root(big_category: str, levels: list[str]) -> bool:
    if not levels:
        return False
    return normalize_translation_key(big_category) == normalize_translation_key(levels[0])


def split_category_path(value) -> list[str]:
    if value in (None, ""):
        return []
    return [part.strip() for part in str(value).split(":") if part and part.strip()]


def normalize_text(value) -> str:
    if value in (None, ""):
        return ""
    return str(value).strip()


def normalize_translation_key(value: str) -> str:
    text = normalize_text(value).lower()
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def load_category_translation_cache(path: Path = CATEGORY_TRANSLATION_CACHE_PATH) -> dict[str, str]:
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError):
        return {}
    if isinstance(data, dict) and isinstance(data.get("translations"), dict):
        raw_translations = data["translations"]
    elif isinstance(data, dict):
        raw_translations = data
    else:
        return {}
    translations: dict[str, str] = {}
    for key, value in raw_translations.items():
        source = normalize_text(key)
        translated = normalize_text(value)
        if source and translated:
            translations[source] = translated
            translations[normalize_translation_key(source)] = translated
    return translations


def cached_category_translation(level: str, translation_cache: dict[str, str]) -> str:
    return translation_cache.get(level) or translation_cache.get(normalize_translation_key(level), "")


def split_category_translation_path(value) -> list[str]:
    text = normalize_text(value)
    if not text:
        return []
    parts = [part.strip() for part in re.split(r"\s*[:：]\s*", text) if part and part.strip()]
    if len(parts) > 1:
        return parts
    return [part.strip() for part in re.split(r"\s*(?:>|＞|›|»)\s*", text) if part and part.strip()]


def category_display_text(level: str, translation: str | None) -> str:
    translation_text = normalize_text(translation)
    if translation_text and translation_text != level:
        return f"{level}\n({translation_text})"
    return level


def collect_translated_levels(
    ws,
    row: int,
    levels: list[str],
    path_translation_col: int | None,
    big_translation_col: int | None,
    leaf_translation_col: int | None,
    translation_cache: dict[str, str],
) -> list[str]:
    translations = [""] * len(levels)
    if path_translation_col:
        path_translations = split_category_translation_path(ws.cell(row, path_translation_col).value)
        for index, translation in enumerate(path_translations[: len(translations)]):
            translations[index] = translation

    if big_translation_col and translations:
        translations[0] = translations[0] or normalize_text(ws.cell(row, big_translation_col).value)

    if leaf_translation_col and translations:
        translations[-1] = translations[-1] or normalize_text(ws.cell(row, leaf_translation_col).value)

    for index, level in enumerate(levels):
        translations[index] = translations[index] or cached_category_translation(level, translation_cache)

    return translations


def merge_translated_levels(existing: list[str], incoming: list[str]) -> list[str]:
    max_len = max(len(existing), len(incoming))
    merged = []
    for index in range(max_len):
        old = existing[index] if index < len(existing) else ""
        new = incoming[index] if index < len(incoming) else ""
        merged.append(old or new)
    return merged


def get_sheet_by_name_or_contains(wb, names=(), contains=()):
    for name in names:
        if name in wb.sheetnames:
            return wb[name]
    for ws in wb.worksheets:
        if any(token in ws.title for token in contains):
            return ws
    return None


def find_source_sheet(wb):
    for ws in wb.worksheets:
        if ws.title.endswith("数据源") and ws.title != "数据源":
            return ws
    return get_sheet_by_name_or_contains(wb, names=("数据源", "US", "Sheet1"))


def infer_product_name(wb):
    for ws in wb.worksheets:
        if ws.title.endswith("市场调研"):
            product_name = ws.title[: -len("市场调研")].strip("-_ ")
            if product_name:
                return product_name
    return None


def ensure_canonical_source_sheet_name(wb, ws):
    if ws.title != "数据源":
        return ws
    product_name = infer_product_name(wb)
    if not product_name:
        return ws
    canonical_name = f"{product_name}-数据源"
    if canonical_name not in wb.sheetnames:
        ws.title = canonical_name
        return ws
    return wb[canonical_name]


def remove_duplicate_plain_source_sheet(wb, selected_ws):
    if selected_ws.title != "数据源" and "数据源" in wb.sheetnames and len(wb.worksheets) > 1:
        del wb["数据源"]


def copy_row_style(ws, source_row: int, target_row: int, max_col: int) -> None:
    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height
    for col in range(1, max_col + 1):
        src = ws.cell(source_row, col)
        dst = ws.cell(target_row, col)
        if src.has_style:
            dst._style = copy(src._style)
        if src.number_format:
            dst.number_format = src.number_format
        if src.alignment:
            dst.alignment = copy(src.alignment)
        if src.font:
            dst.font = copy(src.font)
        if src.fill:
            dst.fill = copy(src.fill)
        if src.border:
            dst.border = copy(src.border)


def copy_cell_format(src, dst) -> None:
    if src.has_style:
        dst._style = copy(src._style)
    if src.number_format:
        dst.number_format = src.number_format
    if src.alignment:
        dst.alignment = copy(src.alignment)
    if src.font:
        dst.font = copy(src.font)
    if src.fill:
        dst.fill = copy(src.fill)
    if src.border:
        dst.border = copy(src.border)


def shift_range_for_insert(merged_range: CellRange, insert_at: int, amount: int) -> CellRange:
    min_row = merged_range.min_row
    max_row = merged_range.max_row
    if merged_range.min_row >= insert_at:
        min_row += amount
        max_row += amount
    elif merged_range.max_row >= insert_at:
        max_row += amount
    return CellRange(
        min_col=merged_range.min_col,
        min_row=min_row,
        max_col=merged_range.max_col,
        max_row=max_row,
    )


def ranges_overlap(left: CellRange, right: CellRange) -> bool:
    return not (
        left.max_row < right.min_row
        or left.min_row > right.max_row
        or left.max_col < right.min_col
        or left.min_col > right.max_col
    )


def merge_range_if_free(ws, candidate: CellRange) -> None:
    if candidate.min_row == candidate.max_row and candidate.min_col == candidate.max_col:
        return
    for existing in ws.merged_cells.ranges:
        if str(existing) == candidate.coord:
            return
        if ranges_overlap(existing, candidate):
            return
    ws.merge_cells(
        start_row=candidate.min_row,
        start_column=candidate.min_col,
        end_row=candidate.max_row,
        end_column=candidate.max_col,
    )


def insert_rows_preserving_merges(ws, insert_at: int, amount: int, template_row: int) -> None:
    if amount <= 0:
        return
    merged_ranges = [CellRange(str(merged_range)) for merged_range in ws.merged_cells.ranges]
    row_heights = {
        row_index: dimension.height
        for row_index, dimension in ws.row_dimensions.items()
    }
    template_row_merges = [
        CellRange(str(merged_range))
        for merged_range in merged_ranges
        if merged_range.min_row == template_row and merged_range.max_row == template_row
    ]

    for merged_range in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(merged_range))

    ws.insert_rows(insert_at, amount)
    shift_floating_objects_for_insert(ws, insert_at, amount)

    max_row = ws.max_row
    for row in range(max_row, insert_at + amount - 1, -1):
        source_row = row - amount
        ws.row_dimensions[row].height = row_heights.get(source_row)
    template_height = row_heights.get(template_row)
    for row in range(insert_at, insert_at + amount):
        ws.row_dimensions[row].height = template_height

    for merged_range in merged_ranges:
        merge_range_if_free(ws, shift_range_for_insert(merged_range, insert_at, amount))

    for row_offset in range(amount):
        target_row = insert_at + row_offset
        for merged_range in template_row_merges:
            merge_range_if_free(
                ws,
                CellRange(
                    min_col=merged_range.min_col,
                    min_row=target_row,
                    max_col=merged_range.max_col,
                    max_row=target_row,
                ),
            )


def shift_floating_objects_for_insert(ws, insert_at: int, amount: int) -> None:
    for drawing in list(getattr(ws, "_images", [])) + list(getattr(ws, "_charts", [])):
        shift_anchor_for_insert(getattr(drawing, "anchor", None), insert_at, amount)


def shift_anchor_for_insert(anchor, insert_at: int, amount: int) -> None:
    if anchor is None or isinstance(anchor, str):
        return
    for marker_name in ("_from", "_to"):
        marker = getattr(anchor, marker_name, None)
        if marker is not None and marker.row + 1 >= insert_at:
            marker.row += amount


def find_research_sheet(wb):
    for ws in wb.worksheets:
        if ws.title.endswith("市场调研") or "市场调研" in ws.title:
            return ws
    return None


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


def collect_category_summaries(ws) -> list[CategorySummary]:
    columns = header_columns(ws)
    path_col = columns.get("类目路径")
    big_col = columns.get("大类目")
    sales_col = columns.get(MONTHLY_SALES_HEADER)
    path_translation_col = first_header_column(columns, CATEGORY_PATH_TRANSLATION_HEADERS)
    big_translation_col = first_header_column(columns, CATEGORY_BIG_TRANSLATION_HEADERS)
    leaf_translation_col = first_header_column(columns, CATEGORY_LEAF_TRANSLATION_HEADERS)
    translation_cache = load_category_translation_cache()
    if not path_col or not big_col:
        return []

    summaries: dict[tuple[str, str], CategorySummary] = {}
    total_products = 0
    for row in range(2, ws.max_row + 1):
        path_cell = ws.cell(row, path_col)
        big_cell = ws.cell(row, big_col)
        path_text = str(path_cell.value).strip() if path_cell.value not in (None, "") else ""
        big_category = str(big_cell.value).strip() if big_cell.value not in (None, "") else ""
        levels = split_category_path(path_text)
        if not path_text or not big_category or not levels:
            continue
        translated_levels = collect_translated_levels(
            ws,
            row,
            levels,
            path_translation_col,
            big_translation_col,
            leaf_translation_col,
            translation_cache,
        )

        total_products += 1
        path_url = cell_hyperlink_target(path_cell)
        node_id = extract_node_id(path_url)
        key = ("node", node_id) if node_id else ("path", category_path_key(path_text))
        candidate_link = build_leaf_category_url(
            cell_hyperlink_target(big_cell),
            path_url,
            big_category,
        )
        if key not in summaries:
            summaries[key] = CategorySummary(
                big_category=big_category,
                path_text=path_text,
                levels=levels,
                translated_levels=translated_levels,
                final_link=candidate_link,
            )
        else:
            summary = summaries[key]
            summary.translated_levels = merge_translated_levels(
                summary.translated_levels,
                translated_levels,
            )
            current_is_canonical = big_category_matches_path_root(summary.big_category, summary.levels)
            candidate_is_canonical = big_category_matches_path_root(big_category, levels)
            if candidate_link and (not summary.final_link or (candidate_is_canonical and not current_is_canonical)):
                summary.big_category = big_category
                summary.final_link = candidate_link
        summaries[key].product_count += 1
        if sales_col:
            summaries[key].monthly_sales += parse_number(ws.cell(row, sales_col).value)

    if total_products:
        for summary in summaries.values():
            summary.product_share = summary.product_count / total_products

    return sorted(
        summaries.values(),
        key=lambda item: (-item.product_count, -item.monthly_sales, item.path_text.lower()),
    )


def clone_category_row_format(ws, template_row: int, target_row: int, max_col: int) -> None:
    for col in range(1, max_col + 1):
        source_col = col if col <= max(ws.max_column, CATEGORY_SALES_COL) else CATEGORY_SALES_COL
        src = ws.cell(template_row, source_col)
        dst = ws.cell(target_row, col)
        if isinstance(dst, MergedCell):
            continue
        copy_cell_format(src, dst)


def ensure_category_rows(ws, data_row_count: int, first_data_row: int, next_section_row: int) -> int:
    existing_rows = max(0, next_section_row - first_data_row)
    if data_row_count > existing_rows:
        insert_rows_preserving_merges(
            ws,
            next_section_row,
            data_row_count - existing_rows,
            first_data_row,
        )
    return max(existing_rows, data_row_count)


def clear_category_area(ws, first_data_row: int, row_count: int, max_col: int) -> None:
    for row in range(first_data_row, first_data_row + row_count):
        for col in range(1, max_col + 1):
            cell = ws.cell(row, col)
            if isinstance(cell, MergedCell):
                continue
            cell.value = None
            cell.hyperlink = None


def apply_category_table_borders(ws, first_data_row: int, row_count: int) -> None:
    thin = CATEGORY_BORDER_SIDE
    inner_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    merge_inner = Border(top=thin, bottom=thin)
    merge_left = Border(left=thin, right=thin, top=thin, bottom=thin)
    merge_right = Border(right=thin, top=thin, bottom=thin)

    for row in range(first_data_row, first_data_row + row_count):
        for col in range(1, CATEGORY_TABLE_END_COL + 1):
            ws.cell(row, col).border = copy(inner_border)

        # The "类目表现" area is F:H. WPS/Excel can render copied
        # merged-cell borders inconsistently after row insertion, so set the
        # visual edges on each cell in the merged range explicitly.
        ws.cell(row, 6).border = copy(merge_left)
        ws.cell(row, 7).border = copy(merge_inner)
        ws.cell(row, 8).border = copy(merge_right)


def set_category_link(cell, link: str | None) -> None:
    if not link:
        return
    cell.hyperlink = link
    cell.font = copy(CATEGORY_LINK_FONT)


def apply_category_share_fill(cell, share: float) -> None:
    if share > CATEGORY_PRIMARY_SHARE_THRESHOLD:
        cell.fill = copy(CATEGORY_PRIMARY_FILL)
    else:
        cell.fill = copy(CATEGORY_TRANSFER_FILL)


def fill_category_section(template_wb, source_wb) -> None:
    src_ws = find_source_sheet(source_wb) or source_wb.worksheets[0]
    research_ws = find_research_sheet(template_wb)
    if research_ws is None:
        return

    section_row = find_row_containing(research_ws, CATEGORY_SECTION_MARKER)
    header_row = find_row_containing(research_ws, CATEGORY_HEADER_MARKER)
    if section_row is None or header_row is None:
        return

    summaries = collect_category_summaries(src_ws)
    first_data_row = header_row + 1
    next_section_row = find_next_section_row(research_ws, first_data_row)
    max_extra_levels = max((max(0, len(item.levels) - CATEGORY_LEVEL_FIXED_COLS) for item in summaries), default=0)
    max_col = max(CATEGORY_SALES_COL, CATEGORY_EXTRA_LEVEL_START_COL + max_extra_levels - 1, research_ws.max_column)
    row_count = ensure_category_rows(research_ws, max(len(summaries), 1), first_data_row, next_section_row)

    for row in range(first_data_row, first_data_row + row_count):
        clone_category_row_format(research_ws, first_data_row, row, max_col)
    clear_category_area(research_ws, first_data_row, row_count, max_col)
    apply_category_table_borders(research_ws, first_data_row, row_count)

    for row_offset, summary in enumerate(summaries):
        row = first_data_row + row_offset
        levels = summary.levels
        translated_levels = summary.translated_levels
        leaf_col = CATEGORY_LEVEL_START_COL
        for index, level in enumerate(levels):
            col = CATEGORY_LEVEL_START_COL + index
            if index >= CATEGORY_LEVEL_FIXED_COLS:
                col = CATEGORY_EXTRA_LEVEL_START_COL + index - CATEGORY_LEVEL_FIXED_COLS
            cell = research_ws.cell(row, col)
            if isinstance(cell, MergedCell):
                continue
            translation = translated_levels[index] if index < len(translated_levels) else ""
            cell.value = category_display_text(level, translation)
            cell.font = copy(CATEGORY_TEXT_FONT)
            if translation:
                alignment = copy(cell.alignment)
                alignment.wrap_text = True
                cell.alignment = alignment
            leaf_col = col

        set_category_link(research_ws.cell(row, leaf_col), summary.final_link)
        apply_category_share_fill(research_ws.cell(row, CATEGORY_SALES_COL), summary.product_share)


def write_missing_category_translation_report(source_wb, output_path: Path) -> None:
    """Record untranslated category levels for the agent post-processing step."""
    src_ws = find_source_sheet(source_wb) or source_wb.worksheets[0]
    missing: dict[str, set[str]] = {}
    for summary in collect_category_summaries(src_ws):
        for index, level in enumerate(summary.levels):
            translation = summary.translated_levels[index] if index < len(summary.translated_levels) else ""
            if not normalize_text(translation):
                missing.setdefault(level, set()).add(summary.path_text)

    payload = {
        "description": "Categories without a source-column or local-cache Chinese translation.",
        "missing_translations": [
            {"english": level, "category_paths": sorted(paths)}
            for level, paths in sorted(missing.items(), key=lambda item: item[0].lower())
        ],
    }
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def fill_data_source(template_wb, source_wb):
    src_ws = find_source_sheet(source_wb) or source_wb.worksheets[0]
    dst_ws = find_source_sheet(template_wb)
    if dst_ws is None:
        raise KeyError("Template workbook does not contain a 数据源 sheet")
    dst_ws = ensure_canonical_source_sheet_name(template_wb, dst_ws)
    src_headers, src_rows = read_rows_by_header(src_ws)
    src_header_cols = header_columns(src_ws)
    dst_headers = [dst_ws.cell(1, col).value for col in range(1, dst_ws.max_column + 1)]

    max_col = len(dst_headers)
    template_style_row = 2 if dst_ws.max_row >= 2 else 1
    for row_idx, row_data in enumerate(src_rows, start=2):
        if row_idx > dst_ws.max_row:
            copy_row_style(dst_ws, template_style_row, row_idx, max_col)
        for col_idx, header in enumerate(dst_headers, start=1):
            value = None
            if header == "图片":
                value = None
            elif header == "year":
                value = extract_year(row_data.get("上架时间"))
            elif header in row_data:
                value = row_data.get(header)
            dst_cell = dst_ws.cell(row_idx, col_idx)
            dst_cell.value = clean(value)
            dst_cell.hyperlink = None
            src_col = src_header_cols.get(str(header)) if header is not None else None
            if src_col:
                link = cell_hyperlink_target(src_ws.cell(row_idx, src_col))
                if link:
                    dst_cell.hyperlink = link

    for row_idx in range(len(src_rows) + 2, dst_ws.max_row + 1):
        for col_idx in range(1, max_col + 1):
            dst_cell = dst_ws.cell(row_idx, col_idx)
            dst_cell.value = None
            dst_cell.hyperlink = None
    remove_duplicate_plain_source_sheet(template_wb, dst_ws)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--template", type=Path, required=True)
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument(
        "--translation-report",
        type=Path,
        help="Optional JSON report listing category levels without a Chinese translation.",
    )
    args = parser.parse_args()

    template_wb = load_workbook(args.template, rich_text=True)
    source_wb = load_workbook(args.source, data_only=False, rich_text=True)

    fill_data_source(template_wb, source_wb)
    fill_category_section(template_wb, source_wb)

    if args.translation_report:
        write_missing_category_translation_report(source_wb, args.translation_report)

    args.output.parent.mkdir(parents=True, exist_ok=True)
    template_wb.save(args.output)
    print(args.output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
