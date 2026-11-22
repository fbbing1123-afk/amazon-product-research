from __future__ import annotations

import argparse
from copy import copy
import json
import re
import shutil
import sys
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

MODULES_DIR = Path(__file__).resolve().parents[1]
if str(MODULES_DIR) not in sys.path:
    sys.path.insert(0, str(MODULES_DIR))

from competitor_analysis.prepare_review_analysis import (
    SCHEMA_VERSION,
    competitor_columns,
    file_sha256,
    find_label_row,
    find_sheet,
    machine_digest,
    normalize_text,
)
from shared.quality_gates import validate_competitor_review_analysis


def visible_length(value: str) -> int:
    return len(re.sub(r"\s+", "", value))


def strip_numeric_prefix(value: Any) -> str:
    return re.sub(r"^\s*\d+[.、)]\s*", "", str(value or "").strip())


def advantage_text(items: list[Any]) -> str:
    values: list[str] = []
    for item in items:
        value = strip_numeric_prefix(item.get("text")) if isinstance(item, dict) else strip_numeric_prefix(item)
        if not value or value in values:
            continue
        if visible_length(value) > 20:
            raise ValueError(f"优点超过 20 字：{value}")
        values.append(value)
    return "\n".join(values)


def disadvantage_text(items: list[Any], max_items: int = 8) -> str:
    values: list[str] = []
    for item in items:
        if isinstance(item, dict):
            value = strip_numeric_prefix(item.get("display"))
        else:
            value = strip_numeric_prefix(item)
        if value.startswith("其他体验问题"):
            continue
        if value and value not in values:
            values.append(value)
        if len(values) >= max_items:
            break
    return "\n".join(values)


def summary_and_improvement_text(analysis: dict[str, Any]) -> str:
    summary = strip_numeric_prefix(analysis.get("overall_summary"))
    improvements: list[tuple[str, str]] = []
    for item in analysis.get("improvements") or []:
        if not isinstance(item, dict):
            continue
        pain_point = strip_numeric_prefix(item.get("pain_point"))
        solution = strip_numeric_prefix(item.get("solution"))
        pair = (pain_point, solution)
        if pain_point and solution and pair not in improvements:
            improvements.append(pair)

    has_review_data = any(product.get("has_review_data") for product in analysis.get("products", []))
    has_disadvantages = any(product.get("disadvantages") for product in analysis.get("products", []))
    if has_review_data and not summary:
        raise ValueError("已有竞品评价数据，但 overall_summary 尚未填写。")
    if has_disadvantages and not improvements:
        raise ValueError("已有竞品缺点，但 improvements 尚未填写。")
    if not summary and not improvements:
        return "/"

    lines = [f"总结：{summary}"] if summary else []
    if improvements:
        if lines:
            lines.append("")
        lines.append("改进方式：")
        lines.extend(f"{pain_point} -> {solution}" for pain_point, solution in improvements)
    return "\n".join(lines)


MINIMUM_COMPETITOR_SLOTS = 5
FIRST_COMPETITOR_COLUMN = 3
COMPARABLE_NUMBER_PATTERN = re.compile(
    r"(?<![A-Z0-9.])"
    r"\d+(?:\.\d+)?"
    r"(?:\s*(?:[x×*]\s*\d+(?:\.\d+)?){1,2}"
    r"|(?:\s*/\s*\d+(?:\.\d+)?)+"
    r"|\s*(?:-|–|—|~|至)\s*\d+(?:\.\d+)?)?"
    r"\s*(?:-|\s)?\s*"
    r"(?:D|K|LBS?|LB|磅|英尺|英寸|FEET|FT|INCH(?:ES)?|CM|MM|M|ML|L|OZ|KG|G|HZ|W|V|A|%|"
    r"个|件|层|档|种|片|枚|张|包|套|次|小时|分钟|秒|度|年|月|天|岁)"
    r"(?![A-Z])",
    re.I,
)


def competitor_layout_last_column(products: list[dict[str, Any]]) -> int:
    actual_last = max(
        (int(item["column_index"]) for item in products),
        default=FIRST_COMPETITOR_COLUMN - 1,
    )
    minimum_last = FIRST_COMPETITOR_COLUMN + MINIMUM_COMPETITOR_SLOTS - 1
    return max(actual_last, minimum_last)


def _copy_cell_style(source, target) -> None:
    target._style = copy(source._style)
    target.number_format = source.number_format
    target.protection = copy(source.protection)
    target.alignment = copy(source.alignment)


def _replace_horizontal_merge(worksheet, row: int, start_column: int, end_column: int) -> None:
    intersecting = [
        merged
        for merged in list(worksheet.merged_cells.ranges)
        if merged.min_row <= row <= merged.max_row
        and merged.min_col <= end_column
        and merged.max_col >= start_column
    ]
    for merged in intersecting:
        worksheet.unmerge_cells(str(merged))
    worksheet.merge_cells(
        start_row=row,
        start_column=start_column,
        end_row=row,
        end_column=end_column,
    )


def comparison_rich_text(value: Any) -> str | CellRichText:
    text = re.sub(r"\s*[；;]\s*", "\n", normalize_text(value))
    text = re.sub(r"\n{2,}", "\n", text).strip()
    matches = list(COMPARABLE_NUMBER_PATTERN.finditer(text))
    if not matches:
        return text
    runs: list[str | TextBlock] = []
    cursor = 0
    red_bold = InlineFont(b=True, color="FFFF0000")

    def append_plain(value: str) -> None:
        if not value:
            return
        # Excel/WPS 可能丢弃单独占一个富文本 run 的纯换行，将其附着到前一个 run。
        if not value.strip() and runs and isinstance(runs[-1], TextBlock):
            previous = runs.pop()
            runs.append(TextBlock(previous.font, previous.text + value))
        else:
            runs.append(value)

    for match in matches:
        if match.start() > cursor:
            append_plain(text[cursor:match.start()])
        runs.append(TextBlock(red_bold, match.group(0)))
        cursor = match.end()
    if cursor < len(text):
        append_plain(text[cursor:])
    return CellRichText(runs)


def format_comparison_dimension_cells(
    worksheet,
    *,
    first_row: int,
    last_row: int,
    last_column: int,
) -> None:
    for row in range(first_row, last_row + 1):
        for column in range(FIRST_COMPETITOR_COLUMN, last_column + 1):
            cell = worksheet.cell(row, column)
            if cell.value in (None, ""):
                continue
            rich_value = comparison_rich_text(cell.value)
            cell.value = rich_value
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def format_selling_point_cells(worksheet, *, row: int, last_column: int) -> None:
    for column in range(FIRST_COMPETITOR_COLUMN, last_column + 1):
        cell = worksheet.cell(row, column)
        if cell.value in (None, ""):
            continue
        cell.value = comparison_rich_text(cell.value)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def restore_summary_template_style(
    worksheet,
    *,
    summary_row: int,
    last_column: int,
    border: Border,
) -> None:
    label = worksheet.cell(summary_row, 1)
    label.font = Font(name="微软雅黑", size=11, bold=True, color="FF000000")
    label.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    content = worksheet.cell(summary_row, FIRST_COMPETITOR_COLUMN)
    content.font = Font(name="微软雅黑", size=11, bold=True, color="FFFF0000")
    content.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    for column in range(1, last_column + 1):
        worksheet.cell(summary_row, column).border = border


def restore_product_images_to_original_size(
    worksheet,
    products: list[dict[str, Any]],
) -> int:
    product_columns = {
        worksheet[f"{item['column']}1"].column
        for item in products
    }
    restored = 0
    for image in getattr(worksheet, "_images", []):
        marker = getattr(getattr(image, "anchor", None), "_from", None)
        if marker is None or marker.row != 1 or marker.col + 1 not in product_columns:
            continue
        image.anchor = f"{get_column_letter(marker.col + 1)}2"
        restored += 1
    return restored


def extend_competitor_layout(
    worksheet,
    products: list[dict[str, Any]],
    *,
    summary_row: int,
    border: Border,
) -> int:
    normalized_products = [
        {
            **item,
            "column_index": worksheet[f"{item['column']}1"].column,
        }
        for item in products
    ]
    actual_last = max(
        (item["column_index"] for item in normalized_products),
        default=FIRST_COMPETITOR_COLUMN - 1,
    )
    layout_last = competitor_layout_last_column(normalized_products)

    _replace_horizontal_merge(worksheet, 1, 1, layout_last)
    _replace_horizontal_merge(worksheet, summary_row, FIRST_COMPETITOR_COLUMN, layout_last)

    if actual_last < layout_last:
        style_source_column = max(actual_last, FIRST_COMPETITOR_COLUMN)
        for column in range(actual_last + 1, layout_last + 1):
            for row in range(2, summary_row):
                source = worksheet.cell(row, style_source_column)
                target = worksheet.cell(row, column)
                _copy_cell_style(source, target)
                target.value = None
                target.hyperlink = None
                target.comment = None
    if layout_last > FIRST_COMPETITOR_COLUMN + MINIMUM_COMPETITOR_SLOTS - 1:
        template_column = FIRST_COMPETITOR_COLUMN + MINIMUM_COMPETITOR_SLOTS - 1
        for column in range(template_column + 1, layout_last + 1):
            for row in range(2, summary_row):
                _copy_cell_style(worksheet.cell(row, template_column), worksheet.cell(row, column))

    product_width = worksheet.column_dimensions[
        get_column_letter(FIRST_COMPETITOR_COLUMN)
    ].width
    for column in range(FIRST_COMPETITOR_COLUMN, layout_last + 1):
        worksheet.column_dimensions[get_column_letter(column)].width = product_width

    for column in range(1, layout_last + 1):
        worksheet.cell(1, column).border = border
    for column in range(FIRST_COMPETITOR_COLUMN, layout_last + 1):
        worksheet.cell(summary_row, column).border = border
    return layout_last


def backfill_review_analysis(analysis_path: Path, workbook_path: Path, output: Path) -> dict[str, int]:
    analysis = json.loads(analysis_path.read_text(encoding="utf-8"))
    schema_version = analysis.get("schema_version")
    if schema_version != SCHEMA_VERSION:
        raise ValueError(f"不支持的竞品分析 handoff 版本：{schema_version}")
    source_digest = str(analysis.get("source_digest") or "").strip()
    if not source_digest or file_sha256(workbook_path) != source_digest:
        raise ValueError("输入工作簿已在 handoff 生成后发生变化，请重新运行 prepare_review_analysis.py。")
    expected_machine_digest = str(analysis.get("machine_digest") or "").strip()
    if not expected_machine_digest or machine_digest(analysis) != expected_machine_digest:
        raise ValueError("竞品分析 handoff 的只读机器字段已发生变化，请重新准备分析 JSON。")
    if workbook_path.resolve() != output.resolve():
        output.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(workbook_path, output)

    workbook = load_workbook(output, rich_text=True)
    try:
        worksheet = find_sheet(workbook, "竞品分析", "竞品分析")
        advantages_row = find_label_row(worksheet, "优点", (2,))
        disadvantages_row = find_label_row(worksheet, "缺点", (2,))
        summary_row = find_label_row(worksheet, "总结及改进方式", (1, 2))
        products = competitor_columns(worksheet)
        columns = {item["asin"]: item["column"] for item in products}
        updated_advantages = 0
        updated_disadvantages = 0

        for product in analysis.get("products", []):
            asin = str(product.get("asin") or "").upper()
            column = columns.get(asin)
            if not column or not product.get("has_review_data"):
                continue
            advantages = advantage_text(product.get("advantages") or [])
            if product.get("positive_review_count", 0) and not advantages:
                raise ValueError(f"{asin} 已抽取好评，但 advantages 尚未填写。")
            worksheet[f"{column}{advantages_row}"] = advantages or "/"
            disadvantages = disadvantage_text(product.get("disadvantages") or [])
            worksheet[f"{column}{disadvantages_row}"] = disadvantages or "/"
            updated_advantages += 1
            updated_disadvantages += 1

        thin = Side(style="thin", color="000000")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        last_column = extend_competitor_layout(
            worksheet,
            products,
            summary_row=summary_row,
            border=border,
        )
        restore_product_images_to_original_size(worksheet, products)
        comparison_row = find_label_row(worksheet, "对比维度", (1, 2))
        format_comparison_dimension_cells(
            worksheet,
            first_row=comparison_row,
            last_row=advantages_row - 1,
            last_column=last_column,
        )
        selling_points_row = find_label_row(worksheet, "产品卖点", (1, 2))
        format_selling_point_cells(
            worksheet,
            row=selling_points_row,
            last_column=last_column,
        )
        for row in (advantages_row, disadvantages_row):
            for column in range(2, last_column + 1):
                cell = worksheet.cell(row, column)
                cell.border = border
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        summary_text = summary_and_improvement_text(analysis)
        summary_cell = worksheet.cell(summary_row, 3)
        summary_cell.value = summary_text
        restore_summary_template_style(
            worksheet,
            summary_row=summary_row,
            last_column=last_column,
            border=border,
        )

        workbook.save(output)
        validate_competitor_review_analysis(output)
        return {
            "advantages_cells": updated_advantages,
            "disadvantages_cells": updated_disadvantages,
            "summary_cells": 1,
        }
    finally:
        workbook.close()


def main() -> int:
    parser = argparse.ArgumentParser(description="把 Agent 总结的竞品优点和差评覆盖率写入竞品分析页。")
    parser.add_argument("--analysis", type=Path, required=True)
    parser.add_argument("--workbook", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    analysis = json.loads(args.analysis.read_text(encoding="utf-8"))
    workbook = args.workbook or Path(analysis["source_workbook"])
    result = backfill_review_analysis(args.analysis, workbook, args.output)
    print(f"advantages_cells={result['advantages_cells']}")
    print(f"disadvantages_cells={result['disadvantages_cells']}")
    print(f"summary_cells={result['summary_cells']}")
    print(args.output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
