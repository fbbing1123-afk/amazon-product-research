from __future__ import annotations

import argparse
import json
import re
import sys
from copy import copy, deepcopy
from io import BytesIO
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.drawing.image import Image
from openpyxl.styles import Border, Font, Side
from openpyxl.utils import get_column_letter, column_index_from_string

MODULES_DIR = Path(__file__).resolve().parents[1]
if str(MODULES_DIR) not in sys.path:
    sys.path.insert(0, str(MODULES_DIR))

from shared.quality_gates import validate_competitor_basics  # noqa: E402
from shared.worksheet_rows import insert_rows_preserving_merges
from competitor_analysis.collect_basic_facts import normalize_value
from competitor_analysis.collect_product_images import validate_asset


ASIN_PATTERN = re.compile(r"(?<![A-Z0-9])([A-Z0-9]{10})(?![A-Z0-9])", re.I)
MAX_COMPARISON_DIMENSIONS = 5
COMPARABLE_NUMBER_PATTERN = re.compile(
    r"(?<![A-Z0-9.])"
    r"\d+(?:\.\d+)?"
    r"(?:\s*(?:[x×*]\s*\d+(?:\.\d+)?){1,2}"
    r"|(?:\s*/\s*\d+(?:\.\d+)?)+"
    r"|\s*(?:-|–|—|~|至)\s*\d+(?:\.\d+)?)?"
    r"\s*(?:-|\s)?\s*"
    r"(?:(?i:KWH|WH|MAH|AH|LBS?|LB|FEET|FT|INCH(?:ES)?|CM|MM|ML|KG|HZ|"
    r"D|K|M|L|OZ|G|W|V|A)(?![A-Za-z])|磅|英尺|英寸|%|"
    r"个|件|层|档|种|片|枚|张|包|套|次|小时|分钟|秒|度|年|月|天|岁)"
)


def normalize_text(value: Any) -> str:
    return str(value or "").strip()


def find_sheet(workbook, suffix: str):
    return next(
        (worksheet for worksheet in workbook.worksheets if worksheet.title.strip().endswith(suffix) and (suffix != "数据源" or not any(word in worksheet.title for word in ("评价", "评论")))),
        None,
    )


def find_label_row(worksheet, label: str) -> int:
    for row in range(1, worksheet.max_row + 1):
        if any(normalize_text(worksheet.cell(row, column).value) == label for column in (1, 2)):
            return row
    raise ValueError(f"工作表“{worksheet.title}”中找不到“{label}”行。")


def competitor_asins(worksheet) -> list[str]:
    found: list[str] = []
    seen: set[str] = set()
    for row in worksheet.iter_rows(min_row=1, max_row=min(5, worksheet.max_row)):
        for cell in row:
            values = [cell.value]
            if cell.hyperlink and cell.hyperlink.target:
                values.append(cell.hyperlink.target)
            for value in values:
                for match in ASIN_PATTERN.finditer(normalize_text(value).upper()):
                    asin = match.group(1)
                    if asin not in seen:
                        seen.add(asin)
                        found.append(asin)
    if not found:
        raise ValueError("竞品分析页顶部没有找到竞品 ASIN。")
    return found


def source_records(worksheet, asins: list[str], fallback_products=None) -> dict[str, dict[str, Any]]:
    headers = {
        normalize_text(cell.value): cell.column
        for cell in worksheet[1]
        if normalize_text(cell.value)
    }
    required = {"ASIN"}
    if missing := required - headers.keys():
        raise ValueError(f"数据源缺少字段：{', '.join(sorted(missing))}")
    targets = set(asins)
    records: dict[str, dict[str, Any]] = {}
    for row in worksheet.iter_rows(min_row=2, values_only=True):
        asin = normalize_text(row[headers["ASIN"] - 1]).upper()
        if asin not in targets:
            continue
        records[asin] = {
            header: row[column - 1] if column - 1 < len(row) else None
            for header, column in headers.items()
        }
    for asin in targets - records.keys():
        asset = (fallback_products or {}).get(asin, {}).get('main_image')
        if not asset:
            raise ValueError(f'数据源缺少竞品：{asin}，请先补取网页主图和基础证据')
        validate_asset(asset, asin)
        rating, review_count = asset.get('rating'), asset.get('review_count')
        if ((rating is not None and (type(rating) not in (int, float) or not 0 <= rating <= 5))
                or (review_count is not None and (type(review_count) is not int or review_count < 0))):
            raise ValueError(f'{asin}网页评分/评论数值无效')
        records[asin] = {'ASIN':asin, '商品详情页链接':asset['page_url'], '评分':rating, '评分数':review_count}
    return records


def rating_summary(record):
    """评分字段可自然为空；保留真实零值，不把未知值改成零。"""
    values = [str(record[key]).strip() if record.get(key) is not None else '' for key in ('评分', '评分数')]
    return ' / '.join(value or '—' for value in values) if any(values) else '/'


def fill_basic_fact_rows(worksheet, column, record, fallback=None):
    """按模板标签定位新增三项，优先数据源，缺项才用已核验网页结果。"""
    definitions = (
        ("monthly_sales", "月销量", ("销量", "月销量")),
        ("price", "价格($)", ("价格", "价格()", "价格($)")),
        ("listed_at", "上架时间", ("上架时间",)),
    )
    for field, header, labels in definitions:
        matches = [cell for row in worksheet.iter_rows(min_col=1, max_col=2)
                   for cell in row if normalize_text(cell.value) in labels]
        if not matches:
            continue
        if len(matches) != 1:
            raise ValueError(f"基础字段标签不唯一：{header}")
        value = normalize_value(field, record.get(header))
        if value is None:
            item = (fallback or {}).get(field)
            value = normalize_value(field, item.get("value") if isinstance(item, dict) else item)
        if value is None:
            raise ValueError(f"{record.get('ASIN', '')} 缺少{header}，请先完成网页补取")
        label = matches[0]
        label.value = header
        worksheet.cell(label.row, column).value = value


def source_main_images(worksheet, asins: list[str]) -> dict[str, bytes]:
    """按数据源行读取产品主图，并重新挂载字节流以避免保存时流被关闭。"""
    headers = {
        normalize_text(cell.value): cell.column
        for cell in worksheet[1]
        if normalize_text(cell.value)
    }
    asin_column = headers.get("ASIN")
    if not asin_column:
        raise ValueError("数据源缺少 ASIN 列。")
    targets = set(asins)
    images: dict[str, bytes] = {}
    for image in list(getattr(worksheet, "_images", [])):
        marker = getattr(image.anchor, "_from", None)
        if marker is None:
            continue
        asin = normalize_text(worksheet.cell(marker.row + 1, asin_column).value).upper()
        if asin not in targets:
            continue
        payload = image._data()
        images[asin] = payload
        image.ref = BytesIO(payload)
    return images


def card_files(cards_dir: Path, asins: list[str]) -> dict[str, Path]:
    files: dict[str, Path] = {}
    for asin in asins:
        matches = sorted(cards_dir.glob(f"*_{asin}.png"))
        if not matches:
            raise ValueError(f"缺少竞品 {asin} 的商品卡 PNG。")
        files[asin] = matches[0]
    return files


def remove_existing_product_images(worksheet, image_row: int, columns: set[int]) -> None:
    kept = []
    for image in worksheet._images:
        anchor = image.anchor
        marker = getattr(anchor, "_from", None)
        if marker is not None and marker.row + 1 == image_row and marker.col + 1 in columns:
            continue
        kept.append(image)
    worksheet._images = kept


def excel_column_width_to_pixels(width: float | None) -> int:
    value = float(width or 13.0)
    return int(value * 12 + 0.5) if value < 1 else int(value * 7 + 5)


def row_height_to_pixels(height: float | None) -> int:
    return int(round(float(height or 15.0) * 96 / 72))


def fit_image_to_cell(image: Image, worksheet, row: int, column: int, margin_pixels: int = 6) -> None:
    """仅缩小显示框，保留嵌入图片的原始像素。"""
    column_letter = worksheet.cell(row, column).column_letter
    cell_width = excel_column_width_to_pixels(worksheet.column_dimensions[column_letter].width)
    cell_height = row_height_to_pixels(worksheet.row_dimensions[row].height)
    available_width = max(1, cell_width - 2 * margin_pixels)
    available_height = max(1, cell_height - 2 * margin_pixels)
    scale = min(1.0, available_width / image.width, available_height / image.height)
    image.width = max(1, int(round(image.width * scale)))
    image.height = max(1, int(round(image.height * scale)))


def apply_thin_borders(worksheet, min_row: int, max_row: int, min_column: int, max_column: int) -> None:
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in worksheet.iter_rows(
        min_row=min_row,
        max_row=max_row,
        min_col=min_column,
        max_col=max_column,
    ):
        for cell in row:
            cell.border = border


def comparison_rich_text(value: Any) -> str | CellRichText:
    """只突出可比较的规格数值，不标红说明文字或列表编号。"""
    text_value = re.sub(r"\s*[；;]\s*", "\n", normalize_text(value))
    text_value = re.sub(r"\n{2,}", "\n", text_value).strip()
    matches = list(COMPARABLE_NUMBER_PATTERN.finditer(text_value))
    if not matches:
        return text_value
    runs: list[str | TextBlock] = []
    cursor = 0
    red_bold = InlineFont(b=True, color="FFFF0000")
    for match in matches:
        if match.start() > cursor:
            runs.append(text_value[cursor:match.start()])
        runs.append(TextBlock(red_bold, match.group(0)))
        cursor = match.end()
    if cursor < len(text_value):
        runs.append(text_value[cursor:])
    return CellRichText(runs)


def replace_horizontal_merge(worksheet, row: int, start_column: int, end_column: int) -> None:
    """把同一行包含起始单元格的横向合并区域扩展到最后一个竞品列。"""
    anchor = worksheet.cell(row, start_column)
    matching_merges = [
        merged
        for merged in list(worksheet.merged_cells.ranges)
        if (
            merged.min_row <= row <= merged.max_row
            and merged.min_col <= start_column <= merged.max_col
        )
    ]
    if (
        len(matching_merges) == 1
        and matching_merges[0].min_row == row
        and matching_merges[0].max_row == row
        and matching_merges[0].min_col == start_column
        and matching_merges[0].max_col == end_column
    ):
        # 合并区域未变化时不得取消再重建，否则会丢失首行富文本和模板样式。
        return

    source_value = deepcopy(anchor.value)
    source_style = copy(anchor._style)
    for merged in matching_merges:
        worksheet.unmerge_cells(str(merged))
    worksheet.merge_cells(
        start_row=row,
        start_column=start_column,
        end_row=row,
        end_column=end_column,
    )
    # 合并范围为新增竞品列时，恢复原标题内容和首行单元格样式。
    target = worksheet.cell(row, start_column)
    target.value = source_value
    target._style = source_style


def prepare_competitor_layout(
    worksheet,
    product_columns: list[int],
    *,
    summary_row: int,
) -> None:
    """统一产品列列宽和新列样式，并扩展标题、总结的合并区域。"""
    first_column = product_columns[0]
    last_column = product_columns[-1]
    product_width = (
        worksheet.column_dimensions[get_column_letter(first_column)].width
        or 30.625
    )
    # 拆开模板的分组列宽，只扩展产品列，避免把 H:P 整组一起变宽或产生重叠 cols。
    dimensions = list(worksheet.column_dimensions.items())
    worksheet.column_dimensions.clear()
    for key, dimension in dimensions:
        start = dimension.min or column_index_from_string(key)
        end = dimension.max or start
        segments = ([(start, end)] if end < first_column or start > last_column else
                    [(start, first_column - 1), (last_column + 1, end)])
        for left, right in segments:
            if left > right:
                continue
            cloned = copy(dimension)
            cloned.min, cloned.max = left, right
            cloned.index = get_column_letter(left)
            worksheet.column_dimensions[cloned.index] = cloned
    for column in range(first_column, last_column + 1):
        original = next((dimension for key, dimension in dimensions
                         if (dimension.min or column_index_from_string(key)) <= column
                         <= (dimension.max or dimension.min or column_index_from_string(key))), None)
        key = get_column_letter(column)
        if original is not None:
            worksheet.column_dimensions[key] = copy(original)
        dimension = worksheet.column_dimensions[key]
        dimension.index = key
        dimension.min = dimension.max = column
        dimension.width = product_width

    # 模板保留 C:G 五个产品栏位；第 6 个起复制 G 列格式，不继承 H:P 的窄列样式。
    style_source_column = min(7, last_column)
    for column in range(8, last_column + 1):
        for row in range(1, summary_row + 1):
            target = worksheet.cell(row, column)
            target._style = copy(worksheet.cell(row, style_source_column)._style)

    replace_horizontal_merge(worksheet, 1, 1, last_column)
    replace_horizontal_merge(worksheet, summary_row, first_column, last_column)


def comparison_display_value(value, label: str) -> str:
    """未知不等于不支持；保留有证据的答案，缺值改为明确的未说明提示。"""
    text = "" if value is None else str(value).strip()
    if text.casefold() not in {"", "/", "-", "--", "n/a", "暂无", "未知", "未说明"}:
        return text
    label = normalize_text(label).rstrip("?？：:")
    if label.startswith("支持"):
        label = "是否" + label
    return f"页面未明确说明{label}" if label else "页面未明确说明该项信息"


def ensure_comparison_rows(worksheet, count: int) -> None:
    """按实际维度数补足行，保持下方评价、总结的内容和格式。"""
    if not 1 <= count <= MAX_COMPARISON_DIMENSIONS:
        raise ValueError(f"对比维度要求 1 到 {MAX_COMPARISON_DIMENSIONS} 个。")
    start = find_label_row(worksheet, "对比维度")
    following = find_label_row(worksheet, "评价分析")
    available = following - start
    if available < 1:
        raise ValueError("对比维度与评价分析区域顺序异常")
    if count <= available:
        return
    label_area = next((area for area in worksheet.merged_cells.ranges
                       if worksheet.cell(start, 1).coordinate in area), None)
    if (normalize_text(worksheet.cell(start, 1).value) != "对比维度"
            or (label_area is not None and
                (label_area.min_col != 1 or label_area.max_col != 1
                 or label_area.min_row != start or label_area.max_row != following - 1))):
        raise ValueError("对比维度标题合并范围异常，不能自动增行")
    if any(worksheet.cell(row, 1).value not in (None, "") for row in range(start + 1, following)):
        raise ValueError("对比维度标题列含其它内容，不能自动合并")
    extra = count - available
    previous_label_area = str(label_area) if label_area is not None else None
    insert_rows_preserving_merges(worksheet, following, extra, following - 1)
    # 在原区域末端插行不会自动延长标题合并区域，因此单独扩展此纵向标题。
    if previous_label_area is not None:
        worksheet.unmerge_cells(previous_label_area)
    worksheet.merge_cells(start_row=start, end_row=start + count - 1, start_column=1, end_column=1)


def populate(
    workbook_path: Path,
    content_path: Path,
    cards_dir: Path,
    output_path: Path,
    images_path: Path | None = None,
    image_display: str = 'original',
) -> None:
    content = json.loads(content_path.read_text(encoding="utf-8-sig"))
    workbook = load_workbook(workbook_path, rich_text=True)
    try:
        competitor = find_sheet(workbook, "竞品分析")
        datasource = find_sheet(workbook, "数据源")
        if competitor is None or datasource is None:
            raise ValueError("工作簿缺少竞品分析页或数据源页。")

        asins = competitor_asins(competitor)
        products = content.get("products") or {}
        if images_path is not None:
            assets = json.loads(images_path.read_text(encoding='utf-8-sig'))['products']
            for asin, asset in assets.items():
                if asin in products:
                    products[asin]['main_image'] = asset
        if image_display not in ('fit','original'):
            raise ValueError('图片显示方式只能为fit或original')
        comparisons = content.get("comparisons") or []
        if set(asins) - set(products):
            raise ValueError("内容 JSON 没有覆盖全部竞品 ASIN。")
        if not 1 <= len(comparisons) <= MAX_COMPARISON_DIMENSIONS:
            raise ValueError(f"对比维度要求 1 到 {MAX_COMPARISON_DIMENSIONS} 个。")

        records = source_records(datasource, asins, products)
        main_images = source_main_images(datasource, asins)
        # 新版优先采用对应商品页的高清原图，旧内容仍可复用数据源图片。
        for asin in asins:
            if products[asin].get('main_image'):
                main_images[asin] = validate_asset(products[asin]['main_image'], asin)
        fallback_asins = [asin for asin in asins if asin not in main_images]
        cards = card_files(cards_dir, fallback_asins) if fallback_asins else {}
        product_columns = list(range(3, 3 + len(asins)))
        if product_columns[-1] > 16384:
            raise ValueError("竞品数量超过 Excel 可用列数。")

        ensure_comparison_rows(competitor, len(comparisons))
        image_row = find_label_row(competitor, "产品图")
        link_row = find_label_row(competitor, "详情页链接")
        rating_row = find_label_row(competitor, "评分/评论数")
        selling_points_row = find_label_row(competitor, "产品卖点")
        comparison_start = find_label_row(competitor, "对比维度")
        evaluation_row = find_label_row(competitor, "评价分析")
        summary_row = find_label_row(competitor, "总结及改进方式")
        available_comparison_rows = list(range(comparison_start, evaluation_row))

        prepare_competitor_layout(
            competitor,
            product_columns,
            summary_row=summary_row,
        )
        remove_existing_product_images(competitor, image_row, set(product_columns))
        blue = "0563C1"
        for asin, column in zip(asins, product_columns):
            record = records[asin]
            product = products[asin]
            fill_basic_fact_rows(competitor, column, record, product.get("basic_facts"))
            link = normalize_text(record.get("商品详情页链接")) or f"https://www.amazon.com/dp/{asin}"
            link_cell = competitor.cell(link_row, column, link)
            link_cell.hyperlink = link
            link_cell.font = Font(
                name=link_cell.font.name or "微软雅黑",
                size=link_cell.font.sz or 10,
                color=blue,
                underline="single",
            )

            competitor.cell(rating_row, column, rating_summary(record))

            points = [normalize_text(item) for item in product.get("selling_points", []) if normalize_text(item)]
            competitor.cell(
                selling_points_row,
                column,
                "\n".join(f"{index}. {point}" for index, point in enumerate(points, start=1)) or "/",
            )
            competitor.cell(
                selling_points_row,
                column,
            ).value = comparison_rich_text(
                competitor.cell(selling_points_row, column).value
            )

            image = (
                Image(BytesIO(main_images[asin]))
                if asin in main_images
                else Image(cards[asin])
            )
            if image_display == 'fit':
                fit_image_to_cell(image, competitor, image_row, column)
            image.anchor = competitor.cell(image_row, column).coordinate
            competitor.add_image(image)

        # 市场调研首格使用竞品页第一款主图，原尺寸浮动，重复运行只替换该位置。
        market = find_sheet(workbook, '市场调研')
        if market is not None:
            first = asins[0]
            remove_existing_product_images(market, 1, {1})
            cover = Image(BytesIO(main_images[first])) if first in main_images else Image(cards[first])
            market.add_image(cover, 'A1')

        for offset, row in enumerate(available_comparison_rows):
            if offset >= len(comparisons):
                competitor.cell(row, 2).value = None
                for column in product_columns:
                    competitor.cell(row, column).value = None
                continue
            comparison = comparisons[offset]
            competitor.cell(row, 2, normalize_text(comparison.get("label")))
            values = comparison.get("values") or {}
            for asin, column in zip(asins, product_columns):
                competitor.cell(
                    row,
                    column,
                    comparison_rich_text(comparison_display_value(values.get(asin), comparison.get("label"))),
                )

        for row in range(image_row, evaluation_row + 2):
            for column in product_columns:
                alignment = copy(competitor.cell(row, column).alignment)
                alignment.wrap_text = True
                alignment.vertical = "center"
                competitor.cell(row, column).alignment = alignment
        apply_thin_borders(
            competitor,
            1,
            competitor.max_row,
            product_columns[0],
            product_columns[-1],
        )

        output_path.parent.mkdir(parents=True, exist_ok=True)
        workbook.save(output_path)
        validate_competitor_basics(output_path, require_evaluation_blank=True)
    finally:
        workbook.close()


def main() -> int:
    parser = argparse.ArgumentParser(description="按竞品顺序填充竞品分析基础信息和原始商品卡。")
    parser.add_argument("--workbook", type=Path, required=True)
    parser.add_argument("--content", type=Path, required=True)
    parser.add_argument("--cards-dir", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument('--images', type=Path, help='已核验的网页主图清单，优先于数据源图片')
    parser.add_argument('--image-display', choices=['fit','original'], default='original')
    args = parser.parse_args()
    populate(args.workbook, args.content, args.cards_dir, args.output, args.images, args.image_display)
    print(args.output.resolve())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
