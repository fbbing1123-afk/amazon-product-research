from __future__ import annotations

import argparse
import io
import re
import shutil
import zipfile
from copy import copy
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from PIL import Image as PILImage, ImageStat


VIEW_SHEET = "差评透视"
LIBRARY_SHEET = "差评透视图片库"
SOURCE_SHEET = "差评痛点"
TABLE_NAME = "tblPainReviewView"
HEADER_ROW = 4
DATA_START_ROW = HEADER_ROW + 1
LIBRARY_BLOCK_ROWS = 84
LIBRARY_BLOCK_COLUMNS = 4
LIBRARY_ROW_HEIGHT = 14.1
ORANGE_FILL = "FFED7D31"

DISPLAY_HEADERS = [
    "星级",
    "评论日期",
    "ASIN",
    "属性/型号",
    "差评词/短语",
    "归一化痛点（中文）",
    "Helpful数",
    "图片数量",
    "评论图片",
    "评价内容",
    "评论链接",
]
HELPER_HEADER = "可见行"
MISSING_MODEL_LABEL = "未注明型号"
MISSING_PHRASE_LABEL = "未提取词/短语"
UNCLASSIFIED_PAIN_LABEL = "未归类痛点"


def normalize(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def split_lines(value: Any) -> list[str]:
    text = str(value or "").replace("；", "\n").replace(";", "\n")
    return [
        normalized
        for item in text.splitlines()
        if (normalized := normalize(item)) and normalized != "/"
    ]


def extract_asin(value: Any) -> str:
    match = re.search(r"(?<![A-Z0-9])([A-Z0-9]{10})(?![A-Z0-9])", str(value or "").upper())
    return match.group(1) if match else ""


def find_sheet(workbook, name: str):
    if name in workbook.sheetnames:
        return workbook[name]
    for worksheet in workbook.worksheets:
        if worksheet.title.endswith(name):
            return worksheet
    raise ValueError(f"工作簿中缺少“{name}”sheet。")


def find_label_row(worksheet, label: str) -> int:
    for row in range(1, worksheet.max_row + 1):
        for column in (1, 2):
            if normalize(worksheet.cell(row, column).value) == label:
                return row
    raise ValueError(f"{worksheet.title} 中没有找到“{label}”行。")


def header_map(worksheet) -> dict[str, int]:
    return {
        normalize(worksheet.cell(1, column).value): column
        for column in range(1, worksheet.max_column + 1)
        if normalize(worksheet.cell(1, column).value)
    }


def workbook_has_slicers(path: Path) -> bool:
    with zipfile.ZipFile(path) as archive:
        return any(name.startswith("xl/slicers/") for name in archive.namelist())


def read_image_payload(image) -> bytes:
    payload = image._data()
    # openpyxl closes an image's source stream while reading it. Resetting the
    # source keeps the original drawing saveable in the same workbook pass.
    image.ref = io.BytesIO(payload)
    return payload


def is_meaningful_review_image(payload: bytes) -> bool:
    """过滤下载失败后形成的近乎纯白占位图。"""
    try:
        with PILImage.open(io.BytesIO(payload)) as picture:
            grayscale = picture.convert("L")
            minimum, maximum = grayscale.getextrema()
            deviation = ImageStat.Stat(grayscale).stddev[0]
            return maximum - minimum >= 24 or deviation >= 12
    except Exception:
        return False


def image_payloads_by_row(worksheet, image_column: int) -> dict[int, list[bytes]]:
    payloads: dict[int, list[bytes]] = {}
    for image in getattr(worksheet, "_images", []):
        marker = getattr(getattr(image, "anchor", None), "_from", None)
        if marker is None or marker.col + 1 != image_column:
            continue
        payload = read_image_payload(image)
        if is_meaningful_review_image(payload):
            payloads.setdefault(marker.row + 1, []).append(payload)
    return payloads


def product_card_payloads(workbook) -> tuple[list[str], dict[str, bytes]]:
    competitor = next(
        (worksheet for worksheet in workbook.worksheets if worksheet.title.endswith("竞品分析")),
        None,
    )
    if competitor is None:
        raise ValueError("工作簿中缺少竞品分析 sheet。")
    link_row = find_label_row(competitor, "详情页链接")
    asin_by_column: dict[int, str] = {}
    ordered_asins: list[str] = []
    for column in range(3, competitor.max_column + 1):
        cell = competitor.cell(link_row, column)
        asin = extract_asin(cell.hyperlink.target if cell.hyperlink else cell.value)
        if not asin:
            continue
        asin_by_column[column] = asin
        ordered_asins.append(asin)

    # 某些模板只在标题中列出 ASIN，详情页链接行为空；沿用竞品分析页的标题顺序作为回退。
    if not ordered_asins:
        title_text = " ".join(
            str(cell.value or "")
            for row in competitor.iter_rows(min_row=1, max_row=min(2, competitor.max_row))
            for cell in row
        )
        ordered_asins = list(dict.fromkeys(re.findall(r"(?<![A-Z0-9])([A-Z0-9]{10})(?![A-Z0-9])", title_text.upper())))

    cards: dict[str, bytes] = {}
    for image in getattr(competitor, "_images", []):
        marker = getattr(getattr(image, "anchor", None), "_from", None)
        if marker is None or marker.row != 1:
            continue
        asin = asin_by_column.get(marker.col + 1)
        if asin:
            cards[asin] = read_image_payload(image)

    # 竞品页可能没有浮动商品卡，但数据源页仍保留原始商品图；按 ASIN 回退读取，避免切片器图片库为空。
    if len(cards) < len(ordered_asins):
        source = next(
            (worksheet for worksheet in workbook.worksheets if worksheet.title.endswith("数据源")),
            None,
        )
        if source is not None:
            headers = header_map(source)
            asin_column = headers.get("ASIN")
            image_column = headers.get("图片")
            if asin_column and image_column:
                asin_by_row = {
                    row: extract_asin(source.cell(row, asin_column).value)
                    for row in range(2, source.max_row + 1)
                }
                for image in getattr(source, "_images", []):
                    marker = getattr(getattr(image, "anchor", None), "_from", None)
                    if marker is None or marker.col + 1 != image_column:
                        continue
                    asin = asin_by_row.get(marker.row + 1)
                    if asin in ordered_asins and asin not in cards:
                        cards[asin] = read_image_payload(image)

    # 竞品页缺少产品卡时，将数据源原图补到对应产品图单元格，供后续 Excel 切片器校验和展示。
    existing_columns = {
        getattr(getattr(image, "anchor", None), "_from", None).col + 1
        for image in getattr(competitor, "_images", [])
        if getattr(getattr(image, "anchor", None), "_from", None) is not None
    }
    for offset, asin in enumerate(ordered_asins):
        column = 3 + offset
        payload = cards.get(asin)
        if payload is None or column in existing_columns:
            continue
        image = clone_image_original(payload)
        image.anchor = f"{get_column_letter(column)}2"
        competitor.add_image(image)
        existing_columns.add(column)
    return ordered_asins, cards


def expanded_pain_rows(source) -> tuple[list[dict[str, Any]], dict[int, list[bytes]]]:
    headers = header_map(source)
    missing = [header for header in DISPLAY_HEADERS if header not in headers and header != "ASIN"]
    asin_header = "竞品ASIN" if "竞品ASIN" in headers else "ASIN"
    if asin_header not in headers:
        missing.append("竞品ASIN")
    if missing:
        raise ValueError(f"差评痛点 sheet 缺少字段：{', '.join(missing)}")

    image_payloads = image_payloads_by_row(source, headers["评论图片"])
    rows: list[dict[str, Any]] = []
    for source_row in range(2, source.max_row + 1):
        asin = extract_asin(source.cell(source_row, headers[asin_header]).value)
        if not asin:
            continue
        raw_phrases = split_lines(source.cell(source_row, headers["差评词/短语"]).value)
        pain_points = split_lines(source.cell(source_row, headers["归一化痛点（中文）"]).value)
        if not pain_points:
            pain_points = [UNCLASSIFIED_PAIN_LABEL]
        model = normalize(source.cell(source_row, headers["属性/型号"]).value)
        if not model or model == "/":
            model = MISSING_MODEL_LABEL
        row_images = image_payloads.get(source_row, [])
        image_count = source.cell(source_row, headers["图片数量"]).value or 0
        if not row_images:
            image_count = 0
        for index, pain_point in enumerate(pain_points):
            if len(raw_phrases) == len(pain_points):
                phrase = raw_phrases[index]
            else:
                phrase = "\n".join(raw_phrases)
            row = {
                "星级": source.cell(source_row, headers["星级"]).value,
                "评论日期": source.cell(source_row, headers["评论日期"]).value,
                "ASIN": asin,
                "属性/型号": model,
                "差评词/短语": phrase or MISSING_PHRASE_LABEL,
                "归一化痛点（中文）": pain_point,
                "Helpful数": source.cell(source_row, headers["Helpful数"]).value or 0,
                "图片数量": image_count,
                "评论图片": "",
                "评价内容": source.cell(source_row, headers["评价内容"]).value or "/",
                "评论链接": source.cell(source_row, headers["评论链接"]).value,
                "_source_row": source_row,
            }
            rows.append(row)
    return rows, image_payloads


def empty_pain_row(asin: str) -> dict[str, Any]:
    return {
        "星级": "",
        "评论日期": None,
        "ASIN": asin,
        "属性/型号": MISSING_MODEL_LABEL,
        "差评词/短语": MISSING_PHRASE_LABEL,
        "归一化痛点（中文）": UNCLASSIFIED_PAIN_LABEL,
        "Helpful数": 0,
        "图片数量": 0,
        "评论图片": "",
        "评价内容": "暂无可信差评",
        "评论链接": "",
        "_source_row": -1,
    }


def clone_image(payload: bytes, max_width: int, max_height: int) -> ExcelImage:
    buffer = io.BytesIO(payload)
    image = ExcelImage(buffer)
    with PILImage.open(io.BytesIO(payload)) as picture:
        scale = min(max_width / picture.width, max_height / picture.height, 1.0)
        image.width = round(picture.width * scale)
        image.height = round(picture.height * scale)
    return image


def clone_image_original(payload: bytes) -> ExcelImage:
    return ExcelImage(io.BytesIO(payload))


def set_border(worksheet, min_row: int, max_row: int, min_column: int, max_column: int) -> None:
    thin = Side(style="thin", color="FF7F8C99")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in worksheet.iter_rows(
        min_row=min_row,
        max_row=max_row,
        min_col=min_column,
        max_col=max_column,
    ):
        for cell in row:
            cell.border = border


def set_outline_border(worksheet, cell_range: str, color: str = "FF000000") -> None:
    cells = worksheet[cell_range]
    min_row = cells[0][0].row
    max_row = cells[-1][0].row
    min_column = cells[0][0].column
    max_column = cells[0][-1].column
    side = Side(style="thin", color=color)
    for row in range(min_row, max_row + 1):
        for column in range(min_column, max_column + 1):
            cell = worksheet.cell(row, column)
            cell.border = Border(
                left=side if column == min_column else Side(style=None),
                right=side if column == max_column else Side(style=None),
                top=side if row == min_row else Side(style=None),
                bottom=side if row == max_row else Side(style=None),
            )


def create_view_sheet(workbook, rows: list[dict[str, Any]], review_images: dict[int, list[bytes]]) -> None:
    for name in (VIEW_SHEET, LIBRARY_SHEET):
        if name in workbook.sheetnames:
            del workbook[name]
    view = workbook.create_sheet(VIEW_SHEET)
    view.sheet_view.showGridLines = False
    view.freeze_panes = f"A{DATA_START_ROW}"

    view.merge_cells("A1:K1")
    title = view["A1"]
    title.value = "差评透视"
    title.fill = PatternFill("solid", fgColor=ORANGE_FILL)
    title.font = Font(name="微软雅黑", size=16, bold=True, color="FF000000")
    title.alignment = Alignment(horizontal="center", vertical="center")
    for row in view["A1:K1"]:
        for cell in row:
            cell.fill = PatternFill("solid", fgColor=ORANGE_FILL)

    view.merge_cells("A2:D2")
    view.merge_cells("E2:K2")
    view["A2"] = "产品图片"
    view["E2"] = "维度筛选"
    for row in view["A2:K2"]:
        for cell in row:
            cell.fill = PatternFill("solid", fgColor=ORANGE_FILL)
            cell.font = Font(name="微软雅黑", size=16, bold=True, color="FF000000")
            cell.alignment = Alignment(horizontal="center", vertical="center")
    view.merge_cells("A3:D3")
    view.merge_cells("E3:K3")
    view.row_dimensions[1].height = 36
    view.row_dimensions[2].height = 19
    view.row_dimensions[3].height = 409
    set_outline_border(view, "A1:K1")
    set_outline_border(view, "A2:D2")
    set_outline_border(view, "E2:K2")
    set_outline_border(view, "A3:D3", "FF999999")
    set_outline_border(view, "E3:K3", "FF999999")

    last_row = DATA_START_ROW + max(len(rows), 1) - 1
    view["M1"] = f'=IFERROR(INDEX($C${DATA_START_ROW}:$C${last_row},MATCH(1,$L${DATA_START_ROW}:$L${last_row},0)),"")'

    for column, header in enumerate(DISPLAY_HEADERS + [HELPER_HEADER], start=1):
        cell = view.cell(HEADER_ROW, column)
        cell.value = header
        cell.fill = PatternFill("solid", fgColor=ORANGE_FILL)
        cell.font = Font(name="微软雅黑", size=10, bold=True, color="FFFFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    if not rows:
        rows = [empty_pain_row("")]
    for offset, item in enumerate(rows):
        row_number = DATA_START_ROW + offset
        for column, header in enumerate(DISPLAY_HEADERS, start=1):
            cell = view.cell(row_number, column)
            value = item.get(header)
            if header == "评论链接":
                if value:
                    cell.value = "查看评论"
                    cell.hyperlink = str(value)
                    cell.font = Font(name="微软雅黑", size=10, color="FF0563C1", underline="single")
                else:
                    cell.value = "/"
            elif header != "评论图片":
                cell.value = value
            if header == "评论日期":
                cell.number_format = "yyyy-mm-dd"
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        view.cell(row_number, len(DISPLAY_HEADERS) + 1).value = f"=SUBTOTAL(103,$C{row_number})"
        payloads = review_images.get(int(item.get("_source_row", -1)), [])
        for payload in payloads:
            image = clone_image(payload, 110, 76)
            image.anchor = f"I{row_number}"
            view.add_image(image)
        view.row_dimensions[row_number].height = 66 if payloads else 42

    widths = [8, 16, 15, 24, 25, 22, 10, 10, 18, 68, 14, 10]
    for column, width in enumerate(widths, start=1):
        view.column_dimensions[get_column_letter(column)].width = width
    view.column_dimensions["L"].hidden = True
    view.column_dimensions["M"].hidden = True
    view.row_dimensions[HEADER_ROW].height = 34
    set_border(view, HEADER_ROW, last_row, 1, len(DISPLAY_HEADERS))
    view.auto_filter.ref = f"A{HEADER_ROW}:K{last_row}"
    view.print_title_rows = f"1:{HEADER_ROW}"
    view.sheet_properties.pageSetUpPr.fitToPage = True
    view.page_setup.fitToWidth = 1
    view.page_setup.fitToHeight = 0


def create_image_library(workbook, ordered_asins: list[str], cards: dict[str, bytes]) -> None:
    library = workbook.create_sheet(LIBRARY_SHEET)
    library.sheet_view.showGridLines = False
    library["A1"] = "ASIN"
    library["B1"] = "起始行"
    library["C1"] = "状态"
    for column in range(4, 4 + LIBRARY_BLOCK_COLUMNS):
        library.column_dimensions[get_column_letter(column)].width = 8.5

    for index, asin in enumerate(ordered_asins):
        index_row = index + 2
        start_row = 2 + index * LIBRARY_BLOCK_ROWS
        library.cell(index_row, 1).value = asin
        library.cell(index_row, 2).value = start_row
        payload = cards.get(asin)
        library.cell(index_row, 3).value = "ready" if payload else "missing"
        for row in range(start_row, start_row + LIBRARY_BLOCK_ROWS):
            library.row_dimensions[row].height = LIBRARY_ROW_HEIGHT
            for column in range(4, 4 + LIBRARY_BLOCK_COLUMNS):
                library.cell(row, column).fill = PatternFill("solid", fgColor="FFFFFFFF")
        if payload:
            image = clone_image_original(payload)
            # Keep one row of padding so Excel linked pictures never include
            # the first pixels of the next product card at a block boundary.
            image.anchor = f"D{start_row + 1}"
            library.add_image(image)
        else:
            library.cell(start_row + 12, 5).value = "暂无商品图"
            library.cell(start_row + 12, 5).alignment = Alignment(horizontal="center")
    library.sheet_state = "hidden"


def build_pain_slicer_view(source_path: Path, output_path: Path) -> dict[str, Any]:
    source_path = source_path.resolve()
    output_path = output_path.resolve()
    if workbook_has_slicers(source_path):
        raise ValueError("输入工作簿已经包含切片器。请使用创建切片器之前的工作副本重新生成，避免 openpyxl 删除其它切片器。")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if source_path != output_path:
        shutil.copy2(source_path, output_path)

    workbook = load_workbook(output_path, rich_text=True)
    try:
        source = find_sheet(workbook, SOURCE_SHEET)
        rows, review_images = expanded_pain_rows(source)
        review_row_count = len(rows)
        ordered_asins, cards = product_card_payloads(workbook)
        if not rows:
            rows = [empty_pain_row(asin) for asin in ordered_asins]
        del workbook[source.title]
        create_view_sheet(workbook, rows, review_images)
        create_image_library(workbook, ordered_asins, cards)
        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
        workbook.calculation.calcMode = "auto"
        workbook.save(output_path)
    finally:
        workbook.close()
    return {
        "output": str(output_path),
        "review_rows": review_row_count,
        "asin_count": len(ordered_asins),
        "product_cards": len(cards),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="生成差评透视页及切片器所需的数据和图片库。")
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    result = build_pain_slicer_view(args.source, args.output)
    for key, value in result.items():
        print(f"{key}: {value}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
