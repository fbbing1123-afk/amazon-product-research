from __future__ import annotations

import argparse
import io
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable

from openpyxl import load_workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.utils import get_column_letter
from PIL import Image
import xml.etree.ElementTree as ET


ERROR_PATTERN = re.compile(
    r"#(?:REF!|DIV/0!|VALUE!|NAME\?|N/A|NUM!|NULL!|SPILL!|CALC!)",
    re.IGNORECASE,
)
ASIN_PATTERN = re.compile(r"(?<![A-Z0-9])([A-Z0-9]{10})(?![A-Z0-9])", re.I)


class QualityGateError(RuntimeError):
    """模块输出未通过硬门禁。"""

    def __init__(self, gate: str, failures: Iterable[str]):
        self.gate = gate
        self.failures = list(dict.fromkeys(str(item) for item in failures))
        message = "\n".join(f"- {item}" for item in self.failures)
        super().__init__(f"质量门禁失败 [{gate}]：\n{message}")


@dataclass
class GateReport:
    gate: str
    workbook: Path
    metrics: dict[str, int] = field(default_factory=dict)

    def as_dict(self) -> dict[str, object]:
        return {
            "gate": self.gate,
            "status": "passed",
            "workbook": str(self.workbook),
            "metrics": self.metrics,
        }


@dataclass(frozen=True)
class ImageEvidence:
    payload_size: tuple[int, int] | None
    display_size: tuple[int, int] | None
    real: bool


def text(value: object) -> str:
    return str(value or "").strip()


def find_sheet(workbook, suffix: str):
    return next(
        (sheet for sheet in workbook.worksheets if sheet.title.endswith(suffix)),
        None,
    )


def find_row(sheet, label: str, columns: tuple[int, ...] = (1, 2)) -> int | None:
    for row in range(1, sheet.max_row + 1):
        if any(text(sheet.cell(row, column).value) == label for column in columns):
            return row
    return None


def find_row_containing(
    sheet,
    marker: str,
    columns: tuple[int, ...] = tuple(range(1, 11)),
) -> int | None:
    for row in range(1, sheet.max_row + 1):
        if any(marker in text(sheet.cell(row, column).value) for column in columns):
            return row
    return None


def image_anchor(image) -> tuple[int, int] | None:
    marker = getattr(getattr(image, "anchor", None), "_from", None)
    if marker is None:
        return None
    return marker.row + 1, marker.col + 1


def image_display_size(image) -> tuple[int, int] | None:
    extent = getattr(getattr(image, "anchor", None), "ext", None)
    if extent is None:
        return None
    cx = int(getattr(extent, "width", getattr(extent, "cx", 0)) or 0)
    cy = int(getattr(extent, "height", getattr(extent, "cy", 0)) or 0)
    if cx <= 0 or cy <= 0:
        return None
    return max(1, round(cx / 9525)), max(1, round(cy / 9525))


def image_payload_size(image) -> tuple[int, int] | None:
    try:
        payload = image._data()
        with Image.open(io.BytesIO(payload)) as picture:
            return picture.width, picture.height
    except Exception:
        return None


def image_is_real(image, minimum_nonwhite: float = 0.001) -> bool:
    try:
        payload = image._data()
        with Image.open(io.BytesIO(payload)) as picture:
            preview = picture.convert("RGB")
            preview.thumbnail((64, 64))
            pixels = list(preview.getdata())
        nonwhite = sum(1 for pixel in pixels if min(pixel) < 245)
        return bool(pixels) and nonwhite / len(pixels) >= minimum_nonwhite
    except Exception:
        return False


def payload_image_info(payload: bytes) -> tuple[tuple[int, int] | None, bool]:
    try:
        with Image.open(io.BytesIO(payload)) as picture:
            size = (picture.width, picture.height)
            preview = picture.convert("RGB")
            preview.thumbnail((64, 64))
            pixels = list(preview.getdata())
        nonwhite = sum(1 for pixel in pixels if min(pixel) < 245)
        return size, bool(pixels) and nonwhite / len(pixels) >= 0.001
    except Exception:
        return None, False


def wps_cell_image_data(path: Path) -> dict[str, ImageEvidence]:
    """读取 WPS cellimages 中的编号、原始像素和显示尺寸。"""
    if not zipfile.is_zipfile(path):
        return {}
    rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    rel_id_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xdr_ns = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
    a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        if "xl/cellimages.xml" not in names or "xl/_rels/cellimages.xml.rels" not in names:
            return {}
        root = ET.fromstring(archive.read("xl/cellimages.xml"))
        rel_root = ET.fromstring(archive.read("xl/_rels/cellimages.xml.rels"))
        relations = {
            item.attrib.get("Id", ""): item.attrib.get("Target", "")
            for item in rel_root.findall(f"{{{rel_ns}}}Relationship")
        }
        result: dict[str, ImageEvidence] = {}
        for item in root:
            props = item.find(f".//{{{xdr_ns}}}cNvPr")
            blip = item.find(f".//{{{a_ns}}}blip")
            extent = item.find(f".//{{{a_ns}}}ext")
            if props is None or blip is None:
                continue
            image_id = props.attrib.get("name", "")
            rel_id = blip.attrib.get(f"{{{rel_id_ns}}}embed", "")
            target = relations.get(rel_id, "")
            media_path = (
                target.lstrip("/")
                if target.startswith("/")
                else f"xl/{target}"
            )
            if not image_id or media_path not in names:
                continue
            payload = archive.read(media_path)
            payload_size, real = payload_image_info(payload)
            display_size = None
            if extent is not None:
                cx = int(extent.attrib.get("cx", "0") or 0)
                cy = int(extent.attrib.get("cy", "0") or 0)
                if cx > 0 and cy > 0:
                    display_size = (round(cx / 9525), round(cy / 9525))
            result[image_id] = ImageEvidence(payload_size, display_size, real)
    return result


def sheet_image_evidence(path: Path, workbook, sheet) -> dict[tuple[int, int], ImageEvidence]:
    """合并普通浮动图片和 WPS DISPIMG 单元格图片证据。"""
    result: dict[tuple[int, int], ImageEvidence] = {}
    for image in getattr(sheet, "_images", []):
        anchor = image_anchor(image)
        if anchor is None:
            continue
        try:
            payload = image._data()
            payload_size, real = payload_image_info(payload)
        except Exception:
            payload_size, real = None, False
        result[anchor] = ImageEvidence(
            payload_size,
            image_display_size(image),
            real,
        )
    wps_images = wps_cell_image_data(path)
    if not wps_images:
        return result
    for row in sheet.iter_rows():
        for cell in row:
            if not isinstance(cell.value, str):
                continue
            match = re.search(r'DISPIMG\("([^"]+)"', cell.value, re.IGNORECASE)
            if match and match.group(1) in wps_images:
                result[(cell.row, cell.column)] = wps_images[match.group(1)]
    return result


def duplicate_anchors(images: Iterable[object]) -> set[tuple[int, int]]:
    seen: set[tuple[int, int]] = set()
    duplicates: set[tuple[int, int]] = set()
    for image in images:
        anchor = image_anchor(image)
        if anchor is None:
            continue
        if anchor in seen:
            duplicates.add(anchor)
        seen.add(anchor)
    return duplicates


def cell_formula_errors(workbook) -> list[str]:
    errors: list[str] = []
    for sheet in workbook.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and ERROR_PATTERN.search(cell.value):
                    errors.append(f"{sheet.title}!{cell.coordinate}: {cell.value}")
    return errors


def color_signature(color) -> tuple | None:
    """把主题色和 RGB 色统一转换成可比较的稳定值。"""
    if color is None:
        return None
    color_type = getattr(color, "type", None)
    return (
        color_type,
        getattr(color, "rgb", None) if color_type == "rgb" else None,
        getattr(color, "theme", None) if color_type == "theme" else None,
        round(float(getattr(color, "tint", 0) or 0), 6),
    )


def font_signature(font) -> tuple:
    return (
        font.name,
        font.sz,
        font.b,
        font.i,
        font.u,
        font.strike,
        font.vertAlign,
        font.scheme,
        color_signature(font.color),
    )


def cell_format_signature(cell) -> tuple:
    """提取竞品标题首行需要逐项保持的模板格式。"""
    border = tuple(
        (
            side,
            getattr(getattr(cell.border, side), "style", None),
            color_signature(getattr(getattr(cell.border, side), "color", None)),
        )
        for side in ("left", "right", "top", "bottom")
    )
    return (
        font_signature(cell.font),
        cell.fill.fill_type,
        cell.fill.patternType,
        color_signature(cell.fill.fgColor),
        color_signature(cell.fill.bgColor),
        cell.alignment.horizontal,
        cell.alignment.vertical,
        cell.alignment.wrap_text,
        cell.alignment.shrink_to_fit,
        cell.alignment.indent,
        border,
        cell.number_format,
    )


def rich_text_signature(value) -> tuple:
    if not isinstance(value, CellRichText):
        return ("plain", text(value))
    runs = []
    for item in value:
        if isinstance(item, TextBlock):
            inline_font = item.font
            runs.append(
                (
                    "rich",
                    text(item.text),
                    inline_font.rFont,
                    inline_font.sz,
                    inline_font.b,
                    inline_font.i,
                    inline_font.u,
                    inline_font.strike,
                    inline_font.vertAlign,
                    color_signature(inline_font.color),
                )
            )
        else:
            runs.append(("plain", text(item)))
    return ("rich", tuple(runs))


def validate_competitor_title_format(path: Path, template_path: Path) -> GateReport:
    """确保竞品标题首行没有因列扩展或合并操作丢失模板格式。"""
    failures: list[str] = []
    target_wb = load_workbook(path, data_only=False, rich_text=True)
    template_wb = load_workbook(template_path, data_only=False, rich_text=True)
    try:
        target = find_sheet(target_wb, "竞品分析")
        template = find_sheet(template_wb, "竞品分析")
        if target is None or template is None:
            failures.append("缺少竞品分析工作表，无法核对标题格式")
            raise QualityGateError("competitor_title_format", failures)

        target_title = target["A1"]
        template_title = template["A1"]
        if text(target_title.value) != text(template_title.value):
            failures.append("竞品分析标题文本与原模板不一致")
        if rich_text_signature(target_title.value) != rich_text_signature(template_title.value):
            failures.append("竞品分析标题富文本字体或分段格式与原模板不一致")
        if cell_format_signature(target_title) != cell_format_signature(template_title):
            failures.append("竞品分析标题首行字体、填充、对齐或边框与原模板不一致")
        if target.row_dimensions[1].height != template.row_dimensions[1].height:
            failures.append(
                "竞品分析首行行高不一致："
                f"{target.row_dimensions[1].height} != {template.row_dimensions[1].height}"
            )

        template_merge = next(
            (
                merged
                for merged in template.merged_cells.ranges
                if merged.min_row == 1 and merged.min_col == 1
            ),
            None,
        )
        target_merge = next(
            (
                merged
                for merged in target.merged_cells.ranges
                if merged.min_row == 1 and merged.min_col == 1
            ),
            None,
        )
        if template_merge is None or target_merge is None:
            failures.append("竞品分析标题合并区域缺失")
        elif target_merge.max_col < template_merge.max_col:
            failures.append("竞品分析标题合并区域比原模板更窄")

        if failures:
            raise QualityGateError("competitor_title_format", failures)
        return GateReport(
            "competitor_title_format",
            path,
            {"cells": 1, "rich_text": int(isinstance(target_title.value, CellRichText))},
        )
    finally:
        target_wb.close()
        template_wb.close()


def validate_keyword_research(path: Path) -> GateReport:
    failures: list[str] = []
    workbook = load_workbook(path, data_only=False, rich_text=True)
    try:
        sheet = find_sheet(workbook, "市场调研")
        if sheet is None:
            failures.append("缺少市场调研工作表")
            raise QualityGateError("keyword_research", failures)
        header_row = find_row(sheet, "关键词", tuple(range(1, 10)))
        next_section = find_row_containing(sheet, "所属平台分类及常用分类节点")
        if header_row is None or next_section is None:
            failures.append("找不到关键词表头或类目分区起点")
            raise QualityGateError("keyword_research", failures)
        rows = list(range(header_row + 1, next_section))
        if not rows:
            failures.append("关键词区域没有数据行")
        image_evidence = sheet_image_evidence(path, workbook, sheet)

        for row in rows:
            keyword = text(sheet.cell(row, 1).value)
            if not keyword:
                failures.append(f"第 {row} 行关键词为空")
                continue
            link = sheet.cell(row, 1).hyperlink
            target = text(link.target if link else "")
            if "amazon." not in target.lower() or "/s?" not in target.lower():
                failures.append(f"关键词“{keyword}”缺少 Amazon 搜索超链接")
            for column, label in (
                (5, "趋势旺季"),
                (6, "站内结果呈现"),
                (7, "淡季月搜索量"),
                (8, "旺季月搜索量"),
                (9, "旺季竞价"),
            ):
                if sheet.cell(row, column).value in (None, ""):
                    failures.append(f"关键词“{keyword}”的{label}为空")
            image = image_evidence.get((row, 2))
            if image is None:
                failures.append(f"关键词“{keyword}”趋势图缺失")
                continue
            if not image.real:
                failures.append(f"关键词“{keyword}”趋势图为空白或无法读取")
            payload = image.payload_size
            display = image.display_size
            if payload and display and payload[0] <= display[0]:
                failures.append(
                    f"关键词“{keyword}”趋势图未保留高清原图："
                    f"原图 {payload[0]}px，显示 {display[0]}px"
                )
        if errors := cell_formula_errors(workbook):
            failures.extend(f"公式错误：{item}" for item in errors[:20])
        if failures:
            raise QualityGateError("keyword_research", failures)
        return GateReport(
            "keyword_research",
            path,
            {"rows": len(rows), "images": sum(1 for row in rows if (row, 2) in image_evidence)},
        )
    finally:
        workbook.close()


def validate_category_research(
    path: Path,
    *,
    screenshot_limit: int | None = None,
) -> GateReport:
    failures: list[str] = []
    workbook = load_workbook(path, data_only=False, rich_text=True)
    try:
        sheet = find_sheet(workbook, "市场调研")
        if sheet is None:
            failures.append("缺少市场调研工作表")
            raise QualityGateError("category_research", failures)
        header_row = find_row(sheet, "一级类目")
        # 新模板把整体销售数据放在市场调研页，类目表以新旧两种标题为边界。
        boundaries = [find_row_containing(sheet, label)
                      for label in ("数据分析概览", "市场数据分析")]
        next_section = min((row for row in boundaries if row is not None
                            and header_row is not None and row > header_row), default=None)
        if header_row is None or next_section is None:
            failures.append("找不到类目表头或市场数据分析/数据分析概览起点")
            raise QualityGateError("category_research", failures)
        rows = [
            row
            for row in range(header_row + 1, next_section)
            if text(sheet.cell(row, 1).value)
        ]
        if not rows:
            failures.append("类目区域没有数据行")
        screenshot_rows = (
            rows
            if screenshot_limit is None
            else rows[: max(0, screenshot_limit)]
        )
        image_evidence = sheet_image_evidence(path, workbook, sheet)

        for row in rows:
            levels = [
                column
                for column in (*range(1, 6), 10)
                if sheet.cell(row, column).value not in (None, "")
            ]
            if not levels:
                failures.append(f"第 {row} 行没有类目名称")
                continue
            leaf_column = max(levels)
            leaf = sheet.cell(row, leaf_column)
            target = text(leaf.hyperlink.target if leaf.hyperlink else "")
            if "/gp/bestsellers/" not in target:
                failures.append(
                    f"第 {row} 行末级类目“{text(leaf.value)}”未绑定 BSR 榜单链接"
                )
            if row not in screenshot_rows:
                continue
            if sheet.cell(row, 9).value in (None, ""):
                failures.append(f"第 {row} 行月总销量为空")
            image = image_evidence.get((row, 6))
            if image is None:
                failures.append(f"第 {row} 行类目截图缺失")
                continue
            if not image.real:
                failures.append(f"第 {row} 行类目截图为空白或无法读取")
            payload = image.payload_size
            display = image.display_size
            # 原尺寸浮动截图允许1:1显示，缩略图仍须保留高清原像素。
            original_size = bool(payload and display and all(abs(a - b) <= 1 for a, b in zip(payload, display)))
            if payload and display and payload[0] <= display[0] and not original_size:
                failures.append(
                    f"第 {row} 行类目截图未保留高清原图："
                    f"原图 {payload[0]}px，显示 {display[0]}px"
                )
            if display and not original_size and (display[0] > 420 or display[1] > 170):
                failures.append(
                    f"第 {row} 行类目截图显示框过大：{display[0]}×{display[1]}"
                )
        if failures:
            raise QualityGateError("category_research", failures)
        return GateReport(
            "category_research",
            path,
            {
                "rows": len(rows),
                "images": sum(
                    1 for row in screenshot_rows if (row, 6) in image_evidence
                ),
            },
        )
    finally:
        workbook.close()


def validate_overview_charts(path: Path) -> GateReport:
    failures: list[str] = []
    workbook = load_workbook(path, data_only=False, rich_text=True)
    try:
        sheet = find_sheet(workbook, "市场调研")
        if sheet is None:
            failures.append("缺少市场调研工作表")
            raise QualityGateError("market_overview", failures)
        expected: list[tuple[int, int, str]] = []
        for label, columns in (
            ("品牌分析", (2, 4, 6)),
            ("上架时间分析", (2, 4, 6)),
            ("Review数量分析", (2,)),
            ("Rating分析", (2,)),
            ("价格分析", (2,)),
        ):
            row = find_row_containing(sheet, label)
            if row is None:
                failures.append(f"找不到图表标题“{label}”")
                continue
            expected.extend((row, column, label) for column in columns)
        if len(expected) != 9:
            failures.append(f"数据分析概览锚点应为 9 个，实际 {len(expected)}")
        images_by_anchor = sheet_image_evidence(path, workbook, sheet)
        for row, column, label in expected:
            image = images_by_anchor.get((row, column))
            if image is None:
                failures.append(
                    f"“{label}”图表缺失，期望锚点 {get_column_letter(column)}{row}"
                )
            elif not image.real:
                failures.append(f"“{label}”图表为空白或无法读取")
        if failures:
            raise QualityGateError("market_overview", failures)
        return GateReport("market_overview", path, {"images": len(expected)})
    finally:
        workbook.close()


def validate_market_analysis(path: Path) -> GateReport:
    failures: list[str] = []
    workbook = load_workbook(path, data_only=False, rich_text=True)
    try:
        sheet = find_sheet(workbook, "市场分析")
        if sheet is None:
            failures.append("缺少市场分析工作表")
            raise QualityGateError("market_analysis", failures)
        for coordinate in ("B3", "D3", "F3", "B5", "D5", "F5", "B6", "D6", "F6"):
            if sheet[coordinate].value in (None, ""):
                failures.append(f"市场分析 {coordinate} 为空")
        expected = {(12, 1), (24, 1), (38, 1)}
        images = sheet_image_evidence(path, workbook, sheet)
        for row, column in expected:
            image = images.get((row, column))
            if image is None:
                failures.append(f"市场分析图缺失：{get_column_letter(column)}{row}")
            elif not image.real:
                failures.append(f"市场分析图空白：{get_column_letter(column)}{row}")
        if failures:
            raise QualityGateError("market_analysis", failures)
        return GateReport("market_analysis", path, {"images": 3})
    finally:
        workbook.close()


def validate_review_research(path: Path) -> GateReport:
    """验收评论抓取结果，不在此阶段检查后续翻译和痛点归类。"""
    failures: list[str] = []
    workbook = load_workbook(path, data_only=False, rich_text=True)
    try:
        summary = find_sheet(workbook, "评论抓取摘要")
        details = find_sheet(workbook, "商品评论")
        pain = find_sheet(workbook, "差评痛点")
        if summary is None or details is None or pain is None:
            failures.append("评论阶段必须包含评论抓取摘要、商品评论和差评痛点")
            raise QualityGateError("review_research", failures)
        summary_headers = {
            text(cell.value): cell.column
            for cell in summary[1]
            if text(cell.value)
        }
        detail_headers = {
            text(cell.value): cell.column
            for cell in details[1]
            if text(cell.value)
        }
        required_summary = {"竞品ASIN", "评论类型", "抓取状态", "抓取评论数"}
        required_details = {"竞品ASIN", "评论类型"}
        if missing := required_summary - summary_headers.keys():
            failures.append(f"评论摘要缺少字段：{sorted(missing)}")
        if missing := required_details - detail_headers.keys():
            failures.append(f"商品评论缺少字段：{sorted(missing)}")
        if failures:
            raise QualityGateError("review_research", failures)
        detail_counts: dict[tuple[str, str], int] = {}
        for row in details.iter_rows(min_row=2, values_only=True):
            asin = text(row[detail_headers["竞品ASIN"] - 1]).upper()
            review_type = text(row[detail_headers["评论类型"] - 1])
            if asin and review_type:
                detail_counts[(asin, review_type)] = detail_counts.get((asin, review_type), 0) + 1
        allowed = {"success", "insufficient_reviews", "no_reviews"}
        checked = 0
        for row in range(2, summary.max_row + 1):
            asin = text(summary.cell(row, summary_headers["竞品ASIN"]).value).upper()
            review_type = text(summary.cell(row, summary_headers["评论类型"]).value)
            if not asin or review_type not in {"Positive reviews", "Critical reviews"}:
                continue
            checked += 1
            status = text(summary.cell(row, summary_headers["抓取状态"]).value)
            if status not in allowed:
                failures.append(f"{asin} {review_type} 抓取状态不合格：{status or '空'}")
            expected = int(summary.cell(row, summary_headers["抓取评论数"]).value or 0)
            actual = detail_counts.get((asin, review_type), 0)
            if expected != actual:
                failures.append(
                    f"{asin} {review_type} 摘要 {expected} 条与明细 {actual} 条不一致"
                )
        if checked == 0:
            failures.append("评论摘要没有 Positive/Critical 任务")
        if errors := cell_formula_errors(workbook):
            failures.extend(f"公式错误：{item}" for item in errors[:20])
        if failures:
            raise QualityGateError("review_research", failures)
        return GateReport("review_research", path, {"review_tasks": checked})
    finally:
        workbook.close()


def rich_text_red_text(value: object) -> list[str]:
    if not isinstance(value, CellRichText):
        return []
    result = []
    for item in value:
        if not isinstance(item, TextBlock):
            continue
        color = item.font.color if item.font is not None else None
        if color is not None and color.type == "rgb" and color.rgb == "FFFF0000":
            result.append(text(item.text))
    return result


def validate_competitor_basics(
    path: Path,
    *,
    require_evaluation_blank: bool = True,
) -> GateReport:
    failures: list[str] = []
    workbook = load_workbook(path, data_only=False, rich_text=True)
    try:
        sheet = find_sheet(workbook, "竞品分析")
        if sheet is None:
            failures.append("缺少竞品分析工作表")
            raise QualityGateError("competitor_basics", failures)
        title_asins = [
            match.group(1).upper()
            for match in ASIN_PATTERN.finditer(text(sheet["A1"].value))
        ]
        if not title_asins:
            failures.append("竞品分析标题没有 ASIN")
            raise QualityGateError("competitor_basics", failures)
        link_row = find_row(sheet, "详情页链接")
        image_row = find_row(sheet, "产品图")
        rating_row = find_row(sheet, "评分/评论数")
        selling_row = find_row(sheet, "产品卖点")
        comparison_row = find_row(sheet, "对比维度")
        evaluation_row = find_row(sheet, "评价分析")
        if None in (link_row, image_row, rating_row, selling_row, comparison_row, evaluation_row):
            failures.append("竞品分析基础字段行不完整")
            raise QualityGateError("competitor_basics", failures)
        columns = list(range(3, 3 + len(title_asins)))
        images = sheet_image_evidence(path, workbook, sheet)
        red_comparison_values: list[str] = []
        for asin, column in zip(title_asins, columns):
            link_cell = sheet.cell(link_row, column)
            link_target = text(link_cell.hyperlink.target if link_cell.hyperlink else "")
            if asin not in link_target.upper():
                failures.append(f"{asin} 详情链接缺失或错列")
            image = images.get((image_row, column))
            if image is None:
                failures.append(
                    f"{asin} 产品图缺失：{get_column_letter(column)}{image_row}"
                )
            elif not image.real:
                failures.append(f"{asin} 产品图为空白")
            if sheet.cell(rating_row, column).value in (None, ""):
                failures.append(f"{asin} 评分/评论数为空")
            selling = text(sheet.cell(selling_row, column).value)
            if not selling:
                failures.append(f"{asin} 产品卖点为空")
        for row in range(comparison_row, evaluation_row):
            if all(sheet.cell(row, column).value in (None, "") for column in columns):
                continue
            for column in columns:
                red_comparison_values.extend(
                    rich_text_red_text(sheet.cell(row, column).value)
                )
        if any(
            re.search(r"\d", text(sheet.cell(row, column).value))
            for row in range(comparison_row, evaluation_row)
            for column in columns
        ) and not red_comparison_values:
            failures.append("对比维度含数值，但没有检测到红色加粗富文本")
        if require_evaluation_blank:
            for row in range(evaluation_row, evaluation_row + 3):
                for column in columns:
                    if sheet.cell(row, column).value not in (None, ""):
                        failures.append(
                            f"已要求跳过评价分析，但 {sheet.cell(row, column).coordinate} 非空"
                        )
        if errors := cell_formula_errors(workbook):
            failures.extend(f"公式错误：{item}" for item in errors[:20])
        if failures:
            raise QualityGateError("competitor_basics", failures)
        return GateReport(
            "competitor_basics",
            path,
            {
                "asins": len(title_asins),
                "images": len(title_asins),
                "red_runs": len(red_comparison_values),
            },
        )
    finally:
        workbook.close()


def validate_competitor_review_analysis(path: Path) -> GateReport:
    """验收评价分析已经回填，同时复用竞品基础字段门禁。"""
    report = validate_competitor_basics(path, require_evaluation_blank=False)
    failures: list[str] = []
    workbook = load_workbook(path, data_only=False, rich_text=True)
    try:
        sheet = find_sheet(workbook, "竞品分析")
        evaluation_row = find_row(sheet, "评价分析") if sheet is not None else None
        summary_row = find_row(sheet, "总结及改进方式") if sheet is not None else None
        if sheet is None or evaluation_row is None or summary_row is None:
            failures.append("竞品评价分析区域不完整")
        else:
            title_asins = [
                match.group(1).upper()
                for match in ASIN_PATTERN.finditer(text(sheet["A1"].value))
            ]
            columns = list(range(3, 3 + len(title_asins)))
            for column in columns:
                if all(
                    sheet.cell(row, column).value in (None, "")
                    for row in (evaluation_row, evaluation_row + 1)
                ):
                    failures.append(
                        f"{get_column_letter(column)} 列优点/缺点均未回填"
                    )
            if sheet.cell(summary_row, 3).value in (None, ""):
                failures.append("总结及改进方式为空")
        if failures:
            raise QualityGateError("competitor_review_analysis", failures)
    finally:
        workbook.close()
    return GateReport(
        "competitor_review_analysis",
        path,
        report.metrics,
    )


def validate_zip_package(path: Path) -> GateReport:
    failures: list[str] = []
    if not zipfile.is_zipfile(path):
        failures.append("输出文件不是有效 xlsx ZIP 包")
    else:
        with zipfile.ZipFile(path) as archive:
            broken = archive.testzip()
            if broken:
                failures.append(f"ZIP 包损坏：{broken}")
    if failures:
        raise QualityGateError("zip_package", failures)
    return GateReport("zip_package", path)


def validate_final_package(
    path: Path,
    *,
    require_evaluation_blank: bool = True,
    template_path: Path | None = None,
) -> list[GateReport]:
    reports = [validate_zip_package(path)]
    reports.append(validate_keyword_research(path))
    reports.append(validate_category_research(path))
    reports.append(validate_overview_charts(path))
    reports.append(validate_market_analysis(path))
    reports.append(
        validate_competitor_basics(
            path,
            require_evaluation_blank=require_evaluation_blank,
        )
    )
    if template_path is not None:
        reports.append(validate_competitor_title_format(path, template_path))
    return reports


def run_step_gate(step: str, path: Path) -> GateReport:
    if step in {"keyword_research", "keyword_backfill", "keyword_trends"}:
        return validate_keyword_research(path)
    if step in {"category_research", "product_research"}:
        return validate_category_research(path)
    if step == "market_analysis":
        overview = validate_overview_charts(path)
        market = validate_market_analysis(path)
        return GateReport(
            "market_analysis",
            path,
            {**overview.metrics, **market.metrics},
        )
    if step == "review_research":
        return validate_review_research(path)
    raise ValueError(f"未配置质量门禁：{step}")


def main() -> int:
    parser = argparse.ArgumentParser(description="执行模块输出质量门禁。")
    parser.add_argument(
        "--step",
        required=True,
        choices=(
            "keyword_research",
            "category_research",
            "market_analysis",
            "review_research",
            "competitor_basics",
            "competitor_review_analysis",
            "final",
        ),
    )
    parser.add_argument("--workbook", type=Path, required=True)
    parser.add_argument(
        "--template",
        type=Path,
        help="最终门禁使用原始模板核对竞品标题首行格式。",
    )
    parser.add_argument(
        "--allow-evaluation",
        action="store_true",
        help="竞品评价分析已执行时允许评价区域非空",
    )
    args = parser.parse_args()
    if args.step == "competitor_basics":
        reports = [
            validate_competitor_basics(
                args.workbook,
                require_evaluation_blank=not args.allow_evaluation,
            )
        ]
    elif args.step == "competitor_review_analysis":
        reports = [validate_competitor_review_analysis(args.workbook)]
    elif args.step == "final":
        reports = validate_final_package(
            args.workbook,
            require_evaluation_blank=not args.allow_evaluation,
            template_path=args.template,
        )
    else:
        reports = [run_step_gate(args.step, args.workbook)]
    for report in reports:
        print(report.as_dict())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
