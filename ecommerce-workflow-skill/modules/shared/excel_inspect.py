from __future__ import annotations

import shutil
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook


SOURCE_SHEET_EXACT_NAMES = ("数据源", "US", "Sheet1")
SOURCE_SHEET_SUFFIX = "数据源"


@dataclass(frozen=True)
class ImageRisk:
    sheet_name: str
    has_dispimg: bool
    floating_image_count: int

    @property
    def needs_warning(self) -> bool:
        return self.has_dispimg and self.floating_image_count == 0


def find_sheet(wb, exact_names: tuple[str, ...] = (), contains: tuple[str, ...] = ()):
    for name in exact_names:
        if name in wb.sheetnames:
            return wb[name]
    for ws in wb.worksheets:
        if any(token in ws.title for token in contains):
            return ws
    return None


def find_source_sheet(wb):
    for ws in wb.worksheets:
        if ws.title.endswith(SOURCE_SHEET_SUFFIX) and ws.title != SOURCE_SHEET_SUFFIX:
            return ws
    return find_sheet(wb, exact_names=SOURCE_SHEET_EXACT_NAMES)


def inspect_image_risk(path: Path) -> ImageRisk:
    wb = load_workbook(path, data_only=False)
    ws = find_source_sheet(wb) or wb.worksheets[0]
    headers = [ws.cell(1, col).value for col in range(1, ws.max_column + 1)]
    image_col = None
    for index, header in enumerate(headers, start=1):
        if str(header).strip() == "图片":
            image_col = index
            break
    has_dispimg = False
    if image_col:
        for row in range(2, ws.max_row + 1):
            value = ws.cell(row, image_col).value
            if isinstance(value, str) and "DISPIMG" in value.upper():
                has_dispimg = True
                break
    return ImageRisk(ws.title, has_dispimg, len(getattr(ws, "_images", [])))


def normalize_market_source(source: Path, work_dir: Path) -> Path:
    wb = load_workbook(source)
    if find_source_sheet(wb) is not None:
        return source

    raise ValueError("Market analysis requires a source sheet named 产品名-数据源, 数据源, US, or Sheet1.")
