from __future__ import annotations

import argparse
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.drawing.image import Image


DISPLAY_IMAGES = {
    "B10": (210, 281),
    "D10": (190, 281),
    "F10": (190, 281),
    "B11": (210, 253),
    "D11": (190, 253),
    "F11": (190, 253),
    "B12": (501, 253),
    "B13": (501, 253),
    "B14": (501, 253),
}

CELL_OFFSETS = {
    "B10": ("B", 1),
    "D10": ("D", 1),
    "F10": ("F", 1),
    "B11": ("B", 2),
    "D11": ("D", 2),
    "F11": ("F", 2),
    "B12": ("B", 3),
    "B13": ("B", 4),
    "B14": ("B", 5),
}

LABEL_POSITIONS = {
    "B10": ("品牌分析", "B"),
    "D10": ("品牌分析", "D"),
    "F10": ("品牌分析", "F"),
    "B11": ("上架时间分析", "B"),
    "D11": ("上架时间分析", "D"),
    "F11": ("上架时间分析", "F"),
    "B12": ("Review数量分析", "B"),
    "B13": ("Rating分析", "B"),
    "B14": ("价格分析", "B"),
}

SECTION_MARKER = "卖家精灵工具抓取主关键词前5页"


def find_research_sheet(workbook, requested: str | None):
    if requested:
        return workbook[requested]
    for ws in workbook.worksheets:
        if "市场调研" in ws.title:
            return ws
    return workbook.worksheets[0]


def find_section_row(ws) -> int:
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and SECTION_MARKER in str(cell.value):
                return cell.row
    raise ValueError(f"Cannot find section marker: {SECTION_MARKER}")


def find_label_row(ws, label: str) -> int | None:
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and label in str(cell.value):
                return cell.row
    return None


def chart_anchors(ws) -> dict[str, str]:
    anchors = {}
    for logical_cell, (label, col) in LABEL_POSITIONS.items():
        row = find_label_row(ws, label)
        if row is not None:
            anchors[logical_cell] = f"{col}{row}"

    if len(anchors) == len(LABEL_POSITIONS):
        return anchors

    section_row = find_section_row(ws)
    for logical_cell, (col, row_offset) in CELL_OFFSETS.items():
        anchors.setdefault(logical_cell, f"{col}{section_row + row_offset}")
    return anchors


def clear_existing_chart_images(ws) -> None:
    chart_anchor_cells = set(chart_anchors(ws).values())
    kept = []
    for image in getattr(ws, "_images", []):
        name = str(getattr(image, "path", "") or getattr(image, "_id", ""))
        anchor_cell = None
        try:
            anchor_cell = f"{chr(image.anchor._from.col + 65)}{image.anchor._from.row + 1}"
        except Exception:
            pass
        if "hires_section3" not in name and anchor_cell not in chart_anchor_cells:
            kept.append(image)
    ws._images = kept


def insert_images(workbook_path: Path, image_dir: Path, output_path: Path, sheet_name: str | None) -> None:
    wb = load_workbook(workbook_path, rich_text=True)
    ws = find_research_sheet(wb, sheet_name)
    anchors = chart_anchors(ws)
    clear_existing_chart_images(ws)

    for logical_cell, anchor_cell in anchors.items():
        image_path = image_dir / f"{logical_cell}_hires.png"
        if not image_path.exists():
            raise FileNotFoundError(image_path)
        img = Image(str(image_path))
        img.width, img.height = DISPLAY_IMAGES[logical_cell]
        ws.add_image(img, anchor_cell)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", type=Path, required=True)
    parser.add_argument("--image-dir", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--sheet")
    args = parser.parse_args()

    insert_images(args.input, args.image_dir, args.output, args.sheet)
    print(args.output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
