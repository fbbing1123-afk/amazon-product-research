from __future__ import annotations

import argparse
import math
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import Image, ImageDraw, ImageFont


def load_font(size: int, bold: bool = False):
    names = ["msyhbd.ttc", "msyh.ttc", "simhei.ttf", "arial.ttf"] if bold else ["msyh.ttc", "simhei.ttf", "arial.ttf"]
    for name in names:
        try:
            return ImageFont.truetype(name, size)
        except OSError:
            continue
    return ImageFont.load_default()


def find_market_analysis_sheet(workbook):
    for name in workbook.sheetnames:
        if name.endswith("市场分析"):
            return workbook[name]
    raise RuntimeError("No sheet ending with 市场分析 found.")


def read_chart_rows(ws, header_row: int, first_data_row: int, last_data_row: int):
    rows = []
    for row in range(first_data_row, last_data_row + 1):
        label = ws.cell(row, 1).value
        value = ws.cell(row, 3).value
        count = ws.cell(row, 2).value
        if label in (None, ""):
            continue
        try:
            numeric = float(value or 0)
        except (TypeError, ValueError):
            numeric = 0
        try:
            sample_count = int(float(count or 0))
        except (TypeError, ValueError):
            sample_count = 0
        rows.append((str(label), numeric, sample_count))
    title = str(ws.cell(header_row, 1).value or "")
    return title, rows


def compact_number(value: float) -> str:
    if abs(value) >= 10000:
        return f"{value / 10000:.1f}万"
    return f"{value:,.0f}"


def draw_bar_chart(title: str, rows: list[tuple[str, float, int]], output: Path, width: int = 1450, height: int = 520) -> None:
    image = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(image)
    label_font = load_font(22)
    value_font = load_font(21, bold=True)

    margin_left = 90
    margin_right = 50
    margin_top = 38
    margin_bottom = 110
    plot_w = width - margin_left - margin_right
    plot_h = height - margin_top - margin_bottom
    baseline = margin_top + plot_h

    if title:
        title_font = load_font(32, bold=True)
        draw.text((margin_left, 22), title, fill="#17324D", font=title_font)
    draw.line((margin_left, baseline, width - margin_right, baseline), fill="#CBD5E1", width=2)

    max_value = max([value for _, value, _ in rows] + [1])
    nice_max = math.ceil(max_value / 1000) * 1000 if max_value > 1000 else max_value
    grid_count = 4
    for i in range(1, grid_count + 1):
        y = baseline - plot_h * i / grid_count
        draw.line((margin_left, y, width - margin_right, y), fill="#E5E7EB", width=1)

    count = max(1, len(rows))
    slot = plot_w / count
    bar_w = min(110, slot * 0.58)
    color = "#4472C4"
    for index, (label, value, sample_count) in enumerate(rows):
        center = margin_left + slot * index + slot / 2
        bar_h = 0 if nice_max == 0 else plot_h * value / nice_max
        x1 = center - bar_w / 2
        x2 = center + bar_w / 2
        y1 = baseline - bar_h
        draw.rounded_rectangle((x1, y1, x2, baseline), radius=8, fill=color)
        value_text = f"{compact_number(value)} · {sample_count}样本"
        vb = draw.textbbox((0, 0), value_text, font=value_font)
        draw.text((center - (vb[2] - vb[0]) / 2, y1 - 30), value_text, fill="#17324D", font=value_font)

        display_label = label if len(label) <= 11 else label[:10] + "…"
        lb = draw.textbbox((0, 0), display_label, font=label_font)
        draw.text((center - (lb[2] - lb[0]) / 2, baseline + 18), display_label, fill="#475569", font=label_font)

    output.parent.mkdir(parents=True, exist_ok=True)
    image.save(output)


def replace_charts(workbook_path: Path, output: Path, image_dir: Path) -> dict:
    workbook = load_workbook(workbook_path, rich_text=True)
    ws = find_market_analysis_sheet(workbook)
    ws._charts = []

    specs = [
        ("sales_layers", 13, 15, 18, "A12"),
        ("review_buckets", 25, 27, 32, "A24"),
        ("price_buckets", 39, 41, 46, "A38"),
    ]
    target_anchors = {anchor for *_rest, anchor in specs}
    kept_images = []
    removed_images = 0
    for image in list(getattr(ws, "_images", [])):
        marker = getattr(getattr(image, "anchor", None), "_from", None)
        if marker is None:
            kept_images.append(image)
            continue
        anchor = ws.cell(marker.row + 1, marker.col + 1).coordinate
        if anchor in target_anchors:
            removed_images += 1
        else:
            kept_images.append(image)
    ws._images = kept_images
    generated = []
    for name, title_row, first_row, last_row, anchor in specs:
        _, rows = read_chart_rows(ws, title_row, first_row, last_row)
        image_path = image_dir / f"{name}.png"
        draw_bar_chart("", rows, image_path)
        img = XLImage(image_path)
        img.width = 1070
        img.height = 384
        ws.add_image(img, anchor)
        generated.append(str(image_path))

    output.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output)
    return {
        "output": str(output),
        "sheet": ws.title,
        "images": generated,
        "removed_previous_chart_images": removed_images,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Replace native market analysis charts with rendered PNG images.")
    parser.add_argument("--input", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--image-dir", type=Path, required=True)
    args = parser.parse_args()
    result = replace_charts(args.input, args.output, args.image_dir)
    for key, value in result.items():
        print(f"{key}: {value}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
