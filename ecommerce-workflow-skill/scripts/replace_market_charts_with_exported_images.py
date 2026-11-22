from __future__ import annotations

import argparse
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage


def find_market_analysis_sheet(workbook):
    for name in workbook.sheetnames:
        if name.endswith("市场分析"):
            return workbook[name]
    raise RuntimeError("No sheet ending with 市场分析 found.")


def replace(workbook_path: Path, image_dir: Path, output: Path) -> dict:
    workbook = load_workbook(workbook_path)
    ws = find_market_analysis_sheet(workbook)
    ws._charts = []

    anchors = ["A12", "A24", "A38"]
    inserted = []
    for index, anchor in enumerate(anchors, start=1):
        image_path = image_dir / f"market_chart_{index}.png"
        if not image_path.exists():
            raise FileNotFoundError(f"Missing exported chart image: {image_path}")
        image = XLImage(image_path)
        image.width = 1070
        image.height = 384
        ws.add_image(image, anchor)
        inserted.append(str(image_path))

    output.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output)
    return {
        "output": str(output),
        "sheet": ws.title,
        "inserted": inserted,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Replace market analysis native charts with COM-exported chart PNGs.")
    parser.add_argument("--input", type=Path, required=True)
    parser.add_argument("--image-dir", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    result = replace(args.input, args.image_dir, args.output)
    for key, value in result.items():
        print(f"{key}: {value}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
