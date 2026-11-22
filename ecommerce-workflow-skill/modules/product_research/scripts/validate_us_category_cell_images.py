from __future__ import annotations

import argparse
import hashlib
import io
import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from PIL import Image as PILImage


MODULES_DIR = Path(__file__).resolve().parents[2]
if str(MODULES_DIR) not in sys.path:
    sys.path.insert(0, str(MODULES_DIR))

from shared.image_conversion import read_cellimage_map


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="验证美国类目节点截图工作簿。")
    parser.add_argument("--original", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    return parser.parse_args()


def sha256(payload: bytes) -> str:
    return hashlib.sha256(payload).hexdigest()


def main() -> None:
    args = parse_args()
    allowed_changed = {
        "xl/worksheets/sheet8.xml",
        "xl/cellimages.xml",
        "xl/_rels/cellimages.xml.rels",
    }
    expected_extra = {f"xl/media/codex_us_category_{index:02d}.png" for index in range(1, 20)}
    target_rows = list(range(5, 42, 2))
    expected_ids = {
        "A3": "ID_599B51B5B2CE4DD89C80320AF3A16E67",
        **{
            f"A{row}": f"ID_CODEX_US_CATEGORY_{index:04d}"
            for index, row in enumerate(target_rows, start=1)
        },
    }

    with zipfile.ZipFile(args.original, "r") as original, zipfile.ZipFile(args.output, "r") as output:
        original_names = set(original.namelist())
        output_names = set(output.namelist())
        missing_parts = sorted(original_names - output_names)
        extra_parts = output_names - original_names
        unexpected_extra = sorted(extra_parts - expected_extra)
        missing_extra = sorted(expected_extra - extra_parts)
        unexpected_changed = [
            name
            for name in sorted(original_names & output_names)
            if name not in allowed_changed and sha256(original.read(name)) != sha256(output.read(name))
        ]

        sheet_root = ET.fromstring(output.read("xl/worksheets/sheet8.xml"))
        actual_formulas = {}
        for coordinate in expected_ids:
            cell = sheet_root.find(f".//{{{MAIN_NS}}}c[@r='{coordinate}']")
            formula = cell.find(f"{{{MAIN_NS}}}f") if cell is not None else None
            actual_formulas[coordinate] = formula.text if formula is not None else None

        formula_mismatches = {
            coordinate: actual_formulas.get(coordinate)
            for coordinate, image_id in expected_ids.items()
            if image_id not in (actual_formulas.get(coordinate) or "")
        }
        error_markers = {
            marker: sum(
                output.read(name).count(marker.encode("utf-8"))
                for name in output_names
                if name.endswith(".xml")
            )
            for marker in ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A")
        }

    media, mapping = read_cellimage_map(args.output)
    unresolved = []
    bad_dimensions = []
    for coordinate, image_id in expected_ids.items():
        info = mapping.get(image_id)
        payload = media.get(info.media_path) if info else None
        if not info or not payload:
            unresolved.append({"cell": coordinate, "image_id": image_id})
            continue
        with PILImage.open(io.BytesIO(payload)) as image:
            if coordinate != "A3" and image.size != (1896, 821):
                bad_dimensions.append({"cell": coordinate, "size": image.size})

    result = {
        "output": str(args.output.resolve()),
        "size": args.output.stat().st_size,
        "formula_count": len(expected_ids) - len(formula_mismatches),
        "cellimage_count": len(mapping),
        "missing_parts": missing_parts,
        "unexpected_extra": unexpected_extra,
        "missing_extra": missing_extra,
        "unexpected_changed": unexpected_changed,
        "formula_mismatches": formula_mismatches,
        "unresolved": unresolved,
        "bad_dimensions": bad_dimensions,
        "formula_error_markers": error_markers,
    }
    result["ok"] = all(
        not result[key]
        for key in (
            "missing_parts",
            "unexpected_extra",
            "missing_extra",
            "unexpected_changed",
            "formula_mismatches",
            "unresolved",
            "bad_dimensions",
        )
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))
    if not result["ok"]:
        raise SystemExit(2)


if __name__ == "__main__":
    main()
