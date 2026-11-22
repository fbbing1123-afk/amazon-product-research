from __future__ import annotations

import argparse
import sys
from pathlib import Path

MODULES_DIR = Path(__file__).resolve().parents[1] / "modules"
if str(MODULES_DIR) not in sys.path:
    sys.path.insert(0, str(MODULES_DIR))

from shared.image_conversion import convert_wps_dispimg_to_floating


def main() -> int:
    parser = argparse.ArgumentParser(description="Convert WPS DISPIMG cell images to standard floating images.")
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--sheet", default=None, help="Sheet to scan. Omit to scan all sheets.")
    parser.add_argument("--image-column", default="图片")
    parser.add_argument("--width", type=int, default=200)
    parser.add_argument("--height", type=int, default=200)
    args = parser.parse_args()

    result = convert_wps_dispimg_to_floating(args.source, args.output, args.sheet, args.width, args.height)
    for key, value in result.items():
        print(f"{key}: {value}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
