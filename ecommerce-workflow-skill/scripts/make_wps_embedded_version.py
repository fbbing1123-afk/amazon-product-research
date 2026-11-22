from __future__ import annotations

import argparse
import sys
from pathlib import Path

MODULES_DIR = Path(__file__).resolve().parents[1] / "modules"
if str(MODULES_DIR) not in sys.path:
    sys.path.insert(0, str(MODULES_DIR))

from shared.image_conversion import make_wps_embedded_version


def main() -> int:
    parser = argparse.ArgumentParser(description="Create a WPS embedded-cell-image version from a floating-image workbook.")
    parser.add_argument("--original", type=Path, required=True, help="Original workbook that contains WPS DISPIMG data.")
    parser.add_argument("--floating-workbook", type=Path, required=True, help="Processed workbook with floating images.")
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--temp-output", type=Path)
    parser.add_argument(
        "--embed-extra-sheet",
        action="append",
        default=[],
        help="Also convert remaining floating images on this sheet into WPS DISPIMG cell images. Can be repeated.",
    )
    args = parser.parse_args()

    temp_output = args.temp_output or args.output.with_name(args.output.stem + "_temp.xlsx")
    result = make_wps_embedded_version(
        args.original,
        args.floating_workbook,
        args.output,
        temp_output,
        set(args.embed_extra_sheet),
    )
    for key, value in result.items():
        print(f"{key}: {value}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
