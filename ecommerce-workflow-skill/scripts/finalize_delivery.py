from __future__ import annotations

import argparse
import sys
from pathlib import Path


MODULES_DIR = Path(__file__).resolve().parents[1] / "modules"
if str(MODULES_DIR) not in sys.path:
    sys.path.insert(0, str(MODULES_DIR))

from shared.image_conversion import make_wps_embedded_version
from shared.quality_gates import validate_final_package


def main() -> int:
    parser = argparse.ArgumentParser(
        description="生成最终 WPS 工作簿并强制执行完整交付门禁。"
    )
    parser.add_argument("--original", type=Path, required=True)
    parser.add_argument("--floating-workbook", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--temp-output", type=Path)
    parser.add_argument(
        "--embed-extra-sheet",
        action="append",
        default=[],
        help="需要转换为 WPS 单元格图片的工作表，可重复传入。",
    )
    parser.add_argument(
        "--allow-evaluation",
        action="store_true",
        help="竞品评价分析已执行时允许评价区域非空。",
    )
    args = parser.parse_args()

    temp_output = args.temp_output or args.output.with_name(
        args.output.stem + "_temp.xlsx"
    )
    make_wps_embedded_version(
        original=args.original,
        floating_workbook=args.floating_workbook,
        output=args.output,
        temp_output=temp_output,
        embed_extra_sheets=set(args.embed_extra_sheet),
    )
    reports = validate_final_package(
        args.output,
        require_evaluation_blank=not args.allow_evaluation,
        template_path=args.original,
    )
    for report in reports:
        print(report.as_dict())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
