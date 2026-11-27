from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

MODULES_DIR = Path(__file__).resolve().parents[1] / "modules"
if str(MODULES_DIR) not in sys.path:
    sys.path.insert(0, str(MODULES_DIR))

from shared.image_conversion import make_wps_embedded_version
from shared.cell_image_embedding import embed_final_images


def main() -> int:
    parser = argparse.ArgumentParser(description="最终将浮动图片嵌入单元格，保留行高、列宽及合并区域。")
    parser.add_argument("--original", type=Path, help="仅历史原图恢复模式使用；当前最终嵌入无需此参数")
    parser.add_argument("--floating-workbook", type=Path, required=True, help="全部填数结束后的浮动图片检查点")
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--temp-output", type=Path)
    parser.add_argument('--manifest',type=Path,help='保存本批嵌入位置、原媒体和指纹清单')
    parser.add_argument(
        "--embed-extra-sheet",
        action="append",
        default=[],
        help="Also convert remaining floating images on this sheet into WPS DISPIMG cell images. Can be repeated.",
    )
    args = parser.parse_args()
    if args.manifest and args.manifest.resolve() in {args.floating_workbook.resolve(),args.output.resolve()}:
        parser.error('嵌入清单不能覆盖输入或输出工作簿')

    if args.original is None:
        # 当前流程直接对最终检查点做ZIP定点嵌入，不再整体重存工作簿。
        if args.embed_extra_sheet or args.temp_output:
            parser.error('--embed-extra-sheet和--temp-output仅供带--original的历史恢复模式使用')
        result = embed_final_images(args.floating_workbook,args.output)
        if args.manifest:
            args.manifest.parent.mkdir(parents=True,exist_ok=True)
            args.manifest.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
        print(json.dumps({k:v for k,v in result.items() if k != 'images'},ensure_ascii=False,indent=2))
        return 0

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
