from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from PIL import Image as PILImage

from competitor_analysis.build_pain_slicer_view import (
    DATA_START_ROW,
    DISPLAY_HEADERS,
    HEADER_ROW,
    LIBRARY_ROW_HEIGHT,
    MISSING_MODEL_LABEL,
    MISSING_PHRASE_LABEL,
    ORANGE_FILL,
    UNCLASSIFIED_PAIN_LABEL,
    build_pain_slicer_view,
    is_meaningful_review_image,
)


class PainSlicerViewTests(unittest.TestCase):
    def test_rejects_nearly_blank_review_image(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            blank_path = Path(directory) / "blank.png"
            real_path = Path(directory) / "real.png"
            PILImage.new("RGB", (120, 100), (249, 249, 249)).save(blank_path)
            real = PILImage.new("RGB", (120, 100), "white")
            for x in range(20, 100):
                for y in range(20, 80):
                    real.putpixel((x, y), (20, 80, 160))
            real.save(real_path)
            self.assertFalse(is_meaningful_review_image(blank_path.read_bytes()))
            self.assertTrue(is_meaningful_review_image(real_path.read_bytes()))

    def test_builds_exploded_filter_rows_and_image_library(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "source.xlsx"
            output = root / "output.xlsx"
            picture_path = root / "card.png"
            PILImage.new("RGB", (80, 240), "white").save(picture_path)

            workbook = Workbook()
            competitor = workbook.active
            competitor.title = "demo-竞品分析"
            competitor["A2"] = "产品图"
            competitor["A3"] = "详情页链接"
            competitor["C3"] = "https://www.amazon.com/dp/B0ABCDEFGH"
            card = ExcelImage(picture_path)
            card.anchor = "C2"
            competitor.add_image(card)

            pain = workbook.create_sheet("差评痛点")
            headers = [
                "竞品ASIN", "实际筛选", "星级", "评论日期", "属性/型号", "差评词/短语",
                "归一化痛点（中文）", "Helpful数", "图片数量", "评论图片", "评价内容", "评论链接",
            ]
            pain.append(headers)
            pain.append([
                "B0ABCDEFGH", "All reviewers", 2, "2026-09-01", "Gray", "fade\npoor material",
                "日晒褪色\n材料质量问题", 3, 0, "", "review body", "https://example.com/review/1",
            ])
            pain.append([
                "B0ABCDEFGH", "All reviewers", 1, "2026-09-02", "/", "/",
                "/", 0, 0, "", "unclassified review", "https://example.com/review/2",
            ])
            workbook.save(source)
            workbook.close()

            result = build_pain_slicer_view(source, output)
            self.assertEqual(result["review_rows"], 3)
            self.assertEqual(result["product_cards"], 1)

            checked = load_workbook(output, data_only=False)
            view = checked["差评透视"]
            self.assertEqual([view.cell(HEADER_ROW, column).value for column in range(1, 12)], DISPLAY_HEADERS)
            self.assertEqual(view.cell(DATA_START_ROW, 6).value, "日晒褪色")
            self.assertEqual(view.cell(DATA_START_ROW + 1, 6).value, "材料质量问题")
            self.assertEqual(view.cell(DATA_START_ROW, 5).value, "fade")
            self.assertEqual(view.cell(DATA_START_ROW + 1, 5).value, "poor material")
            self.assertEqual(view.cell(DATA_START_ROW + 2, 4).value, MISSING_MODEL_LABEL)
            self.assertEqual(view.cell(DATA_START_ROW + 2, 5).value, MISSING_PHRASE_LABEL)
            self.assertEqual(view.cell(DATA_START_ROW + 2, 6).value, UNCLASSIFIED_PAIN_LABEL)
            self.assertEqual(view["A1"].fill.fgColor.rgb, ORANGE_FILL)
            self.assertEqual(view["A2"].fill.fgColor.rgb, ORANGE_FILL)
            self.assertEqual(view["A4"].fill.fgColor.rgb, ORANGE_FILL)
            self.assertEqual(view.cell(DATA_START_ROW, 12).value, f"=SUBTOTAL(103,$C{DATA_START_ROW})")
            self.assertNotIn("差评痛点", checked.sheetnames)
            library = checked["差评透视图片库"]
            self.assertEqual(library["A2"].value, "B0ABCDEFGH")
            self.assertEqual(library.row_dimensions[2].height, LIBRARY_ROW_HEIGHT)
            self.assertEqual(len(library._images), 1)
            self.assertEqual(library._images[0].anchor._from.row + 1, 3)
            checked.close()


if __name__ == "__main__":
    unittest.main()
