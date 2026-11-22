from __future__ import annotations

import tempfile
import unittest
from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Alignment, Border, PatternFill, Side
from PIL import Image as PillowImage

from competitor_analysis.populate_basics import (
    comparison_rich_text,
    excel_column_width_to_pixels,
    fit_image_to_cell,
    populate,
    prepare_competitor_layout,
    row_height_to_pixels,
)


class PopulateCompetitorBasicsTests(unittest.TestCase):
    def test_comparison_values_are_red_without_highlighting_explanations(self):
        value = comparison_rich_text("219Wh/60000mAh；最长约 10 小时；支持直通供电")
        self.assertIsInstance(value, CellRichText)
        highlighted = [
            item.text
            for item in value
            if isinstance(item, TextBlock)
            and item.font.color is not None
            and item.font.color.rgb == "FFFF0000"
        ]
        self.assertEqual(highlighted, ["219Wh", "60000mAh", "10 小时"])

    def test_layout_uniformizes_added_columns_and_extends_merges(self):
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.merge_cells("A1:G1")
        worksheet["A1"] = CellRichText(
            [
                TextBlock(InlineFont(rFont="Arial", b=True, color="FFFF0000"), "竞品分析"),
                "-US",
            ]
        )
        worksheet["A1"].fill = PatternFill("solid", fgColor="FFD9EAD3")
        worksheet["A1"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        thin = Side(style="thin", color="FF000000")
        worksheet["A1"].border = Border(left=thin, right=thin, top=thin, bottom=thin)
        worksheet.merge_cells("A11:B11")
        worksheet.merge_cells("C11:G11")
        worksheet.column_dimensions["C"].width = 30.625
        worksheet.column_dimensions["H"].width = 9
        worksheet.row_dimensions[2].height = 377
        worksheet.row_dimensions[11].height = 198

        prepare_competitor_layout(
            worksheet,
            list(range(3, 12)),
            summary_row=11,
        )

        self.assertTrue(
            all(
                worksheet.column_dimensions[chr(64 + column)].width == 30.625
                for column in range(3, 12)
            )
        )
        merges = {str(item) for item in worksheet.merged_cells.ranges}
        self.assertIn("A1:K1", merges)
        self.assertIn("A11:B11", merges)
        self.assertIn("C11:K11", merges)
        self.assertEqual(worksheet.row_dimensions[2].height, 377)
        self.assertEqual(worksheet.row_dimensions[11].height, 198)
        self.assertIsInstance(worksheet["A1"].value, CellRichText)
        self.assertEqual(worksheet["A1"].fill.fgColor.rgb, "FFD9EAD3")
        self.assertEqual(worksheet["A1"].alignment.horizontal, "left")
        self.assertEqual(worksheet["A1"].border.top.style, "thin")
        workbook.close()

    def test_populate_uses_source_main_images_and_preserves_dimensions(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            workbook_path = root / "input.xlsx"
            content_path = root / "content.json"
            output_path = root / "output.xlsx"
            cards_dir = root / "cards"
            cards_dir.mkdir()
            asins = [f"B0AAAA000{index}" for index in range(1, 7)]

            workbook = Workbook()
            competitor = workbook.active
            competitor.title = "US-竞品分析"
            competitor["A1"] = "竞品分析-US：" + "/".join(asins)
            competitor["A2"] = "产品图"
            competitor["A3"] = "详情页链接"
            competitor["A4"] = "评分/评论数"
            competitor["A5"] = "产品卖点"
            competitor["A6"] = "对比维度"
            competitor["A9"] = "评价分析"
            competitor["A11"] = "总结及改进方式"
            competitor.merge_cells("A1:G1")
            competitor.merge_cells("A11:B11")
            competitor.merge_cells("C11:G11")
            for column in range(3, 8):
                competitor.column_dimensions[chr(64 + column)].width = 30.625
            for column in range(8, 17):
                competitor.column_dimensions[chr(64 + column)].width = 9
            expected_heights = {
                1: 18,
                2: 377,
                3: 54,
                4: 52,
                5: 261,
                6: 127,
                7: 124,
                8: 108,
                9: 267,
                10: 233,
                11: 198,
            }
            for row, height in expected_heights.items():
                competitor.row_dimensions[row].height = height

            datasource = workbook.create_sheet("US-数据源")
            datasource.append(["图片", "ASIN", "商品详情页链接", "评分数", "评分"])
            image_file = root / "main.png"
            PillowImage.new("RGB", (200, 200), "blue").save(image_file)
            for index, asin in enumerate(asins, start=2):
                datasource.cell(index, 2, asin)
                datasource.cell(index, 3, f"https://www.amazon.com/dp/{asin}")
                datasource.cell(index, 4, 100)
                datasource.cell(index, 5, 4.5)
                datasource.add_image(ExcelImage(image_file), f"A{index}")

            product_payload = {
                "selling_points": ["200Wh 电池", "最长 10 小时续航", "USB-C 100W"],
                "comparisons": [],
            }
            content_path.write_text(
                __import__("json").dumps(
                    {
                        "products": {asin: product_payload for asin in asins},
                        "comparisons": [
                            {
                                "label": "容量与续航",
                                "values": {asin: "200Wh；最长约 10 小时" for asin in asins},
                            }
                        ],
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )
            workbook.save(workbook_path)
            workbook.close()

            populate(workbook_path, content_path, cards_dir, output_path)

            checked = load_workbook(output_path, rich_text=True)
            try:
                worksheet = checked["US-竞品分析"]
                self.assertTrue(
                    all(
                        worksheet.column_dimensions[chr(64 + column)].width == 30.625
                        for column in range(3, 9)
                    )
                )
                self.assertEqual(
                    {row: worksheet.row_dimensions[row].height for row in expected_heights},
                    expected_heights,
                )
                merges = {str(item) for item in worksheet.merged_cells.ranges}
                self.assertIn("A1:H1", merges)
                self.assertIn("C11:H11", merges)
                self.assertEqual(len(worksheet._images), 6)
                self.assertEqual(
                    {(image.width, image.height) for image in worksheet._images},
                    {(200, 200)},
                )
                self.assertIsInstance(worksheet["C6"].value, CellRichText)
                self.assertIn(
                    "FFFF0000",
                    {
                        item.font.color.rgb
                        for item in worksheet["C6"].value
                        if isinstance(item, TextBlock) and item.font.color
                    },
                )
            finally:
                checked.close()

    def test_product_card_display_fits_cell_without_resampling_media(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            image_path = root / "card.png"
            workbook_path = root / "output.xlsx"
            source = PillowImage.new("RGB", (249, 1266), "white")
            source.save(image_path)
            source.close()

            workbook = Workbook()
            worksheet = workbook.active
            worksheet.row_dimensions[2].height = 377
            worksheet.column_dimensions["C"].width = 30.625
            card = ExcelImage(image_path)
            fit_image_to_cell(card, worksheet, 2, 3)
            worksheet.add_image(card, "C2")

            available_width = excel_column_width_to_pixels(30.625) - 12
            available_height = row_height_to_pixels(377) - 12
            self.assertLessEqual(card.width, available_width)
            self.assertLessEqual(card.height, available_height)
            self.assertAlmostEqual(card.width / card.height, 249 / 1266, places=2)
            workbook.save(workbook_path)
            workbook.close()

            namespaces = {
                "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
            }
            with ZipFile(workbook_path) as archive:
                drawing = ET.fromstring(archive.read("xl/drawings/drawing1.xml"))
                extent = drawing.find("xdr:oneCellAnchor/xdr:ext", namespaces)
                self.assertIsNotNone(extent)
                self.assertLessEqual(int(extent.attrib["cx"]) / 9525, available_width)
                self.assertLessEqual(int(extent.attrib["cy"]) / 9525, available_height)

                media_name = next(
                    name for name in archive.namelist()
                    if name.startswith("xl/media/") and name.endswith(".png")
                )
                with PillowImage.open(BytesIO(archive.read(media_name))) as embedded:
                    self.assertEqual(embedded.size, (249, 1266))


if __name__ == "__main__":
    unittest.main()
