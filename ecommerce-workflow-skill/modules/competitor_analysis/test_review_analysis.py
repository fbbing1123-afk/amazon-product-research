from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Border
from PIL import Image as PillowImage

from competitor_analysis.backfill_review_analysis import (
    advantage_text,
    backfill_review_analysis,
    comparison_rich_text,
    competitor_layout_last_column,
    extend_competitor_layout,
)
from competitor_analysis.prepare_review_analysis import prepare_review_analysis


class CompetitorReviewAnalysisTests(unittest.TestCase):
    def test_highlights_comparable_values_without_highlighting_list_numbers(self):
        value = comparison_rich_text(
            "1. 80/250英尺双档\n"
            "2. 59g轻量机身\n"
            "3. 35分钟充电发射25-40次\n"
            "4. 3秒延时启动\n"
            "5. 配8张贴纸"
        )

        self.assertIsInstance(value, CellRichText)
        highlighted = [
            item for item in value
            if isinstance(item, TextBlock)
            and item.font.color is not None
            and item.font.color.rgb == "FFFF0000"
        ]
        self.assertEqual(
            [item.text.strip() for item in highlighted],
            ["80/250英尺", "59g", "35分钟", "25-40次", "3秒", "8张"],
        )
        self.assertTrue(all(item.font.b for item in highlighted))
        self.assertTrue(all(item.font.color.rgb == "FFFF0000" for item in highlighted))
        self.assertNotIn("1.", [item.text for item in highlighted])

    def test_highlights_values_in_comparison_dimensions(self):
        value = comparison_rich_text("80英尺/250英尺；35分钟25-40次；3秒延时")

        self.assertEqual(str(value), "80英尺/250英尺\n35分钟25-40次\n3秒延时")
        highlighted = [
            item for item in value
            if isinstance(item, TextBlock)
            and item.font.color is not None
            and item.font.color.rgb == "FFFF0000"
        ]
        self.assertEqual(
            [item.text.strip() for item in highlighted],
            ["80英尺", "250英尺", "35分钟", "25-40次", "3秒"],
        )

    def make_workbook(self, path: Path) -> None:
        workbook = Workbook()
        competitor = workbook.active
        competitor.title = "demo-竞品分析"
        competitor["A1"] = "竞品分析：B0AAAA0001/B0BBBB0002/B0CCCC0003"
        competitor["A2"] = "产品图"
        competitor["A3"] = "详情页链接"
        competitor["C3"] = "https://www.amazon.com/dp/B0AAAA0001"
        competitor["D3"] = "https://www.amazon.com/dp/B0BBBB0002"
        competitor["E3"] = "https://www.amazon.com/dp/B0CCCC0003"
        competitor["C3"].hyperlink = competitor["C3"].value
        competitor["D3"].hyperlink = competitor["D3"].value
        competitor["E3"].hyperlink = competitor["E3"].value
        competitor["A4"] = "评分/评论数"
        competitor["C4"] = "4.5 / 100"
        competitor["D4"] = "4.2 / 50"
        competitor["E4"] = "4.0 / 10"
        competitor["A5"] = "产品卖点"
        competitor["C5"] = "600D材质\n操作简单"
        competitor["D5"] = "配件丰富"
        competitor["E5"] = "价格实惠"
        competitor["A6"] = "对比维度"
        competitor["B6"] = "成像清晰度"
        competitor["C6"] = "600D牛津布；抗撕裂"
        competitor["D6"] = "一般"
        competitor["E6"] = "一般"
        competitor["A9"] = "评价分析"
        competitor["B9"] = "优点"
        competitor["B10"] = "缺点"
        competitor["A11"] = "总结及改进方式"
        competitor.merge_cells("A11:B11")
        competitor.merge_cells("C11:E11")
        competitor.row_dimensions[11].height = 198
        product_image_path = path.with_suffix(".png")
        PillowImage.new("RGB", (40, 100), "blue").save(product_image_path)
        for cell in ("C2", "D2", "E2"):
            product_image = ExcelImage(product_image_path)
            product_image.width = 20
            product_image.height = 50
            competitor.add_image(product_image, cell)

        summary = workbook.create_sheet("评论抓取摘要")
        summary.append(
            [
                "抓取时间", "站点", "竞品ASIN", "评论类型", "实际筛选",
                "抓取状态", "抓取评论数",
                "Verified Purchase数量", "绿标评论数量", "含图片评论数",
                "含视频评论数", "商品链接", "错误/说明",
            ]
        )
        summary.append([None, "US", "B0AAAA0001", "Positive reviews", "All reviewers", "success", 25, 25, 0, 0, 0, "https://www.amazon.com/dp/B0AAAA0001", None])
        summary.append([None, "US", "B0AAAA0001", "Critical reviews", "All reviewers", "success", 2, 2, 0, 0, 0, "https://www.amazon.com/dp/B0AAAA0001", None])
        summary.append([None, "US", "B0BBBB0002", "Positive reviews", "All reviewers", "success", 1, 0, 0, 0, 0, "https://www.amazon.com/dp/B0BBBB0002", None])
        summary.append([None, "US", "B0BBBB0002", "Critical reviews", "All reviewers", "success", 2, 0, 0, 0, 0, "https://www.amazon.com/dp/B0BBBB0002", None])
        summary.append([None, "US", "B0CCCC0003", "Positive reviews", "All reviewers", "no_reviews", 0, 0, 0, 0, 0, "https://www.amazon.com/dp/B0CCCC0003", None])
        summary.append([None, "US", "B0CCCC0003", "Critical reviews", "All reviewers", "no_reviews", 0, 0, 0, 0, 0, "https://www.amazon.com/dp/B0CCCC0003", None])

        details = workbook.create_sheet("商品评论")
        details.append(
            [
                "站点", "竞品ASIN", "实际筛选", "评论日期", "星级",
                "评论类型", "评论标题", "评论正文",
                "购买变体", "Verified Purchase", "绿标评论", "Helpful数", "图片URL",
                "含视频", "评论链接",
            ]
        )
        for index in range(25):
            details.append(["US", "B0AAAA0001", "All reviewers", None, 5, "Positive reviews", f"Good {index}", f"Body {index}", "A", "是", "否", 0, None, "否", f"https://review/a{index}"])
        for index in range(2):
            details.append(["US", "B0AAAA0001", "All reviewers", None, 1, "Critical reviews", f"Bad {index}", f"Bad body {index}", "A", "是", "否", 0, None, "否", f"https://review/ac{index}"])
        details.append(["US", "B0BBBB0002", "All reviewers", None, 5, "Positive reviews", "Good B", "Body B", "B", "否", "否", 0, None, "否", "https://review/b"])
        for index in range(2):
            details.append(["US", "B0BBBB0002", "All reviewers", None, 1, "Critical reviews", f"Bad B{index}", f"Bad B body {index}", "B", "否", "否", 0, None, "否", f"https://review/bc{index}"])

        pain = workbook.create_sheet("差评痛点")
        pain.append(["竞品ASIN", "实际筛选", "星级", "评论日期", "属性/型号", "差评词/短语", "归一化痛点（中文）", "Helpful数", "图片数量", "评论图片", "评价内容", "评论链接"])
        pain.append(["B0AAAA0001", "All reviewers", 1, None, "A", None, "产品失效\n连接不稳"])
        pain.append(["B0AAAA0001", "All reviewers", 1, None, "A", None, "产品失效"])
        pain.append(["B0BBBB0002", "All reviewers", 1, None, "B", None, "尺寸不合"])
        pain.append(["B0BBBB0002", "All reviewers", 1, None, "B", None, "产品失效\n尺寸不合"])
        workbook.save(path)
        workbook.close()

    def test_prepares_stable_samples_and_per_asin_disadvantages(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            workbook_path = Path(temp_dir) / "input.xlsx"
            analysis_path = Path(temp_dir) / "analysis.json"
            self.make_workbook(workbook_path)
            first = prepare_review_analysis(workbook_path, analysis_path, sample_size=20)
            second = prepare_review_analysis(workbook_path, analysis_path, sample_size=20)
            default_sample = prepare_review_analysis(workbook_path, analysis_path)

            product_a, product_b, product_c = first["products"]
            self.assertEqual(first["schema_version"], "competitor-analysis.v1")
            self.assertEqual(len(first["source_digest"]), 64)
            self.assertEqual(len(first["machine_digest"]), 64)
            self.assertEqual(len(product_a["positive_sample"]), 20)
            self.assertEqual(product_a["positive_sample"], second["products"][0]["positive_sample"])
            self.assertEqual(default_sample["sample_size"], 25)
            self.assertEqual(len(default_sample["products"][0]["positive_sample"]), 25)
            self.assertEqual(len(product_b["positive_sample"]), 1)
            self.assertEqual(product_a["rating_and_review_count"], "4.5 / 100")
            self.assertEqual(product_a["selling_points"], "600D材质\n操作简单")
            self.assertEqual(
                product_a["comparison_dimensions"],
                [{"dimension": "成像清晰度", "value": "600D牛津布；抗撕裂"}],
            )
            self.assertEqual(
                [item["display"] for item in product_a["disadvantages"]],
                ["产品失效（100%）", "连接不稳（50%）"],
            )
            self.assertEqual(
                [item["display"] for item in product_b["disadvantages"]],
                ["尺寸不合（100%）", "产品失效（50%）"],
            )
            self.assertTrue(product_c["has_review_data"])
            self.assertEqual(product_c["positive_review_count"], 0)
            self.assertEqual(product_c["disadvantages"], [])

    def test_backfills_unnumbered_review_analysis_and_preserves_following_row(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            workbook_path = Path(temp_dir) / "input.xlsx"
            analysis_path = Path(temp_dir) / "analysis.json"
            output_path = Path(temp_dir) / "output.xlsx"
            self.make_workbook(workbook_path)
            report = prepare_review_analysis(workbook_path, analysis_path, sample_size=20)
            report["products"][0]["advantages"] = [
                {"text": "1. 容易使用", "evidence_review_urls": ["https://review/a1"]},
                {"text": "2、清洁效果好", "evidence_review_urls": ["https://review/a2"]},
            ]
            report["products"][1]["advantages"] = ["1) 操作方便"]
            report["overall_summary"] = "此类产品整体操作简单，但连接稳定性仍有提升空间。"
            report["improvements"] = [
                {"pain_point": "1. 产品失效", "solution": "加强耐久测试与出厂老化检测"},
                {"pain_point": "2、连接不稳", "solution": "优化配网流程和系统兼容性"},
            ]
            analysis_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

            result = backfill_review_analysis(analysis_path, workbook_path, output_path)
            self.assertEqual(
                result,
                {"advantages_cells": 3, "disadvantages_cells": 3, "summary_cells": 1},
            )
            checked = load_workbook(output_path, data_only=True, rich_text=True)
            try:
                worksheet = checked["demo-竞品分析"]
                self.assertEqual(worksheet["C9"].value, "容易使用\n清洁效果好")
                self.assertEqual(worksheet["C10"].value, "产品失效（100%）\n连接不稳（50%）")
                self.assertEqual(worksheet["D10"].value, "尺寸不合（100%）\n产品失效（50%）")
                self.assertEqual(worksheet["E9"].value, "/")
                self.assertEqual(worksheet["E10"].value, "/")
                self.assertEqual(worksheet["A11"].value, "总结及改进方式")
                self.assertEqual(
                    worksheet["C11"].value,
                    "总结：此类产品整体操作简单，但连接稳定性仍有提升空间。\n\n"
                    "改进方式：\n"
                    "产品失效 -> 加强耐久测试与出厂老化检测\n"
                    "连接不稳 -> 优化配网流程和系统兼容性",
                )
                self.assertEqual(worksheet["C9"].border.left.style, "thin")
                self.assertEqual(worksheet["C11"].border.left.style, "thin")
                self.assertIn("A1:G1", {str(item) for item in worksheet.merged_cells.ranges})
                self.assertIn("C11:G11", {str(item) for item in worksheet.merged_cells.ranges})
                self.assertIsNone(worksheet["F9"].value)
                self.assertIsNone(worksheet["G10"].value)
                self.assertEqual(worksheet["F9"].border.left.style, "thin")
                self.assertEqual(worksheet["G10"].border.right.style, "thin")
                self.assertIsInstance(worksheet["C6"].value, CellRichText)
                self.assertEqual(str(worksheet["C6"].value), "600D牛津布\n抗撕裂")
                highlighted = [
                    item for item in worksheet["C6"].value
                    if isinstance(item, TextBlock)
                ]
                self.assertEqual([item.text for item in highlighted], ["600D"])
                self.assertTrue(highlighted[0].font.b)
                self.assertEqual(highlighted[0].font.color.rgb, "FFFF0000")
                self.assertIsInstance(worksheet["C5"].value, CellRichText)
                selling_point_highlights = [
                    item for item in worksheet["C5"].value
                    if isinstance(item, TextBlock)
                ]
                self.assertEqual([item.text for item in selling_point_highlights], ["600D"])
                self.assertEqual(worksheet["A11"].font.name, "微软雅黑")
                self.assertTrue(worksheet["A11"].font.bold)
                self.assertEqual(worksheet["A11"].alignment.horizontal, "center")
                self.assertEqual(worksheet["C11"].font.name, "微软雅黑")
                self.assertTrue(worksheet["C11"].font.bold)
                self.assertEqual(worksheet["C11"].font.color.rgb, "FFFF0000")
                self.assertEqual(worksheet["C11"].alignment.vertical, "top")
                self.assertEqual(worksheet.row_dimensions[11].height, 198)
                self.assertEqual(len(worksheet._images), 3)
                for image in worksheet._images:
                    self.assertAlmostEqual(image.anchor.ext.cx / 9525, 40, delta=1)
                    self.assertAlmostEqual(image.anchor.ext.cy / 9525, 100, delta=1)
            finally:
                checked.close()

    def test_competitor_layout_keeps_five_slots_and_extends_past_five(self):
        self.assertEqual(
            competitor_layout_last_column(
                [{"column_index": 3}, {"column_index": 4}, {"column_index": 5}]
            ),
            7,
        )
        self.assertEqual(
            competitor_layout_last_column(
                [{"column_index": column} for column in range(3, 9)]
            ),
            8,
        )

        workbook = Workbook()
        worksheet = workbook.active
        worksheet.column_dimensions["C"].width = 30
        worksheet.column_dimensions["G"].width = 12.375
        extend_competitor_layout(
            worksheet,
            [{"column": chr(64 + column)} for column in range(3, 9)],
            summary_row=11,
            border=Border(),
        )
        self.assertEqual(
            [worksheet.column_dimensions[chr(64 + column)].width for column in range(3, 9)],
            [30, 30, 30, 30, 30, 30],
        )
        workbook.close()

    def test_rejects_advantages_longer_than_twenty_characters(self):
        with self.assertRaisesRegex(ValueError, "超过 20 字"):
            advantage_text(["这是一个明显超过二十个汉字长度限制的产品优点总结"])

    def test_rejects_tampered_machine_fields(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            workbook_path = Path(temp_dir) / "input.xlsx"
            analysis_path = Path(temp_dir) / "analysis.json"
            output_path = Path(temp_dir) / "output.xlsx"
            self.make_workbook(workbook_path)
            report = prepare_review_analysis(workbook_path, analysis_path, sample_size=20)
            report["products"][0]["positive_review_count"] = 999
            analysis_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

            with self.assertRaisesRegex(ValueError, "只读机器字段"):
                backfill_review_analysis(analysis_path, workbook_path, output_path)


if __name__ == "__main__":
    unittest.main()
