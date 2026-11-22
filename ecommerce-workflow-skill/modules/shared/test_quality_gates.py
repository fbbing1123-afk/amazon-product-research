from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from PIL import Image as PillowImage, ImageDraw

from shared.quality_gates import (
    QualityGateError,
    validate_keyword_research,
)


class QualityGateTests(unittest.TestCase):
    def make_keyword_workbook(self, root: Path, include_link: bool) -> Path:
        image_path = root / "trend.png"
        picture = PillowImage.new("RGB", (800, 300), "white")
        ImageDraw.Draw(picture).rectangle((40, 40, 760, 260), fill="blue")
        picture.save(image_path)
        picture.close()
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "demo-市场调研"
        sheet["A2"] = "1. 关键词搜索量及趋势"
        sheet["A4"] = "关键词"
        sheet["A5"] = "starlink mount\n(星链支架)"
        if include_link:
            sheet["A5"].hyperlink = "https://www.amazon.com/s?k=starlink+mount"
        sheet["E5"] = "无明显趋势"
        sheet["F5"] = 3000
        sheet["G5"] = 100
        sheet["H5"] = 500
        sheet["I5"] = 1.2
        sheet["A6"] = "2. 所属平台分类及常用分类节点"
        image = ExcelImage(image_path)
        image.width = 200
        image.height = 75
        sheet.add_image(image, "B5")
        path = root / "keyword.xlsx"
        workbook.save(path)
        workbook.close()
        return path

    def test_keyword_gate_requires_search_link_and_high_resolution_image(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            missing_link = self.make_keyword_workbook(root, include_link=False)
            with self.assertRaisesRegex(QualityGateError, "Amazon 搜索超链接"):
                validate_keyword_research(missing_link)

            valid = self.make_keyword_workbook(root, include_link=True)
            report = validate_keyword_research(valid)
            self.assertEqual(report.metrics["rows"], 1)
            self.assertEqual(report.metrics["images"], 1)


if __name__ == "__main__":
    unittest.main()
