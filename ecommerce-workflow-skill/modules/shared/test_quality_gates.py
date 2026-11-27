from __future__ import annotations

import tempfile
import io
import unittest
from pathlib import Path
from zipfile import ZipFile
from unittest.mock import patch

from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from PIL import Image as PillowImage, ImageDraw

from shared.quality_gates import (
    QualityGateError,
    validate_keyword_research,
    validate_category_research,
    validate_final_package,
)


class QualityGateTests(unittest.TestCase):
    def test_current_final_gate_accepts_no_legacy_market_sheet_and_keeps_category_check(self):
        # 隔离无关模块，使用真实类目门禁验证：当前布局不需要旧页，缺图仍必须失败。
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'current.xlsx'
            book = Workbook()
            sheet = book.active
            sheet.title = '示例-市场调研'
            sheet['A2'], sheet['A3'], sheet['A4'] = '一级类目', '手工', '3.市场数据分析'
            sheet['A3'].hyperlink = 'https://www.amazon.com/gp/bestsellers/arts-crafts/1'
            sheet['I3'] = '100\n数量占比：100.0%'
            image_path = Path(folder) / 'category.png'
            with PillowImage.new('RGB', (1600, 940), 'blue') as picture:
                picture.save(image_path)
            sheet.add_image(ExcelImage(str(image_path)), 'F3')
            book.save(path)
            with patch('shared.quality_gates.validate_keyword_research'), \
                 patch('shared.quality_gates.validate_overview_charts'), \
                 patch('shared.quality_gates.validate_competitor_basics'):
                self.assertTrue(validate_final_package(path))
                sheet._images.clear()
                book.save(path)
                with self.assertRaisesRegex(QualityGateError, '截图缺失'):
                    validate_final_package(path)
            book.close()

    def test_embedded_category_uses_existing_large_cell_without_resizing(self):
        from shared.cell_image_embedding import embed_final_images
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            image = root/'category.png'
            with PillowImage.new('RGB',(1600,940),'blue') as picture:
                picture.save(image)
            book = Workbook()
            sheet = book.active
            sheet.title = '示例-市场调研'
            sheet['A2'],sheet['A3'],sheet['A4'] = '一级类目','手工','3.市场数据分析'
            sheet['A3'].hyperlink = 'https://www.amazon.com/gp/bestsellers/arts-crafts/1'
            sheet['I3'] = '100\n数量占比：100.0%'
            sheet.row_dimensions[3].height = 200
            sheet.column_dimensions['F'].width = 50
            sheet.merge_cells('F3:H3')
            sheet.add_image(ExcelImage(str(image)),'F3')
            source,output = root/'source.xlsx',root/'embedded.xlsx'
            book.save(source)
            book.close()
            embed_final_images(source,output)
            self.assertEqual(validate_category_research(output).metrics['images'],1)

    def test_embedded_keyword_uses_cell_fit_and_still_rejects_low_resolution(self):
        from shared.cell_image_embedding import embed_final_images
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = self.make_keyword_workbook(root,include_link=True)
            embedded = root/'embedded.xlsx'
            embed_final_images(source,embedded)
            self.assertEqual(validate_keyword_research(embedded).metrics['images'],1)
            # 模式1也必须保留高清原图，不能只因已嵌入就绕过低分辨率检查。
            payload = io.BytesIO()
            with PillowImage.new('RGB',(8,8),'blue') as picture:
                picture.save(payload,format='PNG')
            low_source = root/'low-source.xlsx'
            with ZipFile(source) as before,ZipFile(low_source,'w') as after:
                for info in before.infolist():
                    after.writestr(info,payload.getvalue() if info.filename.startswith('xl/media/') else before.read(info.filename))
            low_output = root/'low-embedded.xlsx'
            embed_final_images(low_source,low_output)
            with self.assertRaisesRegex(QualityGateError,'高清原图'):
                validate_keyword_research(low_output)

    def test_category_native_size_passes_but_stretching_fails(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            image_path = root / 'category.png'
            with PillowImage.new('RGB', (1600, 940), 'white') as picture:
                ImageDraw.Draw(picture).rectangle((20, 20, 1580, 920), fill='blue')
                picture.save(image_path)
            book = Workbook()
            sheet = book.active
            sheet.title = '示例-市场调研'
            sheet['A2'], sheet['A3'], sheet['A4'] = '一级类目', '手工', '3.市场数据分析'
            sheet['A3'].hyperlink = 'https://www.amazon.com/gp/bestsellers/arts-crafts/1'
            sheet['I3'] = '100\n数量占比：100.0%'
            image = ExcelImage(image_path)
            sheet.add_image(image, 'F3')
            path = root / 'category.xlsx'
            book.save(path)
            self.assertEqual(1, validate_category_research(path).metrics['images'])
            image.width = 1800
            book.save(path)
            with self.assertRaises(QualityGateError):
                validate_category_research(path)
            book.close()

    def test_category_boundary_supports_both_templates_and_still_checks_links(self):
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / 'category.xlsx'
            for title in ('3.数据分析概览', '3.市场数据分析'):
                book = Workbook()
                sheet = book.active
                sheet.title = '示例-市场调研'
                sheet['A2'], sheet['A3'], sheet['A4'] = '一级类目', '手工', title
                sheet['A3'].hyperlink = 'https://www.amazon.com/gp/bestsellers/arts-crafts/1'
                book.save(path)
                self.assertEqual(1, validate_category_research(path, screenshot_limit=0).metrics['rows'])
                sheet['A3'].hyperlink = 'https://www.amazon.com/b?node=1'
                book.save(path)
                with self.assertRaisesRegex(QualityGateError, 'BSR'):
                    validate_category_research(path, screenshot_limit=0)
                book.close()

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
