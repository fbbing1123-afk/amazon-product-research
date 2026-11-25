"""检查新版整体指标随模板增行定位，并保护右侧结论合并区。"""
import unittest
import tempfile
from pathlib import Path
from zipfile import ZipFile
from copy import copy
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from product_research.scripts.populate_market_summary import METRICS, fill_summary, refresh_summary
from shared.worksheet_rows import insert_rows_preserving_merges


class MarketSummaryTests(unittest.TestCase):
    def test_zip_metrics_write_blank_cells_and_preserve_other_parts(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            book = Workbook()
            book.active.title = '市场调研'
            for index, label in enumerate(METRICS):
                book.active.cell(14 + index//3, 1 + index%3*2, label)
            book.active.merge_cells('G14:I17')
            book.active['G14'] = '原有结论'
            source = book.create_sheet('数据源')
            source.append(['ASIN', '月销量', '价格($)', '月销售额($)'])
            source.append(['B012345678', 10, 2, 20])
            source.append(['B087654321', 30, 6, 180])
            book.create_sheet('评价数据源').append(['分析商品 ASIN', '原评论'])
            before, after = root/'before.xlsx', root/'after.xlsx'
            book.save(before)
            book.close()
            result = refresh_summary(before, after, root/'data')
            checked = load_workbook(after, rich_text=True)
            expected = [2,4,4,1,1,1,40,20,20,200,100,100]
            for label, value in zip(METRICS, expected):
                self.assertEqual(checked['市场调研'][result['metrics'][label]].value, value)
            self.assertEqual(checked['市场调研']['G14'].value, '原有结论')
            checked.close()
            with ZipFile(before) as old, ZipFile(after) as new:
                self.assertEqual(old.namelist(), new.namelist())
                self.assertEqual({n for n in old.namelist() if old.read(n)!=new.read(n)},
                                 {'xl/styles.xml','xl/worksheets/sheet1.xml'})

    def test_expanded_rows_preserve_conclusion_and_styles(self):
        ws = Workbook().active
        for index, label in enumerate(METRICS):
            row, col = 14 + index // 3, 1 + index % 3 * 2
            ws.cell(row, col, label)
            ws.cell(row, col + 1).fill = PatternFill("solid", fgColor="FFABCDEF")
        ws.merge_cells("G14:I17")
        ws["G14"] = "原有结论"
        before = copy(ws["F14"].fill)
        insert_rows_preserving_merges(ws, 10, 11, 9)
        metrics = {key: index + 1 for index, key in enumerate(METRICS.values())}
        written = fill_summary(ws, metrics)
        self.assertEqual("F25", written["中位价格($)"])
        self.assertEqual(3, ws["F25"].value)
        self.assertEqual("原有结论", ws["G25"].value)
        self.assertIn("G25:I28", ws.merged_cells)
        self.assertEqual(before, copy(ws["F25"].fill))

    def test_conflicting_merge_is_rejected(self):
        ws = Workbook().active
        for index, label in enumerate(METRICS):
            ws.cell(14 + index // 3, 1 + index % 3 * 2, label)
        ws.merge_cells("F14:I17")
        with self.assertRaisesRegex(ValueError, "合并区域冲突"):
            fill_summary(ws, dict.fromkeys(METRICS.values(), 1))


if __name__ == "__main__":
    unittest.main()
