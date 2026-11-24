"""检查新版整体指标随模板增行定位，并保护右侧结论合并区。"""
import unittest
from copy import copy
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from product_research.scripts.populate_market_summary import METRICS, fill_summary
from shared.worksheet_rows import insert_rows_preserving_merges


class MarketSummaryTests(unittest.TestCase):
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
