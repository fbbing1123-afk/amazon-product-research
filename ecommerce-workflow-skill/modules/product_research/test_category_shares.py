from __future__ import annotations

import tempfile
import unittest
import zipfile
from pathlib import Path

from openpyxl import Workbook, load_workbook

from modules.product_research.scripts.fill_custom_template import (
    apply_category_share_display, collect_category_summaries, fill_category_section,
)
from modules.product_research.scripts.refresh_category_shares import (
    collect_updates, refresh_category_shares,
)


class CategoryShareTests(unittest.TestCase):
    def fixture(self):
        source = Workbook()
        data = source.active
        data.title = "数据源"
        data.append(["类目路径", "大类目", "月销量"])
        for node, name, sales in [(11, "A", 1), (11, "A", 2), (22, "B", 9999)]:
            data.append([f"Electronics:{name}", "Electronics", sales])
            data.cell(data.max_row, 1).hyperlink = f"https://www.amazon.com/b?node={node}"
        data.append(["", "Electronics", 10000])
        report = Workbook()
        sheet = report.active
        sheet.title = "市场调研"
        sheet["A1"] = "2. 所属平台分类及常用分类节点\n注：蓝色为常见类目 红色为可探索类目"
        sheet["A2"] = "一级类目"
        sheet["I2"] = "月总销量"
        # 报告的行序与数据源相反，必须按节点匹配。
        for row, name, node in [(3, "B", 22), (4, "A", 11)]:
            sheet.cell(row, 1, "Electronics")
            sheet.cell(row, 2, name).hyperlink = f"https://www.amazon.com/gp/bestsellers/electronics/{node}"
            sheet.cell(row, 9, 42095)
        sheet["A5"] = "3. 数据分析概览"
        report.create_sheet("其它页")["A1"] = "保留"
        return report, source

    def test_counts_not_sales_and_invalid_rows_excluded(self):
        report, source = self.fixture()
        _, _, _, updates = collect_updates(report, source)
        self.assertAlmostEqual(1 / 3, updates[0]["share"])
        self.assertAlmostEqual(2 / 3, updates[1]["share"])
        self.assertEqual([1, 2], [u["count"] for u in updates])
        self.assertAlmostEqual(1, sum(s.product_share for s in collect_category_summaries(source.active)))

    def test_unknown_node_does_not_guess_by_name(self):
        report, source = self.fixture()
        report.active["B3"].hyperlink = "https://www.amazon.com/gp/bestsellers/electronics/999"
        with self.assertRaisesRegex(ValueError, "无法唯一匹配"):
            collect_updates(report, source)

    def test_template_and_later_sales_refresh_keep_display(self):
        report, source = self.fixture()
        fill_category_section(report, source)
        sheet = report.active
        self.assertIn("蓝色 >50%", sheet["A1"].value)
        sheet["I3"] = 12345
        self.assertEqual(12345, sheet["I3"].value)
        self.assertIn("66.7%", sheet["I3"].number_format)
        self.assertTrue(sheet["I3"].alignment.wrap_text)
        self.assertEqual("D9EAF7", sheet["I3"].fill.fgColor.rgb[-6:])
        # 真实零销量不转成缺失值，100% 和 0% 均可显示。
        sheet["I3"] = 0
        for share, color in ((0, "FCE4D6"), (0.5, "FCE4D6"), (1, "D9EAF7")):
            apply_category_share_display(sheet["I3"], share)
            self.assertEqual(0, sheet["I3"].value)
            self.assertIn(f"{share:.1%}", sheet["I3"].number_format)
            self.assertEqual(color, sheet["I3"].fill.fgColor.rgb[-6:])

    def test_zip_refresh_preserves_values_other_parts_and_newline(self):
        report, source = self.fixture()
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            before, data, after = [root / name for name in ("report.xlsx", "data.xlsx", "output.xlsx")]
            report.save(before)
            source.save(data)
            refresh_category_shares(before, data, after)
            with zipfile.ZipFile(before) as old, zipfile.ZipFile(after) as new:
                self.assertEqual(old.namelist(), new.namelist())
                changed = {name for name in old.namelist() if old.read(name) != new.read(name)}
                self.assertEqual({"xl/worksheets/sheet1.xml", "xl/styles.xml"}, changed)
            reopened = load_workbook(after, rich_text=True)
            self.assertEqual(42095, reopened.active["I3"].value)
            self.assertEqual('n', reopened.active["I3"].data_type)
            self.assertIn("\n产品占比：33.3%", reopened.active["I3"].number_format)
            self.assertIn("\n产品占比：66.7%", reopened.active["I4"].number_format)
            self.assertEqual("FCE4D6", reopened.active["I3"].fill.fgColor.rgb[-6:])
            self.assertEqual("D9EAF7", reopened.active["I4"].fill.fgColor.rgb[-6:])
            reopened.close()
            with self.assertRaisesRegex(ValueError, "不能覆盖"):
                refresh_category_shares(before, data, before)


if __name__ == "__main__":
    unittest.main()
