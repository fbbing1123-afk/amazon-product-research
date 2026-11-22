from __future__ import annotations

import unittest

from openpyxl import Workbook

from modules.product_research.scripts.fill_custom_template import collect_category_summaries


class CategoryDedupTests(unittest.TestCase):
    def make_sheet(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["类目路径", "大类目", "月销量"])
        return sheet

    def append_category(self, sheet, path, big_category, node_id, big_slug, monthly_sales):
        sheet.append([path, big_category, monthly_sales])
        row = sheet.max_row
        if node_id:
            sheet.cell(row, 1).hyperlink = f"https://www.amazon.com/b?node={node_id}"
        if big_slug:
            sheet.cell(row, 2).hyperlink = f"https://www.amazon.com/gp/bestsellers/{big_slug}"

    def test_same_node_merges_and_prefers_matching_big_category(self):
        sheet = self.make_sheet()
        path = "Health & Household:Health Care:Ear Care:Earwax Removal"
        self.append_category(sheet, path, "Beauty & Personal Care", "3779891", "beauty", 100)
        self.append_category(sheet, path, "Health & Household", "3779891", "hpc", 200)

        summaries = collect_category_summaries(sheet)

        self.assertEqual(1, len(summaries))
        self.assertEqual(2, summaries[0].product_count)
        self.assertEqual(300, summaries[0].monthly_sales)
        self.assertEqual("Health & Household", summaries[0].big_category)
        self.assertEqual("https://www.amazon.com/gp/bestsellers/hpc/3779891", summaries[0].final_link)

    def test_same_path_without_node_merges(self):
        sheet = self.make_sheet()
        path = "Health & Household:Health Care:Ear Care:Earwax Removal"
        self.append_category(sheet, path, "Beauty & Personal Care", None, "beauty", 100)
        self.append_category(sheet, path, "Health & Household", None, "hpc", 200)

        summaries = collect_category_summaries(sheet)

        self.assertEqual(1, len(summaries))
        self.assertEqual(2, summaries[0].product_count)

    def test_missing_big_category_link_falls_back_to_bsr_node(self):
        sheet = self.make_sheet()
        path = "Home & Kitchen:Event & Party Supplies:Party Packs"
        self.append_category(sheet, path, "Home & Kitchen", "723480011", None, 100)

        summaries = collect_category_summaries(sheet)

        self.assertEqual(1, len(summaries))
        self.assertEqual(
            "https://www.amazon.com/gp/bestsellers/home-garden/723480011",
            summaries[0].final_link,
        )

    def test_different_nodes_with_same_path_remain_separate(self):
        sheet = self.make_sheet()
        path = "Health & Household:Health Care:Ear Care:Earwax Removal"
        self.append_category(sheet, path, "Health & Household", "111", "hpc", 100)
        self.append_category(sheet, path, "Health & Household", "222", "hpc", 200)

        summaries = collect_category_summaries(sheet)

        self.assertEqual(2, len(summaries))


if __name__ == "__main__":
    unittest.main()
