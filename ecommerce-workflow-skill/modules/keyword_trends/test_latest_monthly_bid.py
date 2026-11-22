import importlib.util
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.drawing.image import Image
from PIL import Image as PILImage

sys.path.insert(0, str(Path(__file__).parent))
from keyword_rows import latest_monthly_bid, update_keyword_rows

spec = importlib.util.spec_from_file_location("trends_latest_bid_tests", Path(__file__).with_name("excel-sellersprite-trends.py"))
module = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = module
spec.loader.exec_module(module)


class LatestMonthTests(unittest.TestCase):
    def test_latest_month_across_year_and_unsorted_input(self):
        self.assertEqual(latest_monthly_bid([
            {"month": "2026-1", "bid": "1.20"},
            {"month": "2025-12", "bid": 9},
            {"month": "2026-01", "bid": 1.5},
        ]), ("2026-01", 1.5))

    def test_latest_missing_or_zero_does_not_use_old_bid(self):
        for missing in [0, None, "/", float("nan"), float("inf")]:
            with self.subTest(missing=missing):
                self.assertEqual(latest_monthly_bid([
                    {"month": "2026-09", "bid": 4}, {"month": "2026-10", "bid": missing},
                ]), ("2026-10", None))

    def test_week_dates_and_invalid_months_are_not_monthly_data(self):
        self.assertEqual(latest_monthly_bid([
            {"month": "2026-13", "bid": 7}, {"month": "2026-10-02", "bid": 9},
            {"week": "2026-11-01", "bid": 8},
        ]), (None, None))

    def test_monthly_bid_does_not_depend_on_search_history(self):
        book = Workbook()
        sheet = book.active
        labels = ["关键词", "趋势旺季", "淡季月搜索量", "旺季月搜索量", "旺季竞价"]
        sheet.append(labels)
        sheet.append(["example"])
        module.apply_trend_data_fields(sheet, [module.KeywordItem("example", "A2", "F2")],
            {label: sheet.cell(1, col) for col, label in enumerate(labels, 1)},
            {"example": {"ppcMonthly": [{"month": "2026-10", "bid": 1.25}],
                         "ppcWeekly": [{"week": "2026-11-01", "bid": 99}]}}, 2025, ppc_bid_mode="latest-month")
        self.assertEqual(sheet["E2"].value, 1.25)
        self.assertEqual(sheet["D2"].value, "/")

    def test_monthly_and_old_mode_remain_selectable(self):
        for mode, expected in [("latest-month", 1.25), ("peak-period", 9)]:
            book = Workbook()
            sheet = book.active
            labels = ["关键词", "趋势旺季", "淡季月搜索量", "旺季月搜索量", "旺季竞价"]
            sheet.append(labels)
            sheet.append(["example"])
            module.apply_trend_data_fields(sheet, [module.KeywordItem("example", "A2", "F2")],
                {label: sheet.cell(1, col) for col, label in enumerate(labels, 1)},
                {"example": {"trendMonthly": [{"month": f"2025-{m:02d}", "searchVolume": 500 if m == 6 else 100} for m in range(1, 13)],
                             "ppcMonthly": [{"month": "2026-10", "bid": 1.25}],
                             "ppcWeekly": [{"week": "2025-06-01", "bid": 9}]}}, 2025, ppc_bid_mode=mode)
            self.assertEqual(sheet["E2"].value, expected)


class KeywordSortTests(unittest.TestCase):
    def test_sort_preserves_links_rich_text_images_merges_and_other_parts(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source, output, png = root / "in.xlsx", root / "out.xlsx", root / "picture.png"
            PILImage.new("RGB", (100, 30), "blue").save(png)
            book = Workbook()
            sheet = book.active
            sheet.title = "市场调研"
            for row, (name, peak) in enumerate([("low", 10), ("high", "1,000"), ("tie", 1000), ("missing", "/")], 5):
                sheet.cell(row, 1, CellRichText([TextBlock(InlineFont(b=True), name)]))
                sheet.cell(row, 1).hyperlink = "https://www.amazon.com/s?k=" + name
                sheet.cell(row, 8, peak)
                sheet.cell(row, 9, 4)
                sheet.cell(row, 10, f"=H{row}*2")
                sheet.cell(row, 2, f'=DISPIMG("image-{name}",1)')
                sheet.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
                sheet.row_dimensions[row].height = row * 10
            sheet.add_image(Image(str(png)), "E5")
            sheet["A12"] = "类目区不变"
            book.create_sheet("其它页")["A1"] = "保留内容"
            book.save(source)
            with zipfile.ZipFile(source, "a") as archive:
                archive.writestr("xl/cellimages.xml", b"unchanged-cellimage-fixture")
            source_bytes = source.read_bytes()
            report = update_keyword_rows(source, output, sheet_name=sheet.title, rows=[5, 6, 7, 8],
                peak_column=8, bid_column=9, bids={5: .5, 6: .6, 7: .7, 8: None})
            self.assertEqual(source.read_bytes(), source_bytes)
            result = load_workbook(output, rich_text=True)
            ws = result["市场调研"]
            self.assertEqual([str(ws.cell(r, 1).value) for r in range(5, 9)], ["high", "tie", "low", "missing"])
            self.assertEqual([ws.cell(r, 9).value for r in range(5, 9)], [.6, .7, .5, "/"])
            self.assertEqual(ws["A7"].hyperlink.target, "https://www.amazon.com/s?k=low")
            self.assertTrue(ws["A7"].value[0].font.b)
            self.assertEqual(ws["B7"].value, '=DISPIMG("image-low",1)')
            self.assertEqual(ws["J7"].value, "=H7*2")
            self.assertEqual(ws.row_dimensions[7].height, 50)
            self.assertEqual(ws._images[0].anchor._from.row, 6)
            self.assertEqual(ws["A12"].value, "类目区不变")
            self.assertEqual({str(m) for m in ws.merged_cells.ranges}, {"B5:D5", "B6:D6", "B7:D7", "B8:D8"})
            with zipfile.ZipFile(source) as before, zipfile.ZipFile(output) as after:
                self.assertEqual(before.namelist(), after.namelist())
                for name in before.namelist():
                    if name not in report["patched_parts"]:
                        self.assertEqual(before.read(name), after.read(name), name)
            result.close()

    def test_cross_row_merge_is_rejected_without_output(self):
        with tempfile.TemporaryDirectory() as directory:
            source, output = Path(directory)/"in.xlsx", Path(directory)/"out.xlsx"
            book = Workbook()
            sheet = book.active
            sheet["H5"], sheet["H6"] = 1, 2
            sheet.merge_cells("B5:D6")
            book.save(source)
            with self.assertRaisesRegex(ValueError, "跨行合并"):
                update_keyword_rows(source, output, sheet_name=sheet.title, rows=[5, 6], peak_column=8)
            self.assertFalse(output.exists())


if __name__ == "__main__":
    unittest.main()
