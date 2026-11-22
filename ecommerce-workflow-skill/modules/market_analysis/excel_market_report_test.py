import base64
import io
import tempfile
import unittest
from copy import copy
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Border, PatternFill, Side

from excel_market_report import (
    build_dynamic_price_buckets,
    compute_metrics,
    ensure_rows_before_merged_title,
    generate_report,
)


class ComputeMetricsTests(unittest.TestCase):
    def test_computes_market_metrics_and_anomalies(self):
        rows = [
            {"ASIN": "A1", "品牌": "BrandA", "商品标题": "Item A", "月销量": 100, "月销售额($)": 1000, "价格($)": 10, "评分数": 5, "评分": 4.5, "上架天数": 20, "搜索排名": "自然位：第1页第1位"},
            {"ASIN": "A2", "品牌": "BrandB", "商品标题": "Item B", "月销量": 50, "月销售额($)": 1000, "价格($)": 20, "评分数": 100, "评分": 4.2, "上架天数": 500, "搜索排名": "自然位：第1页第2位"},
            {"ASIN": "A3", "品牌": "BrandC", "商品标题": "Item C", "月销量": 25, "月销售额($)": 750, "价格($)": 30, "评分数": 40, "评分": 4.0, "上架天数": 300, "搜索排名": "广告位：第1页第3位"},
        ]

        metrics = compute_metrics(rows)

        self.assertEqual(metrics["sample_count"], 3)
        self.assertEqual(metrics["total_sales"], 175)
        self.assertEqual(metrics["total_revenue"], 2750)
        self.assertAlmostEqual(metrics["top5_share"], 1.0)
        self.assertEqual(metrics["median_sales"], 50)
        self.assertEqual(metrics["median_price"], 20)
        self.assertEqual(metrics["sales_layers"]["Top 10"]["sample_count"], 3)
        self.assertEqual(metrics["price_buckets"]["$10-20"]["total_sales"], 100)
        self.assertEqual(metrics["price_buckets"]["$20-30"]["total_sales"], 75)
        self.assertEqual(metrics["review_buckets"]["≤30评"]["median_sales"], 100)
        self.assertEqual(metrics["anomalies"][0]["ASIN"], "A1")
        self.assertIn("低评分数高销量", metrics["anomalies"][0]["异常原因"])

    def test_computes_metrics_from_euro_currency_headers(self):
        rows = [
            {"ASIN": "A1", "品牌": "BrandA", "商品标题": "Item A", "月销量": 100, "月销售额(€)": "1,000", "价格(€)": "10.5", "评分数": 5, "评分": 4.5, "上架天数": 20},
            {"ASIN": "A2", "品牌": "BrandB", "商品标题": "Item B", "月销量": 50, "月销售额(€)": "€750", "价格(€)": "€15", "评分数": 100, "评分": 4.2, "上架天数": 500},
        ]

        metrics = compute_metrics(rows)

        self.assertEqual(metrics["currency_symbol"], "€")
        self.assertEqual(metrics["total_revenue"], 1750)
        self.assertEqual(metrics["median_price"], 12.75)
        self.assertEqual(metrics["price_buckets"]["€10-15"]["total_sales"], 150)

    def test_review_bucket_includes_product_without_monthly_sales(self):
        rows = [
            {"ASIN": "A1", "月销量": 100, "价格($)": 10, "评分数": 5, "评分": 4.5},
            {"ASIN": "A2", "月销量": None, "价格($)": 20, "评分数": 175, "评分": 4.2},
        ]

        metrics = compute_metrics(rows)

        self.assertEqual(metrics["sample_count"], 1)
        self.assertEqual(metrics["review_buckets"]["101-300评"]["sample_count"], 1)
        self.assertEqual(metrics["review_buckets"]["101-300评"]["total_sales"], 0)
        self.assertAlmostEqual(metrics["avg_reviews"], 90)

    def test_price_segments_follow_actual_price_gaps(self):
        prices = [9.99, 10.99, 11.99, 12.99, 19.99, 20.99, 22.99, 39.99]

        definitions = build_dynamic_price_buckets(prices, "$")

        self.assertEqual([item["label"] for item in definitions], ["$10-13", "$20-23", "$40"])
        self.assertEqual([item["upper_exclusive"] for item in definitions], [19.99, 39.99, None])

    def test_price_segment_labels_do_not_overlap(self):
        prices = [80.0, 85.99, 89.95, 99.0, 104.99, 121.83, 219.99, 397.99]

        definitions = build_dynamic_price_buckets(prices, "$")

        labels = [item["label"] for item in definitions]
        self.assertEqual(labels, ["$80-122", "$220", "$398"])


class WorkbookReportTests(unittest.TestCase):
    def test_inserted_rows_copy_price_table_format(self):
        wb = Workbook()
        ws = wb.active
        ws.row_dimensions[4].height = 18
        ws.cell(4, 1).value = "$10-20"
        ws.cell(4, 1).fill = PatternFill(fill_type="solid", fgColor="FFFF00")
        font = copy(ws.cell(4, 1).font)
        font.bold = True
        ws.cell(4, 1).font = font
        thin = Side(style="thin", color="000000")
        ws.cell(4, 1).border = Border(left=thin, right=thin, top=thin, bottom=thin)
        ws["A5"] = "3.价格分析"

        inserted = ensure_rows_before_merged_title(ws, 5, 4, 1)

        self.assertEqual(inserted, 3)
        for row in range(5, 8):
            self.assertEqual(ws.row_dimensions[row].height, 18)
            self.assertEqual(ws.cell(row, 1)._style, ws.cell(4, 1)._style)

    def test_fills_existing_market_analysis_template_without_changing_layout(self):
        with tempfile.TemporaryDirectory() as tmp:
            input_path = Path(tmp) / "input.xlsx"
            output_path = Path(tmp) / "output.xlsx"
            wb = Workbook()
            ws_research = wb.active
            ws_research.title = "Demo市场调研"
            ws_template = wb.create_sheet("Demo市场分析")
            ws_template.column_dimensions["A"].width = 31
            ws_template.row_dimensions[5].height = 44
            ws_template.merge_cells("A1:G1")
            ws_template.merge_cells("F3:G3")
            ws_template.merge_cells("F4:G4")
            ws_template.merge_cells("F5:G5")
            ws_template.merge_cells("F6:G6")
            ws_template["A1"] = "市场数据分析"
            for row, labels in {
                3: ["样本数", None, "平均价格", None, "中位价格"],
                4: ["Top5销量占比", None, "Top10销量占比", None, "Top20销量占比"],
                5: ["总销量", None, "平均月销量", None, "中位月销量"],
                6: ["总销售额", None, "平均销售额", None, "中位销售额"],
            }.items():
                for col, value in enumerate(labels, 1):
                    ws_template.cell(row, col, value)
            for col, value in enumerate(["分层", "样本数", "总月销量", "平均月销量", "中位月销量", "平均价格($)", "中位评分数"], 1):
                ws_template.cell(14, col, value)
            for col, value in enumerate(["价格段", "样本数", "总月销量", "平均月销量", "中位月销量", "中位价格($)"], 1):
                ws_template.cell(40, col, value)
            for row, value in enumerate(["$0-20", "$20-40", "$40-60", "$60-80", "$80-100"], 41):
                ws_template.cell(row, 1, value)
            ws_template["A46"] = "3.价格分析"
            for col, value in enumerate(["评分数区间", "样本数", "总月销量", "平均月销量", "中位月销量", "中位评分数", "代表评分"], 1):
                ws_template.cell(32, col, value)
            for col, value in enumerate(["ASIN", "品牌", "月销量", "价格($)", "评分数", "评分", "上架天数"], 1):
                ws_template.cell(43, col, value)
            ws_source = wb.create_sheet("数据源")
            ws_source.append(["ASIN", "品牌", "商品标题", "月销量", "月销售额($)", "价格($)", "评分数", "评分", "上架天数", "搜索排名"])
            ws_source.append(["A1", "BrandA", "Item A", 100.7, 1000.8, 10.9, 5.2, 4.56, 20.9, "自然位：第1页第1位"])
            ws_source.append(["A2", "BrandB", "Item B", 50.4, 999.9, 20.2, 100.8, 4.21, 500.1, "自然位：第1页第2位"])
            ws_source.append(["A3", "BrandC", "Item C", 25.1, 750.2, 30.1, 40.4, 4.0, 300.2, "广告位：第1页第3位"])
            ws_source.append(["A4", "BrandD", "Item D", 10, 1500, 150, 20, 4.3, 100, "自然位：第1页第4位"])
            ws_source.append(["A5", "BrandE", "Item E", None, 500, 50, 60, 4.1, 50, "自然位：第1页第5位"])
            wb.save(input_path)

            generate_report(input_path, output_path)

            out = load_workbook(output_path)
            analysis = out["Demo市场分析"]
            self.assertEqual(analysis.column_dimensions["A"].width, 31)
            self.assertEqual(analysis.row_dimensions[5].height, 44)
            self.assertIn("A1:G1", [str(r) for r in analysis.merged_cells.ranges])
            self.assertIn("F3:G3", [str(r) for r in analysis.merged_cells.ranges])
            self.assertEqual(analysis["B5"].value, 186)
            self.assertIsNotNone(analysis["F3"].value)
            self.assertIsInstance(analysis["B5"].value, int)
            self.assertEqual(analysis["B32"].value, 0)
            self.assertEqual(analysis["A8"].value.startswith("样本共 4 个链接"), True)
            self.assertEqual(analysis["A19"].value, None)
            self.assertTrue(str(analysis["A20"].value).startswith("小结："))
            self.assertEqual(analysis["A33"].value, None)
            self.assertTrue(str(analysis["A34"].value).startswith("小结："))
            self.assertEqual(analysis["A41"].value, "$11-30")
            self.assertEqual(analysis["A42"].value, "$150")
            self.assertEqual(analysis["A46"].value, "3.价格分析")
            self.assertTrue(str(analysis["A47"].value).startswith("小结："))
            self.assertGreaterEqual(len(analysis._charts), 3)
            for chart in analysis._charts:
                self.assertIsNotNone(chart.dataLabels)
                self.assertTrue(chart.dataLabels.showVal)
                self.assertFalse(chart.dataLabels.showSerName)
                self.assertFalse(chart.dataLabels.showCatName)
                self.assertFalse(chart.dataLabels.showLegendKey)
                self.assertIsNone(chart.x_axis.title)
                self.assertIsNone(chart.y_axis.title)

            source = out["Demo-数据源"]
            headers = {source.cell(1, column).value: column for column in range(1, source.max_column + 1)}
            self.assertNotIn("销量分层", headers)
            self.assertNotIn("评分数区间", headers)
            self.assertNotIn("价格段", headers)

    def test_adds_source_filter_columns_only_when_requested(self):
        with tempfile.TemporaryDirectory() as tmp:
            input_path = Path(tmp) / "input.xlsx"
            output_path = Path(tmp) / "output.xlsx"
            wb = Workbook()
            source = wb.active
            source.title = "数据源"
            source.append(["ASIN", "月销量", "价格($)", "评分数"])
            source.append(["A1", 100, 10, 5])
            source.append(["A2", None, 30, 60])
            source["A2"].hyperlink = "https://www.amazon.com/dp/A1"
            image_path = Path(tmp) / "product.png"
            image_path.write_bytes(base64.b64decode(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
            ))
            image_stream = io.BytesIO(image_path.read_bytes())
            test_image = XLImage(image_stream)
            source.add_image(test_image, "A2")
            wb.save(input_path)
            image_stream.close()

            generate_report(input_path, output_path, include_source_slicers=True)

            out = load_workbook(output_path)
            source = out["数据源"]
            source_headers = {source.cell(1, column).value: column for column in range(1, source.max_column + 1)}
            self.assertNotIn("销量分层", source_headers)
            self.assertEqual(len(source._images), 1)
            self.assertIn("切片联动", out.sheetnames)

            slicer_view = out["切片联动"]
            self.assertEqual(len(slicer_view._images), 1)
            self.assertEqual(slicer_view["A1"].value, "切片联动")
            self.assertIsNone(slicer_view["A2"].hyperlink)
            self.assertEqual(slicer_view["A13"].hyperlink.ref, "A13")
            headers = {
                slicer_view.cell(12, column).value: column
                for column in range(1, slicer_view.max_column + 1)
            }
            self.assertEqual(headers["销量分层"], 13)
            self.assertEqual(headers["评分数区间"], 14)
            self.assertEqual(headers["价格段"], 15)
            self.assertTrue(slicer_view.column_dimensions["M"].hidden)
            self.assertTrue(slicer_view.column_dimensions["N"].hidden)
            self.assertTrue(slicer_view.column_dimensions["O"].hidden)
            self.assertEqual(slicer_view.cell(13, headers["销量分层"]).value, "Top 10")
            self.assertEqual(slicer_view.cell(13, headers["评分数区间"]).value, "≤30评")
            self.assertEqual(slicer_view.cell(13, headers["价格段"]).value, "$10")
            self.assertEqual(slicer_view.cell(14, headers["销量分层"]).value, "未纳入市场分析")
            self.assertEqual(slicer_view.cell(14, headers["评分数区间"]).value, "未纳入市场分析")
            self.assertEqual(slicer_view.cell(14, headers["价格段"]).value, "未纳入市场分析")


if __name__ == "__main__":
    unittest.main()
