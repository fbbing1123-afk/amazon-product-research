"""验证字段口径、来源优先级和 ASIN 隔离，浏览器实测另行记录。"""

from datetime import datetime
from pathlib import Path
import tempfile
import unittest

from openpyxl import Workbook

from modules.competitor_analysis.collect_basic_facts import (
    FIELDS, asin_from_input, normalize_value, parse_plugin_panel, read_source, resolve_product,
    normalize_selling_points,
    normalize_alexa_questions, aggregate_question_evidence,
)


ASIN = "B000TEST01"
PANEL = """ASIN: B000TEST01
品牌: 示例
近30天销量(父体): 904
近30天销量(子体): 600+
销售额: $32,535
均价: $34.59
价格: $35.99
评分(评分数): 4.2 (1,020)
上架时间: 2025-07-25 (486天)
"""
URL = "https://www.amazon.com/dp/" + ASIN


class BasicFactsTests(unittest.TestCase):
    def test_alexa_keeps_questions_only_and_deduplicates(self):
        self.assertEqual("Can it seal liquids?\nIs it portable?", normalize_alexa_questions([
            "Submit", "", " Can it seal liquids? ", "CAN IT SEAL LIQUIDS?", "Is it portable?", "Looking for specific info"
        ]))
        self.assertIsNone(normalize_alexa_questions("暂无"))

    def test_question_evidence_counts_product_pages_not_repeated_buttons(self):
        products = [
            {"asin": "B000TEST01", "fields": {"alexa_questions": {"value": "Is it portable?\nIs it portable?"}}},
            {"asin": "B000TEST02", "fields": {"alexa_questions": {"value": "Is it portable?\nCan it seal liquids?"}}},
            {"asin": "B000TEST03", "fields": {"alexa_questions": {"value": None}}},
        ]
        groups = aggregate_question_evidence(products)
        self.assertEqual([2, 1], [group["product_page_count"] for group in groups])
        self.assertEqual(["B000TEST01", "B000TEST02"], groups[0]["asins"])

    def observation(self):
        return parse_plugin_panel(PANEL, ASIN, URL)

    def test_parent_monthly_sales_current_price_and_date(self):
        fields = self.observation()["fields"]
        self.assertEqual(904, fields["monthly_sales"]["value"])
        self.assertEqual(35.99, fields["price"]["value"])
        self.assertEqual("2025-07-25", fields["listed_at"]["value"])

    def test_child_sales_and_average_price_are_not_fallbacks(self):
        text = PANEL.replace("近30天销量(父体): 904", "近30天销量(父体): --").replace("价格: $35.99", "")
        result = parse_plugin_panel(text, ASIN, URL)
        self.assertIsNone(result["fields"]["monthly_sales"]["value"])
        self.assertIsNone(result["fields"]["price"]["value"])

    def test_other_asin_or_wrong_marketplace_rejected(self):
        for url, text in ((URL.replace(ASIN, "B000TEST02"), PANEL),
                          (URL, PANEL.replace(ASIN, "B000TEST02")),
                          (URL.replace("amazon.com", "amazon.ca"), PANEL)):
            with self.assertRaises(ValueError):
                parse_plugin_panel(text, ASIN, url)

    def test_input_urls_are_reduced_to_exact_asin(self):
        self.assertEqual(ASIN, asin_from_input(URL + "/ref=abc?psc=1"))
        self.assertEqual(ASIN, asin_from_input(ASIN.lower()))
        for value in ("my-internal-sku", "https://amazon.com.evil.example/dp/" + ASIN,
                      "https://www.amazon.com/s?k=" + ASIN):
            with self.assertRaises(ValueError):
                asin_from_input(value)

    def test_source_zero_is_preserved_and_browser_is_not_called(self):
        candidates = {field: [{"value": value, "source": "workbook"}] for field, value in
                      zip(FIELDS, [0, 25.0, "2024-01-01"])}
        def unexpected(*args):
            self.fail("完整数据源不应触发网页")
        result = resolve_product(ASIN, candidates, unexpected)
        self.assertEqual("complete", result["status"])
        self.assertEqual(0, result["fields"]["monthly_sales"]["value"])

    def test_web_only_fills_missing_fields(self):
        calls = []
        def fetch(asin, needed):
            calls.append((asin, needed))
            return self.observation()
        result = resolve_product(ASIN, {"price": [{"value": 29.99, "source": "workbook"}]}, fetch)
        self.assertEqual([(ASIN, ["monthly_sales", "listed_at"])], calls)
        self.assertEqual(29.99, result["fields"]["price"]["value"])
        self.assertEqual("workbook", result["fields"]["price"]["source"])
        self.assertEqual(904, result["fields"]["monthly_sales"]["value"])

    def test_conflicting_rows_are_not_silently_overwritten(self):
        result = resolve_product(ASIN, {"monthly_sales": [{"value": 100}, {"value": 200}]},
                                 lambda *args: self.observation())
        self.assertEqual("conflict", result["fields"]["monthly_sales"]["source"])
        self.assertIsNone(result["fields"]["monthly_sales"]["value"])
        self.assertEqual("incomplete", result["status"])

    def test_page_failure_preserves_source_values(self):
        def fail(*args):
            raise RuntimeError("插件未登录")
        result = resolve_product(ASIN, {"price": [{"value": 29.99, "source": "workbook"}]}, fail)
        self.assertEqual(29.99, result["fields"]["price"]["value"])
        self.assertEqual(["月销量", "上架时间"], result["missing_fields"])
        self.assertIn("插件未登录", result["errors"][0])

    def test_missing_ranges_and_invalid_dates_are_not_invented(self):
        for value in (None, "--", "/", "600+", "1k", -1, float("nan"), True):
            self.assertIsNone(normalize_value("monthly_sales", value))
        self.assertEqual(20639, normalize_value("monthly_sales", "20,639"))
        self.assertEqual("2025-07-25", normalize_value("listed_at", datetime(2025, 7, 25)))
        self.assertIsNone(normalize_value("listed_at", "2025-02-31"))
        self.assertIsNone(normalize_value("price", 0))

    def test_actual_workbook_reads_exact_headers_and_asin(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "source.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "产品-数据源"
            ws.append(["ASIN", "月销量", "子体销量", "价格($)", "Prime价格($)", "上架时间"])
            ws.append([ASIN, 0, "600+", 35.99, 25.99, datetime(2025, 7, 25)])
            ws.append(["B000TEST02", 999, 888, 100, 50, None])
            wb.save(path)
            values = read_source(path, [ASIN, "B000TEST03"])
            result = resolve_product(ASIN, values[ASIN])
            self.assertEqual([0, 35.99, "2025-07-25"], [result["fields"][f]["value"] for f in FIELDS])
            self.assertEqual(2, result["fields"]["price"]["row"])
            self.assertEqual("incomplete", resolve_product("B000TEST03", values["B000TEST03"])["status"])

    def test_selling_points_missing_cell_and_missing_asin_both_fetch(self):
        fields = {"selling_points": "产品卖点"}
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "source.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "数据源"
            ws.append(["ASIN", "产品卖点"])
            ws.append([ASIN, " / "])
            ws.append(["B000TEST02", "已有卖点"])
            wb.save(path)
            targets = [ASIN, "B000TEST02", "B000TEST03"]
            source = read_source(path, targets, fields=fields)
            calls = []
            def fetch(asin, needed):
                calls.append(asin)
                return {"url": "https://www.amazon.com/dp/" + asin, "observed_at": "测试",
                        "fields": {"selling_points": {"value": "页面原文", "source": "amazon_product_page"}}}
            results = [resolve_product(asin, source[asin], fetch, fields) for asin in targets]
            self.assertEqual([ASIN, "B000TEST03"], calls)
            self.assertEqual("已有卖点", results[1]["fields"]["selling_points"]["value"])
            self.assertEqual("workbook", results[1]["fields"]["selling_points"]["source"])
            self.assertEqual("amazon_product_page", results[0]["fields"]["selling_points"]["source"])
            self.assertTrue(all(result["status"] == "complete" for result in results))

    def test_missing_selling_points_column_uses_fallback(self):
        fields = {"selling_points": "产品卖点"}
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "source.xlsx"
            wb = Workbook()
            wb.active.title = "数据源"
            wb.active.append(["ASIN"])
            wb.active.append([ASIN])
            wb.save(path)
            result = read_source(path, [ASIN], fields=fields)
            self.assertEqual([], result[ASIN]["selling_points"])

    def test_selling_points_preserve_six_items_without_truncation(self):
        items = [f"Point {number}: original text" for number in range(6)]
        value = normalize_selling_points(["About this item", "", *items, items[0], "See more product details"])
        self.assertEqual(items, value.splitlines())
        for value in (None, " \n /\n--", 0, False):
            self.assertIsNone(normalize_selling_points(value))


if __name__ == "__main__":
    unittest.main()
