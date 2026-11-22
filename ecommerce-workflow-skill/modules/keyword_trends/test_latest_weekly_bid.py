import unittest

from test_latest_monthly_bid import module
from keyword_rows import latest_weekly_bid
from openpyxl import Workbook


class LatestWeekTests(unittest.TestCase):
    def test_latest_week_instead_of_maximum_recent_week(self):
        self.assertEqual(latest_weekly_bid([
            {"week":"2026-11-08", "weekEnd":"2026-11-13", "bid":2.1},
            {"week":"2026-11-01", "weekEnd":"2026-11-06", "bid":9},
        ]), ("2026-11-08~2026-11-13", 2.1))

    def test_latest_missing_week_does_not_fall_back(self):
        for value in [None, 0, "/", float("nan"), float("inf")]:
            with self.subTest(value=value):
                self.assertEqual(latest_weekly_bid([
                    {"week":"2026-11-01", "bid":9}, {"week":"2026-11-08", "bid":value},
                ]), ("2026-11-08", None))

    def test_cross_year_unsorted_and_duplicate_rows(self):
        self.assertEqual(latest_weekly_bid([
            {"week":"2025-12-28", "weekEnd":"2026-01-02", "bid":2},
            {"week":"2025-12-21", "weekEnd":"2025-12-26", "bid":8},
            {"week":"2025-12-28", "weekEnd":"2026-01-02", "bid":3},
        ]), ("2025-12-28~2026-01-02", 3))

    def test_rejects_invalid_dates_and_month_only_labels(self):
        self.assertEqual(latest_weekly_bid([
            {"week":"2026-11", "bid":9}, {"week":"2026-02-30", "bid":9},
            {"week":"2026-11-08", "weekEnd":"2026-11-01", "bid":9},
        ]), (None, None))

    def test_default_weekly_bid_with_and_without_trend_history(self):
        for history in [[], [{"month":f"2025-{m:02d}","searchVolume":100} for m in range(1,13)]]:
            book = Workbook()
            sheet = book.active
            labels = ["关键词", "趋势旺季", "淡季月搜索量", "旺季月搜索量", "旺季竞价"]
            sheet.append(labels)
            sheet.append(["example"])
            module.apply_trend_data_fields(sheet, [module.KeywordItem("example", "A2", "F2")],
                {label:sheet.cell(1,col) for col,label in enumerate(labels,1)},
                {"example":{"trendMonthly":history, "ppcMonthly":[{"month":"2026-11","bid":99}],
                            "ppcWeekly":[{"week":"2026-11-01","bid":8},{"week":"2026-11-08","bid":2.1}]}}, 2025)
            self.assertEqual(sheet['E2'].value, 2.1)


if __name__ == "__main__":
    unittest.main()
