import importlib.util
import sys
import unittest
from pathlib import Path


SCRIPT_PATH = Path(__file__).with_name("excel-sellersprite-trends.py")
SPEC = importlib.util.spec_from_file_location("keyword_trends_script", SCRIPT_PATH)
assert SPEC and SPEC.loader
MODULE = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)


class SeasonalityWindowTests(unittest.TestCase):
    def test_keeps_a_complete_preferred_year(self):
        points = [MODULE.MonthlyTrendPoint(f"2025-{month:02d}", 100 + month) for month in range(1, 13)]
        points.append(MODULE.MonthlyTrendPoint("2026-01", 999))

        selected, ppc_year, mode = MODULE.select_seasonality_points(points, 2025)

        self.assertEqual(mode, "complete-year")
        self.assertEqual(ppc_year, 2025)
        self.assertEqual(len(selected), 12)
        self.assertEqual(selected[-1].month, "2025-12")

    def test_uses_all_data_since_entry_for_an_incomplete_year(self):
        points = [MODULE.MonthlyTrendPoint(f"2025-{month:02d}", 100 + month) for month in range(2, 13)]
        points.extend([
            MODULE.MonthlyTrendPoint("2026-01", 300),
            MODULE.MonthlyTrendPoint("2026-02", 500),
        ])

        selected, ppc_year, mode = MODULE.select_seasonality_points(points, 2025)
        result = MODULE.analyze_seasonality(selected)

        self.assertEqual(mode, "since-keyword-entry")
        self.assertEqual(ppc_year, 2026)
        self.assertEqual(selected[0].month, "2025-02")
        self.assertEqual(result.low_search_volume, 102)
        self.assertEqual(result.peak_search_volume, 500)

    def test_ignores_zero_and_one_when_finding_keyword_entry(self):
        points = [
            MODULE.MonthlyTrendPoint("2025-01", 1),
            MODULE.MonthlyTrendPoint("2025-02", 0),
            MODULE.MonthlyTrendPoint("2025-03", 25),
            MODULE.MonthlyTrendPoint("2025-04", 40),
        ]

        selected, _ppc_year, mode = MODULE.select_seasonality_points(points, 2025)

        self.assertEqual(mode, "since-keyword-entry")
        self.assertEqual([point.month for point in selected], ["2025-03", "2025-04"])

    def test_returns_no_bid_when_only_zero_ppc_data_exists(self):
        zero_weekly = [{"week": "2026-05-10", "weekEnd": "2026-05-16", "bid": 0}]
        zero_monthly = [{"month": "2026-05", "bid": 0}]

        bid = MODULE.max_bid_for_peak_period(zero_weekly, zero_monthly, 2026, [5])

        self.assertIsNone(bid)

    def test_keeps_positive_ppc_bid_when_zero_is_also_returned(self):
        weekly = [
            {"week": "2026-05-03", "weekEnd": "2026-05-09", "bid": 0},
            {"week": "2026-05-10", "weekEnd": "2026-05-16", "bid": 1.35},
        ]

        bid = MODULE.max_bid_for_peak_period(weekly, [], 2026, [5])

        self.assertEqual(bid, 1.35)


if __name__ == "__main__":
    unittest.main()
