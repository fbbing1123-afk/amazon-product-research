"""用独立空白页验证 PPC 页面解析；通过环境变量显式启用本机 CDP。"""

import os
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import sellersprite_playwright_runner as runner


@unittest.skipUnless(os.environ.get("KEYWORD_TEST_CDP"), "未启用本机浏览器测试")
class PpcBrowserTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.playwright = runner.require_sync_playwright()().start()
        cls.browser = cls.playwright.chromium.connect_over_cdp(os.environ["KEYWORD_TEST_CDP"], timeout=20000)
        cls.page = cls.browser.contexts[0].new_page()

    @classmethod
    def tearDownClass(cls):
        cls.page.close()
        cls.playwright.stop()

    def setUp(self):
        self.page.set_content('''<div id="search-anlysis-drawer-echarts-trends">
          <label style="display:block"><input type="checkbox">查看历史月数据</label>
          <div id="ppc-trends" style="width:600px;height:300px"></div></div>''')

    def set_chart(self, labels, values):
        self.page.evaluate('''([labels, values]) => {
          document.getElementById('ppc-trends').__vue__ = {chart: {getOption: () => ({
            xAxis:[{data:labels}],series:[{name:'最大值',data:values}]
          })}};
        }''', [labels, values])

    def test_monthly_checkbox_is_idempotent(self):
        self.assertTrue(runner.click_ppc_monthly_history(self.page))
        self.assertTrue(self.page.locator('input').is_checked())
        self.assertTrue(runner.click_ppc_monthly_history(self.page))
        self.assertTrue(self.page.locator('input').is_checked())

    def test_week_dates_do_not_become_monthly_points(self):
        self.set_chart(["2026-10-01", "2026第46周(11/08~11/13)"], [9, 8])
        result = runner.extract_ppc_data(self.page)
        self.assertEqual(result["ppcMonthly"], [])
        self.assertEqual(len(result["ppcWeekly"]), 2)

    def test_weekly_mode_unchecks_monthly_and_stays_unchecked(self):
        self.page.locator('input').check()
        self.assertTrue(runner.set_ppc_history_mode(self.page, monthly=False))
        self.assertFalse(self.page.locator('input').is_checked())
        self.assertTrue(runner.set_ppc_history_mode(self.page, monthly=False))
        self.assertFalse(self.page.locator('input').is_checked())

    def test_latest_null_week_is_preserved(self):
        self.set_chart(["2026第45周(11/01~11/06)", "2026第46周(11/08~11/13)"], [4, None])
        result = runner.extract_ppc_data(self.page)
        self.assertEqual(result["ppcWeekly"][-1], {"week":"2026-11-08", "weekEnd":"2026-11-13", "bid":None})
        self.assertEqual(result["ppcMonthly"], [])

    def test_week_year_handles_first_week_and_year_end_week(self):
        self.set_chart(["2026第01周(12/28~01/02)", "2025第53周(12/28~01/02)", "2026第46周(11/08~11/13)"], [9, 8, 2])
        result = runner.extract_ppc_data(self.page)
        self.assertEqual(result["ppcWeekly"][0], {"week":"2025-12-28", "weekEnd":"2026-01-02", "bid":9})
        self.assertEqual(result["ppcWeekly"][1], {"week":"2025-12-28", "weekEnd":"2026-01-02", "bid":8})
        from keyword_rows import latest_weekly_bid
        self.assertEqual(latest_weekly_bid(result["ppcWeekly"]), ("2026-11-08~2026-11-13", 2))

    def test_latest_null_month_is_preserved(self):
        self.set_chart(["2026-09", "2026-10"], [4, None])
        result = runner.extract_ppc_data(self.page)
        self.assertEqual(result["ppcMonthly"], [{"month":"2026-09", "bid":4}, {"month":"2026-10", "bid":None}])
        self.assertEqual(result["ppcWeekly"], [])


if __name__ == "__main__":
    unittest.main()
