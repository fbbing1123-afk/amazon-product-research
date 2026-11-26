"""检查记录端口的兼容性、各阶段默认值以及切换后的连接目标。"""
import argparse
import json
from pathlib import Path
import runpy
import tempfile
import unittest
from unittest.mock import Mock, patch

from shared import browser_runtime


class BrowserRuntimeTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.config = Path(self.temporary.name) / 'browser-config.json'
        self.override = patch.object(browser_runtime, 'CONFIG_PATH', self.config)
        self.override.start()
        self.addCleanup(self.override.stop)

    def save(self, value):
        self.config.write_text(json.dumps(value), encoding='utf-8-sig')

    def test_old_config_and_new_endpoint(self):
        self.assertEqual(browser_runtime.get_browser_port(), 9223)
        self.save({'profile_dir': 'existing-profile'})
        self.assertEqual(browser_runtime.get_browser_port(), 9223)
        self.save({'port': 19231, 'profile_dir': 'existing-profile'})
        self.assertEqual(browser_runtime.get_cdp_url(), 'http://127.0.0.1:19231')
        self.assertEqual(browser_runtime.get_debugger_address(), '127.0.0.1:19231')
        self.save({'cdp_url': 'http://127.0.0.1:19232'})
        self.assertEqual(browser_runtime.get_browser_port(), 19232)
        for value in (True, 0, 65536, 'bad'):
            self.save({'port': value})
            with self.assertRaises(ValueError):
                browser_runtime.get_browser_port()

    def test_current_cli_defaults_share_recorded_port(self):
        self.save({'port': 19233})
        skill = Path(__file__).resolve().parents[2]
        scripts = {
            'scripts/run_workflow.py': {'keyword_backfill_debugger_address': '127.0.0.1:19233', 'keyword_trends_debug_port': 19233},
            'scripts/run_review_prefetch.py': {'cdp': 'http://127.0.0.1:19233'},
            'scripts/amazon_human_gate.py': {'debugger_address': '127.0.0.1:19233'},
            'scripts/repair_keyword_amazon_results.py': {'port': 19233},
            'modules/competitor_analysis/collect_basic_facts.py': {'cdp': 'http://127.0.0.1:19233'},
            'modules/competitor_analysis/collect_product_images.py': {'cdp': 'http://127.0.0.1:19233'},
            'modules/review_research/download_sellersprite_reviews.py': {'cdp': 'http://127.0.0.1:19233'},
            'modules/product_research/scripts/capture_sellersprite_categories.py': {'cdp': 'http://127.0.0.1:19233'},
            'modules/product_research/scripts/capture_sellersprite_report_charts.py': {'cdp': 'http://127.0.0.1:19233'},
            'modules/keyword_trends/refresh_ppc_bids.py': {'debug_port': 19233},
            'modules/keyword_trends/excel-sellersprite-trends.py': {'debug_port': 19233},
        }

        class Parsed(Exception):
            pass

        for name, defaults in scripts.items():
            with self.subTest(script=name):
                def inspect(parser, *args, **kwargs):
                    for key, expected in defaults.items():
                        self.assertEqual(parser.get_default(key), expected, (name, key))
                    raise Parsed()
                with patch.object(argparse.ArgumentParser, 'parse_args', inspect), \
                     patch.dict('os.environ', {'ECOMMERCE_WORKFLOW_SKIP_VENV_REEXEC': '1'}), \
                     self.assertRaises(Parsed):
                    runpy.run_path(str(skill / name), run_name='__main__')

    def test_preflight_connects_to_selected_port(self):
        from competitor_analysis import preflight_browser as module
        self.save({'port': 19234})
        seller = Mock(url='https://www.sellersprite.com/v2/welcome')
        amazon = Mock(url='https://www.amazon.com/')
        browser = Mock(contexts=[Mock(pages=[seller, amazon])])
        runtime = Mock()
        runtime.chromium.connect_over_cdp.return_value = browser
        manager = Mock(__enter__=Mock(return_value=runtime), __exit__=Mock(return_value=False))
        with patch('playwright.sync_api.sync_playwright', return_value=manager), \
             patch.object(module, 'debugger_ready', return_value=True), \
             patch.object(module.subprocess, 'run') as launch, \
             patch.object(module, 'ensure_sellersprite_logged_in'), \
             patch.object(module, 'preflight_amazon'):
            result = module.run_preflight(debugger_address='127.0.0.1:9223', site='US', postal_code=None)
        launch.assert_called_once()
        runtime.chromium.connect_over_cdp.assert_called_once_with('http://127.0.0.1:19234')
        self.assertEqual(result['debugger_address'], '127.0.0.1:19234')


if __name__ == '__main__':
    unittest.main()
