"""核验跨站链接、200条以上完整性、商品源选择及不改原文件。"""
import hashlib
from pathlib import Path
import tempfile
import unittest
from unittest.mock import Mock, patch
from urllib.parse import parse_qs, urlparse

from openpyxl import Workbook

from shared.asin_collection_links import build_collection_url, collection_from_workbook, normalize_asins, normalize_site


class CollectionLinkTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)

    def workbook(self, domain='amazon.co.jp', rows=None):
        book = Workbook()
        book.active.title = '评价数据源'
        book.active.append(['ASIN'])
        book.active.append(['B099999999'])
        source = book.create_sheet('产品-数据源')
        source.append(['父ASIN', 'ASIN', '商品详情页链接', '月销量', '子体销量'])
        for asin, sales in (rows or [('B012345678', None), ('B087654321', 0)]):
            source.append(['B000000000', asin, '商品页面', 0, sales])
            if domain:
                source.cell(source.max_row, 3).hyperlink = f'https://www.{domain}/dp/{asin}'
        path = self.root / 'source.xlsx'
        book.save(path)
        book.close()
        return path

    def test_more_than_200_round_trip_keeps_all_asins_in_order(self):
        asins = ['B0' + f'{index:08d}' for index in range(229)]
        url = build_collection_url([None, *asins, asins[3].lower(), ''], 'US')
        self.assertEqual(parse_qs(urlparse(url).query)['rh'][0], 'p_78:' + '|'.join(asins))
        self.assertEqual(len(url.split('|')), 229)
        self.assertNotIn('%25', url)

    def test_country_domains_and_aliases(self):
        expected = {'US':'amazon.com','JP':'amazon.co.jp','UK':'amazon.co.uk','DE':'amazon.de',
                    'FR':'amazon.fr','IT':'amazon.it','ES':'amazon.es','CA':'amazon.ca',
                    'IN':'amazon.in','MX':'amazon.com.mx','BR':'amazon.com.br','AU':'amazon.com.au'}
        for site, domain in expected.items():
            with self.subTest(site=site):
                self.assertEqual(build_collection_url(['B012345678', 'B087654321'], site),
                                 f'https://www.{domain}/s?rh=p_78:B012345678|B087654321')
        self.assertEqual(normalize_site('GB'), 'UK')
        self.assertEqual(normalize_site('日本站'), 'JP')
        self.assertEqual(normalize_site('澳大利亚'), 'AU')

    def test_invalid_values_and_unknown_site_fail_instead_of_silent_skip(self):
        for values in ([], [' '], ['B012345678', 'BAD'], ['B012345678|B087654321']):
            with self.subTest(values=values), self.assertRaises(ValueError):
                normalize_asins(values)
        with self.assertRaises(ValueError):
            build_collection_url(['B012345678'], 'UNKNOWN')

    def test_source_asins_include_zero_missing_sales_and_use_hyperlink_site(self):
        source = self.workbook(rows=[('b012345678', None), ('B087654321', 0), ('B012345678', 40)])
        fingerprint = hashlib.sha256(source.read_bytes()).hexdigest()
        result = collection_from_workbook(source)
        self.assertEqual(result['site'], 'JP')
        self.assertEqual(result['asins'], ['B012345678', 'B087654321'])
        self.assertEqual(result['source_asin_rows'], 3)
        self.assertEqual(result['duplicate_rows'], 1)
        self.assertNotIn('B000000000', result['url'])
        self.assertNotIn('B099999999', result['url'])
        self.assertEqual(hashlib.sha256(source.read_bytes()).hexdigest(), fingerprint)

    def test_unknown_and_conflicting_workbook_site_are_explicit(self):
        with self.assertRaisesRegex(ValueError, '不一致'):
            collection_from_workbook(self.workbook(), site='US')
        source = self.workbook(domain=None)
        self.assertEqual(collection_from_workbook(source)['site'], 'US')
        self.assertEqual(collection_from_workbook(source)['site_source'], 'default_us')
        self.assertEqual(collection_from_workbook(source, site='DE')['site'], 'DE')

    def test_mixed_sites_or_ambiguous_source_are_rejected(self):
        book = Workbook()
        source = book.active
        source.title = '商品-数据源'
        source.append(['ASIN', '商品详情页链接'])
        source.append(['B012345678', 'https://www.amazon.com/dp/B012345678'])
        source.append(['B087654321', 'https://www.amazon.de/dp/B087654321'])
        path = self.root / 'mixed.xlsx'
        book.save(path)
        with self.assertRaisesRegex(ValueError, '多个站点'):
            collection_from_workbook(path)
        book.create_sheet('另一商品-数据源').append(['ASIN'])
        book.save(path)
        book.close()
        with self.assertRaisesRegex(ValueError, '--sheet'):
            collection_from_workbook(path)

    def test_invalid_nonempty_asin_is_reported_with_source_cell(self):
        source = self.workbook(rows=[('NOT-ASIN', 0)])
        with self.assertRaisesRegex(ValueError, 'B2'):
            collection_from_workbook(source)


class PreflightTests(unittest.TestCase):
    def test_preflight_no_longer_opens_link_builder(self):
        from competitor_analysis import preflight_browser as module
        seller = Mock(url='https://www.sellersprite.com/v3/keyword-miner/')
        amazon = Mock(url='https://www.amazon.com/')
        browser = Mock(contexts=[Mock(pages=[seller, amazon])])
        runtime = Mock()
        runtime.chromium.connect_over_cdp.return_value = browser
        manager = Mock()
        manager.__enter__ = Mock(return_value=runtime)
        manager.__exit__ = Mock(return_value=False)
        with patch('playwright.sync_api.sync_playwright', return_value=manager), \
             patch.object(module, 'debugger_ready', return_value=True), \
             patch.object(module, 'start_project_browser', return_value='127.0.0.1:9223'), \
             patch.object(module, 'ensure_sellersprite_logged_in') as login, \
             patch.object(module, 'preflight_amazon'):
            result = module.run_preflight(debugger_address='127.0.0.1:9223', site='US', postal_code='10001')
        self.assertEqual(result['status'], 'ready')
        login.assert_called_once_with(seller, 'https://www.sellersprite.com/v2/welcome')
        seller.goto.assert_not_called()


if __name__ == '__main__':
    unittest.main()
