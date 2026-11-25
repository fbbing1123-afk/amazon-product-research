"""回归同事填表场景：页名差异、自然空值和自动站点选择。"""
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from shared.excel_inspect import find_source_sheet, infer_product_name
from shared.asin_collection_links import collection_from_workbook, infer_workbook_site, source_sheet
from shared.marketplaces import get_marketplace
from product_research.scripts.extract_source_records import choose_sheet, read_records
from product_research.scripts.fill_custom_template import find_source_sheet as category_source
from competitor_analysis.populate_basics import find_sheet, source_records, rating_summary
from competitor_analysis.collect_basic_facts import read_source


class InputDefaultsTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.path = Path(self.temporary.name) / 'input.xlsx'
        self.book = Workbook()
        self.addCleanup(self.book.close)
        self.book.active.title = '模板示例-市场调研'
        self.competitor = self.book.create_sheet('模板示例-竞品分析')
        self.competitor['A1'] = '4. 竞品分析-US：B012345678'
        self.book.create_sheet('评价数据源').append(['ASIN'])
        self.source = self.book.create_sheet('Water Bottle Organizer-数据源 ')
        self.source.append(['ASIN', '商品详情页链接', '月销量', '价格($)', '上架时间', '评分', '评分数'])
        self.source.append(['B012345678', None, None, None, None, None, None])

    def test_current_readers_accept_different_prefix_and_trailing_space(self):
        original_names = self.book.sheetnames[:]
        for selected in (find_source_sheet(self.book), source_sheet(self.book),
                         choose_sheet(self.book, None), category_source(self.book),
                         find_sheet(self.book, '数据源'), source_sheet(self.book, self.source.title.strip())):
            self.assertIs(selected, self.source)
        self.assertEqual(infer_product_name(self.book), 'Water Bottle Organizer')
        self.assertEqual(self.book.sheetnames, original_names)
        self.book.save(self.path)
        result = read_source(self.path, ['B012345678'])
        self.assertEqual(result['B012345678']['monthly_sales'], [])

    def test_missing_values_remain_missing_and_do_not_drop_product(self):
        rows = read_records(self.source)
        self.assertEqual(len(rows), 1)
        self.assertIsNone(rows[0]['月销量'])
        records = source_records(self.source, ['B012345678'])
        self.assertIsNone(records['B012345678']['评分'])
        self.assertEqual(rating_summary(records['B012345678']), '/')
        self.assertEqual(rating_summary({'评分': 0, '评分数': 0}), '0 / 0')
        self.book.save(self.path)
        self.assertEqual(collection_from_workbook(self.path)['asins'], ['B012345678'])

    def test_source_site_then_competitor_links_then_us(self):
        self.assertEqual(infer_workbook_site(self.book), ('US', 'default_us'))
        self.competitor['C3'] = '商品页面'
        self.competitor['C3'].hyperlink = 'https://www.amazon.co.uk/dp/B012345678'
        self.assertEqual(infer_workbook_site(self.book), ('UK', 'competitor_links'))
        self.book.save(self.path)
        self.assertEqual(collection_from_workbook(self.path)['site'], 'UK')
        self.source['B2'].hyperlink = 'https://www.amazon.de/dp/B012345678'
        self.assertEqual(infer_workbook_site(self.book), ('DE', 'source_product_links'))

    def test_explicit_site_and_configured_postal_codes(self):
        self.book.save(self.path)
        result = collection_from_workbook(self.path, site='法国')
        self.assertEqual((result['site'], result['site_source']), ('FR', 'explicit'))
        for code, postal in {'US':'10001', 'DE':'10115', 'UK':'B100AB', 'CA':'M4Y1M7', 'FR':'75000'}.items():
            self.assertEqual(get_marketplace(code).postal_code, postal)


if __name__ == '__main__':
    unittest.main()
