"""验证同页取数的源优先、ASIN隔离和动态定位，保护既有评价与样式。"""
from pathlib import Path
import tempfile
import unittest
from unittest.mock import AsyncMock, Mock
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from competitor_analysis.prefetch_basics import collect_same_page, refresh_basic_cells


class SamePageTests(unittest.IsolatedAsyncioTestCase):
    def source(self):
        return {key:[{'value':value,'source':'workbook'}] for key,value in
                {'monthly_sales':0,'price':12.5,'listed_at':'2026-01-01','selling_points':'原有卖点'}.items()}

    async def test_existing_source_uses_no_plugin_or_navigation(self):
        page = Mock(url='https://www.amazon.com/dp/B012345678')
        page.locator.return_value = Mock(count=AsyncMock(return_value=0))
        result = await collect_same_page(page,'B012345678',self.source())
        self.assertEqual(result['status'],'complete')
        self.assertEqual(result['fields']['monthly_sales']['value'],0)
        page.goto.assert_not_called()
        page.locator.assert_called_once_with('#dpx-rex-nice-widget-container')

    async def test_missing_price_reads_same_product_plugin(self):
        candidates = self.source()
        candidates['price'] = []
        page = Mock(url='https://www.amazon.com/dp/B012345678',wait_for_timeout=AsyncMock())
        panel = Mock(inner_text=AsyncMock(return_value='ASIN: B012345678\n价格: $19.99\n'))
        page.locator.side_effect = lambda selector: panel if 'quick-view' in selector else Mock(count=AsyncMock(return_value=0))
        result = await collect_same_page(page,'B012345678',candidates)
        self.assertEqual(result['fields']['price']['value'],19.99)
        self.assertEqual(result['fields']['monthly_sales']['source'],'workbook')
        self.assertEqual(panel.inner_text.await_count,2)
        page.goto.assert_not_called()
        page.url='https://www.amazon.com/dp/B087654321'
        with self.assertRaisesRegex(ValueError,'其它 ASIN'):
            await collect_same_page(page,'B012345678',candidates)


class BasicWriterTests(unittest.TestCase):
    def test_dynamic_rows_preserve_other_parts_and_evaluation(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder)
            book=Workbook()
            sheet=book.active
            sheet.title='竞品分析'
            sheet['A1']='B012345678'
            for row,label in [(5,'详情页链接'),(9,'月销量'),(10,'价格($)'),(11,'上架时间')]:
                sheet.cell(row,1,label)
                sheet.cell(row,3).fill=PatternFill('solid',fgColor='FF123456')
            sheet['C5']='https://www.amazon.com/dp/B012345678'
            sheet['C12']='原中文卖点和评价'
            book.create_sheet('评价数据源').append(['原评论'])
            before,after=root/'before.xlsx',root/'after.xlsx'
            book.save(before)
            book.close()
            product={'asin':'B012345678','fields':{k:{'value':v} for k,v in
                    {'monthly_sales':0,'price':19.99,'listed_at':'2026-01-01'}.items()}}
            refresh_basic_cells(before,after,{'B012345678':product})
            checked=load_workbook(after,rich_text=True)
            self.assertEqual([checked.active[c].value for c in ['C9','C10','C11']],[0,19.99,'2026-01-01'])
            self.assertEqual(checked.active['C9'].fill.fgColor.rgb,'FF123456')
            self.assertEqual(checked.active['C12'].value,'原中文卖点和评价')
            checked.close()
            with ZipFile(before) as old,ZipFile(after) as new:
                self.assertEqual({n for n in old.namelist() if old.read(n)!=new.read(n)},{'xl/worksheets/sheet1.xml'})


if __name__=='__main__':
    unittest.main()
