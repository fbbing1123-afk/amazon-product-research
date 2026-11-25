"""验证批量评价扩行保留完整结论、后方格式及源数据隔离。"""
from copy import copy
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment, PatternFill
from PIL import Image as PillowImage

from competitor_analysis.fill_review_batch import fill_batch
from competitor_analysis.sellersprite_review_analysis import make_packet


class FillReviewBatchTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.asin = 'B012345678'
        self.source = self.root / 'reviews.xlsx'
        book = Workbook()
        sheet = book.active
        sheet.title = self.asin + '-Reviews(2)'
        sheet.append(['ASIN','标题','内容','VP评论','Vine Voice评论','星级','评论链接','评论时间','赞同数','标题(翻译)','内容(翻译)'])
        for number, rating in enumerate((5, 1), start=1):
            sheet.append([self.asin,'真实标题','真实评价','Y','',rating,
                          f'https://www.amazon.com/gp/customer-reviews/R{number}',
                          '2026-11-24',number,'中文标题','中文评价'])
        book.save(self.source)
        book.close()
        packet = make_packet(self.source, self.asin)
        self.analysis = {key:packet[key] for key in ('schema','asin','packet_digest','source_sha256','source_file')}
        self.analysis['advantages'] = [{'text':'操作方便','review_numbers':[1]}]
        self.analysis['disadvantages'] = [{'text':f'问题{number}需改善','review_numbers':[1]} for number in range(12)]
        self.template, self.output = self.root / 'template.xlsx', self.root / 'result.xlsx'
        book = Workbook()
        sheet = book.active
        sheet.title = '示例-竞品分析'
        sheet['A1'] = self.asin
        sheet['A2'], sheet['C2'] = '详情页链接', 'https://www.amazon.com/dp/' + self.asin
        sheet['A5'], sheet['B5'], sheet['B6'] = '评价分析', '优点', '缺点'
        sheet.merge_cells('A5:A6')
        sheet['A7'], sheet['C7'] = '总结及改进方式', '旧总结'
        sheet.merge_cells('A7:B7')
        sheet.merge_cells('C7:D7')
        sheet.row_dimensions[5].height = 100
        sheet.row_dimensions[6].height = 80
        sheet.row_dimensions[7].height = 198
        sheet.column_dimensions['C'].width = 30
        for coord in ('C5','C6','C7'):
            sheet[coord].alignment = Alignment(wrap_text=True,vertical='center')
            sheet[coord].fill = PatternFill('solid',fgColor='FFCCDD')
        self.summary_style = copy(sheet['C7']._style)
        picture = self.root / 'picture.png'
        PillowImage.new('RGB',(20,20),'blue').save(picture)
        sheet.add_image(Image(picture),'C8')
        book.save(self.template)
        book.close()

    def test_long_reviews_keep_every_point_and_move_summary_and_image(self):
        # 该测试聚焦扩行；完整基础门禁由真实最终报告另行验证。
        with patch('competitor_analysis.fill_review_batch.validate_competitor_review_analysis'):
            result = fill_batch(self.template,self.output,[self.analysis],'本批改进')
        book = load_workbook(self.output,rich_text=True)
        sheet = book.active
        summary = result['summary_row']
        self.assertGreater(summary,7)
        self.assertEqual('本批改进',sheet.cell(summary,3).value)
        self.assertEqual(self.summary_style,sheet.cell(summary,3)._style)
        self.assertEqual(198,sheet.row_dimensions[summary].height)
        self.assertIn(f'C{summary}:D{summary}',sheet.merged_cells)
        actual = '\n'.join(sheet.cell(row,3).value or '' for row in range(6,summary))
        self.assertEqual('\n'.join(f'问题{number}需改善（100%）' for number in range(12)),actual)
        self.assertTrue(all(sheet.row_dimensions[row].height == 80 for row in range(6,summary)))
        self.assertEqual(summary,sheet._images[0].anchor._from.row)
        self.assertEqual(2,result['review_source_rows'])
        self.assertEqual('评价数据源',result['review_source_sheet'])
        self.assertEqual(self.asin,book['评价数据源']['A2'].value)
        book.close()

    def test_changed_source_rejects_old_analysis_without_output(self):
        book = load_workbook(self.source)
        book.active['C3'] = '源评论已经变化'
        book.save(self.source)
        book.close()
        with self.assertRaisesRegex(ValueError,'另一批次'):
            fill_batch(self.template,self.output,[self.analysis],'本批改进')
        self.assertFalse(self.output.exists())

    def test_wrong_product_rejects_batch_without_output(self):
        book = load_workbook(self.template)
        book.active['C2'] = 'https://www.amazon.com/dp/B087654321'
        book.save(self.template)
        book.close()
        with self.assertRaisesRegex(ValueError,'一一对应'):
            fill_batch(self.template,self.output,[self.analysis],'本批改进')
        self.assertFalse(self.output.exists())


if __name__ == '__main__':
    unittest.main()
