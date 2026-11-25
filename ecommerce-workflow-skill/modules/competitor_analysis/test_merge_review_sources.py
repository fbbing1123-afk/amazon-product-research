"""验证合并评论的商品归属、原文保留、证据对齐和原报告部件保护。"""
from pathlib import Path
import tempfile
import unittest
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook

from competitor_analysis.merge_review_sources import merge_sources, SHEET_NAME, ASIN_HEADER
from competitor_analysis.sellersprite_review_analysis import make_packet


class MergeReviewSourcesTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.asins = ['B012345678', 'B087654321']
        self.headers = ['ASIN', '标题', '内容', 'VP评论', 'Vine Voice评论', '星级', '评论链接', '评论时间', '赞同数']
        self.analyses = []
        for index, asin in enumerate(self.asins):
            source = self.root / f'{asin}.xlsx'
            book = Workbook()
            sheet = book.active
            sheet.title = f'{asin}-Reviews(3)'
            headers = self.headers if index == 0 else list(reversed(self.headers)) + ['新源字段']
            sheet.append(headers)
            for number, (rating, trusted, row_asin) in enumerate(((1, 'Y', None), (5, 'Y', 'B099999999'), (2, '', asin)), 1):
                record = dict(zip(self.headers, [row_asin, f'标题{number}', '=1+1' if number == 1 else f'{asin}原文{number}',
                    trusted, '', rating, f'https://www.amazon.com/gp/customer-reviews/RSHARED{number}', '2026-11-25', 10-number]))
                record['新源字段'] = '额外原始值'
                sheet.append([record.get(h) for h in headers])
                # 模拟以等号开头的评论正文，源文件也明确存成文字。
                sheet.cell(number+1, headers.index('内容')+1).data_type = 's'
            book.save(source)
            book.close()
            packet = make_packet(source, asin)
            analysis = {k: packet[k] for k in ('schema','asin','packet_digest','source_sha256','source_file')}
            analysis.update(advantages=[{'text': f'优点{index}', 'review_numbers': [1]}],
                            disadvantages=[{'text': f'问题{index}', 'review_numbers': [1]},
                                           {'text': '配件缺失', 'review_numbers': [1]}])
            self.analyses.append(analysis)
        self.input, self.output = self.root/'report.xlsx', self.root/'merged.xlsx'
        book = Workbook()
        sheet = book.active
        sheet.title = '示例-竞品分析'
        sheet['A1'] = '竞品报告'
        sheet['A2'] = '详情页链接'
        for col, asin in enumerate(reversed(self.asins), 3):
            sheet.cell(2, col, 'https://www.amazon.com/dp/' + asin)
        book.create_sheet('数据源')['A1'] = '原产品数据'
        book.save(self.input)
        book.close()
        with ZipFile(self.input, 'a') as archive:
            archive.writestr('xl/drawings/native-unchanged.xml', b'opaque-native-object')

    def test_all_reviews_group_by_target_keep_blank_variant_asins_and_annotations(self):
        result = merge_sources(self.input, self.output, self.analyses)
        self.assertEqual(6, result['reviews'])
        book = load_workbook(self.output)
        sheet = book[SHEET_NAME]
        headers = [c.value for c in sheet[1]]
        self.assertEqual(ASIN_HEADER, headers[0])
        self.assertEqual(['优点归纳','差评原因','分析说明'], headers[-3:])
        self.assertNotIn('分析入选', headers)
        self.assertNotIn('分析编号', headers)
        self.assertEqual([self.asins[1]]*3+[self.asins[0]]*3, [sheet.cell(r,1).value for r in range(2,8)])
        for start, index in ((2,1),(5,0)):
            value = lambda row, header: sheet.cell(row, headers.index(header)+1).value
            self.assertIsNone(value(start,'ASIN'))
            self.assertEqual('B099999999', value(start+1,'ASIN'))
            self.assertEqual('=1+1', value(start,'内容'))
            self.assertEqual('s', sheet.cell(start,headers.index('内容')+1).data_type)
            self.assertEqual(f'问题{index}\n配件缺失', value(start,'差评原因'))
            self.assertNotIn('%', value(start,'差评原因'))
            self.assertEqual(f'优点{index}', value(start+1,'优点归纳'))
            self.assertTrue(all(value(start+2,h) is None for h in headers[-3:]))
            self.assertEqual('https://www.amazon.com/gp/customer-reviews/RSHARED1', sheet.cell(start,headers.index('评论链接')+1).hyperlink.target)
        self.assertEqual('B2',sheet.freeze_panes)
        self.assertEqual(f'A1:N7',sheet.auto_filter.ref)
        book.close()

    def test_original_zip_parts_and_source_files_preserved(self):
        originals = {a['source_file']:Path(a['source_file']).read_bytes() for a in self.analyses}
        before = self.input.read_bytes()
        merge_sources(self.input,self.output,self.analyses)
        self.assertEqual(before,self.input.read_bytes())
        self.assertTrue(all(Path(p).read_bytes()==data for p,data in originals.items()))
        with ZipFile(self.input) as old, ZipFile(self.output) as new:
            changed = {p for p in old.namelist() if old.read(p)!=new.read(p)}
            self.assertEqual({'xl/workbook.xml','xl/_rels/workbook.xml.rels','xl/styles.xml','[Content_Types].xml'},changed)
            self.assertEqual(2,len(set(new.namelist())-set(old.namelist())))

    def test_missing_duplicate_or_wrong_product_rejected_before_writing(self):
        for analyses in (self.analyses[:1], [*self.analyses,self.analyses[0]]):
            with self.assertRaisesRegex(ValueError,'一一对应'):
                merge_sources(self.input,self.output,analyses)
            self.assertFalse(self.output.exists())

    def test_changed_source_rejects_stale_evidence(self):
        source = Path(self.analyses[0]['source_file'])
        book = load_workbook(source)
        book.active['C2'] = '新的评论'
        book.save(source)
        book.close()
        with self.assertRaisesRegex(ValueError,'另一批次'):
            merge_sources(self.input,self.output,self.analyses)
        self.assertFalse(self.output.exists())

    def test_existing_output_or_source_sheet_not_silently_overwritten(self):
        merge_sources(self.input,self.output,self.analyses)
        with self.assertRaisesRegex(ValueError,'独立文件'):
            merge_sources(self.input,self.output,self.analyses)
        with self.assertRaisesRegex(ValueError,'已有评价数据源'):
            merge_sources(self.output,self.root/'again.xlsx',self.analyses)
        self.assertFalse((self.root/'again.xlsx').exists())


if __name__ == '__main__':
    unittest.main()
