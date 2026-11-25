"""验证每商品一个评价单元格，保留长短结论、其它原生部件与后方位置。"""
from pathlib import Path
import tempfile
import unittest
from zipfile import ZipFile

from openpyxl import Workbook,load_workbook
from openpyxl.styles import Alignment,PatternFill
from competitor_analysis.merge_review_cells import merge_review_cells


class MergeReviewCellsTests(unittest.TestCase):
    def setUp(self):
        self.temporary=tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root=Path(self.temporary.name)
        self.before,self.after=self.root/'before.xlsx',self.root/'after.xlsx'
        book=Workbook()
        sheet=book.active
        sheet.title='示例-竞品分析'
        sheet['A1']='B012345678/B087654321'
        sheet['A3']='详情页链接'
        sheet['C3']='https://www.amazon.com/dp/B012345678'
        sheet['D3']='https://www.amazon.com/dp/B087654321'
        sheet['A8'],sheet['B8'],sheet['B9']='评价分析','优点','缺点'
        sheet.merge_cells('A8:A11')
        sheet.merge_cells('B9:B11')
        sheet['C8'],sheet['D8']='操作方便','做工好'
        sheet['C9'],sheet['C10'],sheet['C11']='缺少配件（50%）','板材变形（25%）','包装破损（25%）'
        sheet['D9']='/'
        sheet['A12'],sheet['C12']='总结及改进方式','原总结'
        for row in range(8,13):
            sheet.row_dimensions[row].height=99
            for col in (3,4):
                sheet.cell(row,col).alignment=Alignment(wrap_text=True,vertical='center')
                sheet.cell(row,col).fill=PatternFill('solid',fgColor='FFAAFF')
        book.create_sheet('评价数据源').append(['原始评论和空白分析不变'])
        book.save(self.before)
        book.close()
        with ZipFile(self.before,'a') as archive:
            archive.writestr('xl/slicers/preservation.xml',b'preserve-native-part')

    def test_merges_full_text_and_slash_without_moving_summary_or_other_parts(self):
        original=self.before.read_bytes()
        result=merge_review_cells(self.before,self.after)
        self.assertEqual(result['merged_regions'],['C9:C11','D9:D11'])
        book=load_workbook(self.after,rich_text=True)
        sheet=book.active
        self.assertEqual(sheet['C9'].value,'缺少配件（50%）\n板材变形（25%）\n包装破损（25%）')
        self.assertEqual(sheet['D9'].value,'/')
        self.assertIsNone(sheet['C10'].value)
        self.assertEqual(sheet['C12'].value,'原总结')
        self.assertEqual(sheet.row_dimensions[12].height,99)
        self.assertEqual(sheet['C9'].fill.fgColor.rgb,'00FFAAFF')
        book.close()
        self.assertEqual(self.before.read_bytes(),original)
        with ZipFile(self.before) as old,ZipFile(self.after) as new:
            self.assertEqual({n for n in old.namelist() if old.read(n)!=new.read(n)},{'xl/worksheets/sheet1.xml'})
        second=self.root/'second.xlsx'
        merge_review_cells(self.after,second)
        book=load_workbook(second,rich_text=True)
        self.assertEqual(book.active['C9'].value.count('缺少配件'),1)
        self.assertEqual(len(book.active.merged_cells.ranges),4)
        book.close()

    def test_rejects_cross_product_merge_and_overwrite(self):
        with self.assertRaisesRegex(ValueError,'独立输出'):
            merge_review_cells(self.before,self.before)
        book=load_workbook(self.before,rich_text=True)
        book.active.merge_cells('C9:D9')
        book.save(self.root/'cross.xlsx')
        book.close()
        with self.assertRaisesRegex(ValueError,'跨商品'):
            merge_review_cells(self.root/'cross.xlsx',self.after)
        self.assertFalse(self.after.exists())


if __name__=='__main__':
    unittest.main()
