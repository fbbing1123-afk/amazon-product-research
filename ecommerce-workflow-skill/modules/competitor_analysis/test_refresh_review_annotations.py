"""核验分析纠错只改同批目标证据列，保护原评论、其它商品和人工修改。"""
from copy import deepcopy
from pathlib import Path
import tempfile
import unittest
from zipfile import ZipFile
from openpyxl import Workbook,load_workbook
from competitor_analysis.sellersprite_review_analysis import make_packet
from competitor_analysis.merge_review_sources import merge_sources
from competitor_analysis.refresh_review_annotations import refresh_annotations


class RefreshAnnotationsTests(unittest.TestCase):
    def setUp(self):
        self.temporary=tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root=Path(self.temporary.name)
        self.asin='B012345678'
        self.source=self.root/'source.xlsx'
        book=Workbook();sheet=book.active;sheet.title=self.asin+'-Reviews(3)'
        sheet.append(['ASIN','标题','内容','VP评论','Vine Voice评论','星级','评论链接','评论时间','赞同数'])
        for number,(rating,vp) in enumerate([(5,'Y'),(1,'Y'),(1,'')],1):
            sheet.append([None,'标题','真实评价',vp,'',rating,f'https://www.amazon.com/gp/customer-reviews/R{number}','2026-11-25',0])
        book.save(self.source);book.close()
        packet=make_packet(self.source,self.asin)
        self.previous={k:packet[k] for k in ('schema','asin','packet_digest','source_file','source_sha256')}
        self.previous.update(advantages=[{'text':'使用方便','review_numbers':[1]}],disadvantages=[],empty_reasons={'disadvantages':'没有具体原因'})
        self.updated=deepcopy(self.previous)
        self.updated['disadvantages']=[{'text':'压花效果不佳','review_numbers':[1]}]
        book=Workbook();book.active.title='竞品分析';book.active['A1']=self.asin
        book.active['C2']='https://www.amazon.com/dp/'+self.asin
        base=self.root/'base.xlsx';book.save(base);book.close()
        self.merged,self.output=self.root/'merged.xlsx',self.root/'updated.xlsx'
        merge_sources(base,self.merged,[self.previous])

    def test_updates_only_source_annotations_and_preserves_unselected_blank(self):
        result=refresh_annotations(self.merged,self.output,self.previous,self.updated)
        book=load_workbook(self.output,rich_text=True);sheet=book['评价数据源']
        headers={c.value:c.column for c in sheet[1]}
        self.assertEqual(sheet.cell(3,headers['差评原因']).value,'压花效果不佳')
        self.assertEqual(sheet.cell(3,headers['分析说明']).value,'已分析')
        self.assertIsNone(sheet.cell(4,headers['分析说明']).value)
        self.assertIsNone(sheet.cell(3,headers['ASIN']).value)
        self.assertEqual(sheet.cell(3,headers['赞同数']).value,0)
        book.close()
        self.assertEqual(len(result['updated_cells']),2)
        with ZipFile(self.merged) as old,ZipFile(self.output) as new:
            self.assertEqual({n for n in old.namelist() if old.read(n)!=new.read(n)},{'xl/worksheets/sheet2.xml'})

    def test_manual_note_and_raw_zero_change_are_rejected(self):
        for field,value in [('分析说明','人工备注'),('赞同数',None)]:
            with self.subTest(field=field):
                book=load_workbook(self.merged,rich_text=True);sheet=book['评价数据源']
                headers={c.value:c.column for c in sheet[1]}
                sheet.cell(3,headers[field]).value=value
                modified=self.root/(field+'.xlsx');book.save(modified);book.close()
                with self.assertRaises(ValueError):
                    refresh_annotations(modified,self.output,self.previous,self.updated)
                self.assertFalse(self.output.exists())

    def test_unchanged_short_row_keeps_its_height(self):
        book=load_workbook(self.merged,rich_text=True)
        book['评价数据源'].row_dimensions[2].height=15
        book.save(self.merged)
        book.close()
        refresh_annotations(self.merged,self.output,self.previous,self.updated)
        book=load_workbook(self.output,rich_text=True)
        self.assertEqual(book['评价数据源'].row_dimensions[2].height,15)
        book.close()


if __name__=='__main__':
    unittest.main()
