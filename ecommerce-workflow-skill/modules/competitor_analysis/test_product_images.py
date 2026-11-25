"""核验高清主图归属、数据源缺失回退及ZIP写图的原生部件保护。"""
import hashlib
from io import BytesIO
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import AsyncMock,Mock
from zipfile import ZipFile

from PIL import Image as PillowImage
from openpyxl import Workbook,load_workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import PatternFill

from competitor_analysis.collect_product_images import image_candidates,validate_asset,selected_asin,collect_main_image
from competitor_analysis.refresh_product_images import refresh_product_images
from competitor_analysis.populate_basics import source_records,populate


class ProductImageTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root=Path(self.temp.name)
        self.asin='B012345678'
        image=PillowImage.new('RGB',(1200,900),'white')
        image.paste('blue',(200,150,1000,750))
        image.save(self.root/'main.png')
        image.close()
        payload=(self.root/'main.png').read_bytes()
        self.asset={'status':'ok','asin':self.asin,'observed_asin':self.asin,
            'page_url':'https://www.amazon.com/dp/'+self.asin,
            'image_url':'https://m.media-amazon.com/images/I/main.png',
            'image_path':str(self.root/'main.png'),'sha256':hashlib.sha256(payload).hexdigest(),
            'width':1200,'height':900,'rating':4.2,'review_count':43}

    def workbook(self,with_image=False):
        book=Workbook()
        sheet=book.active
        sheet.title='测试-竞品分析'
        sheet['A1']=self.asin
        sheet['A6']='产品图'
        sheet['A7']='详情页链接'
        sheet['C7']='https://www.amazon.com/dp/'+self.asin
        sheet['C12']='原有中文总结'
        sheet['C12'].fill=PatternFill('solid',fgColor='FF123456')
        sheet.row_dimensions[6].height=200
        sheet.column_dimensions['C'].width=30
        sheet.merge_cells('C12:F12')
        # 故意不建商品数据源，主图模块应只依赖竞品链接和网页证据。
        book.create_sheet('评价数据源').append(['原评论'])
        if with_image:
            sheet.add_image(Image(BytesIO((self.root/'main.png').read_bytes())),'C6')
        source=self.root/('existing.xlsx' if with_image else 'empty.xlsx')
        book.save(source)
        book.close()
        return source

    def test_observed_hires_first_and_reject_foreign_hosts(self):
        high='https://m.media-amazon.com/images/W/experiment/images/I/high.jpg'
        low='https://m.media-amazon.com/images/I/low.jpg'
        observation={'hires':high,'dynamic':json.dumps({low:[200,200],'https://evil.example/images/I/a.jpg':[3000,3000]}),'src':low}
        self.assertEqual(image_candidates(observation),[high,low])

    def test_wrong_product_and_changed_file_are_rejected(self):
        for key,value in [('asin','B087654321'),('observed_asin','B087654321'),('sha256','wrong'),('width',200)]:
            with self.subTest(field=key):
                asset={**self.asset,key:value}
                with self.assertRaises(ValueError):
                    validate_asset(asset,self.asin)

    def test_missing_source_record_uses_verified_page_rating(self):
        book=Workbook()
        book.active.append(['ASIN','商品详情页链接','评分','评分数'])
        records=source_records(book.active,[self.asin],{self.asin:{'main_image':self.asset}})
        self.assertEqual(records[self.asin]['评分数'],43)
        self.assertEqual(records[self.asin]['评分'],4.2)
        with self.assertRaises(ValueError):
            source_records(book.active,[self.asin],{self.asin:{'main_image':{**self.asset,'review_count':None}}})
        book.close()

    def test_create_and_replace_without_source_or_resampling(self):
        for existing in (False,True):
            for mode in ('fit','original'):
                with self.subTest(existing=existing,mode=mode):
                    source=self.workbook(existing)
                    output=self.root/f'{existing}-{mode}.xlsx'
                    result=refresh_product_images(source,output,{self.asin:self.asset},display_mode=mode)
                    book=load_workbook(output,rich_text=True)
                    sheet=book['测试-竞品分析']
                    self.assertEqual(sheet['C12'].value,'原有中文总结')
                    self.assertEqual(sheet['C12'].fill.fgColor.rgb,'FF123456')
                    self.assertIn('C12:F12',sheet.merged_cells)
                    self.assertEqual(sheet.row_dimensions[6].height,200)
                    self.assertEqual(len(sheet._images),1)
                    image=sheet._images[0]
                    self.assertEqual(image._data(),(self.root/'main.png').read_bytes())
                    self.assertEqual((image.anchor._from.row,image.anchor._from.col),(5,2))
                    self.assertEqual(result['images'][self.asin]['pixels'],[1200,900])
                    if mode=='original':
                        self.assertEqual(result['images'][self.asin]['display'],(1200,900))
                    else:
                        self.assertLess(result['images'][self.asin]['display'][0],250)
                    book.close()
                    with ZipFile(source) as old,ZipFile(output) as new:
                        for name in old.namelist():
                            if name in ('[Content_Types].xml','xl/worksheets/sheet1.xml','xl/worksheets/_rels/sheet1.xml.rels') or name.startswith('xl/drawings/'):
                                continue
                            self.assertEqual(new.read(name),old.read(name),name)
                        if existing:
                            self.assertEqual(new.read('xl/worksheets/sheet1.xml'),old.read('xl/worksheets/sheet1.xml'))

    def test_wrong_asset_does_not_create_output(self):
        source=self.workbook()
        output=self.root/'failed.xlsx'
        with self.assertRaises(ValueError):
            refresh_product_images(source,output,{self.asin:{**self.asset,'asin':'B087654321'}})
        self.assertFalse(output.exists())

    def test_populate_missing_asin_uses_web_image_and_rating(self):
        book=Workbook()
        sheet=book.active
        sheet.title='测试-竞品分析'
        sheet['A1']=self.asin
        for row,label in [(2,'产品图'),(3,'详情页链接'),(4,'评分/评论数'),(5,'产品卖点'),
                          (6,'对比维度'),(7,'评价分析'),(9,'总结及改进方式')]:
            sheet.cell(row,1,label)
        sheet.merge_cells('A1:G1')
        sheet.merge_cells('C9:G9')
        sheet.row_dimensions[2].height=100
        sheet.column_dimensions['C'].width=30
        source=book.create_sheet('测试-数据源')
        source.append(['ASIN','商品详情页链接','评分','评分数'])
        path=self.root/'missing-source.xlsx'
        book.save(path)
        book.close()
        content={'products':{self.asin:{'selling_points':['10层压花结构'],'main_image':self.asset}},
                 'comparisons':[{'label':'层数','values':{self.asin:'10层'}}]}
        content_path=self.root/'content.json'
        content_path.write_text(json.dumps(content,ensure_ascii=False),encoding='utf-8')
        output=self.root/'populated.xlsx'
        populate(path,content_path,self.root/'no-cards',output)
        checked=load_workbook(output,rich_text=True)
        self.assertEqual(checked['测试-竞品分析']['C4'].value,'4.2 / 43')
        self.assertEqual(checked['测试-数据源'].max_row,1)
        image=checked['测试-竞品分析']._images[0]
        self.assertEqual((image.anchor.ext.cx//9525,image.anchor.ext.cy//9525),(1200,900))
        self.assertEqual(image._data(),(self.root/'main.png').read_bytes())
        checked.close()


class IdentityTests(unittest.IsolatedAsyncioTestCase):
    async def test_duplicate_identical_asin_allowed_but_conflicts_rejected(self):
        for values,expected in [(['B012345678']*2,'B012345678'),(['B012345678','B087654321'],None),([],None)]:
            page=Mock()
            page.locator.return_value=Mock(evaluate_all=AsyncMock(return_value=values))
            self.assertEqual(await selected_asin(page),expected)

    async def test_wrong_page_rejected_before_download(self):
        page=Mock(url='https://www.amazon.com/dp/B087654321')
        with self.assertRaises(ValueError):
            await collect_main_image(page,'B012345678',Path('unused'))
        page.context.request.get.assert_not_called()


if __name__=='__main__':
    unittest.main()
