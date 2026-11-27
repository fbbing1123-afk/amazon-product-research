"""核验嵌入图片不改布局、原媒体或其它内容，并在目标冲突时停止。"""
import hashlib
import json
import tempfile
from pathlib import Path
from xml.dom import minidom
from zipfile import ZipFile
import unittest
from unittest.mock import patch

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AbsoluteAnchor
from openpyxl.styles import Font, PatternFill
from PIL import Image

from shared.cell_image_embedding import embed_final_images, MAIN, XDR, ART, _all


class EmbeddingTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        self.image = self.root/'image.png'
        Image.new('RGBA',(80,40),(30,120,200,170)).save(self.image)

    def fixture(self,occupied=False,duplicate=False,chart=False,absolute=False):
        book = Workbook()
        sheet = book.active
        sheet.title = '市场调研'
        sheet['A1'] = '保留原文'
        sheet['B2'].fill = PatternFill('solid',fgColor='FFC0CB')
        sheet['B2'].font = Font(name='微软雅黑',color='FF0000',bold=True)
        sheet.row_dimensions[2].height = 36
        sheet.column_dimensions['B'].width = 20
        sheet.merge_cells('C3:F4')
        sheet.freeze_panes = 'C3'
        sheet['N1'],sheet['N2'],sheet['N3'] = 1,2,3
        picture = XLImage(str(self.image))
        picture.anchor = AbsoluteAnchor() if absolute else 'B2'
        sheet.add_image(picture)
        sheet.add_image(XLImage(str(self.image)),'D4')
        if occupied:
            sheet['B2'] = '不能覆盖'
        if duplicate:
            sheet.add_image(XLImage(str(self.image)),'B2')
        if chart:
            native = BarChart()
            native.add_data(Reference(sheet,min_col=14,min_row=1,max_row=3))
            sheet.add_chart(native,'H10')
        book.create_sheet('其它页')['A1'] = '=1+2'
        source = self.root/'source.xlsx'
        book.save(source)
        book.close()
        with ZipFile(source,'a') as archive:
            archive.writestr('xl/custom-preserve.bin',b'preserve-native-part')
        return source

    def test_embeds_merged_cells_with_original_pixels_and_unchanged_layout(self):
        source = self.fixture()
        output = self.root/'embedded.xlsx'
        fingerprint = hashlib.sha256(source.read_bytes()).hexdigest()
        result = embed_final_images(source,output)
        self.assertEqual(result['embedded'],2)
        self.assertEqual(['B2','C3'],[p['cell'] for p in result['images']])
        with ZipFile(source) as before,ZipFile(output) as after:
            for part in before.namelist():
                if part not in result['changed_parts']:
                    self.assertEqual(before.read(part),after.read(part),part)
            for entry in result['images']:
                self.assertEqual(after.read(entry['media']),self.image.read_bytes())
                self.assertEqual(entry['pixels'],(80,40))
            a,b=[minidom.parseString(archive.read('xl/worksheets/sheet1.xml')) for archive in (before,after)]
            for tag in ('cols','mergeCells','sheetFormatPr','sheetViews'):
                self.assertEqual([n.toxml() for n in _all(a,MAIN,tag)],[n.toxml() for n in _all(b,MAIN,tag)],tag)
            self.assertEqual([dict(n.attributes.items()) for n in _all(a,MAIN,'row')],
                             [dict(n.attributes.items()) for n in _all(b,MAIN,'row')])
        original,embedded=load_workbook(source),load_workbook(output)
        try:
            self.assertEqual(embedded.active['B2']._style,original.active['B2']._style)
            self.assertEqual(embedded.active['A1'].value,'保留原文')
            self.assertEqual(len(embedded.active._images),0)
            self.assertTrue(embedded.active['B2'].value.startswith('=_xlfn.DISPIMG('))
            self.assertEqual(embedded['其它页']['A1'].value,'=1+2')
        finally:
            original.close()
            embedded.close()
        self.assertEqual(hashlib.sha256(source.read_bytes()).hexdigest(),fingerprint)

    def test_occupied_or_duplicate_cell_fails_without_output(self):
        for options in ({'occupied':True},{'duplicate':True}):
            with self.subTest(options=options):
                source = self.fixture(**options)
                output = self.root/'conflict.xlsx'
                with self.assertRaises(ValueError):
                    embed_final_images(source,output)
                self.assertFalse(output.exists())

    def test_repeated_conversion_preserves_existing_cell_images(self):
        source = self.fixture()
        first,second = self.root/'first.xlsx',self.root/'second.xlsx'
        embed_final_images(source,first)
        result = embed_final_images(first,second)
        self.assertEqual(result['embedded'],0)
        with ZipFile(first) as before,ZipFile(second) as after:
            self.assertEqual(before.namelist(),after.namelist())
            for part in before.namelist():
                self.assertEqual(before.read(part),after.read(part),part)

    def test_native_chart_is_preserved(self):
        source = self.fixture(chart=True)
        output = self.root/'chart.xlsx'
        embed_final_images(source,output)
        with ZipFile(source) as before,ZipFile(output) as after:
            part = 'xl/charts/chart1.xml'
            self.assertEqual(before.read(part),after.read(part))
            drawing = minidom.parseString(after.read('xl/drawings/drawing1.xml'))
            self.assertEqual(len(_all(drawing,XDR,'pic')),0)
            self.assertEqual(len(_all(drawing,XDR,'graphicFrame')),1)

    def test_absolute_anchor_and_existing_output_are_not_overwritten(self):
        source = self.fixture(absolute=True)
        output = self.root/'unsafe.xlsx'
        with self.assertRaisesRegex(ValueError,'图片锚点'):
            embed_final_images(source,output)
        self.assertFalse(output.exists())
        output.write_bytes(b'keep')
        with self.assertRaises(ValueError):
            embed_final_images(source,output)
        self.assertEqual(output.read_bytes(),b'keep')

    def test_existing_command_embeds_final_checkpoint_without_original_argument(self):
        from scripts.make_wps_embedded_version import main
        source = self.fixture()
        output,manifest = self.root/'command.xlsx',self.root/'manifest.json'
        with patch('sys.argv',['make_wps_embedded_version.py','--floating-workbook',str(source),
                               '--output',str(output),'--manifest',str(manifest)]), patch('builtins.print'):
            self.assertEqual(main(),0)
        self.assertEqual(json.loads(manifest.read_text(encoding='utf-8'))['embedded'],2)
        self.assertTrue(output.exists())
        fingerprint = source.read_bytes()
        invalid = self.root/'invalid.xlsx'
        with patch('sys.argv',['make_wps_embedded_version.py','--floating-workbook',str(source),
                               '--output',str(invalid),'--manifest',str(source)]), patch('sys.stderr'):
            with self.assertRaises(SystemExit):
                main()
        self.assertEqual(source.read_bytes(),fingerprint)
        self.assertFalse(invalid.exists())


if __name__ == '__main__':
    unittest.main()
