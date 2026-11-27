"""核验市场结论的来源绑定、动态定位、缺失处理及原生图片/格式保护。"""
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
from xml.dom import minidom
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Alignment, Font, PatternFill
from PIL import Image

from product_research.scripts.fill_market_narrative import (
    SECTIONS, fill_narrative, prepare_evidence, locate_narrative_regions,
)
from product_research.scripts.populate_market_summary import METRICS
from shared.cell_image_embedding import embed_final_images
from shared.quality_gates import QualityGateError, validate_market_narrative
from shared.worksheet_rows import insert_rows_preserving_merges


class MarketNarrativeTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        self.picture = self.root / 'chart.png'
        with Image.new('RGB', (800, 400), 'blue') as image:
            image.save(self.picture)
        self.manifest = self.root / 'report.json'
        self.manifest.write_text(json.dumps({'report_url': 'https://www.sellersprite.com/v3/custom-market-research/1/1',
            'charts': {key: {'sample_count': 153, 'image': str(self.picture)}
                       for key in ('review', 'rating', 'price')}}), encoding='utf-8')

    def fixture(self, shift=0):
        book = Workbook()
        sheet = book.active
        sheet.title = '产品-市场调研'
        for row in range(1, 31):
            sheet.row_dimensions[row].height = 100 if row in (19, 20, 21, 22, 23, 24) else 30
            for column in range(1, 10):
                cell = sheet.cell(row, column)
                cell.font = Font(name='宋体', size=11)
                cell.alignment = Alignment(vertical='center', wrap_text=True)
                cell.fill = PatternFill('solid', fgColor='F0F5FF')
        for index, label in enumerate(METRICS):
            sheet.cell(14 + index // 3, 1 + index % 3 * 2, label)
        sheet.merge_cells('G14:I17')
        for row, label in ((19, '品牌分析'), (20, '上架时间分析')):
            sheet.cell(row, 1, label)
            for columns in ('B:C', 'D:E', 'F:G', 'H:I'):
                left, right = columns.split(':')
                sheet.merge_cells(f'{left}{row}:{right}{row}')
        for row, label in ((21, 'Review数量分析'), (22, 'Rating分析'), (23, '价格分析')):
            sheet.cell(row, 1, label)
            sheet.merge_cells(f'B{row}:F{row}')
            sheet.merge_cells(f'G{row}:I{row}')
        sheet['A24'] = '总结'
        sheet.merge_cells('B24:I24')
        sheet['A30'] = '不变的正文'
        sheet.add_image(ExcelImage(str(self.picture)), 'B21')
        if shift:
            insert_rows_preserving_merges(sheet, 13, shift, 10)
        source = book.create_sheet('产品-数据源')
        source.append(['ASIN', '品牌', '月销量', '价格($)', '月销售额($)', '上架时间'])
        source.append(['B012345678', 'A', 10, 20, 200, '2026-01-01'])
        source.append(['B087654321', 'B', 0, 30, 0, '2025-01-01'])
        source.append(['B000000003', None, None, 40, None, None])
        raw = self.root / f'raw-{shift}.xlsx'
        book.save(raw)
        book.close()
        embedded = self.root / f'embedded-{shift}.xlsx'
        embed_final_images(raw, embedded)
        return embedded

    def packet(self, path, manifest=True):
        return prepare_evidence(path, as_of='2026-11-27', report_manifest=self.manifest if manifest else None)

    def analysis(self, packet):
        return {'packet_digest': packet['packet_digest'], 'sections': {
            key: {'status': 'analyzed' if packet['sources'][primary]['status'] == 'available' else 'unavailable',
                  'text': f'{label}：依据本批数据分析。' if packet['sources'][primary]['status'] == 'available'
                          else f'{label}：本批未取得图表，无法判断。',
                  'highlights': ['本批'],
                  'evidence': [primary]}
            for key, (label, primary) in SECTIONS.items()}}

    def test_source_and_report_populations_and_unknowns_remain_separate(self):
        packet = self.packet(self.fixture())
        self.assertEqual(3, packet['sources']['source.overall']['metrics']['valid_source_asins'])
        self.assertEqual(2, packet['sources']['source.overall']['metrics']['sample_count'])
        self.assertEqual(1, packet['sources']['source.overall']['metrics']['missing_sales_count'])
        self.assertEqual(153, packet['sources']['report.review']['sample_count'])
        unknown = next(g for g in packet['sources']['source.brands']['groups'] if g['label'] == '未知')
        zero = next(g for g in packet['sources']['source.brands']['groups'] if g['label'] == 'B')
        self.assertEqual((0, 1), (unknown['sales_known'], zero['sales_known']))
        self.assertEqual(1, packet['sources']['source.listing']['unknown_dates'])

    def test_shifted_layout_fills_seven_regions_and_preserves_every_other_part(self):
        source = self.fixture(5)
        packet = self.packet(source)
        output = self.root / 'finished.xlsx'
        with self.assertRaises(QualityGateError):
            validate_market_narrative(source)
        result = fill_narrative(source, output, packet, self.analysis(packet))
        self.assertEqual('G19:I22', result['written']['overall'])
        self.assertEqual('B29:I29', result['written']['summary'])
        self.assertEqual(7, validate_market_narrative(output).metrics['sections'])
        from shared.quality_gates import main
        with patch('sys.argv', ['quality_gates.py', '--step', 'market_narrative', '--workbook', str(output)]), patch('builtins.print'):
            self.assertEqual(0, main())
        checked = load_workbook(output, rich_text=True)
        from shared.quality_gates import rich_text_red_text
        self.assertEqual(['本批'], rich_text_red_text(checked.active['G19'].value))
        for run in checked.active['G19'].value:
            self.assertEqual('宋体', run.font.rFont)
            self.assertEqual(11, run.font.sz)
        checked.close()
        with ZipFile(source) as before, ZipFile(output) as after:
            self.assertEqual(before.namelist(), after.namelist())
            self.assertEqual(['xl/worksheets/sheet1.xml'], [n for n in before.namelist() if before.read(n) != after.read(n)])
            original = minidom.parseString(before.read('xl/worksheets/sheet1.xml'))
            written = minidom.parseString(after.read('xl/worksheets/sheet1.xml'))
            targets = {area.split(':')[0] for area in result['written'].values()}
            for doc in (original, written):
                for cell in doc.getElementsByTagName('c'):
                    if cell.getAttribute('r') in targets:
                        for child in list(cell.childNodes):
                            cell.removeChild(child)
                        cell.removeAttribute('t') if cell.hasAttribute('t') else None
            self.assertEqual(original.toxml(), written.toxml())

    def test_missing_charts_must_be_explained_and_never_replaced_with_source_statistics(self):
        source = self.fixture()
        packet = self.packet(source, manifest=False)
        analysis = self.analysis(packet)
        output = self.root / 'missing.xlsx'
        fill_narrative(source, output, packet, analysis)
        book = load_workbook(output)
        self.assertIn('未取得', book.active['G21'].value)
        book.close()
        analysis['sections']['review']['status'] = 'analyzed'
        with self.assertRaisesRegex(ValueError, '状态'):
            fill_narrative(source, self.root / 'invalid.xlsx', packet, analysis)

    def test_incomplete_analysis_wrong_evidence_and_oversize_text_are_rejected(self):
        source = self.fixture()
        packet = self.packet(source)
        for problem in ('missing', 'reference', 'long', 'digest'):
            with self.subTest(problem=problem):
                analysis = self.analysis(packet)
                if problem == 'missing':
                    analysis['sections'].pop('summary')
                elif problem == 'reference':
                    analysis['sections']['review']['evidence'] = ['source.overall']
                elif problem == 'long':
                    analysis['sections']['brand']['text'] = '很长的分析' * 1000
                else:
                    analysis['packet_digest'] = 'another-batch'
                output = self.root / f'{problem}.xlsx'
                with self.assertRaises(ValueError):
                    fill_narrative(source, output, packet, analysis)
                self.assertFalse(output.exists())

    def test_changed_workbook_or_changed_chart_is_rejected(self):
        source = self.fixture()
        packet = self.packet(source)
        analysis = self.analysis(packet)
        original = source.read_bytes()
        source.write_bytes(original + b'changed')
        with self.assertRaisesRegex(ValueError, '工作簿已变化'):
            fill_narrative(source, self.root / 'stale.xlsx', packet, analysis)
        source.write_bytes(original)
        self.picture.write_bytes(self.picture.read_bytes() + b'changed')
        with self.assertRaisesRegex(ValueError, '图像已变化'):
            fill_narrative(source, self.root / 'stale-image.xlsx', packet, analysis)

    def test_existing_text_needs_explicit_replacement(self):
        source = self.fixture()
        packet = self.packet(source)
        first = self.root / 'first.xlsx'
        fill_narrative(source, first, packet, self.analysis(packet))
        packet = self.packet(first)
        with self.assertRaisesRegex(ValueError, '已有内容'):
            fill_narrative(first, self.root / 'refused.xlsx', packet, self.analysis(packet))
        fill_narrative(first, self.root / 'replaced.xlsx', packet, self.analysis(packet), replace_existing=True)
        with self.assertRaisesRegex(ValueError, '独立输出'):
            fill_narrative(source, source, packet, self.analysis(packet))

    def test_ambiguous_or_missing_chart_layout_is_rejected(self):
        source = self.fixture()
        book = load_workbook(source)
        sheet = book.active
        sheet.unmerge_cells('B21:F21')
        with self.assertRaises(ValueError):
            locate_narrative_regions(sheet)
        book.close()

    def test_invalid_and_excessive_highlighting_is_rejected(self):
        source = self.fixture()
        packet = self.packet(source)
        for highlights in ([], ['不存在的内容'], ['本批', '本批'], ['本批', '本批数据'], ['整体销售分析：依据本批数据分析。']):
            with self.subTest(highlights=highlights):
                analysis = self.analysis(packet)
                analysis['sections']['overall']['highlights'] = highlights
                with self.assertRaises(ValueError):
                    fill_narrative(source, self.root / 'invalid-highlight.xlsx', packet, analysis)


if __name__ == '__main__':
    unittest.main()
