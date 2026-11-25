"""把本批全部原始评论及逐行分析合并到报告的一张评价数据源页。"""
from __future__ import annotations

import argparse
from copy import copy
import json
import math
import re
from pathlib import Path
import sys
import tempfile
import unicodedata
from xml.dom import minidom
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from competitor_analysis.annotate_review_export import HEADERS, annotation_rows, children, element
from competitor_analysis.populate_basics import competitor_asins
from competitor_analysis.sellersprite_review_analysis import make_packet, validate_result, MENTION_SCOPE
from review_research.download_sellersprite_reviews import read_export

SHEET_NAME = '评价数据源'
ASIN_HEADER = '分析商品 ASIN'
MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
OFFICE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PACKAGE = 'http://schemas.openxmlformats.org/package/2006/relationships'


def verify_analysis(analysis):
    """重读源文件与证据，不能因为叫checked.json就信任旧结果。"""
    packet = make_packet(Path(analysis['source_file']), analysis['asin'])
    if any(analysis.get(key) != packet.get(key) for key in ('schema', 'packet_digest', 'source_sha256')):
        raise ValueError('评论数据已变化或分析来自另一批次，拒绝回填')
    checked = dict(analysis)
    for kind in ('advantages', 'disadvantages'):
        checked[kind] = validate_result(packet, kind, {
            'schema': packet['schema'], 'asin': packet['asin'], 'packet_digest': packet['packet_digest'],
            'kind': kind, 'reviewed_numbers': [r['number'] for r in packet['samples'][kind]],
            'points': analysis[kind], 'evidence_scope': MENTION_SCOPE,
            'no_supported_points_reason': analysis.get('empty_reasons', {}).get(kind, ''),
        })
    return packet, checked


def merged_data(workbook, analyses):
    """按竞品顺序分组、组内保持源行顺序；共享变体评论不跨商品去重。"""
    book = load_workbook(workbook, rich_text=True)
    try:
        if SHEET_NAME in book.sheetnames:
            raise ValueError('报告已有评价数据源，请使用合并前检查点，避免重复追加或覆盖人工修改')
        sheets = [s for s in book if s.title.endswith('竞品分析')]
        if len(sheets) != 1:
            raise ValueError('必须有唯一的竞品分析页')
        asins = competitor_asins(sheets[0])
    finally:
        book.close()
    by_asin = {a['asin']: a for a in analyses}
    if not asins or len(by_asin) != len(analyses) or set(by_asin) != set(asins):
        raise ValueError('本批分析与模板竞品不一一对应，不能漏商品或混入其它商品')
    sources, headers, records, counts = [], [], [], {}
    for asin in asins:
        packet, analysis = verify_analysis(by_asin[asin])
        source = read_export(Path(packet['source_file']), asin)
        notes = annotation_rows(packet, analysis, source)
        if ASIN_HEADER in source['headers']:
            raise ValueError('原始评论已包含分析商品ASIN列，拒绝再次合并')
        for header in source['headers']:
            if header not in headers:
                headers.append(header)
        sources.append((source, notes))
        counts[asin] = source['review_count']
    columns = [ASIN_HEADER, *headers, *HEADERS]
    for source, notes in sources:
        for review in source['reviews']:
            records.append([source['asin'], *[review['raw'].get(h) for h in headers], *notes[review['source_row']]])
    return columns, records, counts


def add_styles(styles):
    """只追加新页的样式，原报告全部样式编号保持不变。"""
    root = styles.documentElement
    fonts, fills, xfs = [children(root, name)[0] for name in ('fonts', 'fills', 'cellXfs')]
    font_ids = []
    for bold, color in ((False, 'FF222222'), (True, 'FFFFFFFF'), (False, 'FF0563C1')):
        font = element(fonts, 'font')
        font.appendChild(element(font, 'sz', {'val': 11}))
        font.appendChild(element(font, 'name', {'val': '微软雅黑'}))
        font.appendChild(element(font, 'color', {'rgb': color}))
        if bold:
            font.appendChild(element(font, 'b'))
        if color == 'FF0563C1':
            font.appendChild(element(font, 'u'))
        font_ids.append(len(children(fonts, 'font')))
        fonts.appendChild(font)
    fill_id = len(children(fills, 'fill'))
    fill = element(fills, 'fill')
    pattern = element(fill, 'patternFill', {'patternType': 'solid'})
    pattern.appendChild(element(pattern, 'fgColor', {'rgb': 'FFE98A00'}))
    pattern.appendChild(element(pattern, 'bgColor', {'indexed': 64}))
    fill.appendChild(pattern)
    fills.appendChild(fill)
    style_ids = []
    for index, font_id in enumerate(font_ids):
        xf = element(xfs, 'xf', {'numFmtId': 0, 'fontId': font_id, 'fillId': fill_id if index == 1 else 0,
                                'borderId': 0, 'xfId': 0, 'applyFont': 1, 'applyFill': 1, 'applyAlignment': 1})
        xf.appendChild(element(xf, 'alignment', {'vertical': 'center' if index == 1 else 'top',
                                               'horizontal': 'center' if index == 1 else 'left', 'wrapText': 1}))
        style_ids.append(len(children(xfs, 'xf')))
        xfs.appendChild(xf)
    for parent, name in ((fonts, 'font'), (fills, 'fill'), (xfs, 'xf')):
        parent.setAttribute('count', str(len(children(parent, name))))
    return style_ids


def sheet_xml(headers, records, style_ids):
    """字符串一律写为文字，评论中以等号开头的内容不能变成公式。"""
    root = ET.Element('worksheet', {'xmlns': MAIN, 'xmlns:r': OFFICE})
    last = f'{get_column_letter(len(headers))}{len(records) + 1}'
    ET.SubElement(root, 'dimension', {'ref': f'A1:{last}'})
    view = ET.SubElement(ET.SubElement(root, 'sheetViews'), 'sheetView', {'workbookViewId': '0'})
    ET.SubElement(view, 'pane', {'xSplit': '1', 'ySplit': '1', 'topLeftCell': 'B2', 'activePane': 'bottomRight', 'state': 'frozen'})
    ET.SubElement(view, 'selection', {'pane': 'bottomRight', 'activeCell': 'B2', 'sqref': 'B2'})
    ET.SubElement(root, 'sheetFormatPr', {'defaultRowHeight': '30'})
    widths = {'分析商品 ASIN': 18, 'ASIN': 16, '标题': 36, '标题(翻译)': 36, '内容': 65, '内容(翻译)': 65,
              '优点归纳': 34, '差评原因': 34, '分析说明': 46, '评论链接': 40, '评论时间': 18}
    cols = ET.SubElement(root, 'cols')
    for col, header in enumerate(headers, 1):
        ET.SubElement(cols, 'col', {'min': str(col), 'max': str(col), 'width': str(widths.get(header, 18)), 'customWidth': '1'})
    data, links = ET.SubElement(root, 'sheetData'), []
    for number, values in enumerate([headers, *records], 1):
        # 新页按评论与分析文字估算行高，超长原文仍完整存储，遵守Excel行高上限。
        lines = 1
        for header, value in zip(headers, values):
            if header not in ('标题', '标题(翻译)', '内容', '内容(翻译)', *HEADERS):
                continue
            count = sum(max(1, math.ceil(sum(2 if unicodedata.east_asian_width(c) in 'WF' else 1
                                                for c in paragraph) / (widths[header] - 3)))
                        for paragraph in str(value or '').split('\n'))
            lines = max(lines, count)
        height = 32 if number == 1 else min(409, max(48, lines * 16 + 8))
        row = ET.SubElement(data, 'row', {'r': str(number), 'ht': str(height), 'customHeight': '1'})
        for col, value in enumerate(values, 1):
            if value is None or value == '':
                continue
            address = f'{get_column_letter(col)}{number}'
            link = number > 1 and headers[col - 1] == '评论链接'
            cell = ET.SubElement(row, 'c', {'r': address, 's': str(style_ids[1 if number == 1 else 2 if link else 0])})
            if isinstance(value, bool):
                cell.set('t', 'b')
                ET.SubElement(cell, 'v').text = '1' if value else '0'
            elif isinstance(value, (int, float)):
                if not math.isfinite(value):
                    raise ValueError('评论源存在非有限数值')
                cell.set('t', 'n')
                ET.SubElement(cell, 'v').text = str(value)
            else:
                cell.set('t', 'inlineStr')
                ET.SubElement(ET.SubElement(cell, 'is'), 't', {'xml:space': 'preserve'}).text = str(value)
            if link:
                links.append((address, value))
    ET.SubElement(root, 'autoFilter', {'ref': f'A1:{last}'})
    rels = ET.Element('Relationships', {'xmlns': PACKAGE})
    if links:
        hyperlinks = ET.SubElement(root, 'hyperlinks')
        for index, (address, url) in enumerate(links, 1):
            ET.SubElement(hyperlinks, 'hyperlink', {'ref': address, 'r:id': f'rId{index}'})
            ET.SubElement(rels, 'Relationship', {'Id': f'rId{index}', 'Type': f'{OFFICE}/hyperlink', 'Target': url, 'TargetMode': 'External'})
    return ET.tostring(root, encoding='utf-8', xml_declaration=True), ET.tostring(rels, encoding='utf-8', xml_declaration=True)


def merge_sources(workbook, output, analyses):
    if output.exists() or output.resolve() == workbook.resolve():
        raise ValueError('输出必须是新的独立文件')
    headers, records, counts = merged_data(workbook, analyses)
    output.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(workbook) as archive, tempfile.TemporaryDirectory(dir=output.parent) as temporary:
        book = minidom.parseString(archive.read('xl/workbook.xml'))
        rels = minidom.parseString(archive.read('xl/_rels/workbook.xml.rels'))
        types = minidom.parseString(archive.read('[Content_Types].xml'))
        styles = minidom.parseString(archive.read('xl/styles.xml'))
        sheets = book.getElementsByTagNameNS(MAIN, 'sheets')[0]
        sheet_id = max(int(s.getAttribute('sheetId')) for s in children(sheets, 'sheet')) + 1
        index = 1
        while f'xl/worksheets/sheet{index}.xml' in archive.namelist():
            index += 1
        part = f'xl/worksheets/sheet{index}.xml'
        rel_ids = {r.getAttribute('Id') for r in children(rels.documentElement, 'Relationship')}
        rid = 1
        while f'rId{rid}' in rel_ids:
            rid += 1
        sheet = element(sheets, 'sheet', {'name': SHEET_NAME, 'sheetId': sheet_id})
        sheet.setAttributeNS(OFFICE, 'r:id', f'rId{rid}')
        # 某些Excel文件的关系前缀不是r，新元素需显式声明。
        sheet.setAttribute('xmlns:r', OFFICE)
        sheets.appendChild(sheet)
        rels.documentElement.appendChild(element(rels.documentElement, 'Relationship', {
            'Id': f'rId{rid}', 'Type': f'{OFFICE}/worksheet', 'Target': f'worksheets/sheet{index}.xml'}))
        types.documentElement.appendChild(element(types.documentElement, 'Override', {
            'PartName': '/' + part, 'ContentType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'}))
        sheet_data, hyperlinks = sheet_xml(headers, records, add_styles(styles))
        patches = {'xl/workbook.xml': book.toxml(encoding='utf-8'),
                   'xl/_rels/workbook.xml.rels': rels.toxml(encoding='utf-8'),
                   '[Content_Types].xml': types.toxml(encoding='utf-8'),
                   # 数字格式属性里的既有换行必须保留为字符实体。
                   'xl/styles.xml': re.sub(r'formatCode="[^"]*"', lambda m: m[0].replace('\n', '&#10;'),
                                           styles.toxml()).encode('utf-8')}
        staged = Path(temporary) / 'merged.xlsx'
        with ZipFile(staged, 'w') as target:
            for info in archive.infolist():
                target.writestr(copy(info), patches.get(info.filename, archive.read(info.filename)))
            target.writestr(part, sheet_data)
            target.writestr(f'xl/worksheets/_rels/sheet{index}.xml.rels', hyperlinks)
        validate_merged(staged, headers, records)
        staged.replace(output)
    return {'output': str(output), 'sheet': SHEET_NAME, 'reviews': len(records), 'products': counts}


def validate_merged(path, headers, records):
    """保存后逐格核对完整评论、商品归属、分析证据及可点击原评论链接。"""
    with ZipFile(path) as archive:
        if archive.testzip():
            raise ValueError('合并评论后的工作簿ZIP损坏')
    book = load_workbook(path, rich_text=True)
    try:
        sheet = book[SHEET_NAME]
        if sheet.max_row != len(records) + 1 or sheet.max_column != len(headers):
            raise ValueError('合并评论行列数量不一致')
        for number, expected in enumerate([headers, *records], 1):
            for col, value in enumerate(expected, 1):
                cell = sheet.cell(number, col)
                if (cell.value if cell.value is not None else '') != (value if value is not None else ''):
                    raise ValueError(f'合并评论第{number}行第{col}列与来源不符')
                if cell.data_type == 'f':
                    raise ValueError('原评论文字被误写成公式')
                if number > 1 and headers[col - 1] == '评论链接' and (not cell.hyperlink or cell.hyperlink.target != value):
                    raise ValueError('合并评论缺少原评论超链接')
    finally:
        book.close()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--workbook', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--analyses', type=Path, nargs='+', required=True)
    args = parser.parse_args()
    analyses = [json.loads(p.read_text(encoding='utf-8')) for p in args.analyses]
    print(json.dumps(merge_sources(args.workbook, args.output, analyses), ensure_ascii=False))


if __name__ == '__main__':
    main()
