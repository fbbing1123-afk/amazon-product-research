"""准备本批市场分析证据，并将AI核实的七项结论定点写入原模板。"""
from __future__ import annotations

import argparse
from datetime import date, datetime
import hashlib
import json
import math
from pathlib import Path
import re
import sys
import tempfile
import unicodedata
from xml.dom import minidom
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.utils.cell import coordinate_to_tuple

MODULES = Path(__file__).resolve().parents[2]
if str(MODULES) not in sys.path:
    sys.path.insert(0, str(MODULES))
from market_analysis.excel_market_report import compute_metrics, first_field_value, read_rows, to_number
from shared.excel_inspect import find_source_sheet
from product_research.scripts.insert_sellersprite_screenshots import locate_chart_regions, region_pixels, normalize_label
from competitor_analysis.annotate_review_export import worksheet_part, children, element
from product_research.scripts.fill_custom_template import displayed_category_summaries

SECTIONS = {
    'overall': ('整体销售分析', 'source.overall'),
    'brand': ('品牌分析', 'source.brands'),
    'listing': ('上架时间分析', 'source.listing'),
    'review': ('Review数量分析', 'report.review'),
    'rating': ('Rating分析', 'report.rating'),
    'price': ('价格分析', 'report.price'),
    'summary': ('市场总结', 'source.overall'),
}


def digest(value):
    return hashlib.sha256(json.dumps(value, ensure_ascii=False, sort_keys=True, default=str).encode('utf-8')).hexdigest()


def file_sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def market_sheet(book):
    sheets = [sheet for sheet in book if sheet.title.strip().endswith('市场调研')]
    if len(sheets) != 1:
        raise ValueError('市场调研页必须唯一')
    return sheets[0]


def locate_narrative_regions(sheet):
    """按当前标签及实际合并图框找结论区，兼容关键词和类目增行。"""
    def label_cell(label):
        matches = [cell for row in sheet for cell in row if normalize_label(cell.value) == normalize_label(label)]
        if len(matches) != 1:
            raise ValueError(f'{label}标签必须唯一')
        return matches[0]

    def unique(areas, label):
        if len(areas) != 1:
            raise ValueError(f'{label}缺少唯一的结论合并区')
        return areas[0]

    sample = label_cell('样本数')
    last = label_cell('中位销售额($)')
    areas = list(sheet.merged_cells.ranges)
    result = {'overall': unique([a for a in areas if a.min_row == sample.row and a.max_row >= last.row
                                  and a.min_col > last.column + 1], '整体销售分析')}
    for key, label in [('brand', '品牌分析'), ('listing', '上架时间分析')]:
        cell = label_cell(label)
        right = sorted([a for a in areas if a.min_row == a.max_row == cell.row and a.min_col > cell.column],
                       key=lambda a: a.min_col)
        if len(right) != 4 or any(a.max_col >= b.min_col for a, b in zip(right, right[1:])):
            raise ValueError(f'{label}必须有三个图框及一个独立结论区')
        result[key] = right[-1]
    for key, chart in locate_chart_regions(sheet).items():
        result[key] = unique([a for a in areas if a.min_row == chart.min_row and a.max_row == chart.max_row
                              and a.min_col == chart.max_col + 1], SECTIONS[key][0])
    summary = label_cell('总结')
    result['summary'] = unique([a for a in areas if a.min_row == a.max_row == summary.row
                                and a.min_col == summary.column + 1], '市场总结')
    if len({str(area) for area in result.values()}) != len(SECTIONS):
        raise ValueError('结论区域重复，拒绝写入')
    # 结论区不能包含任何浮动图片，防止误把缺失图框识别为文字区域。
    for image in sheet._images:
        anchor = image.anchor
        if isinstance(anchor, str):
            row, column = coordinate_to_tuple(anchor)
        elif hasattr(anchor, '_from'):
            row, column = anchor._from.row + 1, anchor._from.col + 1
        else:
            continue
        if any(a.min_row <= row <= a.max_row and a.min_col <= column <= a.max_col for a in result.values()):
            raise ValueError('结论区存在图片，拒绝覆盖')
    return result


def _number(value):
    number = to_number(value)
    return number if number is not None and math.isfinite(number) else None


def _groups(records, field):
    """保留缺失计数；聚合有值项目，不把未知销量说成真实零销量。"""
    groups = {}
    for record in records:
        label = str(record.get(field) or '未知').strip() or '未知'
        group = groups.setdefault(label, {'label': label, 'products': 0, 'sales': 0, 'revenue': 0,
                                           'sales_known': 0, 'revenue_known': 0})
        group['products'] += 1
        for metric, value in [('sales', _number(record.get('月销量'))),
                              ('revenue', _number(first_field_value(record, '月销售额')))]:
            if value is not None:
                group[metric] += value
                group[metric + '_known'] += 1
    return sorted(groups.values(), key=lambda g: (-g['sales'], -g['products'], g['label']))


def prepare_evidence(workbook, *, as_of, report_manifest=None):
    workbook = Path(workbook)
    observed = date.fromisoformat(as_of)
    book = load_workbook(workbook, rich_text=True)
    try:
        sheet = market_sheet(book)
        regions = locate_narrative_regions(sheet)
        source = find_source_sheet(book)
        if source is None:
            raise ValueError('缺少本批商品数据源')
        records = [r for r in read_rows(source) if re.fullmatch(r'[A-Z0-9]{10}', str(r.get('ASIN', '')).strip())]
        if not records or len({r['ASIN'] for r in records}) != len(records):
            raise ValueError('有效ASIN为空或重复')
        metrics = compute_metrics(records)
        overview = {key: value for key, value in metrics.items()
                    if key not in {'ranked', 'sales_layers', 'price_buckets', 'review_buckets', 'anomalies'}}
        overview.update(valid_source_asins=len(records), missing_sales_count=sum(_number(r.get('月销量')) is None for r in records))
        brands = _groups(records, '品牌')
        listing_records = []
        unknown_dates = 0
        outside_window = 0
        for record in records:
            value = record.get('上架时间')
            try:
                listed = value.date() if isinstance(value, datetime) else value if isinstance(value, date) else date.fromisoformat(str(value)[:10])
            except (ValueError, TypeError):
                unknown_dates += 1
                continue
            if listed > observed or listed.year < observed.year - 5:
                outside_window += 1
                continue
            listing_records.append({**record, '_year': str(listed.year)})
        sources = {
            'source.overall': {'status': 'available', 'metrics': overview},
            'source.brands': {'status': 'available', 'sample_count': len(records), 'groups': brands},
            'source.listing': {'status': 'available' if listing_records else 'unavailable',
                               'window': [observed.year - 5, observed.year], 'sample_count': len(listing_records),
                               'unknown_dates': unknown_dates, 'outside_window': outside_window,
                               'groups': _groups(listing_records, '_year')},
        }
        sources['source.categories'] = {'status': 'available', 'groups': [
            {'path': item.path_text, 'products': item.product_count, 'share': item.product_share}
            for item in displayed_category_summaries(source)]}
        keyword_header = next((cell for row in sheet for cell in row if str(cell.value or '').strip() == '关键词'), None)
        keywords = []
        if keyword_header:
            for row in range(keyword_header.row + 1, sheet.max_row + 1):
                name = sheet.cell(row, keyword_header.column).value
                if re.match(r'^\s*2\.', str(name or '')):
                    break
                if name:
                    keywords.append({'keyword': str(name).splitlines()[0], 'season': sheet.cell(row, 5).value,
                                     'off_season_search': sheet.cell(row, 7).value,
                                     'peak_search': sheet.cell(row, 8).value, 'weekly_bid': sheet.cell(row, 9).value})
        sources['source.keywords'] = {'status': 'available' if keywords else 'unavailable', 'rows': keywords}
        layout = {key: {'range': str(area), 'pixels': region_pixels(sheet, area),
                         'font_points': sheet.cell(area.min_row, area.min_col).font.sz or 11}
                  for key, area in regions.items()}
    finally:
        book.close()
    manifest = json.loads(Path(report_manifest).read_text(encoding='utf-8-sig')) if report_manifest else {}
    for key in ('review', 'rating', 'price'):
        chart = manifest.get('charts', {}).get(key, {})
        image = Path(chart.get('image', ''))
        if report_manifest and not image.is_absolute():
            image = Path(report_manifest).resolve().parent / image
        count = chart.get('sample_count')
        if image.is_file() and type(count) is int and count > 0:
            sources['report.' + key] = {'status': 'available', 'sample_count': count,
                                        'report_url': manifest.get('report_url'), 'image': str(image.resolve()),
                                        'sha256': file_sha(image)}
        else:
            sources['report.' + key] = {'status': 'unavailable', 'reason': '未取得本批有效图表及实际样本数'}
    packet = {'schema_version': 1, 'workbook_sha256': file_sha(workbook), 'sheet': sheet.title,
              'as_of': as_of, 'layout': layout, 'sources': sources,
              'instructions': '由AI逐项阅读本批证据后写七项结论；源表和网页报告样本分开。不得用源表统计冒充报告图数据，不作销量因果或利润保证。缺证据明确写无法判断，不能留空。'}
    packet['packet_digest'] = digest(packet)
    return packet


def check_text_fits(text, width, height, font_points):
    """保守估算原字号下的换行，不缩字号、不扩行列；超长时要求AI精简。"""
    units_per_line = max(1, int((width - 12) / (font_points * 96 / 72 * 0.55)))
    lines = sum(max(1, math.ceil(sum(2 if unicodedata.east_asian_width(c) in 'WF' else 1 for c in paragraph)
                               / units_per_line)) for paragraph in text.split('\n'))
    if lines * font_points * 96 / 72 * 1.3 > height - 8:
        raise ValueError('分析文字超过现有单元格容量，请精简，不调整行高列宽或字号')


def highlight_segments(text, highlights):
    """只标红正文中的明确重点，拒绝不存在、重复、交叠或整段标红。"""
    if not isinstance(highlights, list) or not highlights or any(not isinstance(h, str) or not h.strip() for h in highlights):
        raise ValueError('每节必须指定正文中的重点标红片段')
    if len(set(highlights)) != len(highlights):
        raise ValueError('重点片段重复')
    spans = []
    for phrase in highlights:
        matches = list(re.finditer(re.escape(phrase), text))
        if not matches or '\n' in phrase:
            raise ValueError('重点片段必须是正文中存在的单行文字')
        spans.extend((match.start(), match.end()) for match in matches)
    spans.sort()
    if any(left[1] > right[0] for left, right in zip(spans, spans[1:])):
        raise ValueError('重点片段交叠')
    if sum(end - start for start, end in spans) > len(text) / 2:
        raise ValueError('重点标红超过正文一半，请只标关键数字、判断或建议')
    result, start = [], 0
    for left, right in spans:
        if left > start:
            result.append((text[start:left], False))
        result.append((text[left:right], True))
        start = right
    if start < len(text):
        result.append((text[start:], False))
    return result


def write_rich_text(cell, text, highlights, font):
    """沿用模板字号、字体及其它字形，仅将选中重点改为红色。"""
    inline = element(cell, 'is')
    for content, red in highlight_segments(text, highlights):
        run = element(inline, 'r')
        properties = element(run, 'rPr')
        for child in font.childNodes:
            if child.nodeType != child.ELEMENT_NODE:
                continue
            if child.localName == 'name':
                properties.appendChild(element(properties, 'rFont', dict(child.attributes.items())))
            elif not (red and child.localName == 'color'):
                properties.appendChild(cell.ownerDocument.importNode(child, True))
        if red:
            properties.appendChild(element(properties, 'color', {'rgb': 'FFFF0000'}))
        run.appendChild(properties)
        node = element(run, 't', {'xml:space': 'preserve'})
        node.appendChild(cell.ownerDocument.createTextNode(content))
        run.appendChild(node)
        inline.appendChild(run)
    cell.appendChild(inline)


def validate_analysis(packet, analysis):
    unsigned = {k: v for k, v in packet.items() if k != 'packet_digest'}
    if packet.get('packet_digest') != digest(unsigned) or analysis.get('packet_digest') != packet['packet_digest']:
        raise ValueError('分析与本批证据指纹不一致')
    sections = analysis.get('sections', {})
    if set(sections) != set(SECTIONS):
        raise ValueError('必须实际完成全部七项市场分析及总结')
    for key, (_, primary) in SECTIONS.items():
        section = sections[key]
        text = section.get('text', '')
        refs = section.get('evidence', [])
        if not isinstance(text, str) or not text.strip() or re.search(r'待填写|待分析|TODO|占位文字', text, re.I):
            raise ValueError(f'{key}缺少实际分析正文')
        highlight_segments(text.strip(), section.get('highlights'))
        if not isinstance(refs, list) or primary not in refs or any(ref not in packet['sources'] for ref in refs):
            raise ValueError(f'{key}缺少可追溯的本批来源')
        available = packet['sources'][primary]['status'] == 'available'
        if section.get('status') != ('analyzed' if available else 'unavailable'):
            raise ValueError(f'{key}可用状态与证据不符')
        if not available and not re.search(r'未取得|不足|无法|缺失|未提供', text):
            raise ValueError(f'{key}必须明确说明数据缺失')
        layout = packet['layout'][key]
        check_text_fits(text, *layout['pixels'], layout['font_points'])
    return {key: section['text'].strip() for key, section in sections.items()}


def fill_narrative(workbook, output, packet, analysis, *, replace_existing=False):
    workbook, output = Path(workbook), Path(output)
    if output.exists() or workbook.resolve() == output.resolve():
        raise ValueError('请使用尚不存在的独立输出文件')
    if file_sha(workbook) != packet.get('workbook_sha256'):
        raise ValueError('工作簿已变化，请重新准备本批证据')
    texts = validate_analysis(packet, analysis)
    for source in packet['sources'].values():
        if source.get('image') and file_sha(source['image']) != source['sha256']:
            raise ValueError('报告图像已变化，请重新准备证据')
    book = load_workbook(workbook, rich_text=True)
    try:
        sheet = market_sheet(book)
        regions = locate_narrative_regions(sheet)
        for key, area in regions.items():
            cell = sheet.cell(area.min_row, area.min_col)
            if str(area) != packet['layout'][key]['range']:
                raise ValueError('当前结论区域与本批证据不一致')
            if cell.data_type == 'f' or (cell.value not in (None, '') and not replace_existing):
                raise ValueError(f'{cell.coordinate}已有内容，拒绝覆盖；重写文字需显式--replace-existing')
        sheet_name = sheet.title
    finally:
        book.close()
    output.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(workbook) as original:
        part = worksheet_part(original, sheet_name)
        doc = minidom.parseString(original.read(part))
        styles = minidom.parseString(original.read('xl/styles.xml'))
        style_root = styles.documentElement
        fonts = children(children(style_root, 'fonts')[0], 'font')
        formats = children(children(style_root, 'cellXfs')[0], 'xf')
        data = children(doc.documentElement, 'sheetData')[0]
        rows = {int(row.getAttribute('r')): row for row in children(data, 'row')}
        for key, area in regions.items():
            address = area.start_cell.coordinate
            row = rows[area.min_row]
            cells = children(row, 'c')
            cell = next((c for c in cells if c.getAttribute('r') == address), None)
            if cell is None:
                raise ValueError(f'原模板缺少结论单元格{address}，请先补齐模板格式')
            for child in list(cell.childNodes):
                cell.removeChild(child)
            cell.setAttribute('t', 'inlineStr')
            font_id = int(formats[int(cell.getAttribute('s') or '0')].getAttribute('fontId') or '0')
            write_rich_text(cell, texts[key], analysis['sections'][key]['highlights'], fonts[font_id])
        with tempfile.NamedTemporaryFile(dir=output.parent, suffix='.xlsx', delete=False) as temporary:
            staged = Path(temporary.name)
        try:
            with ZipFile(staged, 'w') as destination:
                for info in original.infolist():
                    destination.writestr(info, doc.toxml(encoding='utf-8') if info.filename == part else original.read(info.filename))
            with ZipFile(staged) as check:
                if check.testzip():
                    raise ValueError('分析版工作簿ZIP校验失败')
            staged.replace(output)
        finally:
            staged.unlink(missing_ok=True)
    return {'output': str(output.resolve()), 'sheet': sheet_name,
            'written': {key: str(area) for key, area in regions.items()}, 'packet_digest': packet['packet_digest']}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest='command', required=True)
    prepare = sub.add_parser('prepare', help='提取本批证据，随后由AI实际分析')
    prepare.add_argument('--workbook', required=True, type=Path)
    prepare.add_argument('--as-of', required=True, help='本批采集日期YYYY-MM-DD')
    prepare.add_argument('--report-manifest', type=Path)
    prepare.add_argument('--output', required=True, type=Path)
    fill = sub.add_parser('fill', help='核对本批指纹并定点写入AI结论')
    fill.add_argument('--workbook', required=True, type=Path)
    fill.add_argument('--packet', required=True, type=Path)
    fill.add_argument('--analysis', required=True, type=Path)
    fill.add_argument('--output', required=True, type=Path)
    fill.add_argument('--replace-existing', action='store_true')
    args = parser.parse_args()
    if args.command == 'prepare':
        if args.output.exists() or args.output.resolve() == args.workbook.resolve() or args.output.suffix.lower() != '.json':
            parser.error('证据输出必须是尚不存在的独立JSON文件')
        packet = prepare_evidence(args.workbook, as_of=args.as_of, report_manifest=args.report_manifest)
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(json.dumps(packet, ensure_ascii=False, indent=2), encoding='utf-8')
        print(json.dumps({'packet': str(args.output.resolve()), 'packet_digest': packet['packet_digest']}, ensure_ascii=False))
    else:
        packet = json.loads(args.packet.read_text(encoding='utf-8-sig'))
        analysis = json.loads(args.analysis.read_text(encoding='utf-8-sig'))
        print(json.dumps(fill_narrative(args.workbook, args.output, packet, analysis,
                                        replace_existing=args.replace_existing), ensure_ascii=False))


if __name__ == '__main__':
    main()
