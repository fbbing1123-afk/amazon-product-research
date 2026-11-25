"""回填本批已校验评价，每款优缺点各一个单元格，长内容纵向合并扩高。"""
from __future__ import annotations

import argparse
import json
import math
import re
import sys
import tempfile
import unicodedata
from copy import copy
from pathlib import Path

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from competitor_analysis.populate_basics import competitor_asins, find_label_row
from competitor_analysis.review_mentions import report_points
from competitor_analysis.merge_review_sources import merge_sources, verify_analysis
from shared.worksheet_rows import insert_rows_preserving_merges
from shared.quality_gates import validate_competitor_review_analysis


def text_units(text):
    """按中西文字符宽度估算换行，留出足够的显示余量。"""
    return sum(2 if unicodedata.east_asian_width(c) in 'WF' else 1 for c in text)


def chunk_points(sheet, row, column, points):
    cell = sheet.cell(row, column)
    size = cell.font.sz or 11
    width = sheet.column_dimensions[cell.column_letter].width or 13
    # 列宽以默认字体字符数表示，换算后预留边距，禁止缩小字体来塞入文字。
    capacity = max(8, int((width * 7 - 12) / (size * 0.8)))
    max_lines = max(1, int(((sheet.row_dimensions[row].height or 15) - 12) / (size * 1.6)))
    chunks, current, lines = [], [], 0
    for point in points:
        needed = max(1, math.ceil(text_units(point) / capacity))
        if needed > max_lines:
            raise ValueError('单条评价超过模板一行容量，请精简结论而非截断文字')
        if current and lines + needed > max_lines:
            chunks.append('\n'.join(current))
            current, lines = [], 0
        current.append(point)
        lines += needed
    if current:
        chunks.append('\n'.join(current))
    return chunks or ['/']


def format_summary(summary):
    """要求实际精简总结和编号改进建议，不把建议列表冒充总结。"""
    lines = [line.rstrip() for line in summary.strip().splitlines()
             if not re.match(r'^\s*依据\s*[:：]', line)]
    text = '\n'.join(lines).strip()
    match = re.fullmatch(
        r'总结[ \t]*[:：][ \t]*(?P<overview>[^\n]+)\n'
        r'[ \t]*改进方式[ \t]*[:：][ \t]*\n?(?P<actions>[\s\S]+)', text)
    if not match or not match['overview'].strip() or not re.match(r'1[.、]\s*\S', match['actions'].strip()):
        raise ValueError('请提供“总结：实际精简总结”及“改进方式：”下的编号建议')
    return '总结：' + match['overview'].strip() + '\n改进方式：\n' + match['actions'].strip()


def fill_batch(workbook, output, analyses, summary):
    if output.exists() or output.resolve() == workbook.resolve():
        raise ValueError('输出必须是新的独立文件')
    # 回填前再核验源文件指纹和证据，不能把旧 checked.json 直接套到新评论。
    analyses = [verify_analysis(analysis)[1] for analysis in analyses]
    book = load_workbook(workbook, rich_text=True)
    try:
        sheet = next(s for s in book if s.title.endswith('竞品分析'))
        asins = competitor_asins(sheet)
        by_asin = {a['asin']: a for a in analyses}
        if len(by_asin) != len(analyses) or set(by_asin) != set(asins):
            raise ValueError('本批分析与模板竞品不一一对应')
        for kind, label in [('advantages', '优点'), ('disadvantages', '缺点')]:
            row = find_label_row(sheet, label)
            blocks = {}
            for column, asin in enumerate(asins, start=3):
                points = report_points(by_asin[asin],kind)
                blocks[column] = chunk_points(sheet, row, column, points)
            extra = max(len(chunks) for chunks in blocks.values()) - 1
            if extra:
                insert_rows_preserving_merges(sheet, row + 1, extra, row)
                # 扩展评价总标签和当前优缺点标签，后续正文与图片已经由公共函数下移。
                evaluation = find_label_row(sheet, '评价分析')
                summary_row = find_label_row(sheet, '总结及改进方式')
                label_style = copy(sheet.cell(row, 2)._style)
                for area in list(sheet.merged_cells.ranges):
                    if area.min_col == area.max_col == 1 and area.min_row == evaluation:
                        sheet.unmerge_cells(str(area))
                sheet.merge_cells(start_row=evaluation, end_row=summary_row - 1, start_column=1, end_column=1)
                sheet.merge_cells(start_row=row, end_row=row + extra, start_column=2, end_column=2)
                sheet.cell(row, 2)._style = label_style
            for column, chunks in blocks.items():
                # 扩行只用于提供高度，完整内容始终写在该商品的同一个单元格里。
                sheet.cell(row, column, '\n'.join(chunks))
                if extra:
                    sheet.merge_cells(start_row=row,end_row=row+extra,start_column=column,end_column=column)
        summary_row = find_label_row(sheet, '总结及改进方式')
        sheet.cell(summary_row, 3, format_summary(summary))
        output.parent.mkdir(parents=True, exist_ok=True)
        # 先保存临时回填结果，评论合并和逐行验收成功后才产生最终文件。
        with tempfile.TemporaryDirectory(dir=output.parent) as temporary:
            staged = Path(temporary) / 'filled.xlsx'
            book.save(staged)
            source_result = merge_sources(staged, output, analyses)
    finally:
        book.close()
    validate_competitor_review_analysis(output)
    return {'output': str(output), 'products': len(asins), 'summary_row': summary_row,
            'review_source_sheet': source_result['sheet'], 'review_source_rows': source_result['reviews']}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--workbook', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--analyses', type=Path, nargs='+', required=True)
    parser.add_argument('--summary', type=Path, required=True)
    args = parser.parse_args()
    analyses = [json.loads(p.read_text(encoding='utf-8')) for p in args.analyses]
    print(json.dumps(fill_batch(args.workbook, args.output, analyses, args.summary.read_text(encoding='utf-8')), ensure_ascii=False))


if __name__ == '__main__':
    main()
