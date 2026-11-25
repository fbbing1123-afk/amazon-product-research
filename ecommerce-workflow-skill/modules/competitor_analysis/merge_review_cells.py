"""把旧报告拆散的评价续行合为每商品一个单元格，保留全部文本与其它部件。"""
from __future__ import annotations

import argparse
from copy import copy
import json
from pathlib import Path
import sys
import tempfile
from xml.dom import minidom
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from competitor_analysis.annotate_review_export import worksheet_part,children,element
from competitor_analysis.populate_basics import competitor_asins,find_label_row


def review_regions(sheet):
    """按实际标签及商品链接定位，不把续行数量或商品列写死。"""
    starts=[find_label_row(sheet,label) for label in ('优点','缺点','总结及改进方式')]
    if not starts[0]<starts[1]<starts[2]:
        raise ValueError('评价区域顺序异常')
    link_row=find_label_row(sheet,'详情页链接')
    columns=[]
    for asin in competitor_asins(sheet):
        matches=[c.column for c in sheet[link_row] if asin in str(c.hyperlink.target if c.hyperlink else c.value or '')]
        if len(matches)!=1:
            raise ValueError(f'{asin}的商品链接列不唯一')
        columns.append(matches[0])
    groups=[]
    for label,start,end in [('优点',starts[0],starts[1]-1),('缺点',starts[1],starts[2]-1)]:
        for column in columns:
            cells=[sheet.cell(row,column) for row in range(start,end+1)]
            if any(c.data_type=='f' for c in cells):
                raise ValueError('评价区域含公式，不能作为普通文字合并')
            values=[str(c.value).strip() for c in cells if c.value not in (None,'')]
            if not values or ('/' in values and len(values)>1):
                raise ValueError('评价内容为空或缺点占位符与正文冲突')
            text='\n'.join(values)
            if len(text)>32767:
                raise ValueError('评价超出Excel单元格文字上限')
            merges=[]
            for area in sheet.merged_cells.ranges:
                if area.min_col<=column<=area.max_col and area.min_row<=end and area.max_row>=start:
                    if area.min_col!=column or area.max_col!=column or area.min_row<start or area.max_row>end:
                        raise ValueError('评价合并区域跨商品或跨模块，停止修改')
                    merges.append(str(area))
            letter=get_column_letter(column)
            groups.append({'label':label,'cells':[c.coordinate for c in cells],'value':text,
                           'merge':f'{letter}{start}:{letter}{end}' if end>start else None,'old_merges':merges})
    return groups


def merge_review_cells(workbook,output):
    if output.exists() or workbook.resolve()==output.resolve():
        raise ValueError('修正版必须使用新的独立输出文件')
    book=load_workbook(workbook,rich_text=True)
    try:
        sheet=next(s for s in book if s.title.endswith('竞品分析'))
        name,groups=sheet.title,review_regions(sheet)
    finally:
        book.close()
    output.parent.mkdir(parents=True,exist_ok=True)
    with ZipFile(workbook) as original,tempfile.TemporaryDirectory(dir=output.parent) as folder:
        part=worksheet_part(original,name)
        doc=minidom.parseString(original.read(part))
        cells={c.getAttribute('r'):c for c in doc.getElementsByTagNameNS(doc.documentElement.namespaceURI,'c')}
        merged=children(doc.documentElement,'mergeCells')
        if len(merged)!=1:
            raise ValueError('模板缺少唯一的合并区域定义')
        merged=merged[0]
        for group in groups:
            for index,address in enumerate(group['cells']):
                cell=cells.get(address)
                if cell is None:
                    if index==0:
                        raise ValueError(f'缺少评价主单元格：{address}')
                    continue
                for child in list(cell.childNodes):
                    cell.removeChild(child)
                cell.setAttribute('t','inlineStr')
                inline=element(cell,'is')
                text=element(inline,'t',{'xml:space':'preserve'})
                text.appendChild(doc.createTextNode(group['value'] if index==0 else ''))
                inline.appendChild(text)
                cell.appendChild(inline)
            for merge in list(children(merged,'mergeCell')):
                if merge.getAttribute('ref') in group['old_merges']:
                    merged.removeChild(merge)
            if group['merge']:
                merged.appendChild(element(merged,'mergeCell',{'ref':group['merge']}))
        merged.setAttribute('count',str(len(children(merged,'mergeCell'))))
        staged=Path(folder)/'merged.xlsx'
        with ZipFile(staged,'w') as result:
            for info in original.infolist():
                result.writestr(copy(info),doc.toxml(encoding='utf-8') if info.filename==part else original.read(info.filename))
        checked=load_workbook(staged,rich_text=True)
        try:
            sheet=checked[name]
            for group in groups:
                if sheet[group['cells'][0]].value!=group['value'] or (group['merge'] and group['merge'] not in sheet.merged_cells):
                    raise ValueError('评价单元格合并后内容不一致')
        finally:
            checked.close()
        with ZipFile(staged) as result:
            if original.namelist()!=result.namelist() or result.testzip():
                raise ValueError('工作簿ZIP结构检查失败')
            if any(original.read(n)!=result.read(n) for n in original.namelist() if n!=part):
                raise ValueError('评价之外的工作簿部件被修改')
        staged.replace(output)
    return {'output':str(output),'sheet':name,'cells':[g['cells'][0] for g in groups],
            'merged_regions':[g['merge'] for g in groups if g['merge']]}


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--workbook',type=Path,required=True)
    parser.add_argument('--output',type=Path,required=True)
    args=parser.parse_args()
    print(json.dumps(merge_review_cells(args.workbook,args.output),ensure_ascii=False))


if __name__=='__main__':
    main()
