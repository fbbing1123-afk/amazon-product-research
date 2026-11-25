"""复核旧源和旧注释后，定点更新一个商品的分析列，原评论及其它部件不变。"""
from copy import copy
from pathlib import Path
import tempfile
from xml.dom import minidom
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from competitor_analysis.annotate_review_export import HEADERS,annotation_rows,children,element,worksheet_part,required_height
from competitor_analysis.merge_review_sources import ASIN_HEADER,SHEET_NAME,verify_analysis
from review_research.download_sellersprite_reviews import read_export


def refresh_annotations(workbook,output,previous,updated):
    """仅更新已核实的同批证据；遇人工改过的原文或分析说明时停止，不能覆盖。"""
    if output.exists() or output.resolve()==workbook.resolve():
        raise ValueError('分析记录修正版必须另存到独立新文件')
    packet,old=verify_analysis(previous)
    new_packet,new=verify_analysis(updated)
    if packet['packet_digest']!=new_packet['packet_digest'] or packet['asin']!=new_packet['asin']:
        raise ValueError('前后分析不属于同一商品和同一评论批次')
    source=read_export(Path(packet['source_file']),packet['asin'])
    old_notes,new_notes=annotation_rows(packet,old,source),annotation_rows(packet,new,source)
    book=load_workbook(workbook,rich_text=True)
    try:
        sheet=book[SHEET_NAME]
        headers={c.value:c.column for c in sheet[1] if c.value}
        if len(headers)!=sheet.max_column or not {ASIN_HEADER,*HEADERS,*source['headers']}<=set(headers):
            raise ValueError('评价数据源表头缺失或重复')
        selected=[r for r in range(2,sheet.max_row+1) if sheet.cell(r,headers[ASIN_HEADER]).value==packet['asin']]
        if len(selected)!=len(source['reviews']):
            raise ValueError('报告中的本商品评论数量与原始源不一致')
        updates,heights={},{}
        for row,review in zip(selected,source['reviews']):
            for header in source['headers']:
                actual=sheet.cell(row,headers[header]).value
                expected=review['raw'].get(header)
                if (actual if actual is not None else '')!=(expected if expected is not None else ''):
                    raise ValueError(f'第{row}行原评论字段已变化，拒绝按旧证据写入')
            for header,expected,value in zip(HEADERS,old_notes[review['source_row']],new_notes[review['source_row']]):
                cell=sheet.cell(row,headers[header])
                if (cell.value or '')!=expected:
                    raise ValueError(f'{cell.coordinate}旧分析已变化，不能覆盖人工修改')
                if expected!=value:
                    updates[cell.coordinate]=(value,cell.style_id if cell.has_style else sheet.cell(row,headers[ASIN_HEADER]).style_id)
            if old_notes[review['source_row']]!=new_notes[review['source_row']]:
                needed=required_height(new_notes[review['source_row']])
                if needed>(sheet.row_dimensions[row].height or 15):
                    heights[row]=needed
    finally:
        book.close()
    output.parent.mkdir(parents=True,exist_ok=True)
    with ZipFile(workbook) as archive,tempfile.TemporaryDirectory(dir=output.parent) as folder:
        part=worksheet_part(archive,SHEET_NAME)
        doc=minidom.parseString(archive.read(part))
        rows={int(r.getAttribute('r')):r for r in doc.getElementsByTagNameNS(doc.documentElement.namespaceURI,'row')}
        for number,height in heights.items():
            rows[number].setAttribute('ht',str(height))
            rows[number].setAttribute('customHeight','1')
        for address,(value,style) in updates.items():
            column=''.join(c for c in address if c.isalpha())
            number=int(address[len(column):])
            row=rows[number]
            cell=next((c for c in children(row,'c') if c.getAttribute('r')==address),None)
            if cell is None:
                cell=element(row,'c',{'r':address,'s':style})
                next_cell=next((c for c in children(row,'c') if column_index_from_string(''.join(x for x in c.getAttribute('r') if x.isalpha()))>column_index_from_string(column)),None)
                row.insertBefore(cell,next_cell) if next_cell else row.appendChild(cell)
            for child in list(cell.childNodes):
                cell.removeChild(child)
            cell.setAttribute('t','inlineStr')
            inline=element(cell,'is')
            text=element(inline,'t',{'xml:space':'preserve'})
            text.appendChild(doc.createTextNode(value))
            inline.appendChild(text)
            cell.appendChild(inline)
        staged=Path(folder)/'updated.xlsx'
        with ZipFile(staged,'w') as target:
            for info in archive.infolist():
                target.writestr(copy(info),doc.toxml(encoding='utf-8') if info.filename==part else archive.read(info.filename))
        checked=load_workbook(staged,rich_text=True)
        try:
            for address,(value,_) in updates.items():
                if checked[SHEET_NAME][address].value!=value:
                    raise ValueError(f'{address}新分析写入校验失败')
        finally:
            checked.close()
        with ZipFile(staged) as result:
            if result.testzip() or any(result.read(n)!=archive.read(n) for n in archive.namelist() if n!=part):
                raise ValueError('分析列之外的工作簿部件被修改')
        staged.replace(output)
    return {'output':str(output),'asin':packet['asin'],'updated_cells':list(updates),'expanded_rows':heights}
