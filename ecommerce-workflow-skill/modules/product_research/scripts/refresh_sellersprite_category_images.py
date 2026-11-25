"""仅替换类目截图和销量占比，保留其它工作簿部件，图片原尺寸浮动显示。"""
from __future__ import annotations

import argparse
from copy import copy
from io import BytesIO
import json
import posixpath
import re
import sys
import tempfile
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl import load_workbook
from PIL import Image

sys.path.insert(0,str(Path(__file__).resolve().parents[2]))
from product_research.scripts.capture_category_overviews import iter_category_rows
from product_research.scripts.refresh_category_shares import refresh_category_shares
from shared.quality_gates import validate_category_research

MAIN='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
OFFICE='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
DRAW='http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
ART='http://schemas.openxmlformats.org/drawingml/2006/main'


def resolve(owner,target):
    return posixpath.normpath(posixpath.join(posixpath.dirname(owner),target)).lstrip('/')


def relations(archive,owner):
    folder,name=posixpath.split(owner)
    root=ET.fromstring(archive.read(f'{folder}/_rels/{name}.rels'))
    return {r.get('Id'):resolve(owner,r.get('Target')) for r in root}


def refresh_images(workbook,output,results):
    if output.exists() or output.resolve()==workbook.resolve():
        raise ValueError('输出须为新的独立文件')
    book=load_workbook(workbook,rich_text=True)
    try:
        sheet=next(s for s in book if s.title.endswith('市场调研'))
        name=sheet.title
        items=list(iter_category_rows(sheet))
    finally:
        book.close()
    selected={item['row']:results[item['url']] for item in items if item['url'] in results}
    if len(selected)!=len(items) or any(r.get('status')!='ok' for r in selected.values()):
        raise ValueError('所有类目必须有本批成功截图')
    sales={re.search(r'/(\d+)(?:/|$)',item['url'])[1]:results[item['url']]['monthly_sales'] for item in items}
    output.parent.mkdir(parents=True,exist_ok=True)
    with ZipFile(workbook) as archive, tempfile.TemporaryDirectory(dir=output.parent) as temp:
        root=ET.fromstring(archive.read('xl/workbook.xml'))
        desc=next(s for s in root.findall(f'.//{{{MAIN}}}sheet') if s.get('name')==name)
        part=relations(archive,'xl/workbook.xml')[desc.get(f'{{{OFFICE}}}id')]
        sheet_root=ET.fromstring(archive.read(part))
        drawing=sheet_root.find(f'{{{MAIN}}}drawing')
        drawing_part=relations(archive,part)[drawing.get(f'{{{OFFICE}}}id')]
        drawing_root=ET.fromstring(archive.read(drawing_part))
        image_rels=relations(archive,drawing_part)
        changes,seen={},set()
        for anchor in drawing_root:
            start=anchor.find(f'{{{DRAW}}}from')
            if start is None:
                continue
            row=int(start.find(f'{{{DRAW}}}row').text)+1
            column=int(start.find(f'{{{DRAW}}}col').text)+1
            if row not in selected or column!=6:
                continue
            if row in seen:
                raise ValueError(f'第{row}行存在重复类目图片')
            seen.add(row)
            result=selected[row]
            content=Path(result['image']).read_bytes()
            with Image.open(BytesIO(content)) as picture:
                width,height=picture.size
            picture_part=image_rels[anchor.find(f'.//{{{ART}}}blip').get(f'{{{OFFICE}}}embed')]
            changes[picture_part]=content
            # 显示框与PNG像素1:1对应，不变更行高、列宽或合并区域。
            extent=anchor.find(f'{{{DRAW}}}ext')
            if extent is None:
                raise ValueError('类目图片必须是单单元格浮动锚点')
            extent.set('cx',str(width*9525))
            extent.set('cy',str(height*9525))
            props=anchor.find(f'.//{{{DRAW}}}cNvPr')
            if props is not None:
                props.set('descr',f"类目原尺寸截图：{result.get('label','')}；{width}×{height}")
        if seen!=set(selected):
            raise ValueError('未找到全部类目原图片锚点')
        changes[drawing_part]=ET.tostring(drawing_root,encoding='utf-8',xml_declaration=True)
        staged=Path(temp)/'images.xlsx'
        with ZipFile(staged,'w') as patched:
            for info in archive.infolist():
                patched.writestr(copy(info),changes.get(info.filename,archive.read(info.filename)))
        refresh_category_shares(staged,staged,output,sales_by_node=sales)
    return {'output':str(output),'images':len(seen),'original_size':True}


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--workbook',type=Path,required=True)
    parser.add_argument('--manifest',type=Path,required=True)
    parser.add_argument('--output',type=Path,required=True)
    args=parser.parse_args()
    result = refresh_images(args.workbook,args.output,json.loads(args.manifest.read_text(encoding='utf-8')))
    result['gate'] = validate_category_research(args.output).as_dict()
    print(json.dumps(result,ensure_ascii=False,default=str))


if __name__=='__main__':
    main()
