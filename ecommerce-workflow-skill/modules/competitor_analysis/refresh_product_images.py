"""按竞品ASIN定点写入高清主图，原始像素保留，不依赖商品数据源中的行。"""
from __future__ import annotations

import argparse
from copy import copy
from io import BytesIO
import json
from pathlib import Path
import posixpath
import sys
import tempfile
from xml.dom import minidom
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.drawing.image import Image

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from competitor_analysis.annotate_review_export import worksheet_part, children
from competitor_analysis.collect_product_images import validate_asset
from competitor_analysis.collect_basic_facts import asin_from_input
from competitor_analysis.populate_basics import competitor_asins, find_label_row, fit_image_to_cell

MAIN='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
OFFICE='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PACKAGE='http://schemas.openxmlformats.org/package/2006/relationships'
DRAW='http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
ART='http://schemas.openxmlformats.org/drawingml/2006/main'
CONTENT='http://schemas.openxmlformats.org/package/2006/content-types'


def relation_part(owner):
    folder, name=posixpath.split(owner)
    return f'{folder}/_rels/{name}.rels'


def related_document(archive, owner):
    part=relation_part(owner)
    return minidom.parseString(archive.read(part)) if part in archive.namelist() else minidom.parseString(f'<Relationships xmlns="{PACKAGE}"/>')


def add_relation(doc, kind, target):
    ids={n.getAttribute('Id') for n in children(doc.documentElement,'Relationship')}
    index=1
    while f'rIdProduct{index}' in ids:
        index+=1
    node=doc.createElementNS(PACKAGE,'Relationship')
    for key,value in {'Id':f'rIdProduct{index}','Type':OFFICE+'/'+kind,'Target':target}.items():
        node.setAttribute(key,value)
    doc.documentElement.appendChild(node)
    return f'rIdProduct{index}'


def append_node(parent, namespace, tag, attrs=None, text=None):
    node=parent.ownerDocument.createElementNS(namespace,tag)
    for key,value in (attrs or {}).items():
        node.setAttribute(key,str(value))
    if text is not None:
        node.appendChild(parent.ownerDocument.createTextNode(str(text)))
    parent.appendChild(node)
    return node


def refresh_product_images(workbook, output, assets, *, display_mode='original', allow_partial=False):
    if output.exists() or output.resolve()==workbook.resolve():
        raise ValueError('高清主图报告必须另存到独立新文件')
    if display_mode not in ('fit','original'):
        raise ValueError('图片显示方式只能为fit或original')
    book=load_workbook(workbook,rich_text=True)
    try:
        sheets=[s for s in book if s.title.endswith('竞品分析')]
        if len(sheets)!=1:
            raise ValueError('必须有唯一的竞品分析页')
        sheet=sheets[0]
        asins=competitor_asins(sheet)
        if not assets or set(assets)-set(asins) or (not allow_partial and set(assets)!=set(asins)):
            raise ValueError('主图与竞品集合不一致')
        row=find_label_row(sheet,'产品图')
        link_row=find_label_row(sheet,'详情页链接')
        prepared={}
        for asin,asset in assets.items():
            columns=[]
            for cell in sheet[link_row]:
                try:
                    if asin_from_input(str(cell.hyperlink.target if cell.hyperlink else cell.value or ''))==asin:
                        columns.append(cell.column)
                except ValueError:
                    continue
            if len(columns)!=1:
                raise ValueError(f'{asin}详情链接列不唯一')
            payload=validate_asset(asset,asin)
            image=Image(BytesIO(payload))
            if display_mode=='fit':
                fit_image_to_cell(image,sheet,row,columns[0])
            prepared[asin]={'asset':asset,'payload':payload,'column':columns[0],
                            'size':(image.width,image.height),'extension':'jpg' if image.format=='jpeg' else 'png'}
        name=sheet.title
        market_sheets=[s for s in book if s.title.endswith('市场调研')]
        if len(market_sheets)>1:
            raise ValueError('市场调研页不唯一')
        link_order=[]
        for cell in sheet[link_row]:
            try:
                link_order.append(asin_from_input(str(cell.hyperlink.target if cell.hyperlink else cell.value or '')))
            except ValueError:
                continue
        update_cover=bool(market_sheets) and bool(link_order) and link_order[0] in prepared
    finally:
        book.close()
    output.parent.mkdir(parents=True,exist_ok=True)
    with ZipFile(workbook) as archive,tempfile.TemporaryDirectory(dir=output.parent) as directory:
        part=worksheet_part(archive,name)
        sheet_doc=minidom.parseString(archive.read(part))
        sheet_rels=related_document(archive,part)
        drawing_nodes=sheet_doc.getElementsByTagNameNS(MAIN,'drawing')
        changes={}
        types=minidom.parseString(archive.read('[Content_Types].xml'))
        def register_type(path,content_type):
            if any(n.getAttribute('PartName')=='/'+path for n in children(types.documentElement,'Override')):
                return
            append_node(types.documentElement,CONTENT,'Override',{'PartName':'/'+path,'ContentType':content_type})
        if len(drawing_nodes)>1:
            raise ValueError('竞品页绘图部件不唯一')
        if drawing_nodes:
            rid=drawing_nodes[0].getAttributeNS(OFFICE,'id')
            relation=next(r for r in children(sheet_rels.documentElement,'Relationship') if r.getAttribute('Id')==rid)
            target=relation.getAttribute('Target')
            drawing_part=target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(part),target))
            drawing_doc=minidom.parseString(archive.read(drawing_part))
        else:
            index=1
            while f'xl/drawings/productImages{index}.xml' in archive.namelist():
                index+=1
            drawing_part=f'xl/drawings/productImages{index}.xml'
            drawing_doc=minidom.parseString(f'<xdr:wsDr xmlns:xdr="{DRAW}" xmlns:a="{ART}" xmlns:r="{OFFICE}"/>')
            rid=add_relation(sheet_rels,'drawing',posixpath.relpath(drawing_part,posixpath.dirname(part)))
            root=sheet_doc.documentElement
            root.setAttribute('xmlns:r',OFFICE)
            drawing=sheet_doc.createElementNS(MAIN,(root.prefix+':' if root.prefix else '')+'drawing')
            drawing.setAttributeNS(OFFICE,'r:id',rid)
            following=next((n for n in root.childNodes if n.nodeType==n.ELEMENT_NODE and n.localName in
                           {'legacyDrawing','legacyDrawingHF','picture','oleObjects','controls','webPublishItems','tableParts','extLst'}),None)
            root.insertBefore(drawing,following) if following else root.appendChild(drawing)
            changes[part]=sheet_doc.toxml(encoding='utf-8')
            changes[relation_part(part)]=sheet_rels.toxml(encoding='utf-8')
            register_type(drawing_part,'application/vnd.openxmlformats-officedocument.drawing+xml')
        drawing_rels=related_document(archive,drawing_part)
        drawing_root=drawing_doc.documentElement
        for prefix,namespace in [('xdr',DRAW),('a',ART),('r',OFFICE)]:
            drawing_root.setAttribute('xmlns:'+prefix,namespace)
        object_id=max([int(n.getAttribute('id') or 0) for n in drawing_doc.getElementsByTagNameNS(DRAW,'cNvPr')]+[0])
        for asin,item in prepared.items():
            matches=[]
            for anchor in children(drawing_root,'oneCellAnchor')+children(drawing_root,'twoCellAnchor'):
                start=children(anchor,'from')
                if start and int(children(start[0],'row')[0].firstChild.data)==row-1 and int(children(start[0],'col')[0].firstChild.data)==item['column']-1:
                    matches.append(anchor)
            if len(matches)>1 or any(not children(anchor,'pic') for anchor in matches):
                raise ValueError(f'{asin}主图位置重复或被其它绘图占用')
            for anchor in matches:
                drawing_root.removeChild(anchor)
            asset=item['asset']
            media=f"xl/media/product-{asin}-{asset['sha256'][:16]}.{item['extension']}"
            if media in archive.namelist() and archive.read(media)!=item['payload']:
                raise ValueError('已有主图媒体文件与记录不一致')
            changes[media]=item['payload']
            register_type(media,'image/jpeg' if item['extension']=='jpg' else 'image/png')
            rid=add_relation(drawing_rels,'image',posixpath.relpath(media,posixpath.dirname(drawing_part)))
            object_id+=1
            anchor=append_node(drawing_root,DRAW,'xdr:oneCellAnchor')
            start=append_node(anchor,DRAW,'xdr:from')
            for tag,value in [('col',item['column']-1),('colOff',0),('row',row-1),('rowOff',0)]:
                append_node(start,DRAW,'xdr:'+tag,text=value)
            width,height=item['size']
            append_node(anchor,DRAW,'xdr:ext',{'cx':round(width*9525),'cy':round(height*9525)})
            pic=append_node(anchor,DRAW,'xdr:pic')
            nv=append_node(pic,DRAW,'xdr:nvPicPr')
            append_node(nv,DRAW,'xdr:cNvPr',{'id':object_id,'name':f'产品主图 {asin}','descr':asset.get('alt') or asin})
            append_node(nv,DRAW,'xdr:cNvPicPr')
            fill=append_node(pic,DRAW,'xdr:blipFill')
            append_node(fill,ART,'a:blip',{'r:embed':rid})
            append_node(append_node(fill,ART,'a:stretch'),ART,'a:fillRect')
            props=append_node(pic,DRAW,'xdr:spPr')
            append_node(append_node(props,ART,'a:prstGeom',{'prst':'rect'}),ART,'a:avLst')
            append_node(anchor,DRAW,'xdr:clientData')
        changes[drawing_part]=drawing_doc.toxml(encoding='utf-8')
        changes[relation_part(drawing_part)]=drawing_rels.toxml(encoding='utf-8')
        changes['[Content_Types].xml']=types.toxml(encoding='utf-8')
        staged=Path(directory)/'product-images.xlsx'
        with ZipFile(staged,'w') as destination:
            for info in archive.infolist():
                destination.writestr(copy(info),changes.get(info.filename,archive.read(info.filename)))
            for path,payload in changes.items():
                if path not in archive.namelist():
                    destination.writestr(path,payload)
        cover=None
        allowed_changes=set(changes)
        if update_cover:
            # 从新写入的竞品主图复用图片，保证两页同款且不重新下载。
            from competitor_analysis.market_cover import copy_first_product_image
            cover=copy_first_product_image(staged,Path(directory)/'with-market-cover.xlsx')
            staged=Path(cover['output'])
            allowed_changes.update(cover['changed_parts'])
        checked=load_workbook(staged,rich_text=True)
        try:
            images={image.anchor._from.col+1:image for image in checked[name]._images if image.anchor._from.row+1==row}
            for item in prepared.values():
                image=images[item['column']]
                if image._data()!=item['payload'] or (round(image.anchor.ext.cx/9525),round(image.anchor.ext.cy/9525))!=item['size']:
                    raise ValueError('保存后的主图像素或显示框与预期不一致')
        finally:
            checked.close()
        with ZipFile(staged) as result:
            if result.testzip() or any(result.read(n)!=archive.read(n) for n in archive.namelist() if n not in allowed_changes):
                raise ValueError('主图之外的工作簿部件被改变')
        staged.replace(output)
    return {'output':str(output),'asins':list(prepared),'display_mode':display_mode,
            'market_cover':{k:v for k,v in cover.items() if k!='output'} if cover else None,
            'images':{asin:{'pixels':[item['asset']['width'],item['asset']['height']],'display':item['size']} for asin,item in prepared.items()}}


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--workbook',type=Path,required=True)
    parser.add_argument('--images',type=Path,required=True)
    parser.add_argument('--output',type=Path,required=True)
    parser.add_argument('--display-mode',choices=['fit','original'],default='original')
    parser.add_argument('--allow-partial',action='store_true')
    args=parser.parse_args()
    assets=json.loads(args.images.read_text(encoding='utf-8-sig'))['products']
    print(json.dumps(refresh_product_images(args.workbook,args.output,assets,display_mode=args.display_mode,allow_partial=args.allow_partial),ensure_ascii=False))


if __name__=='__main__':
    main()
