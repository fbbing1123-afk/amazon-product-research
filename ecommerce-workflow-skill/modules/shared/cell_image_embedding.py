"""将最终浮动图片定点改为WPS单元格图片，不调整任何行高、列宽或合并区域。"""
from __future__ import annotations

import hashlib
import io
import posixpath
import tempfile
from pathlib import Path
from xml.dom import minidom
from zipfile import ZipFile, ZIP_DEFLATED

from openpyxl.utils.cell import get_column_letter, range_boundaries
from PIL import Image


MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
OFFICE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
ART = 'http://schemas.openxmlformats.org/drawingml/2006/main'
ETC = 'http://www.wps.cn/officeDocument/2017/etCustomData'
CONTENT = 'http://schemas.openxmlformats.org/package/2006/content-types'
CELL_REL = 'http://www.wps.cn/officeDocument/2020/cellImage'
CELL_TYPE = 'application/vnd.wps-officedocument.cellimage+xml'
CELL_PART = 'xl/cellimages.xml'
CELL_RELS = 'xl/_rels/cellimages.xml.rels'


def _all(node, namespace, name):
    return list(node.getElementsByTagNameNS(namespace, name))


def _children(node, namespace, name):
    return [n for n in node.childNodes if n.nodeType == n.ELEMENT_NODE
            and n.namespaceURI == namespace and n.localName == name]


def _add(parent, namespace, name, attributes=None, text=None):
    node = parent.ownerDocument.createElementNS(namespace, name)
    for key,value in (attributes or {}).items():
        node.setAttribute(key,str(value))
    if text is not None:
        node.appendChild(parent.ownerDocument.createTextNode(str(text)))
    parent.appendChild(node)
    return node


def _resolve(owner, target):
    return posixpath.normpath(posixpath.join(posixpath.dirname(owner),target)).lstrip('/')


def _rels(part):
    folder,name = posixpath.split(part)
    return f'{folder}/_rels/{name}.rels'


def _text(node):
    return ''.join(n.data if n.nodeType in (n.TEXT_NODE,n.CDATA_SECTION_NODE) else _text(n)
                   for n in node.childNodes)


def _cell(sheet, row, column):
    """合并区域内的图片归到该区域左上格；拒绝覆盖已有文字、数字或公式。"""
    for merge in _all(sheet,MAIN,'mergeCell'):
        left,top,right,bottom = range_boundaries(merge.getAttribute('ref'))
        if left <= column <= right and top <= row <= bottom:
            row,column = top,left
            break
    address = f'{get_column_letter(column)}{row}'
    data = _all(sheet,MAIN,'sheetData')[0]
    rows = _children(data,MAIN,'row')
    target = next((n for n in rows if int(n.getAttribute('r')) == row),None)
    if target is None:
        target = _add(data,MAIN,'row',{'r':row})
        later = next((n for n in rows if int(n.getAttribute('r')) > row),None)
        if later is not None:
            data.insertBefore(target,later)
    cells = _children(target,MAIN,'c')
    cell = next((n for n in cells if n.getAttribute('r') == address),None)
    if cell is None:
        cell = _add(target,MAIN,'c',{'r':address})
        later = next((n for n in cells if range_boundaries(n.getAttribute('r'))[0] > column),None)
        if later is not None:
            target.insertBefore(cell,later)
    if _children(cell,MAIN,'f') or any(_text(n) for name in ('v','is') for n in _children(cell,MAIN,name)):
        raise ValueError(f'图片目标单元格{address}已有内容，停止以免覆盖')
    return cell,address,row,column


def _append_image(root, image_id, relation_id, picture_id, pixels):
    item = _add(root,ETC,'etc:cellImage')
    pic = _add(item,XDR,'xdr:pic')
    properties = _add(pic,XDR,'xdr:nvPicPr')
    _add(properties,XDR,'xdr:cNvPr',{'id':picture_id,'name':image_id,'descr':'单元格图片'})
    _add(properties,XDR,'xdr:cNvPicPr')
    fill = _add(pic,XDR,'xdr:blipFill')
    _add(fill,ART,'a:blip',{'r:embed':relation_id})
    _add(_add(fill,ART,'a:stretch'),ART,'a:fillRect')
    shape = _add(pic,XDR,'xdr:spPr')
    transform = _add(shape,ART,'a:xfrm')
    _add(transform,ART,'a:off',{'x':0,'y':0})
    _add(transform,ART,'a:ext',{'cx':pixels[0]*9525,'cy':pixels[1]*9525})
    _add(_add(shape,ART,'a:prstGeom',{'prst':'rect'}),ART,'a:avLst')


def embed_final_images(source: Path, output: Path) -> dict:
    """复用原媒体字节和现有布局，仅替换图片锚点与对应空单元格。"""
    source,output = Path(source),Path(output)
    if source.resolve() == output.resolve() or output.exists():
        raise ValueError('请使用尚不存在的独立输出文件')
    entries,summary,changes = [],{},{}
    with ZipFile(source) as archive:
        names = set(archive.namelist())
        cell_doc = (minidom.parseString(archive.read(CELL_PART)) if CELL_PART in names else
                    minidom.parseString(f'<etc:cellImages xmlns:etc="{ETC}"/>'))
        cell_root = cell_doc.documentElement
        for prefix,namespace in [('etc',ETC),('xdr',XDR),('a',ART),('r',OFFICE)]:
            cell_root.setAttribute('xmlns:'+prefix,namespace)
        image_ids = {n.getAttribute('name') for n in _all(cell_doc,XDR,'cNvPr')}
        picture_id = max([int(n.getAttribute('id')) for n in _all(cell_doc,XDR,'cNvPr')],default=0)+1
        cell_rels = (minidom.parseString(archive.read(CELL_RELS)) if CELL_RELS in names else
                     minidom.parseString(f'<Relationships xmlns="{REL}"/>'))
        cell_rels.documentElement.setAttribute('xmlns',REL)
        relation_ids = {n.getAttribute('Id') for n in _all(cell_rels,REL,'Relationship')}
        counter,relation_number = 1,1
        book = minidom.parseString(archive.read('xl/workbook.xml'))
        book_rels = minidom.parseString(archive.read('xl/_rels/workbook.xml.rels'))
        relations = {n.getAttribute('Id'):n for n in _all(book_rels,REL,'Relationship')}
        used_drawings = set()
        for sheet_info in _all(book,MAIN,'sheet'):
            title = sheet_info.getAttribute('name')
            sheet_part = _resolve('xl/workbook.xml',relations[sheet_info.getAttributeNS(OFFICE,'id')].getAttribute('Target'))
            sheet = minidom.parseString(archive.read(sheet_part))
            drawing_refs = _all(sheet,MAIN,'drawing')
            if not drawing_refs:
                summary[title] = 0
                continue
            sheet_rels = minidom.parseString(archive.read(_rels(sheet_part)))
            sheet_relations = {n.getAttribute('Id'):n for n in _all(sheet_rels,REL,'Relationship')}
            targets,count = set(),0
            for drawing_ref in drawing_refs:
                drawing_part = _resolve(sheet_part,sheet_relations[drawing_ref.getAttributeNS(OFFICE,'id')].getAttribute('Target'))
                if drawing_part in used_drawings:
                    raise ValueError('多个工作表共享绘图部件，不能安全确定图片单元格')
                used_drawings.add(drawing_part)
                drawing = minidom.parseString(archive.read(drawing_part))
                drawing_relations = {n.getAttribute('Id'):n for n in _all(
                    minidom.parseString(archive.read(_rels(drawing_part))),REL,'Relationship')}
                converted = False
                for anchor in list(drawing.documentElement.childNodes):
                    if anchor.nodeType != anchor.ELEMENT_NODE or not _all(anchor,XDR,'pic'):
                        continue
                    if anchor.localName not in ('oneCellAnchor','twoCellAnchor') or len(_children(anchor,XDR,'pic')) != 1:
                        raise ValueError(f'{title}含无法确定单元格的图片锚点，保留原文件')
                    marker = _children(anchor,XDR,'from')[0]
                    row = int(_children(marker,XDR,'row')[0].firstChild.data)+1
                    column = int(_children(marker,XDR,'col')[0].firstChild.data)+1
                    cell,address,row,column = _cell(sheet,row,column)
                    if address in targets:
                        raise ValueError(f'{title}!{address}有多张图片，停止以免覆盖')
                    targets.add(address)
                    blip = _all(anchor,ART,'blip')[0]
                    relation = drawing_relations[blip.getAttributeNS(OFFICE,'embed')]
                    if relation.getAttribute('TargetMode') == 'External':
                        raise ValueError('无法将外部链接图片作为本地单元格图片嵌入')
                    media = _resolve(drawing_part,relation.getAttribute('Target'))
                    payload = archive.read(media)
                    with Image.open(io.BytesIO(payload)) as image:
                        pixels = image.size
                    image_id = f'ID_CODEX_FINAL_{counter:04d}'
                    while image_id in image_ids:
                        counter += 1
                        image_id = f'ID_CODEX_FINAL_{counter:04d}'
                    image_ids.add(image_id)
                    counter += 1
                    while f'rId{relation_number}' in relation_ids:
                        relation_number += 1
                    rid = f'rId{relation_number}'
                    relation_ids.add(rid)
                    _add(cell_rels.documentElement,REL,'Relationship',{
                        'Id':rid,'Type':OFFICE+'/image','Target':posixpath.relpath(media,'xl')})
                    _append_image(cell_root,image_id,rid,picture_id,pixels)
                    picture_id += 1
                    for name in ('is','v','f'):
                        for node in _children(cell,MAIN,name):
                            cell.removeChild(node)
                    if cell.hasAttribute('t'):
                        cell.removeAttribute('t')
                    _add(cell,MAIN,'f',text=f'_xlfn.DISPIMG("{image_id}",1)')
                    _add(cell,MAIN,'v')
                    # 只扩展实际数据范围，不写行高、列宽、合并或图片显示尺寸到单元格布局。
                    dimensions = _all(sheet,MAIN,'dimension')
                    if dimensions:
                        left,top,right,bottom = range_boundaries(dimensions[0].getAttribute('ref'))
                        bounds = min(left,column),min(top,row),max(right,column),max(bottom,row)
                        if bounds != (left,top,right,bottom):
                            dimensions[0].setAttribute('ref',f'{get_column_letter(bounds[0])}{bounds[1]}:{get_column_letter(bounds[2])}{bounds[3]}')
                    drawing.documentElement.removeChild(anchor)
                    entries.append({'sheet':title,'cell':address,'image_id':image_id,'media':media,
                                    'pixels':pixels,'sha256':hashlib.sha256(payload).hexdigest()})
                    count += 1
                    converted = True
                if converted:
                    changes[drawing_part] = drawing.toxml(encoding='utf-8')
            if count:
                changes[sheet_part] = sheet.toxml(encoding='utf-8')
            summary[title] = count
        if entries:
            changes[CELL_PART] = cell_doc.toxml(encoding='utf-8')
            changes[CELL_RELS] = cell_rels.toxml(encoding='utf-8')
            if not any(n.getAttribute('Type') == CELL_REL for n in relations.values()):
                book_rels.documentElement.setAttribute('xmlns',REL)
                index = 1
                while f'rId{index}' in relations:
                    index += 1
                _add(book_rels.documentElement,REL,'Relationship',{
                    'Id':f'rId{index}','Type':CELL_REL,'Target':'cellimages.xml'})
                changes['xl/_rels/workbook.xml.rels'] = book_rels.toxml(encoding='utf-8')
            types = minidom.parseString(archive.read('[Content_Types].xml'))
            if not any(n.getAttribute('PartName') == '/'+CELL_PART for n in _all(types,CONTENT,'Override')):
                types.documentElement.setAttribute('xmlns',CONTENT)
                _add(types.documentElement,CONTENT,'Override',{'PartName':'/'+CELL_PART,'ContentType':CELL_TYPE})
                changes['[Content_Types].xml'] = types.toxml(encoding='utf-8')
        output.parent.mkdir(parents=True,exist_ok=True)
        with tempfile.NamedTemporaryFile(dir=output.parent,suffix='.xlsx',delete=False) as temporary:
            staged = Path(temporary.name)
        try:
            with ZipFile(staged,'w',ZIP_DEFLATED) as destination:
                for info in archive.infolist():
                    destination.writestr(info,changes.get(info.filename,archive.read(info.filename)))
                for part,payload in changes.items():
                    if part not in names:
                        destination.writestr(part,payload)
            with ZipFile(staged) as checked:
                if checked.testzip() is not None:
                    raise ValueError('嵌入版ZIP校验失败')
            staged.replace(output)
        finally:
            staged.unlink(missing_ok=True)
    return {'output':str(output.resolve()),'embedded':len(entries),'sheets':summary,
            'images':entries,'changed_parts':list(changes)}
