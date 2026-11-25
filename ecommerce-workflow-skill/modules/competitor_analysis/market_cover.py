"""将竞品分析页第一款商品的现有主图原尺寸浮动放到市场调研页A1。"""
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
from PIL import Image

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from competitor_analysis.annotate_review_export import worksheet_part, children
from competitor_analysis.populate_basics import find_label_row
from competitor_analysis.refresh_product_images import (
    MAIN, OFFICE, DRAW, ART, CONTENT, relation_part, related_document, add_relation, append_node,
)


def target_part(owner, relation):
    """解析工作簿内部关系，不接受外链图片作为已保存的商品主图。"""
    if relation.getAttribute('TargetMode') == 'External':
        raise ValueError('商品主图必须是工作簿内已保存的图片')
    target = relation.getAttribute('Target')
    return target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(owner), target))


def drawing_for_sheet(archive, part):
    doc = minidom.parseString(archive.read(part))
    rels = related_document(archive, part)
    nodes = doc.getElementsByTagNameNS(MAIN, 'drawing')
    if len(nodes) > 1:
        raise ValueError('工作表绘图部件不唯一')
    if not nodes:
        return doc, rels, None
    rid = nodes[0].getAttributeNS(OFFICE, 'id')
    relation = next(r for r in children(rels.documentElement, 'Relationship') if r.getAttribute('Id') == rid)
    return doc, rels, target_part(part, relation)


def anchored_at(root, row, column):
    matches = []
    for anchor in children(root, 'oneCellAnchor') + children(root, 'twoCellAnchor'):
        start = children(anchor, 'from')
        if start and int(children(start[0], 'row')[0].firstChild.data) == row - 1 and int(children(start[0], 'col')[0].firstChild.data) == column - 1:
            matches.append(anchor)
    return matches


def copy_first_product_image(workbook, output):
    if output.exists() or output.resolve() == workbook.resolve():
        raise ValueError('首图输出必须另存到独立新文件')
    book = load_workbook(workbook, rich_text=True)
    try:
        markets = [s for s in book if s.title.endswith('市场调研')]
        competitors = [s for s in book if s.title.endswith('竞品分析')]
        if len(markets) != 1 or len(competitors) != 1:
            raise ValueError('必须有唯一的市场调研页和竞品分析页')
        market, competitor = markets[0], competitors[0]
        # 竞品列按从左到右排列，不能使用数据源或图片关系的存储顺序。
        row = find_label_row(competitor, '产品图')
        link_row = find_label_row(competitor, '详情页链接')
        from competitor_analysis.collect_basic_facts import asin_from_input
        links = []
        for cell in competitor[link_row]:
            try:
                asin = asin_from_input(str(cell.hyperlink.target if cell.hyperlink else cell.value or ''))
                links.append((cell.column, asin))
            except ValueError:
                continue
        if not links:
            raise ValueError('竞品详情链接为空')
        column, first_asin = links[0]
        if sum(asin == first_asin for _, asin in links) != 1:
            raise ValueError('第一款竞品的详情链接列不唯一')
        market_name, competitor_name = market.title, competitor.title
    finally:
        book.close()
    output.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(workbook) as archive, tempfile.TemporaryDirectory(dir=output.parent) as temporary:
        source_part = worksheet_part(archive, competitor_name)
        _, _, source_drawing = drawing_for_sheet(archive, source_part)
        if source_drawing is None:
            raise ValueError('第一款竞品尚未保存产品主图')
        source_doc = minidom.parseString(archive.read(source_drawing))
        anchors = anchored_at(source_doc.documentElement, row, column)
        if len(anchors) != 1 or len(children(anchors[0], 'pic')) != 1:
            raise ValueError('第一款竞品主图缺失或重复')
        pic = children(anchors[0], 'pic')[0]
        blip = pic.getElementsByTagNameNS(ART, 'blip')[0]
        source_rels = related_document(archive, source_drawing)
        relation = next(r for r in children(source_rels.documentElement, 'Relationship') if r.getAttribute('Id') == blip.getAttributeNS(OFFICE, 'embed'))
        media = target_part(source_drawing, relation)
        payload = archive.read(media)
        with Image.open(BytesIO(payload)) as picture:
            width, height = picture.size
            picture.verify()
        part = worksheet_part(archive, market_name)
        sheet_doc, sheet_rels, drawing_part = drawing_for_sheet(archive, part)
        changes = {}
        if drawing_part is None:
            index = 1
            while f'xl/drawings/marketCover{index}.xml' in archive.namelist():
                index += 1
            drawing_part = f'xl/drawings/marketCover{index}.xml'
            drawing_doc = minidom.parseString(f'<xdr:wsDr xmlns:xdr="{DRAW}" xmlns:a="{ART}" xmlns:r="{OFFICE}"/>')
            rid = add_relation(sheet_rels, 'drawing', posixpath.relpath(drawing_part, posixpath.dirname(part)))
            root = sheet_doc.documentElement
            root.setAttribute('xmlns:r', OFFICE)
            drawing = sheet_doc.createElementNS(MAIN, (root.prefix + ':' if root.prefix else '') + 'drawing')
            drawing.setAttributeNS(OFFICE, 'r:id', rid)
            following = next((n for n in root.childNodes if n.nodeType == n.ELEMENT_NODE and n.localName in
                {'legacyDrawing', 'legacyDrawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'}), None)
            root.insertBefore(drawing, following) if following else root.appendChild(drawing)
            changes[part] = sheet_doc.toxml(encoding='utf-8')
            changes[relation_part(part)] = sheet_rels.toxml(encoding='utf-8')
            types = minidom.parseString(archive.read('[Content_Types].xml'))
            append_node(types.documentElement, CONTENT, 'Override', {'PartName': '/' + drawing_part, 'ContentType': 'application/vnd.openxmlformats-officedocument.drawing+xml'})
            changes['[Content_Types].xml'] = types.toxml(encoding='utf-8')
        else:
            drawing_doc = minidom.parseString(archive.read(drawing_part))
        root = drawing_doc.documentElement
        for prefix, namespace in [('xdr', DRAW), ('a', ART), ('r', OFFICE)]:
            root.setAttribute('xmlns:' + prefix, namespace)
        existing = anchored_at(root, 1, 1)
        if len(existing) > 1 or any(not children(a, 'pic') for a in existing):
            raise ValueError('市场调研A1存在重复图片或其它绘图，停止替换')
        for anchor in existing:
            root.removeChild(anchor)
        rels = related_document(archive, drawing_part)
        rid = add_relation(rels, 'image', posixpath.relpath(media, posixpath.dirname(drawing_part)))
        object_id = max([int(n.getAttribute('id') or 0) for n in root.getElementsByTagNameNS(DRAW, 'cNvPr')] + [0]) + 1
        anchor = append_node(root, DRAW, 'xdr:oneCellAnchor')
        start = append_node(anchor, DRAW, 'xdr:from')
        for tag in ('col', 'colOff', 'row', 'rowOff'):
            append_node(start, DRAW, 'xdr:' + tag, text=0)
        append_node(anchor, DRAW, 'xdr:ext', {'cx': width * 9525, 'cy': height * 9525})
        copied = drawing_doc.importNode(pic, True)
        copied.setAttribute('xmlns', DRAW)
        copied.getElementsByTagNameNS(ART, 'blip')[0].setAttributeNS(OFFICE, 'r:embed', rid)
        props = copied.getElementsByTagNameNS(DRAW, 'cNvPr')[0]
        props.setAttribute('id', str(object_id))
        props.setAttribute('name', f'市场调研首图 {first_asin}')
        # 去除旧图可能携带的裁切和形变，按原图像素重新放置。
        for namespace, tag in [(ART, 'srcRect'), (ART, 'xfrm')]:
            for node in list(copied.getElementsByTagNameNS(namespace, tag)):
                node.parentNode.removeChild(node)
        anchor.appendChild(copied)
        append_node(anchor, DRAW, 'xdr:clientData')
        changes[drawing_part] = drawing_doc.toxml(encoding='utf-8')
        changes[relation_part(drawing_part)] = rels.toxml(encoding='utf-8')
        staged = Path(temporary) / 'market-cover.xlsx'
        with ZipFile(staged, 'w') as result:
            for info in archive.infolist():
                result.writestr(copy(info), changes.get(info.filename, archive.read(info.filename)))
            for name, data in changes.items():
                if name not in archive.namelist():
                    result.writestr(name, data)
        checked = load_workbook(staged, rich_text=True)
        try:
            images = [im for im in checked[market_name]._images if im.anchor._from.row == 0 and im.anchor._from.col == 0]
            if len(images) != 1 or images[0]._data() != payload or (images[0].anchor.ext.cx, images[0].anchor.ext.cy) != (width * 9525, height * 9525):
                raise ValueError('市场调研首图归属或原尺寸校验失败')
        finally:
            checked.close()
        with ZipFile(staged) as result:
            if result.testzip() or any(result.read(n) != archive.read(n) for n in archive.namelist() if n not in changes):
                raise ValueError('首图以外的原工作簿部件发生变化')
        staged.replace(output)
    return {'output': str(output), 'asin': first_asin, 'anchor': 'A1', 'pixels': [width, height], 'changed_parts': list(changes)}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--workbook', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    print(json.dumps(copy_first_product_image(args.workbook, args.output), ensure_ascii=False))


if __name__ == '__main__':
    main()
