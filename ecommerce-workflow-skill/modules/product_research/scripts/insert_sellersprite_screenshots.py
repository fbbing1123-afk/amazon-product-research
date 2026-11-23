"""按分析名称放置三张浮动截图，只修改绘图部件，保留原工作簿格式。"""

from __future__ import annotations

import argparse
import io
import json
import posixpath
import re
import tempfile
import uuid
import zipfile
from pathlib import Path
from xml.dom import minidom

from openpyxl import load_workbook
from PIL import Image


MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DRAWING = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
ART = "http://schemas.openxmlformats.org/drawingml/2006/main"
CONTENT = "http://schemas.openxmlformats.org/package/2006/content-types"
CHARTS = {"review": "Review数量分析", "rating": "Rating分析", "price": "价格分析"}
PREFIX = "sellersprite-screenshot:"
EMU = 9525


def elements(node, namespace, name):
    return list(node.getElementsByTagNameNS(namespace, name))


def append(parent, namespace, name, attributes=None, text=None):
    doc = parent.ownerDocument
    node = doc.createElementNS(namespace, name)
    for key, value in (attributes or {}).items():
        node.setAttribute(key, str(value))
    if text is not None:
        node.appendChild(doc.createTextNode(str(text)))
    parent.appendChild(node)
    return node


def document(namespace, name):
    return minidom.parseString(f'<{name} xmlns="{namespace}"/>')


def resolve_part(owner, target):
    return posixpath.normpath(posixpath.join(posixpath.dirname(owner), target)).lstrip("/")


def rels_path(part):
    folder, name = posixpath.split(part)
    return f"{folder}/_rels/{name}.rels"


def normalize_label(value):
    return re.sub(r"\s+", "", str(value or "")).casefold()


def locate_chart_regions(ws):
    """每次重新读取标签及右侧合并范围，不依赖行号或固定列宽。"""
    regions = {}
    for key, label in CHARTS.items():
        matches = [cell for row in ws for cell in row
                   if normalize_label(cell.value) == normalize_label(label)]
        if len(matches) != 1:
            raise ValueError(f"{label}必须唯一，实际找到 {len(matches)} 处")
        cell = matches[0]
        label_area = next((area for area in ws.merged_cells.ranges if cell.coordinate in area), None)
        column = label_area.max_col + 1 if label_area else cell.column + 1
        areas = [area for area in ws.merged_cells.ranges
                 if area.min_row == cell.row and area.min_col == column]
        if len(areas) != 1:
            raise ValueError(f"{label}右侧缺少唯一的合并图表区域，停止写入以免覆盖结论")
        regions[key] = areas[0]
    return regions


def region_pixels(ws, area):
    width = 0
    for column in range(area.min_col, area.max_col + 1):
        dimension = next((dim for dim in ws.column_dimensions.values()
                          if dim.min is not None and dim.min <= column <= (dim.max or dim.min)), None)
        if dimension and dimension.hidden:
            continue
        chars = dimension.width if dimension else (ws.sheet_format.defaultColWidth or 8.43)
        width += chars * 7 + 5 if chars >= 1 else chars * 12
    height = sum(0 if ws.row_dimensions[row].hidden else
                 (ws.row_dimensions[row].height or ws.sheet_format.defaultRowHeight or 15) * 96 / 72
                 for row in range(area.min_row, area.max_row + 1))
    if min(width, height) <= 16:
        raise ValueError(f"图表区域 {area} 太小或已隐藏")
    return width, height


def insert_screenshots(source: Path, output: Path, images: dict[str, Path], sheet_name=None):
    if source.resolve() == output.resolve():
        raise ValueError("请输出到新文件，保留输入工作簿")
    wb = load_workbook(source, rich_text=True)
    try:
        candidates = [ws for ws in wb if (ws.title == sheet_name if sheet_name else "市场调研" in ws.title)]
        if len(candidates) != 1:
            raise ValueError("市场调研页不唯一或不存在，请用 --sheet 明确指定")
        ws = candidates[0]
        regions = locate_chart_regions(ws)
        sizes = {key: region_pixels(ws, area) for key, area in regions.items()}
        selected_sheet = ws.title
    finally:
        wb.close()

    # openpyxl 只读布局；不整体保存，以免丢失 WPS 单元格图片等扩展部件。
    with zipfile.ZipFile(source) as original:
        parts = set(original.namelist())
        changes = {}
        workbook_xml = minidom.parseString(original.read("xl/workbook.xml"))
        sheet = next(s for s in elements(workbook_xml, MAIN, "sheet") if s.getAttribute("name") == selected_sheet)
        workbook_rels = minidom.parseString(original.read("xl/_rels/workbook.xml.rels"))
        relationship = next(r for r in elements(workbook_rels, REL, "Relationship")
                            if r.getAttribute("Id") == sheet.getAttributeNS(OFFICE, "id"))
        sheet_part = resolve_part("xl/workbook.xml", relationship.getAttribute("Target"))
        sheet_doc = minidom.parseString(original.read(sheet_part))
        sheet_rel_part = rels_path(sheet_part)
        sheet_rels = (minidom.parseString(original.read(sheet_rel_part)) if sheet_rel_part in parts
                      else document(REL, "Relationships"))
        drawing_nodes = elements(sheet_doc, MAIN, "drawing")
        if drawing_nodes:
            rel_id = drawing_nodes[0].getAttributeNS(OFFICE, "id")
            rel = next(r for r in elements(sheet_rels, REL, "Relationship") if r.getAttribute("Id") == rel_id)
            drawing_part = resolve_part(sheet_part, rel.getAttribute("Target"))
            drawing = minidom.parseString(original.read(drawing_part))
        else:
            suffix = uuid.uuid4().hex
            drawing_part = f"xl/drawings/sellersprite_{suffix}.xml"
            rel_id = f"rIdSellerSprite{suffix}"
            node = append(sheet_doc.documentElement, MAIN, "drawing", {"xmlns:r": OFFICE, "r:id": rel_id})
            # 按工作表架构顺序插入绘图引用，避免放在扩展列表之后。
            following = {"legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls",
                         "webPublishItems", "tableParts", "extLst"}
            for child in list(sheet_doc.documentElement.childNodes):
                if child.nodeType == child.ELEMENT_NODE and child.localName in following:
                    sheet_doc.documentElement.insertBefore(node, child)
                    break
            append(sheet_rels.documentElement, REL, "Relationship", {
                "Id": rel_id, "Type": OFFICE + "/drawing",
                "Target": posixpath.relpath(drawing_part, posixpath.dirname(sheet_part))})
            drawing = minidom.parseString(
                f'<xdr:wsDr xmlns:xdr="{DRAWING}" xmlns:a="{ART}" xmlns:r="{OFFICE}"/>')
            changes[sheet_part] = sheet_doc.toxml(encoding="utf-8")
            changes[sheet_rel_part] = sheet_rels.toxml(encoding="utf-8")

        drawing_rel_part = rels_path(drawing_part)
        drawing_rels = (minidom.parseString(original.read(drawing_rel_part)) if drawing_rel_part in parts
                        else document(REL, "Relationships"))
        old_media = {}
        # 只替换本入口创建的三张图，保留其它截图和图表。
        for node in list(drawing.documentElement.childNodes):
            if node.nodeType != node.ELEMENT_NODE:
                continue
            names = elements(node, DRAWING, "cNvPr")
            if names and names[0].getAttribute("name") in {PREFIX + key for key in CHARTS}:
                key = names[0].getAttribute("name")[len(PREFIX):]
                blips = elements(node, ART, "blip")
                if blips:
                    rid = blips[0].getAttributeNS(OFFICE, "embed")
                    old_media[key] = rid
                drawing.documentElement.removeChild(node)

        ids = [int(node.getAttribute("id")) for node in elements(drawing, DRAWING, "cNvPr")]
        picture_id = max(ids, default=0) + 1
        results = {}
        for key, area in regions.items():
            with Image.open(images[key]) as image:
                iw, ih = image.size
                data = io.BytesIO()
                image.convert("RGB").save(data, format="PNG")
            rid = old_media.get(key)
            existing = next((r for r in elements(drawing_rels, REL, "Relationship")
                             if r.getAttribute("Id") == rid), None)
            still_used = rid and any(b.getAttributeNS(OFFICE, "embed") == rid for b in elements(drawing, ART, "blip"))
            if existing is not None and not still_used and existing.getAttribute("TargetMode") != "External":
                media_part = resolve_part(drawing_part, existing.getAttribute("Target"))
            else:
                suffix = uuid.uuid4().hex
                media_part = f"xl/media/sellersprite_{suffix}.png"
                rid = f"rId{suffix}"
                append(drawing_rels.documentElement, REL, "Relationship", {
                    "Id": rid, "Type": OFFICE + "/image",
                    "Target": posixpath.relpath(media_part, posixpath.dirname(drawing_part))})
            changes[media_part] = data.getvalue()
            width, height = sizes[key]
            scale = min((width - 16) / iw, (height - 16) / ih)
            cx, cy = round(iw * scale * EMU), round(ih * scale * EMU)
            anchor = append(drawing.documentElement, DRAWING, "xdr:oneCellAnchor", {"xmlns:xdr": DRAWING, "xmlns:a": ART, "xmlns:r": OFFICE})
            start = append(anchor, DRAWING, "xdr:from")
            for name, value in [("col", area.min_col - 1), ("colOff", round((width - iw * scale) / 2 * EMU)),
                                ("row", area.min_row - 1), ("rowOff", round((height - ih * scale) / 2 * EMU))]:
                append(start, DRAWING, f"xdr:{name}", text=value)
            append(anchor, DRAWING, "xdr:ext", {"cx": cx, "cy": cy})
            pic = append(anchor, DRAWING, "xdr:pic")
            props = append(pic, DRAWING, "xdr:nvPicPr")
            append(props, DRAWING, "xdr:cNvPr", {"id": picture_id, "name": PREFIX + key})
            append(append(props, DRAWING, "xdr:cNvPicPr"), ART, "a:picLocks", {"noChangeAspect": 1})
            fill = append(pic, DRAWING, "xdr:blipFill")
            append(fill, ART, "a:blip", {"r:embed": rid})
            append(append(fill, ART, "a:stretch"), ART, "a:fillRect")
            geometry = append(append(pic, DRAWING, "xdr:spPr"), ART, "a:prstGeom", {"prst": "rect"})
            append(geometry, ART, "a:avLst")
            append(anchor, DRAWING, "xdr:clientData")
            results[key] = str(area)
            picture_id += 1
        changes[drawing_part] = drawing.toxml(encoding="utf-8")
        changes[drawing_rel_part] = drawing_rels.toxml(encoding="utf-8")
        types = minidom.parseString(original.read("[Content_Types].xml"))
        if not any(n.getAttribute("PartName") == "/" + drawing_part for n in elements(types, CONTENT, "Override")):
            append(types.documentElement, CONTENT, "Override", {"PartName": "/" + drawing_part,
                   "ContentType": "application/vnd.openxmlformats-officedocument.drawing+xml"})
        # 扩展名不区分大小写，防止与原模板中的大写 PNG 重复登记。
        if not any(n.getAttribute("Extension").lower() == "png" for n in elements(types, CONTENT, "Default")):
            append(types.documentElement, CONTENT, "Default", {"Extension": "png", "ContentType": "image/png"})
        changes["[Content_Types].xml"] = types.toxml(encoding="utf-8")
        output.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(dir=output.parent, suffix=".xlsx", delete=False) as temporary:
            temp_path = Path(temporary.name)
        try:
            with zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as target:
                for entry in original.infolist():
                    target.writestr(entry, changes.pop(entry.filename, original.read(entry.filename)))
                for name, data in changes.items():
                    target.writestr(name, data)
            temp_path.replace(output)
        finally:
            temp_path.unlink(missing_ok=True)
    return results


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--sheet")
    for key in CHARTS:
        parser.add_argument(f"--{key}", type=Path, required=True)
    args = parser.parse_args()
    result = insert_screenshots(args.input, args.output, {key: getattr(args, key) for key in CHARTS}, args.sheet)
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
