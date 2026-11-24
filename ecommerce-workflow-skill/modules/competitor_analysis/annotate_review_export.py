"""在原始评论页末尾添加逐行分析结果，保留下载文件的原始数据和原生部件。"""

from copy import copy
from pathlib import Path
import posixpath
import tempfile
import unicodedata
from xml.dom import minidom
from zipfile import ZipFile

from openpyxl.utils import get_column_letter

from review_research.download_sellersprite_reviews import read_export
from competitor_analysis.review_mentions import format_point


HEADERS = ["优点归纳", "差评原因", "分析说明"]
WIDTHS = [34, 56, 46]
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def children(element, name):
    """只选择同一层的元素，避免把扩展部件中的同名节点混进来。"""
    return [node for node in element.childNodes if node.nodeType == node.ELEMENT_NODE and node.localName == name]


def element(parent, name, attributes=None):
    """沿用父节点命名空间和前缀，不改写原文件已有的扩展命名空间。"""
    qualified = f"{parent.prefix}:{name}" if parent.prefix else name
    node = parent.ownerDocument.createElementNS(parent.namespaceURI, qualified)
    for key, value in (attributes or {}).items():
        node.setAttribute(key, str(value))
    return node


def add_cell(row, coordinate, value, style):
    """所有分析文字按文本写入，评论中的公式样式内容不会被执行。"""
    cell = element(row, "c", {"r": coordinate, "s": style, "t": "inlineStr"})
    inline = element(cell, "is")
    text = element(inline, "t", {"xml:space": "preserve"})
    text.appendChild(row.ownerDocument.createTextNode(value))
    inline.appendChild(text)
    cell.appendChild(inline)
    row.appendChild(cell)


def worksheet_part(archive, name):
    """使用关系文件找到导出评论页，保留其它页签与关系。"""
    workbook = minidom.parseString(archive.read("xl/workbook.xml"))
    sheets = workbook.getElementsByTagNameNS(MAIN_NS, "sheet")
    relation_id = next(s.getAttributeNS(REL_NS, "id") for s in sheets if s.getAttribute("name") == name)
    relations = minidom.parseString(archive.read("xl/_rels/workbook.xml.rels"))
    target = next(r.getAttribute("Target") for r in relations.documentElement.childNodes
                  if r.nodeType == r.ELEMENT_NODE and r.getAttribute("Id") == relation_id)
    return target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))


def annotation_rows(packet, analysis, source):
    """一个源行对应一组分析列；同一评论的多个原因在单元格中换行。"""
    if any(analysis.get(key) != packet.get(key) for key in ("asin", "packet_digest", "source_sha256", "source_file")):
        raise ValueError("分析结果与评论源文件不匹配")
    if source["source_sha256"] != packet["source_sha256"]:
        raise ValueError("评论源文件已变化，不能按旧行号加列")
    if set(HEADERS).intersection(source["headers"]):
        raise ValueError("源文件已含分析列，请从原始下载文件生成新的分析副本")
    selected = {}
    for kind in ("advantages", "disadvantages"):
        points = analysis[kind]
        samples = packet["samples"][kind]
        for review in samples:
            conclusions = [format_point(point, kind) for point in points if review["number"] in point["review_numbers"]]
            note = "已分析" if conclusions else "已分析，未被结论引用"
            if not (review["title"] or review["body"]):
                note = "原文为空，未作为结论证据"
            elif not points:
                note = "未形成结论：" + analysis.get("empty_reasons", {}).get(kind, "未找到明确依据")
            text = "\n".join(conclusions)
            selected[review["source_row"]] = [
                text if kind == "advantages" else "", text if kind == "disadvantages" else "", note,
            ]
    # 证据编号只供程序对应原行；未入选评论的新增列全部留空。
    return {review["source_row"]: selected.get(review["source_row"], ["", "", ""])
            for review in source["reviews"]}


def required_height(values):
    """按中英文宽度估算新列换行所需行高，不缩小原表行高。"""
    lines = 1
    for value, width in zip(values, WIDTHS):
        current = 0
        for paragraph in value.split("\n"):
            units = sum(2 if unicodedata.east_asian_width(char) in "WF" else 1 for char in paragraph)
            current += max(1, (units + width - 4) // (width - 3))
        lines = max(lines, current)
    height = max(24, lines * 16 + 8)
    if height > 409:
        raise ValueError("单条评论的归纳文字过多，超过 Excel 可显示行高，请精简后重试")
    return height


def annotate_export(packet, analysis, output: Path):
    """写独立评论表副本，只追加分析列，不重存图形、链接和其它原生部件。"""
    original = Path(packet["source_file"])
    if output.exists() or output.resolve() == original.resolve():
        raise ValueError("分析评论表必须使用不存在的独立输出文件")
    source = read_export(original, packet["asin"])
    records = annotation_rows(packet, analysis, source)
    first_column = len(source["headers"]) + 1
    last_column = first_column + len(HEADERS) - 1
    sheet_name = source["source_sheet"]
    with ZipFile(original) as archive:
        part = worksheet_part(archive, sheet_name)
        sheet_doc = minidom.parseString(archive.read(part))
        root = sheet_doc.documentElement
        if children(root, "tableParts"):
            raise ValueError("当前评论文件含原生表对象，请使用卖家精灵原始导出文件")
        data = children(root, "sheetData")[0]
        rows = {int(r.getAttribute("r")): r for r in children(data, "row")}
        header = rows[1]
        header_style = children(header, "c")[0].getAttribute("s") or "0"
        # 继承正文样式，仅为新列增加换行样式；原有样式编号完全不变。
        sample_cells = children(rows[min(records)], "c") if records else children(header, "c")
        text_column = get_column_letter(source["headers"].index("内容") + 1)
        sample = next((c for c in sample_cells if c.getAttribute("r").rstrip("0123456789") == text_column), sample_cells[0])
        base_style = int(sample.getAttribute("s") or "0")
        styles_doc = minidom.parseString(archive.read("xl/styles.xml"))
        formats = children(styles_doc.documentElement, "cellXfs")[0]
        style_id = len(children(formats, "xf"))
        style = children(formats, "xf")[base_style].cloneNode(True)
        style.setAttribute("numFmtId", "0")
        style.setAttribute("applyAlignment", "1")
        for alignment in children(style, "alignment"):
            style.removeChild(alignment)
        style.appendChild(element(style, "alignment", {"horizontal": "left", "vertical": "top", "wrapText": "1"}))
        formats.appendChild(style)
        formats.setAttribute("count", str(style_id + 1))
        for index, label in enumerate(HEADERS, first_column):
            add_cell(header, f"{get_column_letter(index)}1", label, header_style)
        for number, values in records.items():
            row = rows[number]
            for index, value in enumerate(values, first_column):
                add_cell(row, f"{get_column_letter(index)}{number}", value, style_id)
            if any(values):
                row.setAttribute("ht", str(max(float(row.getAttribute("ht") or 15), required_height(values))))
                row.setAttribute("customHeight", "1")
            row.setAttribute("spans", f"1:{last_column}")
        header.setAttribute("spans", f"1:{last_column}")
        last_row = max(rows)
        dimension = children(root, "dimension")
        if dimension:
            dimension[0].setAttribute("ref", f"A1:{get_column_letter(last_column)}{last_row}")
        columns = children(root, "cols")
        if columns:
            columns = columns[0]
        else:
            columns = element(root, "cols")
            root.insertBefore(columns, data)
        for index, width in enumerate(WIDTHS, first_column):
            columns.appendChild(element(columns, "col", {"min": index, "max": index, "width": width, "customWidth": "1"}))
        filters = children(root, "autoFilter")
        if filters:
            auto_filter = filters[0]
        else:
            auto_filter = element(root, "autoFilter")
            # autoFilter 位于 sheetData 和保护/场景节点之后、合并及绘图节点之前。
            previous = data
            for node in root.childNodes:
                if node.nodeType == node.ELEMENT_NODE and node.localName in ("sheetCalcPr", "sheetProtection", "protectedRanges", "scenarios"):
                    previous = node
            root.insertBefore(auto_filter, previous.nextSibling)
        auto_filter.setAttribute("ref", f"A1:{get_column_letter(last_column)}{last_row}")
        views = children(root, "sheetViews")
        if views:
            view = children(views[0], "sheetView")[0]
            if not children(view, "pane"):
                view.insertBefore(element(view, "pane", {"ySplit": "1", "topLeftCell": "A2", "activePane": "bottomLeft", "state": "frozen"}), view.firstChild)
                for selection in children(view, "selection"):
                    selection.setAttribute("pane", "bottomLeft")
        updates = {part: sheet_doc.toxml(encoding="utf-8"), "xl/styles.xml": styles_doc.toxml(encoding="utf-8")}
        output.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=output.parent) as temporary:
            staged = Path(temporary) / "annotated.xlsx"
            with ZipFile(staged, "w") as destination:
                for info in archive.infolist():
                    destination.writestr(copy(info), updates.get(info.filename, archive.read(info.filename)))
            with ZipFile(staged) as checked:
                if checked.testzip() or checked.namelist() != archive.namelist():
                    raise ValueError("评论分析表 ZIP 验收失败")
                if any(checked.read(name) != archive.read(name) for name in archive.namelist() if name not in updates):
                    raise ValueError("评论分析表修改了非目标部件")
            reread = read_export(staged, packet["asin"])
            if reread["headers"] != source["headers"] + HEADERS:
                raise ValueError("新增分析列表头验收失败")
            for original_review, annotated in zip(source["reviews"], reread["reviews"]):
                if any(annotated["raw"].get(key) != original_review["raw"].get(key) for key in source["headers"]):
                    raise ValueError("原始评论字段发生了变化")
                expected = records[original_review["source_row"]]
                if [annotated["raw"].get(key) or "" for key in HEADERS] != expected:
                    raise ValueError("评论分析列与证据行不一致")
            staged.replace(output)
    return {"output": str(output.resolve()), "sheet": sheet_name, "review_rows": len(records),
            "added_columns": HEADERS, "selected_counts": analysis["selected_counts"]}
