"""按独立数据源刷新类目占比显示，保留 WPS 图片和其它工作表。"""

from __future__ import annotations

import argparse
import json
import posixpath
import re
import tempfile
import zipfile
from pathlib import Path
from xml.dom import minidom

from openpyxl import load_workbook

from .fill_custom_template import (
    CATEGORY_HEADER_MARKER, CATEGORY_SECTION_MARKER, CATEGORY_PRIMARY_COLOR,
    CATEGORY_TRANSFER_COLOR, CATEGORY_PRIMARY_SHARE_THRESHOLD, category_path_key,
    category_section_title, category_share_number_format,
    collect_category_summaries, extract_node_id, find_next_section_row,
    find_research_sheet, find_row_containing, find_source_sheet,
)

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def elements(node, name):
    return list(node.getElementsByTagNameNS(NS, name))


def node_from_link(url):
    node = extract_node_id(url)
    if node:
        return node
    match = re.search(r"/gp/bestsellers/(?:[^/?#]+/)?(\d+)(?:[/?#]|$)", url or "")
    return match[1] if match else None


def share_fill_id(styles, share, current_id):
    """沿用原蓝/粉配色；源数据变化时同步重算底色。"""
    color = CATEGORY_PRIMARY_COLOR if share > CATEGORY_PRIMARY_SHARE_THRESHOLD else CATEGORY_TRANSFER_COLOR
    fills = elements(styles, "fills")[0]
    candidates = list(enumerate(elements(fills, "fill")))
    candidates.sort(key=lambda item: item[0] != current_id)
    for index, fill in candidates:
        patterns = elements(fill, "patternFill")
        foregrounds = elements(fill, "fgColor")
        if (patterns and patterns[0].getAttribute("patternType") == "solid"
                and foregrounds and foregrounds[0].getAttribute("rgb")[-6:].upper() == color):
            return index
    index = len(elements(fills, "fill"))
    fill = styles.createElementNS(NS, "fill")
    pattern = styles.createElementNS(NS, "patternFill")
    pattern.setAttribute("patternType", "solid")
    foreground = styles.createElementNS(NS, "fgColor")
    foreground.setAttribute("rgb", f"FF{color}")
    pattern.appendChild(foreground)
    fill.appendChild(pattern)
    fills.appendChild(fill)
    fills.setAttribute("count", str(index + 1))
    return index


def collect_updates(workbook, source_workbook):
    """按节点关联当前行，不依赖两张表的行序。"""
    source = find_source_sheet(source_workbook)
    sheet = find_research_sheet(workbook)
    if source is None or sheet is None:
        raise ValueError("缺少数据源或市场调研工作表。")
    summaries = collect_category_summaries(source)
    if not summaries:
        raise ValueError("数据源没有可计算占比的有效类目产品。")
    nodes = {node_from_link(s.final_link): s for s in summaries if node_from_link(s.final_link)}
    paths = {}
    for summary in summaries:
        paths.setdefault(category_path_key(summary.path_text), []).append(summary)
    header = find_row_containing(sheet, CATEGORY_HEADER_MARKER)
    section = find_row_containing(sheet, CATEGORY_SECTION_MARKER)
    if header is None or section is None:
        raise ValueError("找不到类目区域。")
    updates = []
    for row in range(header + 1, find_next_section_row(sheet, header + 1)):
        if sheet.cell(row, 1).value in (None, ""):
            continue
        cells = [sheet.cell(row, col) for col in (*range(1, 6), *range(10, sheet.max_column + 1))
                 if sheet.cell(row, col).value not in (None, "")]
        leaf = cells[-1]
        node = node_from_link(leaf.hyperlink.target if leaf.hyperlink else None)
        path = category_path_key(":".join(str(c.value).splitlines()[0] for c in cells))
        candidates = [nodes[node]] if node in nodes else ([] if node else paths.get(path, []))
        if len(candidates) != 1:
            raise ValueError(f"第 {row} 行类目无法唯一匹配数据源，未写入占比。")
        summary = candidates[0]
        updates.append({"cell": f"I{row}", "node": node, "count": summary.product_count,
                        "share": summary.product_share,
                        "format": category_share_number_format(summary.product_share)})
    if not updates:
        raise ValueError("类目区域没有数据行。")
    return sheet.title, section, category_section_title(sheet.cell(section, 1).value), updates


def refresh_category_shares(workbook_path: Path, source_path: Path, output: Path) -> dict:
    if output.resolve() in {workbook_path.resolve(), source_path.resolve()}:
        raise ValueError("输出必须使用独立路径，不能覆盖报告或数据源。")
    workbook = load_workbook(workbook_path, rich_text=True)
    source = load_workbook(source_path, rich_text=True)
    try:
        sheet_name, section, title, updates = collect_updates(workbook, source)
    finally:
        workbook.close()
        source.close()

    with zipfile.ZipFile(workbook_path) as archive:
        book = minidom.parseString(archive.read("xl/workbook.xml"))
        sheet = next(s for s in elements(book, "sheet") if s.getAttribute("name") == sheet_name)
        rels = minidom.parseString(archive.read("xl/_rels/workbook.xml.rels"))
        rel = next(r for r in rels.getElementsByTagName("Relationship")
                   if r.getAttribute("Id") == sheet.getAttributeNS(REL_NS, "id"))
        part = posixpath.normpath(posixpath.join("xl", rel.getAttribute("Target"))).lstrip("/")
        doc = minidom.parseString(archive.read(part))
        styles = minidom.parseString(archive.read("xl/styles.xml"))
        xfs = elements(styles, "cellXfs")[0]
        formats = elements(styles, "numFmts")
        if formats:
            formats = formats[0]
        else:
            formats = styles.createElementNS(NS, "numFmts")
            styles.documentElement.insertBefore(formats, styles.documentElement.firstChild)
        existing_formats = {f.getAttribute("formatCode"): int(f.getAttribute("numFmtId"))
                            for f in elements(formats, "numFmt")}
        next_id = max([163, *existing_formats.values()]) + 1
        cells = {c.getAttribute("r"): c for c in elements(doc, "c")}
        for update in updates:
            cell = cells.get(update["cell"])
            if cell is None:
                raise ValueError(f"缺少目标单元格 {update['cell']}。")
            code = update["format"]
            if code not in existing_formats:
                fmt = styles.createElementNS(NS, "numFmt")
                fmt.setAttribute("numFmtId", str(next_id))
                fmt.setAttribute("formatCode", code)
                formats.appendChild(fmt)
                existing_formats[code] = next_id
                next_id += 1
            xf = elements(xfs, "xf")[int(cell.getAttribute("s") or "0")].cloneNode(True)
            xf.setAttribute("numFmtId", str(existing_formats[code]))
            xf.setAttribute("applyNumberFormat", "1")
            xf.setAttribute("fillId", str(share_fill_id(styles, update["share"], int(xf.getAttribute("fillId") or "0"))))
            xf.setAttribute("applyFill", "1")
            alignment = elements(xf, "alignment")
            if alignment:
                alignment = alignment[0]
            else:
                alignment = styles.createElementNS(NS, "alignment")
                xf.appendChild(alignment)
            alignment.setAttribute("wrapText", "1")
            xf.setAttribute("applyAlignment", "1")
            cell.setAttribute("s", str(len(elements(xfs, "xf"))))
            xfs.appendChild(xf)
        formats.setAttribute("count", str(len(elements(formats, "numFmt"))))
        xfs.setAttribute("count", str(len(elements(xfs, "xf"))))
        title_cell = cells[f"A{section}"]
        for child in list(title_cell.childNodes):
            title_cell.removeChild(child)
        title_cell.setAttribute("t", "inlineStr")
        inline = doc.createElementNS(NS, "is")
        text = doc.createElementNS(NS, "t")
        text.setAttribute("xml:space", "preserve")
        text.appendChild(doc.createTextNode(title))
        inline.appendChild(text)
        title_cell.appendChild(inline)
        # XML 属性中的换行必须转义，否则重开工作簿会归一化为空格。
        style_xml = re.sub(r'formatCode="[^"]*"', lambda m: m[0].replace("\n", "&#10;"), styles.toxml())
        patches = {part: doc.toxml(encoding="utf-8"), "xl/styles.xml": style_xml.encode("utf-8")}
        output.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(dir=output.parent, suffix=".xlsx", delete=False) as handle:
            temporary = Path(handle.name)
        try:
            with zipfile.ZipFile(temporary, "w") as destination:
                for info in archive.infolist():
                    destination.writestr(info, patches.get(info.filename, archive.read(info.filename)))
            with zipfile.ZipFile(temporary) as check:
                if check.testzip():
                    raise ValueError("输出工作簿 ZIP 校验失败。")
            temporary.replace(output)
        finally:
            temporary.unlink(missing_ok=True)
    return {"sheet": sheet_name, "rows": updates}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workbook", type=Path, required=True)
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    result = refresh_category_shares(args.workbook, args.source, args.output)
    from modules.shared.quality_gates import validate_category_research
    gate = validate_category_research(args.output)
    print(json.dumps({"updates": result, "gate": gate.as_dict()}, ensure_ascii=False, default=str))


if __name__ == "__main__":
    main()
