"""更新关键词竞价并排序，保留工作簿的 WPS 图片及其它 ZIP 部件。"""

from __future__ import annotations

import math
import posixpath
import re
import tempfile
import zipfile
from pathlib import Path
from xml.dom import minidom

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import column_index_from_string

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DRAW_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"


def numeric_value(value: object) -> float | None:
    """只接受完整数值，避免把缺失标记或公式中的数字当成结果。"""
    if isinstance(value, bool):
        return None
    try:
        number = float(str(value).strip().replace(",", "").replace("$", "").replace("£", "").replace("€", ""))
    except (TypeError, ValueError):
        return None
    return number if math.isfinite(number) else None


def latest_monthly_bid(points: list[dict]) -> tuple[str | None, float | None]:
    """先确定最新月份，再取该月有效竞价；最新月缺值时不退回旧月。"""
    monthly = []
    for point in points:
        match = re.fullmatch(r"(\d{4})-(\d{1,2})", str(point.get("month", "")).strip())
        if not match or not 1 <= int(match[2]) <= 12:
            continue
        monthly.append((f"{int(match[1]):04d}-{int(match[2]):02d}", numeric_value(point.get("bid"))))
    if not monthly:
        return None, None
    month = max(key for key, _ in monthly)
    values = [bid for key, bid in monthly if key == month and bid is not None and bid > 0]
    return month, max(values) if values else None


def _elements(node, namespace: str, name: str):
    return list(node.getElementsByTagNameNS(namespace, name))


def _target(part: str, target: str) -> str:
    return posixpath.normpath(posixpath.join(posixpath.dirname(part), target)).lstrip("/")


def _relationships(archive, part: str) -> dict:
    rel_path = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
    if rel_path not in archive.namelist():
        return {}
    doc = minidom.parseString(archive.read(rel_path))
    return {r.getAttribute("Id"): r for r in doc.getElementsByTagName("Relationship")}


def _move_ref(reference: str, mapping: dict[int, int]) -> str:
    def replace(match):
        return match[1] + str(mapping.get(int(match[2]), int(match[2])))
    return re.sub(r"(\$?[A-Z]+\$?)(\d+)", replace, reference)


def update_keyword_rows(
    source: Path,
    output: Path,
    *,
    sheet_name: str,
    rows: list[int],
    peak_column: int,
    bid_column: int | None = None,
    bids: dict[int, float | None] | None = None,
) -> dict:
    """只替换目标工作表及需移动的绘图/批注部件，不重存整个工作簿。"""
    rows = sorted(rows)
    if not rows or len(set(rows)) != len(rows):
        raise ValueError("关键词行必须非空且不能重复。")
    if bids is not None and (bid_column is None or set(bids) != set(rows)):
        raise ValueError("竞价更新必须覆盖全部目标关键词，不能留下新旧竞价混合结果。")
    workbook = load_workbook(source, read_only=True, data_only=True)
    try:
        sheet = workbook[sheet_name]
        scores = {row: numeric_value(sheet.cell(row, peak_column).value) for row in rows}
    finally:
        workbook.close()
    ordered = sorted(rows, key=lambda row: (scores[row] is None, -(scores[row] or 0)))
    mapping = dict(zip(ordered, rows))
    patches = {}
    with zipfile.ZipFile(source) as archive:
        book = minidom.parseString(archive.read("xl/workbook.xml"))
        descriptor = next(s for s in _elements(book, MAIN_NS, "sheet") if s.getAttribute("name") == sheet_name)
        rel = _relationships(archive, "xl/workbook.xml")[descriptor.getAttributeNS(REL_NS, "id")]
        part = _target("xl/workbook.xml", rel.getAttribute("Target"))
        doc = minidom.parseString(archive.read(part))
        row_nodes = {int(r.getAttribute("r")): r for r in _elements(doc, MAIN_NS, "row")}
        if not set(rows).issubset(row_nodes):
            raise ValueError("关键词行在 XML 中不存在。")
        for merge in _elements(doc, MAIN_NS, "mergeCell"):
            merge_rows = [int(r) for r in re.findall(r"\d+", merge.getAttribute("ref"))]
            if len(merge_rows) == 2 and merge_rows[0] != merge_rows[1] and any(merge_rows[0] <= r <= merge_rows[1] for r in rows):
                raise ValueError("关键词区域有跨行合并，无法安全排序。")
            merge.setAttribute("ref", _move_ref(merge.getAttribute("ref"), mapping))
        for old in rows:
            row_node = row_nodes[old]
            cells = _elements(row_node, MAIN_NS, "c")
            if bids is not None:
                cell = next((c for c in cells if column_index_from_string(re.match(r"[A-Z]+", c.getAttribute("r"))[0]) == bid_column), None)
                if cell is None:
                    raise ValueError(f"第 {old} 行缺少竞价模板单元格。")
                for child in list(cell.childNodes):
                    cell.removeChild(child)
                value = bids[old]
                if value is None:
                    cell.setAttribute("t", "inlineStr")
                    inline = doc.createElementNS(MAIN_NS, "is")
                    text = doc.createElementNS(MAIN_NS, "t")
                    text.appendChild(doc.createTextNode("/"))
                    inline.appendChild(text)
                    cell.appendChild(inline)
                else:
                    cell.setAttribute("t", "n")
                    v = doc.createElementNS(MAIN_NS, "v")
                    v.appendChild(doc.createTextNode(str(value)))
                    cell.appendChild(v)
            row_node.setAttribute("r", str(mapping[old]))
            for cell in cells:
                old_ref = cell.getAttribute("r")
                new_ref = _move_ref(old_ref, mapping)
                for formula in _elements(cell, MAIN_NS, "f"):
                    if old_ref == new_ref:
                        continue
                    if formula.getAttribute("t") in {"shared", "array", "dataTable"}:
                        raise ValueError("关键词区域含共享、数组或数据表公式，无法安全移动。")
                    expression = "".join(c.data for c in formula.childNodes if c.nodeType == c.TEXT_NODE)
                    if expression and "DISPIMG(" not in expression:
                        translated = Translator("=" + expression, origin=old_ref).translate_formula(new_ref)[1:]
                        for child in list(formula.childNodes):
                            formula.removeChild(child)
                        formula.appendChild(doc.createTextNode(translated))
                cell.setAttribute("r", new_ref)
        sheet_data = _elements(doc, MAIN_NS, "sheetData")[0]
        for node in sorted(row_nodes.values(), key=lambda n: int(n.getAttribute("r"))):
            sheet_data.appendChild(node)
        for link in _elements(doc, MAIN_NS, "hyperlink"):
            link.setAttribute("ref", _move_ref(link.getAttribute("ref"), mapping))
        patches[part] = doc.toxml(encoding="utf-8")
        for relation in _relationships(archive, part).values():
            kind = relation.getAttribute("Type").rsplit("/", 1)[-1]
            if kind not in {"drawing", "comments"}:
                continue
            target = _target(part, relation.getAttribute("Target"))
            linked = minidom.parseString(archive.read(target))
            changed = False
            if kind == "drawing":
                for anchor in linked.documentElement.childNodes:
                    if anchor.nodeType != anchor.ELEMENT_NODE:
                        continue
                    starts = _elements(anchor, DRAW_NS, "from")
                    if not starts:
                        continue
                    marker = _elements(starts[0], DRAW_NS, "row")[0]
                    old_row = int(marker.firstChild.data) + 1
                    if old_row not in mapping or mapping[old_row] == old_row:
                        continue
                    delta = mapping[old_row] - old_row
                    for row_marker in _elements(anchor, DRAW_NS, "row"):
                        row_marker.firstChild.data = str(int(row_marker.firstChild.data) + delta)
                    changed = True
            else:
                for comment in _elements(linked, MAIN_NS, "comment"):
                    ref = comment.getAttribute("ref")
                    moved = _move_ref(ref, mapping)
                    comment.setAttribute("ref", moved)
                    changed |= moved != ref
            if changed:
                patches[target] = linked.toxml(encoding="utf-8")
        output.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(dir=output.parent, suffix=".xlsx", delete=False) as handle:
            temp_path = Path(handle.name)
        try:
            with zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as destination:
                for entry in archive.infolist():
                    destination.writestr(entry, patches.get(entry.filename, archive.read(entry.filename)))
            with zipfile.ZipFile(temp_path) as check:
                if check.testzip() is not None:
                    raise ValueError("输出 ZIP 校验失败。")
        except BaseException:
            temp_path.unlink(missing_ok=True)
            raise
    try:
        temp_path.replace(output)
    finally:
        temp_path.unlink(missing_ok=True)
    return {"rows": len(rows), "moved": sum(old != new for old, new in mapping.items()), "row_mapping": mapping, "patched_parts": list(patches)}
