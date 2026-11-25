"""将整体销售指标及六张本地图填入新版市场调研页，不另建市场分析页。"""

from __future__ import annotations

import argparse
import json
import re
import sys
import tempfile
from copy import copy
from pathlib import Path
from xml.dom import minidom
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.drawing.image import Image
from openpyxl.utils.cell import column_index_from_string

MODULES = Path(__file__).resolve().parents[2]
if str(MODULES) not in sys.path:
    sys.path.insert(0, str(MODULES))
from market_analysis.excel_market_report import compute_metrics, find_source_sheet, read_rows
from product_research.scripts.insert_sellersprite_screenshots import region_pixels
from competitor_analysis.annotate_review_export import worksheet_part, children, element


METRICS = {
    "样本数": "sample_count", "平均价格($)": "avg_price", "中位价格($)": "median_price",
    "Top5销量占比": "top5_share", "Top10销量占比": "top10_share", "Top20销量占比": "top20_share",
    "总销量": "total_sales", "平均月销量": "avg_sales", "中位月销量": "median_sales",
    "总销售额($)": "total_revenue", "平均销售额($)": "avg_revenue", "中位销售额($)": "median_revenue",
}


def prepare_inputs(workbook, output_dir):
    """只统计有效ASIN，排除模板尾部的空行，并保留统计来源。"""
    records = [row for row in read_rows(find_source_sheet(workbook))
               if re.fullmatch(r"[A-Z0-9]{10}", str(row.get("ASIN", "")).strip())]
    asins = [row["ASIN"] for row in records]
    if not asins or len(asins) != len(set(asins)):
        raise ValueError("有效ASIN为空或重复，停止统计")
    metrics = compute_metrics(records)
    metrics["avg_revenue"] = metrics["total_revenue"] / metrics["sample_count"] if metrics["sample_count"] else 0
    output_dir.mkdir(parents=True, exist_ok=True)
    for name, payload in (("source_records.json", {"source_records": records}), ("market_metrics.json", metrics)):
        (output_dir / name).write_text(json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    return metrics


def fill_summary(worksheet, metrics):
    """每次按唯一指标标签找右侧结果格，遇到结论合并区域时拒绝写入。"""
    written = {}
    for label, key in METRICS.items():
        matches = [cell for row in worksheet for cell in row
                   if re.sub(r"\s+", "", str(cell.value or "")) == label]
        if len(matches) != 1:
            raise ValueError(f"整体销售指标标签不唯一：{label}")
        source = matches[0]
        target = worksheet.cell(source.row, source.column + 1)
        if isinstance(target, MergedCell) or any(target.coordinate in area for area in worksheet.merged_cells.ranges):
            raise ValueError(f"{label}右侧与合并区域冲突：{target.coordinate}")
        target.value = metrics[key]
        target.number_format = "0.0%" if key.endswith("_share") else "#,##0.00" if "价格" in label or "销售额" in label else "#,##0"
        written[label] = target.coordinate
    return written


def refresh_summary(input_path, output, data_dir):
    """本地统计并定点写入12项指标，不重存既有图片、评价或原生部件。"""
    if output.exists() or input_path.resolve() == output.resolve():
        raise ValueError('整体指标输出必须是新的独立文件')
    book = load_workbook(input_path, rich_text=True)
    try:
        metrics = prepare_inputs(book, data_dir)
        sheet = next(s for s in book if s.title.endswith('市场调研'))
        locations = fill_summary(sheet, metrics)
        values = {addr: (sheet[addr].value, sheet[addr].number_format) for addr in locations.values()}
        name = sheet.title
    finally:
        book.close()
    output.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(input_path) as archive, tempfile.TemporaryDirectory(dir=output.parent) as temporary:
        part = worksheet_part(archive, name)
        doc = minidom.parseString(archive.read(part))
        styles = minidom.parseString(archive.read('xl/styles.xml'))
        root = styles.documentElement
        formats = children(root, 'numFmts')
        if not formats:
            formats = element(root, 'numFmts', {'count': 0})
            root.insertBefore(formats, root.firstChild)
        else:
            formats = formats[0]
        format_ids = {f.getAttribute('formatCode'): int(f.getAttribute('numFmtId')) for f in children(formats, 'numFmt')}
        format_ids.update({'#,##0': 3, '#,##0.00': 4})
        xfs = children(root, 'cellXfs')[0]
        data = children(doc.documentElement, 'sheetData')[0]
        rows = {int(r.getAttribute('r')): r for r in children(data, 'row')}
        for address, (value, code) in values.items():
            row = rows[int(re.search(r'\d+$', address)[0])]
            matches = [c for c in children(row, 'c') if c.getAttribute('r') == address]
            if matches:
                cell = matches[0]
            else:
                # 未填过且无独立样式的空格可能没有XML，按列序补节点，不移动其它单元格。
                cell = element(row, 'c', {'r': address})
                column = column_index_from_string(re.match(r'[A-Z]+', address)[0])
                following = next((c for c in children(row, 'c') if column_index_from_string(
                    re.match(r'[A-Z]+', c.getAttribute('r'))[0]) > column), None)
                row.insertBefore(cell, following) if following is not None else row.appendChild(cell)
            for child in list(cell.childNodes):
                cell.removeChild(child)
            cell.setAttribute('t', 'n')
            node = element(cell, 'v')
            node.appendChild(doc.createTextNode(str(value)))
            cell.appendChild(node)
            if code not in format_ids:
                format_ids[code] = max([163, *format_ids.values()]) + 1
                formats.appendChild(element(formats, 'numFmt', {'numFmtId': format_ids[code], 'formatCode': code}))
            xf = children(xfs, 'xf')[int(cell.getAttribute('s') or '0')].cloneNode(True)
            xf.setAttribute('numFmtId', str(format_ids[code]))
            xf.setAttribute('applyNumberFormat', '1')
            cell.setAttribute('s', str(len(children(xfs, 'xf'))))
            xfs.appendChild(xf)
        formats.setAttribute('count', str(len(children(formats, 'numFmt'))))
        xfs.setAttribute('count', str(len(children(xfs, 'xf'))))
        style_xml = re.sub(r'formatCode="[^"]*"', lambda m: m[0].replace('\n', '&#10;'), styles.toxml())
        changes = {part: doc.toxml(encoding='utf-8'), 'xl/styles.xml': style_xml.encode('utf-8')}
        staged = Path(temporary) / 'metrics.xlsx'
        with ZipFile(staged, 'w') as result:
            for info in archive.infolist():
                result.writestr(copy(info), changes.get(info.filename, archive.read(info.filename)))
        checked = load_workbook(staged, rich_text=True)
        try:
            for address, (value, code) in values.items():
                cell = checked[name][address]
                if abs(cell.value - value) > max(1, abs(value)) * 1e-12 or cell.number_format != code:
                    raise ValueError(f'整体指标写入校验失败：{address}')
        finally:
            checked.close()
        staged.replace(output)
    return {'output': str(output), 'metrics': locations, 'sample_count': metrics['sample_count']}


def insert_local_charts(worksheet, chart_dir):
    """按品牌和上架时间所在行的合并图框等比例放置，保留原行高和图片像素。"""
    for label, logical_row in (("品牌分析", 10), ("上架时间分析", 11)):
        cells = [cell for row in worksheet for cell in row if str(cell.value or "").strip() == label]
        if len(cells) != 1:
            raise ValueError(f"图表标签不唯一：{label}")
        row = cells[0].row
        targets = {1, 3, 5}
        worksheet._images = [img for img in worksheet._images
                             if not (hasattr(img.anchor, "_from") and img.anchor._from.row == row - 1
                                     and img.anchor._from.col in targets)]
        for col in ("B", "D", "F"):
            anchor = f"{col}{row}"
            areas = [area for area in worksheet.merged_cells.ranges if area.start_cell.coordinate == anchor]
            if len(areas) != 1:
                raise ValueError(f"缺少图表合并区域：{anchor}")
            width, height = region_pixels(worksheet, areas[0])
            image = Image(chart_dir / f"{col}{logical_row}_hires.png")
            scale = min((width - 8) / image.width, (height - 8) / image.height)
            image.width, image.height = image.width * scale, image.height * scale
            worksheet.add_image(image, anchor)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, required=True)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--data-dir", type=Path, required=True)
    parser.add_argument("--chart-dir", type=Path)
    args = parser.parse_args()
    if args.output and args.input.resolve() == args.output.resolve():
        raise ValueError("请使用独立输出文件")
    wb = load_workbook(args.input, rich_text=True)
    try:
        metrics = prepare_inputs(wb, args.data_dir)
        if args.output:
            worksheet = next(ws for ws in wb if ws.title.endswith("市场调研"))
            locations = fill_summary(worksheet, metrics)
            if args.chart_dir:
                insert_local_charts(worksheet, args.chart_dir)
            args.output.parent.mkdir(parents=True, exist_ok=True)
            wb.save(args.output)
            print(json.dumps({"output": str(args.output), "metrics": locations}, ensure_ascii=False))
        print(f"整体销售数据有效样本：{metrics['sample_count']}")
    finally:
        wb.close()


if __name__ == "__main__":
    main()
