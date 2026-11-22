from __future__ import annotations

import argparse
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from openpyxl import load_workbook

XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
EMU_PER_CM = 360000


def find_market_analysis_sheet(workbook):
    for name in workbook.sheetnames:
        if name.endswith("市场分析"):
            return workbook[name]
    raise RuntimeError("No sheet ending with 市场分析 found.")


def sheet_xml_path(workbook, worksheet) -> str:
    index = workbook.worksheets.index(worksheet) + 1
    return f"xl/worksheets/sheet{index}.xml"


def drawing_target(path: Path, workbook, worksheet) -> str:
    sheet_path = sheet_xml_path(workbook, worksheet)
    rels_path = f"xl/worksheets/_rels/{Path(sheet_path).name}.rels"
    with zipfile.ZipFile(path) as archive:
        sheet_xml = archive.read(sheet_path).decode("utf-8", "ignore")
        rels_xml = archive.read(rels_path).decode("utf-8", "ignore")
    marker = "r:id=\""
    start = sheet_xml.find(marker)
    if start == -1:
        raise RuntimeError(f"No drawing relationship found for {worksheet.title}")
    start += len(marker)
    end = sheet_xml.find("\"", start)
    rel_id = sheet_xml[start:end]
    rel_marker = f'Id="{rel_id}"'
    rel_pos = rels_xml.find(rel_marker)
    if rel_pos == -1:
        raise RuntimeError(f"Drawing relationship {rel_id} not found.")
    target_marker = 'Target="'
    target_start = rels_xml.rfind(target_marker, 0, rel_pos)
    if target_start == -1:
        target_start = rels_xml.find(target_marker, rel_pos)
    target_start += len(target_marker)
    target_end = rels_xml.find("\"", target_start)
    target = rels_xml[target_start:target_end]
    if target.startswith("/"):
        return target.lstrip("/")
    return str(Path(sheet_path).parent / target).replace("\\", "/")


def patch_drawing_xml(xml_bytes: bytes, width_cm: float, height_cm: float) -> bytes:
    ET.register_namespace("", XDR_NS)
    root = ET.fromstring(xml_bytes)
    width_emu = str(round(width_cm * EMU_PER_CM))
    height_emu = str(round(height_cm * EMU_PER_CM))
    patched = 0
    for anchor in root.findall(f"{{{XDR_NS}}}oneCellAnchor"):
        if anchor.find(f"{{{XDR_NS}}}graphicFrame") is None:
            continue
        ext = anchor.find(f"{{{XDR_NS}}}ext")
        if ext is not None:
            ext.attrib["cx"] = width_emu
            ext.attrib["cy"] = height_emu
            patched += 1
    if patched == 0:
        raise RuntimeError("No chart anchors patched in drawing XML.")
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def fit_charts(path: Path, output: Path, width: float, height: float) -> dict:
    workbook = load_workbook(path)
    worksheet = find_market_analysis_sheet(workbook)
    for chart in getattr(worksheet, "_charts", []):
        chart.width = width
        chart.height = height
    drawing_path = drawing_target(path, workbook, worksheet)
    temp_output = output.with_name(output.stem + "_openpyxl_tmp.xlsx")
    output.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(temp_output)
    with zipfile.ZipFile(temp_output, "r") as zin, zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED, allowZip64=True) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == drawing_path:
                data = patch_drawing_xml(data, width, height)
            zout.writestr(item, data)
    temp_output.unlink(missing_ok=True)
    return {
        "output": str(output),
        "sheet": worksheet.title,
        "charts": len(getattr(worksheet, "_charts", [])),
        "width_cm": width,
        "height_cm": height,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Resize market analysis charts to fit template boxes.")
    parser.add_argument("--input", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--width", type=float, default=18.8)
    parser.add_argument("--height", type=float, default=7.2)
    args = parser.parse_args()

    result = fit_charts(args.input, args.output, args.width, args.height)
    for key, value in result.items():
        print(f"{key}: {value}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
