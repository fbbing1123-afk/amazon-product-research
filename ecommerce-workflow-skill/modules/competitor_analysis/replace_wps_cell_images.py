from __future__ import annotations

import argparse
import os
import posixpath
import re
import struct
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path


SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOCUMENT_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
DRAWING_MAIN_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
DISPIMG_PATTERN = re.compile(r'DISPIMG\("([^"]+)"', re.I)


def parse_replacements(values: list[str]) -> dict[str, Path]:
    """解析重复传入的 单元格=图片路径 参数。"""
    result: dict[str, Path] = {}
    for value in values:
        if "=" not in value:
            raise ValueError(f"替换参数格式错误，应为 C2=图片路径：{value}")
        coordinate, image_path = value.split("=", 1)
        coordinate = coordinate.strip().upper()
        if not re.fullmatch(r"[A-Z]{1,3}[1-9]\d*", coordinate):
            raise ValueError(f"无效单元格坐标：{coordinate}")
        path = Path(image_path.strip()).expanduser().resolve()
        if not path.is_file():
            raise FileNotFoundError(f"图片不存在：{path}")
        result[coordinate] = path
    if not result:
        raise ValueError("至少需要一个 --replace 参数。")
    return result


def png_dimensions(payload: bytes) -> tuple[int, int]:
    """读取 PNG 尺寸并拒绝损坏或空白资源。"""
    if len(payload) < 24 or payload[:8] != b"\x89PNG\r\n\x1a\n":
        raise ValueError("替换图片必须是有效 PNG。")
    width, height = struct.unpack(">II", payload[16:24])
    if width < 100 or height < 100:
        raise ValueError(f"替换图片尺寸无效：{width}x{height}")
    return width, height


def relationship_targets(root: ET.Element) -> dict[str, str]:
    """读取 OOXML 关系编号与目标部件映射。"""
    return {
        item.attrib["Id"]: item.attrib["Target"]
        for item in root.findall(f"{{{PACKAGE_REL_NS}}}Relationship")
    }


def target_sheet_part(archive: zipfile.ZipFile, sheet_suffix: str) -> tuple[str, str]:
    """按工作表名称后缀定位对应 worksheet XML。"""
    workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
    workbook_rels = relationship_targets(
        ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    )
    matches: list[tuple[str, str]] = []
    for sheet in workbook_root.findall(f".//{{{SPREADSHEET_NS}}}sheet"):
        name = sheet.attrib.get("name", "")
        if not name.endswith(sheet_suffix):
            continue
        relation_id = sheet.attrib[f"{{{DOCUMENT_REL_NS}}}id"]
        target = workbook_rels[relation_id]
        part = posixpath.normpath(posixpath.join("xl", target))
        matches.append((name, part))
    if len(matches) != 1:
        raise ValueError(
            f"按后缀 {sheet_suffix!r} 应找到 1 个工作表，实际找到 {len(matches)} 个。"
        )
    return matches[0]


def cell_image_ids(
    archive: zipfile.ZipFile,
    sheet_part: str,
    coordinates: set[str],
) -> dict[str, str]:
    """读取目标单元格 DISPIMG 公式中的图片编号。"""
    root = ET.fromstring(archive.read(sheet_part))
    result: dict[str, str] = {}
    for cell in root.findall(f".//{{{SPREADSHEET_NS}}}c"):
        coordinate = cell.attrib.get("r", "").upper()
        if coordinate not in coordinates:
            continue
        formula = cell.find(f"{{{SPREADSHEET_NS}}}f")
        match = DISPIMG_PATTERN.search(formula.text or "") if formula is not None else None
        if not match:
            raise ValueError(f"{coordinate} 不是 WPS DISPIMG 单元格图片。")
        result[coordinate] = match.group(1)
    missing = coordinates - result.keys()
    if missing:
        raise ValueError(f"工作表中缺少目标单元格：{', '.join(sorted(missing))}")
    return result


def image_media_parts(archive: zipfile.ZipFile) -> dict[str, str]:
    """建立 WPS 图片编号到 xl/media 文件的映射。"""
    cell_root = ET.fromstring(archive.read("xl/cellimages.xml"))
    rel_targets = relationship_targets(
        ET.fromstring(archive.read("xl/_rels/cellimages.xml.rels"))
    )
    result: dict[str, str] = {}
    for item in cell_root:
        properties = item.find(f".//{{{DRAWING_NS}}}cNvPr")
        blip = item.find(f".//{{{DRAWING_MAIN_NS}}}blip")
        if properties is None or blip is None:
            continue
        image_id = properties.attrib.get("name")
        relation_id = blip.attrib.get(f"{{{DOCUMENT_REL_NS}}}embed")
        if not image_id or not relation_id or relation_id not in rel_targets:
            continue
        result[image_id] = posixpath.normpath(
            posixpath.join("xl", rel_targets[relation_id])
        )
    return result


def replace_wps_cell_images(
    workbook: Path,
    output: Path,
    replacements: dict[str, Path],
    sheet_suffix: str,
) -> dict[str, object]:
    """只替换指定 DISPIMG 的媒体字节，保留其它工作簿部件。"""
    workbook = workbook.expanduser().resolve()
    output = output.expanduser().resolve()
    if not workbook.is_file():
        raise FileNotFoundError(f"工作簿不存在：{workbook}")

    with zipfile.ZipFile(workbook, "r") as source:
        sheet_name, sheet_part = target_sheet_part(source, sheet_suffix)
        coordinate_ids = cell_image_ids(source, sheet_part, set(replacements))
        media_parts = image_media_parts(source)
        payloads: dict[str, bytes] = {}
        dimensions: dict[str, tuple[int, int]] = {}
        for coordinate, image_path in replacements.items():
            image_id = coordinate_ids[coordinate]
            media_part = media_parts.get(image_id)
            if not media_part:
                raise ValueError(f"找不到 {coordinate} 对应的媒体资源：{image_id}")
            payload = image_path.read_bytes()
            dimensions[coordinate] = png_dimensions(payload)
            payloads[media_part] = payload

        output.parent.mkdir(parents=True, exist_ok=True)
        handle, temporary_name = tempfile.mkstemp(
            prefix=f".{output.stem}_",
            suffix=".xlsx",
            dir=output.parent,
        )
        os.close(handle)
        temporary = Path(temporary_name)
        try:
            with zipfile.ZipFile(temporary, "w") as target:
                for item in source.infolist():
                    target.writestr(item, payloads.get(item.filename, source.read(item.filename)))
            with zipfile.ZipFile(temporary, "r") as checked:
                broken = checked.testzip()
                if broken:
                    raise ValueError(f"输出工作簿压缩包损坏：{broken}")
            os.replace(temporary, output)
        finally:
            temporary.unlink(missing_ok=True)

    return {
        "sheet": sheet_name,
        "cells": sorted(replacements),
        "dimensions": {key: list(value) for key, value in dimensions.items()},
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="原位替换 WPS DISPIMG 单元格图片资源。")
    parser.add_argument("--workbook", type=Path, required=True)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--sheet-suffix", default="竞品分析")
    parser.add_argument("--replace", action="append", default=[])
    args = parser.parse_args()

    replacements = parse_replacements(args.replace)
    result = replace_wps_cell_images(
        args.workbook,
        args.output or args.workbook,
        replacements,
        args.sheet_suffix,
    )
    print(result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
