from __future__ import annotations

import argparse
import io
import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from PIL import Image as PILImage


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
ETC_NS = "http://www.wps.cn/officeDocument/2017/etCustomData"
XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
EMU_PER_PIXEL = 9525


ET.register_namespace("", MAIN_NS)
ET.register_namespace("etc", ETC_NS)
ET.register_namespace("xdr", XDR_NS)
ET.register_namespace("a", A_NS)
ET.register_namespace("r", R_NS)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="把美国类目 Softime 截图写入 WPS 单元格图片。")
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--image-dir", required=True, type=Path)
    return parser.parse_args()


def append_cell_image(
    cell_root: ET.Element,
    image_id: str,
    rel_id: str,
    payload: bytes,
    picture_id: int,
) -> None:
    cell_image = ET.SubElement(cell_root, f"{{{ETC_NS}}}cellImage")
    picture = ET.SubElement(cell_image, f"{{{XDR_NS}}}pic")
    non_visual = ET.SubElement(picture, f"{{{XDR_NS}}}nvPicPr")
    ET.SubElement(
        non_visual,
        f"{{{XDR_NS}}}cNvPr",
        {"id": str(picture_id), "name": image_id, "descr": "Picture"},
    )
    ET.SubElement(non_visual, f"{{{XDR_NS}}}cNvPicPr")
    blip_fill = ET.SubElement(picture, f"{{{XDR_NS}}}blipFill")
    ET.SubElement(blip_fill, f"{{{A_NS}}}blip", {f"{{{R_NS}}}embed": rel_id})
    stretch = ET.SubElement(blip_fill, f"{{{A_NS}}}stretch")
    ET.SubElement(stretch, f"{{{A_NS}}}fillRect")
    shape_properties = ET.SubElement(picture, f"{{{XDR_NS}}}spPr")
    transform = ET.SubElement(shape_properties, f"{{{A_NS}}}xfrm")
    ET.SubElement(transform, f"{{{A_NS}}}off", {"x": "0", "y": "0"})
    with PILImage.open(io.BytesIO(payload)) as image:
        width_emu = str(image.width * EMU_PER_PIXEL)
        height_emu = str(image.height * EMU_PER_PIXEL)
    ET.SubElement(transform, f"{{{A_NS}}}ext", {"cx": width_emu, "cy": height_emu})
    geometry = ET.SubElement(shape_properties, f"{{{A_NS}}}prstGeom", {"prst": "rect"})
    ET.SubElement(geometry, f"{{{A_NS}}}avLst")


def set_dispimg_formula(sheet_root: ET.Element, coordinate: str, image_id: str) -> None:
    cell = sheet_root.find(f".//{{{MAIN_NS}}}c[@r='{coordinate}']")
    if cell is None:
        raise RuntimeError(f"目标单元格不存在：{coordinate}")
    for child in list(cell):
        if child.tag in {f"{{{MAIN_NS}}}f", f"{{{MAIN_NS}}}v", f"{{{MAIN_NS}}}is"}:
            cell.remove(child)
    cell.attrib.pop("t", None)
    formula = ET.SubElement(cell, f"{{{MAIN_NS}}}f")
    formula.text = f'_xlfn.DISPIMG("{image_id}",1)'


def main() -> None:
    args = parse_args()
    targets = [
        (5, "01_Bath_mats.png"),
        (7, "02_Carpets.png"),
        (9, "03_Indoor_Entryway_Mats.png"),
        (11, "04_Rug_Grip_Tape_&_Pads.png"),
        (13, "05_Rug_Pads.png"),
        (15, "06_Rug_Protectors.png"),
        (17, "07_Rug_Sets.png"),
        (19, "08_Rugs.png"),
        (21, "09_Rugs_&_Carpets.png"),
        (23, "10_Runners.png"),
        (25, "11_Stair_Pads.png"),
        (27, "12_Doormats.png"),
        (29, "13_Kitchen_Rugs_&_Mats.png"),
        (31, "14_Comfort_Mats.png"),
        (33, "15_Shower_Mats.png"),
        (35, "16_Bathtub_Mats.png"),
        (37, "17_Hard-floor_chair_mat.png"),
        (39, "18_Carpet_Chair_Mats.png"),
        (41, "19_Multi_surface_chair_mats.png"),
    ]

    missing = [name for _, name in targets if not (args.image_dir / name).is_file()]
    if missing:
        raise RuntimeError(f"缺少截图文件：{missing}")

    with zipfile.ZipFile(args.input, "r") as source:
        names = set(source.namelist())
        required_parts = {
            "xl/worksheets/sheet8.xml",
            "xl/cellimages.xml",
            "xl/_rels/cellimages.xml.rels",
        }
        if not required_parts.issubset(names):
            raise RuntimeError("工作簿缺少 US Sheet 或 WPS 单元格图片部件。")

        sheet_root = ET.fromstring(source.read("xl/worksheets/sheet8.xml"))
        cell_root = ET.fromstring(source.read("xl/cellimages.xml"))
        rel_root = ET.fromstring(source.read("xl/_rels/cellimages.xml.rels"))

        rel_numbers = [
            int(match)
            for match in re.findall(r"rId(\d+)", ET.tostring(rel_root, encoding="unicode"))
        ]
        picture_numbers = [
            int(value)
            for value in re.findall(r'id="(\d+)"', ET.tostring(cell_root, encoding="unicode"))
        ]
        next_rel_number = max(rel_numbers, default=0) + 1
        next_picture_id = max(picture_numbers, default=0) + 1
        new_media: dict[str, bytes] = {}
        formulas: dict[str, str] = {}

        for index, (row, image_name) in enumerate(targets, start=1):
            image_path = args.image_dir / image_name
            payload = image_path.read_bytes()
            with PILImage.open(io.BytesIO(payload)) as image:
                if image.size != (1896, 821):
                    raise RuntimeError(f"截图尺寸不符合示例：{image_name} {image.size}")

            image_id = f"ID_CODEX_US_CATEGORY_{index:04d}"
            rel_id = f"rId{next_rel_number}"
            media_name = f"xl/media/codex_us_category_{index:02d}.png"
            if media_name in names:
                raise RuntimeError(f"媒体文件名已存在：{media_name}")

            ET.SubElement(
                rel_root,
                f"{{{REL_NS}}}Relationship",
                {
                    "Id": rel_id,
                    "Type": IMAGE_REL_TYPE,
                    "Target": f"media/codex_us_category_{index:02d}.png",
                },
            )
            append_cell_image(cell_root, image_id, rel_id, payload, next_picture_id)
            coordinate = f"A{row}"
            set_dispimg_formula(sheet_root, coordinate, image_id)
            formulas[coordinate] = image_id
            new_media[media_name] = payload
            next_rel_number += 1
            next_picture_id += 1

        sheet_xml = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=True)
        cell_xml = ET.tostring(cell_root, encoding="utf-8", xml_declaration=True)
        rel_xml = ET.tostring(rel_root, encoding="utf-8", xml_declaration=True)

        args.output.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(args.output, "w", zipfile.ZIP_DEFLATED, allowZip64=True) as output:
            for item in source.infolist():
                if item.filename in {
                    "xl/worksheets/sheet8.xml",
                    "xl/cellimages.xml",
                    "xl/_rels/cellimages.xml.rels",
                }:
                    continue
                output.writestr(item, source.read(item.filename))
            output.writestr("xl/worksheets/sheet8.xml", sheet_xml)
            output.writestr("xl/cellimages.xml", cell_xml)
            output.writestr("xl/_rels/cellimages.xml.rels", rel_xml)
            for media_name, payload in new_media.items():
                output.writestr(media_name, payload)

    print(json.dumps({"output": str(args.output), "inserted": len(targets), "formulas": formulas}, ensure_ascii=False))


if __name__ == "__main__":
    main()
