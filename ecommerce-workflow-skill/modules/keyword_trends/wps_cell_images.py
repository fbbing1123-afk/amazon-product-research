from __future__ import annotations

import io
import posixpath
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Iterable

from openpyxl import load_workbook
from PIL import Image as PILImage


ETC_NS = "http://www.wps.cn/officeDocument/2017/etCustomData"
XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
CELLIMAGE_REL_TYPE = "http://www.wps.cn/officeDocument/2020/cellImage"
CELLIMAGE_CONTENT_TYPE = "application/vnd.wps-officedocument.cellimage+xml"
EMU_PER_PIXEL = 9525

ET.register_namespace("etc", ETC_NS)
ET.register_namespace("xdr", XDR_NS)
ET.register_namespace("a", A_NS)
ET.register_namespace("r", R_NS)


def _image_cell(image) -> str | None:
    marker = getattr(getattr(image, "anchor", None), "_from", None)
    if marker is None:
        return None
    return f"{marker.row + 1}:{marker.col + 1}"


def _image_payload(image) -> bytes:
    raw = image._data()
    if raw.startswith(b"\x89PNG\r\n\x1a\n"):
        return raw
    picture = PILImage.open(io.BytesIO(raw))
    buffer = io.BytesIO()
    picture.save(buffer, format="PNG")
    return buffer.getvalue()


def convert_floating_images_to_wps_cell_images(
    floating_workbook: Path,
    output: Path,
    sheet_name: str,
    target_cells: Iterable[str],
    temp_output: Path | None = None,
    source_workbook: Path | None = None,
) -> dict:
    target_cells = set(target_cells)
    temp_output = temp_output or output.with_name(output.stem + "_floating_temp.xlsx")
    workbook = load_workbook(floating_workbook, data_only=False)
    if sheet_name not in workbook.sheetnames:
        raise ValueError(f"找不到工作表: {sheet_name}")
    worksheet = workbook[sheet_name]

    existing_image_ids = _read_existing_cellimage_ids(source_workbook or floating_workbook)
    extra_images: dict[str, bytes] = {}
    kept_images = []
    embedded = 0
    counter = 1
    for image in list(getattr(worksheet, "_images", [])):
        marker = getattr(getattr(image, "anchor", None), "_from", None)
        coord = worksheet.cell(marker.row + 1, marker.col + 1).coordinate if marker is not None else None
        if coord not in target_cells:
            kept_images.append(image)
            continue
        image_id = f"ID_CODEX_TREND_{counter:04d}"
        while image_id in existing_image_ids or image_id in extra_images:
            counter += 1
            image_id = f"ID_CODEX_TREND_{counter:04d}"
        worksheet[coord] = f'=_xlfn.DISPIMG("{image_id}",1)'
        extra_images[image_id] = _image_payload(image)
        counter += 1
        embedded += 1
    worksheet._images = kept_images

    temp_output.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(temp_output)
    _patch_package(temp_output, output, extra_images, source_workbook=source_workbook)
    return {
        "output": str(output),
        "sheet": sheet_name,
        "target_cells": sorted(target_cells),
        "embedded": embedded,
        "kept_floating_images": len(kept_images),
        "size": output.stat().st_size,
    }


def _read_existing_cellimage_ids(workbook_path: Path) -> set[str]:
    if not workbook_path.exists():
        return set()
    with zipfile.ZipFile(workbook_path, "r") as archive:
        if "xl/cellimages.xml" not in archive.namelist():
            return set()
        root = ET.fromstring(archive.read("xl/cellimages.xml"))
    ids: set[str] = set()
    for node in root.findall(f".//{{{XDR_NS}}}cNvPr"):
        image_id = node.attrib.get("name")
        if image_id:
            ids.add(image_id)
    return ids


def _patch_package(
    temp_workbook: Path,
    output: Path,
    extra_images: dict[str, bytes],
    source_workbook: Path | None = None,
) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(temp_workbook, "r") as zin:
        existing_names = set(zin.namelist())
        source_archive_path = source_workbook if source_workbook and source_workbook.exists() else temp_workbook
        with zipfile.ZipFile(source_archive_path, "r") as source_archive:
            cellimages_xml = _read_or_create_cellimages_xml(source_archive)
            cellimages_rels = _read_or_create_cellimages_rels(source_archive)

    cell_root = ET.fromstring(cellimages_xml)
    rel_root = ET.fromstring(cellimages_rels)
    with zipfile.ZipFile(source_archive_path, "r") as source_archive:
        inherited_payloads = _rewrite_existing_cellimage_payloads(source_archive, rel_root, existing_names)
    image_payloads: dict[str, bytes] = {}
    counter = 1
    next_rel_num = _next_rel_num(ET.tostring(rel_root, encoding="unicode"))
    next_pic_id = _next_pic_id(ET.tostring(cell_root, encoding="unicode"))
    for image_id, payload in extra_images.items():
        media_name = f"xl/media/codex_embedded_trend_{counter}.png"
        while media_name in existing_names or media_name in image_payloads:
            counter += 1
            media_name = f"xl/media/codex_embedded_trend_{counter}.png"
        rel_id = f"rId{next_rel_num}"
        next_rel_num += 1
        counter += 1
        image_payloads[media_name] = payload
        rel_root.append(
            ET.Element(
                f"{{{REL_NS}}}Relationship",
                {"Id": rel_id, "Type": IMAGE_REL_TYPE, "Target": posixpath.relpath(media_name, "xl")},
            )
        )
        _append_cellimage(cell_root, image_id, rel_id, payload, next_pic_id)
        next_pic_id += 1

    with zipfile.ZipFile(temp_workbook, "r") as zin, zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED, allowZip64=True) as zout:
        for item in zin.infolist():
            if item.filename in {"xl/cellimages.xml", "xl/_rels/cellimages.xml.rels"}:
                continue
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = _patch_content_types(data.decode("utf-8", "ignore")).encode("utf-8")
            elif item.filename == "xl/_rels/workbook.xml.rels":
                data = _patch_workbook_rels(data.decode("utf-8", "ignore")).encode("utf-8")
            zout.writestr(item, data)
        zout.writestr("xl/cellimages.xml", ET.tostring(cell_root, encoding="utf-8", xml_declaration=True))
        zout.writestr("xl/_rels/cellimages.xml.rels", ET.tostring(rel_root, encoding="utf-8", xml_declaration=True))
        for name, payload in {**inherited_payloads, **image_payloads}.items():
            zout.writestr(name, payload)


def _rewrite_existing_cellimage_payloads(
    archive: zipfile.ZipFile,
    rel_root: ET.Element,
    existing_names: set[str],
) -> dict[str, bytes]:
    if "xl/_rels/cellimages.xml.rels" not in archive.namelist():
        return {}
    payloads: dict[str, bytes] = {}
    counter = 1
    for rel in rel_root.findall(f"{{{REL_NS}}}Relationship"):
        if rel.attrib.get("Type") != IMAGE_REL_TYPE:
            continue
        target = rel.attrib.get("Target")
        if not target:
            continue
        source_name = target if target.startswith("xl/") else f"xl/{target}"
        if source_name not in archive.namelist():
            continue
        suffix = Path(source_name).suffix or ".png"
        media_name = f"xl/media/wps_cell_image_{counter}{suffix}"
        while media_name in existing_names or media_name in payloads:
            counter += 1
            media_name = f"xl/media/wps_cell_image_{counter}{suffix}"
        counter += 1
        payloads[media_name] = archive.read(source_name)
        rel.attrib["Target"] = posixpath.relpath(media_name, "xl")
    return payloads


def _read_or_create_cellimages_xml(archive: zipfile.ZipFile) -> bytes:
    if "xl/cellimages.xml" in archive.namelist():
        return archive.read("xl/cellimages.xml")
    root = ET.Element(f"{{{ETC_NS}}}cellImages")
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _read_or_create_cellimages_rels(archive: zipfile.ZipFile) -> bytes:
    if "xl/_rels/cellimages.xml.rels" in archive.namelist():
        return archive.read("xl/_rels/cellimages.xml.rels")
    root = ET.Element(f"{{{REL_NS}}}Relationships")
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _patch_content_types(xml_text: str) -> str:
    if 'PartName="/xl/cellimages.xml"' in xml_text:
        return xml_text
    marker = f'<Override PartName="/xl/cellimages.xml" ContentType="{CELLIMAGE_CONTENT_TYPE}"/>'
    return xml_text.replace("</Types>", marker + "</Types>")


def _patch_workbook_rels(xml_text: str) -> str:
    if 'Target="cellimages.xml"' in xml_text or "Target='cellimages.xml'" in xml_text:
        return xml_text
    rel_id = f"rId{_next_rel_num(xml_text)}"
    rel = f'<Relationship Id="{rel_id}" Type="{CELLIMAGE_REL_TYPE}" Target="cellimages.xml"/>'
    return xml_text.replace("</Relationships>", rel + "</Relationships>")


def _next_rel_num(xml_text: str) -> int:
    ids = [int(match) for match in re.findall(r'Id="rId(\d+)"', xml_text)]
    return max(ids, default=0) + 1


def _next_pic_id(xml_text: str) -> int:
    ids = [int(match) for match in re.findall(r'id="(\d+)"', xml_text)]
    return max(ids, default=0) + 1


def _append_cellimage(cell_root: ET.Element, image_id: str, rel_id: str, payload: bytes, pic_id: int) -> None:
    cell_image = ET.SubElement(cell_root, f"{{{ETC_NS}}}cellImage")
    pic = ET.SubElement(cell_image, f"{{{XDR_NS}}}pic")
    nv = ET.SubElement(pic, f"{{{XDR_NS}}}nvPicPr")
    ET.SubElement(nv, f"{{{XDR_NS}}}cNvPr", {"id": str(pic_id), "name": image_id, "descr": "Picture"})
    ET.SubElement(nv, f"{{{XDR_NS}}}cNvPicPr")
    blip_fill = ET.SubElement(pic, f"{{{XDR_NS}}}blipFill")
    ET.SubElement(blip_fill, f"{{{A_NS}}}blip", {f"{{{R_NS}}}embed": rel_id})
    stretch = ET.SubElement(blip_fill, f"{{{A_NS}}}stretch")
    ET.SubElement(stretch, f"{{{A_NS}}}fillRect")
    sppr = ET.SubElement(pic, f"{{{XDR_NS}}}spPr")
    xfrm = ET.SubElement(sppr, f"{{{A_NS}}}xfrm")
    ET.SubElement(xfrm, f"{{{A_NS}}}off", {"x": "0", "y": "0"})
    try:
        with PILImage.open(io.BytesIO(payload)) as picture:
            width = picture.width
            height = picture.height
    except Exception:
        width, height = _png_size(payload)
    width_emu = str(int(width * EMU_PER_PIXEL))
    height_emu = str(int(height * EMU_PER_PIXEL))
    ET.SubElement(xfrm, f"{{{A_NS}}}ext", {"cx": width_emu, "cy": height_emu})
    geom = ET.SubElement(sppr, f"{{{A_NS}}}prstGeom", {"prst": "rect"})
    ET.SubElement(geom, f"{{{A_NS}}}avLst")


def _png_size(data: bytes) -> tuple[int, int]:
    if data.startswith(b"\x89PNG\r\n\x1a\n") and len(data) >= 24:
        return int.from_bytes(data[16:20], "big"), int.from_bytes(data[20:24], "big")
    return 200, 200
