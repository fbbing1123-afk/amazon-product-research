from __future__ import annotations

import io
import posixpath
import re
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from PIL import Image as PILImage


NS = {
    "etc": "http://www.wps.cn/officeDocument/2017/etCustomData",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
DRAWING_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
WORKSHEET_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
CELLIMAGE_REL_TYPE = "http://www.wps.cn/officeDocument/2020/cellImage"
CELLIMAGE_CONTENT_TYPE = "application/vnd.wps-officedocument.cellimage+xml"
ETC_NS = NS["etc"]
XDR_NS = NS["xdr"]
A_NS = NS["a"]
R_NS = NS["r"]
EMU_PER_PIXEL = 9525

ET.register_namespace("etc", ETC_NS)
ET.register_namespace("xdr", XDR_NS)
ET.register_namespace("a", A_NS)
ET.register_namespace("r", R_NS)


@dataclass(frozen=True)
class CellImageInfo:
    media_path: str
    width: int | None = None
    height: int | None = None


def _package_path(base_path: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(base_path), target))


def _relationship_map(archive: zipfile.ZipFile, rels_path: str) -> dict[str, tuple[str, str]]:
    if rels_path not in archive.namelist():
        return {}
    root = ET.fromstring(archive.read(rels_path))
    return {
        rel.attrib.get("Id", ""): (rel.attrib.get("Type", ""), rel.attrib.get("Target", ""))
        for rel in root.findall(f"{{{REL_NS}}}Relationship")
        if rel.attrib.get("Id") and rel.attrib.get("Target")
    }


def restore_unloaded_floating_images(workbook, source: Path) -> list[XLImage]:
    """Recover OOXML pictures that openpyxl did not deserialize from the source file."""
    added: list[XLImage] = []
    with zipfile.ZipFile(source) as archive:
        names = set(archive.namelist())
        workbook_rels = _relationship_map(archive, "xl/_rels/workbook.xml.rels")
        workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
        sheet_paths: dict[str, str] = {}
        for sheet in workbook_root.findall(f".//{{{SHEET_NS}}}sheet"):
            rel_id = sheet.attrib.get(f"{{{OFFICE_REL_NS}}}id", "")
            rel_type, target = workbook_rels.get(rel_id, ("", ""))
            if rel_type == WORKSHEET_REL_TYPE and target:
                sheet_paths[sheet.attrib.get("name", "")] = _package_path("xl/workbook.xml", target)

        for sheet_name, sheet_path in sheet_paths.items():
            if sheet_name not in workbook.sheetnames or sheet_path not in names:
                continue
            worksheet = workbook[sheet_name]
            existing = Counter()
            for image in getattr(worksheet, "_images", []):
                marker = getattr(getattr(image, "anchor", None), "_from", None)
                if marker is not None:
                    existing[(marker.row, marker.col)] += 1

            rels_path = posixpath.join(
                posixpath.dirname(sheet_path),
                "_rels",
                posixpath.basename(sheet_path) + ".rels",
            )
            sheet_rels = _relationship_map(archive, rels_path)
            sheet_root = ET.fromstring(archive.read(sheet_path))
            for drawing_node in sheet_root.findall(f"{{{SHEET_NS}}}drawing"):
                rel_id = drawing_node.attrib.get(f"{{{OFFICE_REL_NS}}}id", "")
                rel_type, target = sheet_rels.get(rel_id, ("", ""))
                if rel_type != DRAWING_REL_TYPE or not target:
                    continue
                drawing_path = _package_path(sheet_path, target)
                if drawing_path not in names:
                    continue
                drawing_rels_path = posixpath.join(
                    posixpath.dirname(drawing_path),
                    "_rels",
                    posixpath.basename(drawing_path) + ".rels",
                )
                drawing_rels = _relationship_map(archive, drawing_rels_path)
                drawing_root = ET.fromstring(archive.read(drawing_path))
                anchors = list(drawing_root.findall(f"{{{XDR_NS}}}oneCellAnchor"))
                anchors.extend(drawing_root.findall(f"{{{XDR_NS}}}twoCellAnchor"))
                for anchor in anchors:
                    marker = anchor.find(f"{{{XDR_NS}}}from")
                    blip = anchor.find(f".//{{{A_NS}}}blip")
                    if marker is None or blip is None:
                        continue
                    try:
                        row = int(marker.findtext(f"{{{XDR_NS}}}row", "0"))
                        column = int(marker.findtext(f"{{{XDR_NS}}}col", "0"))
                    except ValueError:
                        continue
                    anchor_key = (row, column)
                    if existing[anchor_key]:
                        existing[anchor_key] -= 1
                        continue
                    image_rel_id = blip.attrib.get(f"{{{R_NS}}}embed", "")
                    image_rel_type, image_target = drawing_rels.get(image_rel_id, ("", ""))
                    image_path = _package_path(drawing_path, image_target) if image_target else ""
                    if image_rel_type != IMAGE_REL_TYPE or image_path not in names:
                        continue
                    image = XLImage(io.BytesIO(archive.read(image_path)))
                    extent = anchor.find(f"{{{XDR_NS}}}ext")
                    if extent is not None:
                        image.width = emu_to_pixels(extent.attrib.get("cx")) or image.width
                        image.height = emu_to_pixels(extent.attrib.get("cy")) or image.height
                    worksheet.add_image(image, f"{get_column_letter(column + 1)}{row + 1}")
                    added.append(image)
    return added


def emu_to_pixels(value: str | None) -> int | None:
    if not value:
        return None
    try:
        return max(1, round(int(value) / EMU_PER_PIXEL))
    except ValueError:
        return None


def read_cellimage_map(xlsx_path: Path) -> tuple[dict[str, bytes], dict[str, CellImageInfo]]:
    with zipfile.ZipFile(xlsx_path) as archive:
        names = set(archive.namelist())
        if "xl/cellimages.xml" not in names or "xl/_rels/cellimages.xml.rels" not in names:
            return {}, {}

        rel_root = ET.fromstring(archive.read("xl/_rels/cellimages.xml.rels"))
        rid_to_target: dict[str, str] = {}
        for rel in rel_root.findall(f"{{{REL_NS}}}Relationship"):
            rid = rel.attrib.get("Id")
            target = rel.attrib.get("Target")
            if rid and target:
                rid_to_target[rid] = target

        cell_root = ET.fromstring(archive.read("xl/cellimages.xml"))
        id_to_info: dict[str, CellImageInfo] = {}
        for cell_image in cell_root.findall("etc:cellImage", NS):
            name_node = cell_image.find(".//xdr:cNvPr", NS)
            blip_node = cell_image.find(".//a:blip", NS)
            ext_node = cell_image.find(".//a:ext", NS)
            if name_node is None or blip_node is None:
                continue
            image_id = name_node.attrib.get("name")
            rel_id = blip_node.attrib.get(f"{{{R_NS}}}embed")
            target = rid_to_target.get(rel_id or "")
            if not image_id or not target:
                continue
            media_path = target if target.startswith("xl/") else f"xl/{target}"
            if media_path in names:
                width = emu_to_pixels(ext_node.attrib.get("cx")) if ext_node is not None else None
                height = emu_to_pixels(ext_node.attrib.get("cy")) if ext_node is not None else None
                id_to_info[image_id] = CellImageInfo(media_path=media_path, width=width, height=height)

        media_bytes = {
            media_path: archive.read(media_path)
            for media_path in {info.media_path for info in id_to_info.values()}
            if media_path in names
        }
        return media_bytes, id_to_info


def normalize_image(data: bytes, width: int, height: int) -> XLImage:
    picture = PILImage.open(io.BytesIO(data))
    if picture.mode not in ("RGB", "RGBA"):
        picture = picture.convert("RGBA")
    buffer = io.BytesIO()
    picture.save(buffer, format="PNG")
    buffer.seek(0)
    image = XLImage(buffer)
    image.width = width
    image.height = height
    return image


def convert_wps_dispimg_to_floating(
    source: Path,
    output: Path,
    sheet_name: str | None = None,
    width: int = 200,
    height: int = 200,
) -> dict:
    media_bytes, id_to_info = read_cellimage_map(source)
    if not id_to_info:
        raise RuntimeError("No WPS cell image mapping found in xl/cellimages.xml.")

    workbook = load_workbook(source, data_only=False, rich_text=True)
    worksheets = [workbook[sheet_name]] if sheet_name else list(workbook.worksheets)

    converted = 0
    per_sheet: dict[str, int] = {}
    missing: list[tuple[str, str, str]] = []
    for worksheet in worksheets:
        for row in worksheet.iter_rows():
            for cell in row:
                value = cell.value
                if not isinstance(value, str):
                    continue
                match = re.search(r'DISPIMG\("([^"]+)"', value)
                if not match:
                    continue
                image_id = match.group(1)
                info = id_to_info.get(image_id)
                data = media_bytes.get(info.media_path) if info else None
                if not data:
                    missing.append((worksheet.title, cell.coordinate, image_id))
                    continue
                worksheet.add_image(normalize_image(data, info.width or width, info.height or height), cell.coordinate)
                cell.value = None
                converted += 1
                per_sheet[worksheet.title] = per_sheet.get(worksheet.title, 0) + 1

    output.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output)
    check = load_workbook(output, data_only=False)
    return {
        "output": str(output),
        "converted": converted,
        "converted_by_sheet": per_sheet,
        "missing": len(missing),
        "missing_samples": missing[:5],
        "floating_images_after": {ws.title: len(getattr(ws, "_images", [])) for ws in check.worksheets},
    }


def collect_dispimg_formulas(path: Path) -> dict[str, dict[str, str]]:
    workbook = load_workbook(path, data_only=False)
    formulas: dict[str, dict[str, str]] = {}
    for worksheet in workbook.worksheets:
        sheet_formulas = {
            cell.coordinate: cell.value
            for row in worksheet.iter_rows()
            for cell in row
            if isinstance(cell.value, str) and "DISPIMG" in cell.value.upper()
        }
        if sheet_formulas:
            formulas[worksheet.title] = sheet_formulas
    return formulas


def restore_dispimg_cells(
    original: Path,
    floating_workbook: Path,
    temp_output: Path,
    embed_extra_sheets: set[str],
) -> tuple[dict[str, dict[str, int]], dict[str, bytes]]:
    formulas = collect_dispimg_formulas(original)
    # 保留竞品对比维度和产品卖点中的富文本红色标记，避免 WPS 图片合并阶段丢失。
    workbook = load_workbook(floating_workbook, data_only=False, rich_text=True)
    summary: dict[str, dict[str, int]] = {}
    extra_images: dict[str, bytes] = {}
    existing_image_ids: set[str] = set()
    # WPS 按 cNvPr.name 解析 DISPIMG；新增图片必须避开原工作簿已有编号。
    if zipfile.is_zipfile(original):
        with zipfile.ZipFile(original) as archive:
            if "xl/cellimages.xml" in archive.namelist():
                cell_root = ET.fromstring(archive.read("xl/cellimages.xml"))
                existing_image_ids = {
                    str(node.attrib.get("name", ""))
                    for node in cell_root.findall(".//xdr:cNvPr", NS)
                    if node.attrib.get("name")
                }
    extra_counter = 1

    for sheet_name, sheet_formulas in formulas.items():
        if sheet_name not in workbook.sheetnames:
            continue
        worksheet = workbook[sheet_name]
        target_coords = set(sheet_formulas)
        kept_images = []
        removed = 0
        for image in list(getattr(worksheet, "_images", [])):
            marker = getattr(getattr(image, "anchor", None), "_from", None)
            if marker is not None:
                coord = worksheet.cell(marker.row + 1, marker.col + 1).coordinate
                if coord in target_coords:
                    removed += 1
                    continue
            kept_images.append(image)
        worksheet._images = kept_images
        for coord, formula in sheet_formulas.items():
            worksheet[coord] = formula
        summary[sheet_name] = {
            "restored_formulas": len(sheet_formulas),
            "removed_floating_images": removed,
            "kept_floating_images": len(kept_images),
        }

    for sheet_name in embed_extra_sheets:
        if sheet_name not in workbook.sheetnames:
            continue
        worksheet = workbook[sheet_name]
        summary.setdefault(sheet_name, {"restored_formulas": 0, "removed_floating_images": 0, "kept_floating_images": 0})
        kept_images = []
        embedded = 0
        for image in list(getattr(worksheet, "_images", [])):
            marker = getattr(getattr(image, "anchor", None), "_from", None)
            if marker is None:
                kept_images.append(image)
                continue
            image_id = f"ID_CODEX_EMBED_{extra_counter:04d}"
            while image_id in existing_image_ids:
                extra_counter += 1
                image_id = f"ID_CODEX_EMBED_{extra_counter:04d}"
            worksheet.cell(marker.row + 1, marker.col + 1).value = f'=_xlfn.DISPIMG("{image_id}",1)'
            raw = image._data()
            picture = PILImage.open(io.BytesIO(raw))
            buffer = io.BytesIO()
            picture.save(buffer, format="PNG")
            extra_images[image_id] = buffer.getvalue()
            existing_image_ids.add(image_id)
            extra_counter += 1
            embedded += 1
        worksheet._images = kept_images
        summary[sheet_name]["extra_embedded_images"] = embedded
        summary[sheet_name]["kept_floating_images"] = len(kept_images)

    temp_output.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(temp_output)
    return summary, extra_images


def patch_content_types(xml_text: str) -> str:
    marker = f'<Override PartName="/xl/cellimages.xml" ContentType="{CELLIMAGE_CONTENT_TYPE}"/>'
    if 'PartName="/xl/cellimages.xml"' in xml_text:
        return xml_text
    return xml_text.replace("</Types>", marker + "</Types>")


def next_workbook_rel_id(rels_text: str) -> str:
    ids = [int(match) for match in re.findall(r'Id="rId(\d+)"', rels_text)]
    return f"rId{max(ids, default=0) + 1}"


def patch_workbook_rels(xml_text: str) -> str:
    if 'Target="cellimages.xml"' in xml_text or "Target='cellimages.xml'" in xml_text:
        return xml_text
    rel_id = next_workbook_rel_id(xml_text)
    rel = f'<Relationship Id="{rel_id}" Type="{CELLIMAGE_REL_TYPE}" Target="cellimages.xml"/>'
    return xml_text.replace("</Relationships>", rel + "</Relationships>")


def rewrite_cellimage_rels(original: Path, existing_names: set[str], extra_images: dict[str, bytes]) -> tuple[bytes, dict[str, bytes], bytes]:
    with zipfile.ZipFile(original, "r") as archive:
        names = set(archive.namelist())
        if "xl/cellimages.xml" in names and "xl/_rels/cellimages.xml.rels" in names:
            cell_root = ET.fromstring(archive.read("xl/cellimages.xml"))
            root = ET.fromstring(archive.read("xl/_rels/cellimages.xml.rels"))
        else:
            cell_root = ET.Element(f"{{{ETC_NS}}}cellImages")
            root = ET.Element(f"{{{REL_NS}}}Relationships")
        image_payloads: dict[str, bytes] = {}
        counter = 1

        for rel in root.findall(f"{{{REL_NS}}}Relationship"):
            if rel.attrib.get("Type") != IMAGE_REL_TYPE:
                continue
            target = rel.attrib.get("Target")
            if not target:
                continue
            source_name = target if target.startswith("xl/") else f"xl/{target}"
            if source_name not in archive.namelist():
                continue
            suffix = Path(source_name).suffix or ".png"
            new_name = f"xl/media/wps_cell_image_{counter}{suffix}"
            while new_name in existing_names or new_name in image_payloads:
                counter += 1
                new_name = f"xl/media/wps_cell_image_{counter}{suffix}"
            counter += 1
            image_payloads[new_name] = archive.read(source_name)
            rel.attrib["Target"] = posixpath.relpath(new_name, "xl")

        next_rel_num = max([int(match) for match in re.findall(r"rId(\d+)", ET.tostring(root, encoding="unicode"))], default=0) + 1
        next_pic_id = max([int(value) for value in re.findall(r'id="(\d+)"', ET.tostring(cell_root, encoding="unicode"))], default=0) + 1
        for image_id, payload in extra_images.items():
            new_name = f"xl/media/codex_embedded_chart_{counter}.png"
            while new_name in existing_names or new_name in image_payloads:
                counter += 1
                new_name = f"xl/media/codex_embedded_chart_{counter}.png"
            rel_id = f"rId{next_rel_num}"
            next_rel_num += 1
            counter += 1
            image_payloads[new_name] = payload
            rel = ET.Element(f"{{{REL_NS}}}Relationship", {"Id": rel_id, "Type": IMAGE_REL_TYPE, "Target": posixpath.relpath(new_name, "xl")})
            root.append(rel)
            append_cellimage(cell_root, image_id, rel_id, payload, next_pic_id)
            next_pic_id += 1

        return (
            ET.tostring(root, encoding="utf-8", xml_declaration=True),
            image_payloads,
            ET.tostring(cell_root, encoding="utf-8", xml_declaration=True),
        )


def append_cellimage(cell_root: ET.Element, image_id: str, rel_id: str, payload: bytes, pic_id: int) -> None:
    cell_image = ET.SubElement(cell_root, f"{{{ETC_NS}}}cellImage")
    pic = ET.SubElement(cell_image, f"{{{XDR_NS}}}pic")
    nv = ET.SubElement(pic, f"{{{XDR_NS}}}nvPicPr")
    cnv = ET.SubElement(nv, f"{{{XDR_NS}}}cNvPr", {"id": str(pic_id), "name": image_id, "descr": "Picture"})
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
            width_emu = str(int(picture.width * EMU_PER_PIXEL))
            height_emu = str(int(picture.height * EMU_PER_PIXEL))
    except Exception:
        width_emu = str(200 * EMU_PER_PIXEL)
        height_emu = str(200 * EMU_PER_PIXEL)
    ET.SubElement(xfrm, f"{{{A_NS}}}ext", {"cx": width_emu, "cy": height_emu})
    geom = ET.SubElement(sppr, f"{{{A_NS}}}prstGeom", {"prst": "rect"})
    ET.SubElement(geom, f"{{{A_NS}}}avLst")


def patch_package(original: Path, temp_workbook: Path, output: Path, extra_images: dict[str, bytes]) -> None:
    with zipfile.ZipFile(temp_workbook, "r") as temp_zip:
        existing_names = set(temp_zip.namelist())
    cellimages_rels, image_payloads, cellimages_xml = rewrite_cellimage_rels(original, existing_names, extra_images)
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(temp_workbook, "r") as zin, zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED, allowZip64=True) as zout:
        for item in zin.infolist():
            if item.filename in {"xl/cellimages.xml", "xl/_rels/cellimages.xml.rels"}:
                continue
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = patch_content_types(data.decode("utf-8", "ignore")).encode("utf-8")
            elif item.filename == "xl/_rels/workbook.xml.rels":
                data = patch_workbook_rels(data.decode("utf-8", "ignore")).encode("utf-8")
            zout.writestr(item, data)
        zout.writestr("xl/cellimages.xml", cellimages_xml)
        zout.writestr("xl/_rels/cellimages.xml.rels", cellimages_rels)
        for name, payload in image_payloads.items():
            zout.writestr(name, payload)


def make_wps_embedded_version(original: Path, floating_workbook: Path, output: Path, temp_output: Path, embed_extra_sheets: set[str]) -> dict:
    restore_summary, extra_images = restore_dispimg_cells(original, floating_workbook, temp_output, embed_extra_sheets)
    patch_package(original, temp_output, output, extra_images)
    output_summary = inspect_image_state(output)
    return {
        "output": str(output),
        "restore_summary": restore_summary,
        "output_summary": output_summary,
        "size": output.stat().st_size,
    }


def inspect_image_state(path: Path) -> dict[str, dict[str, int]]:
    workbook = load_workbook(path, data_only=False)
    summary: dict[str, dict[str, int]] = {}
    for worksheet in workbook.worksheets:
        dispimg_count = sum(
            1
            for row in worksheet.iter_rows()
            for cell in row
            if isinstance(cell.value, str) and "DISPIMG" in cell.value.upper()
        )
        summary[worksheet.title] = {
            "floating_images": len(getattr(worksheet, "_images", [])),
            "charts": len(getattr(worksheet, "_charts", [])),
            "dispimg_formulas": dispimg_count,
        }
    return summary
