from __future__ import annotations

import hashlib
import io
import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from PIL import Image, ImageDraw, ImageFont, ImageStat


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
ETC_NS = "http://www.wps.cn/officeDocument/2017/etCustomData"
XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def sha256(payload: bytes) -> str:
    return hashlib.sha256(payload).hexdigest()


def main() -> None:
    project = Path(__file__).resolve().parents[4]
    source_path = project.parent / "英国站地毯地垫市场调研-261010-美国加拿大类目节点已完成.xlsx"
    output_path = project / "outputs/20261016-us-category-nodes/英国站地毯地垫市场调研-261010-美国加拿大日本类目节点20项已完成.xlsx"
    contact_path = project / "outputs/20261016-us-category-nodes/jp_category_20_contact_sheet.png"
    expected_coordinates = [f"A{row}" for row in range(3, 42, 2)]
    expected_media = [f"xl/media/codex_jp_category_{index:02d}.png" for index in range(1, 21)]

    with zipfile.ZipFile(source_path) as source, zipfile.ZipFile(output_path) as output:
        source_names = set(source.namelist())
        output_names = set(output.namelist())
        allowed_changed = {
            "xl/worksheets/sheet10.xml",
            "xl/cellimages.xml",
            "xl/_rels/cellimages.xml.rels",
        }
        unexpected_changed = sorted(
            name
            for name in source_names & output_names
            if name not in allowed_changed and sha256(source.read(name)) != sha256(output.read(name))
        )
        unexpected_added = sorted(output_names - source_names - set(expected_media))
        unexpected_removed = sorted(source_names - output_names)

        sheet_root = ET.fromstring(output.read("xl/worksheets/sheet10.xml"))
        cell_root = ET.fromstring(output.read("xl/cellimages.xml"))
        rel_root = ET.fromstring(output.read("xl/_rels/cellimages.xml.rels"))

        formulas = {}
        for coordinate in expected_coordinates:
            cell = sheet_root.find(f".//{{{MAIN_NS}}}c[@r='{coordinate}']")
            formula = cell.find(f"{{{MAIN_NS}}}f") if cell is not None else None
            formulas[coordinate] = formula.text if formula is not None else None

        rel_targets = {
            relation.attrib["Id"]: relation.attrib["Target"]
            for relation in rel_root.findall(f"{{{REL_NS}}}Relationship")
        }
        image_relations = {}
        for cell_image in cell_root.findall(f"{{{ETC_NS}}}cellImage"):
            named = cell_image.find(f".//{{{XDR_NS}}}cNvPr")
            blip = cell_image.find(f".//{{{A_NS}}}blip")
            if named is None or blip is None:
                continue
            image_id = named.attrib.get("name", "")
            if not image_id.startswith("ID_CODEX_JP_CATEGORY_"):
                continue
            rel_id = blip.attrib.get(f"{{{R_NS}}}embed", "")
            image_relations[image_id] = rel_targets.get(rel_id)

        image_checks = []
        images = []
        for index, media_name in enumerate(expected_media, start=1):
            payload = output.read(media_name)
            with Image.open(io.BytesIO(payload)) as image:
                image.load()
                gray = image.convert("L")
                minimum, maximum = gray.getextrema()
                deviation = ImageStat.Stat(gray).stddev[0]
                image_checks.append(
                    {
                        "index": index,
                        "size": image.size,
                        "contrast": maximum - minimum,
                        "stddev": round(deviation, 2),
                    }
                )
                images.append(image.convert("RGB"))

        # 生成缩略拼图，只用于人工确认截图内容完整。
        thumb_width = 570
        thumb_height = 277
        label_height = 28
        columns = 2
        rows = (len(images) + columns - 1) // columns
        canvas = Image.new("RGB", (columns * thumb_width, rows * (thumb_height + label_height)), "white")
        draw = ImageDraw.Draw(canvas)
        font = ImageFont.load_default()
        for index, image in enumerate(images):
            row, column = divmod(index, columns)
            thumb = image.resize((thumb_width, thumb_height))
            x = column * thumb_width
            y = row * (thumb_height + label_height)
            canvas.paste(thumb, (x, y + label_height))
            draw.text((x + 8, y + 7), f"JP category {index + 1:02d}", fill="black", font=font)
        canvas.save(contact_path)

        workbook_formula_errors = []
        error_pattern = re.compile(r"#REF!|#DIV/0!|#VALUE!|#NAME\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!")
        for name in sorted(output_names):
            if not name.startswith("xl/worksheets/") or not name.endswith(".xml"):
                continue
            text = output.read(name).decode("utf-8", errors="ignore")
            if error_pattern.search(text):
                workbook_formula_errors.append(name)

    expected_formulas = {
        coordinate: f'_xlfn.DISPIMG("ID_CODEX_JP_CATEGORY_{index:04d}",1)'
        for index, coordinate in enumerate(expected_coordinates, start=1)
    }
    expected_relations = {
        f"ID_CODEX_JP_CATEGORY_{index:04d}": f"media/codex_jp_category_{index:02d}.png"
        for index in range(1, 21)
    }
    result = {
        "output": str(output_path),
        "file_size": output_path.stat().st_size,
        "formula_count": len(formulas),
        "formulas_ok": formulas == expected_formulas,
        "relations_ok": image_relations == expected_relations,
        "images_ok": all(
            item["size"] == (1898, 921) and item["contrast"] > 20 and item["stddev"] > 5
            for item in image_checks
        ),
        "image_checks": image_checks,
        "unexpected_changed_parts": unexpected_changed,
        "unexpected_added_parts": unexpected_added,
        "unexpected_removed_parts": unexpected_removed,
        "formula_error_parts": workbook_formula_errors,
        "contact_sheet": str(contact_path),
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
