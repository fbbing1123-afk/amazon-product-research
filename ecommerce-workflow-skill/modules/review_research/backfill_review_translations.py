from __future__ import annotations

import argparse
import json
import math
import os
import unicodedata
import zipfile
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree as ET


DEFAULT_CACHE_PATH = Path(__file__).resolve().parents[2] / "resources" / "review_translations.json"
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
XML_NS = "http://www.w3.org/XML/1998/namespace"
PAIN_POINT_SHEET = "差评痛点"
PAIN_FREQUENCY_SHEET = "差评词频分析"
SUMMARY_SHEET = "评论抓取摘要"


def tag(name: str) -> str:
    return f"{{{MAIN_NS}}}{name}"


def load_translation_mapping(path: Path) -> dict[str, str]:
    payload = json.loads(path.read_text(encoding="utf-8-sig"))
    if isinstance(payload, dict) and isinstance(payload.get("translations"), dict):
        payload = payload["translations"]
    if not isinstance(payload, dict):
        raise ValueError("翻译文件必须是 JSON 对象，或包含 translations 对象。")
    return {
        str(source).strip(): str(translation).strip()
        for source, translation in payload.items()
        if str(source).strip() and str(translation).strip()
    }


def update_cache(cache_path: Path, translations: dict[str, str]) -> None:
    existing: dict[str, object] = {}
    if cache_path.exists():
        payload = json.loads(cache_path.read_text(encoding="utf-8-sig"))
        if isinstance(payload, dict):
            existing = payload
    cached = existing.get("translations")
    if not isinstance(cached, dict):
        cached = {}
    cached.update(translations)
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    cache_path.write_text(
        json.dumps({"translations": cached}, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def read_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    return ["".join(item.itertext()) for item in root.findall(tag("si"))]


def cell_text(cell: ET.Element, shared_strings: list[str]) -> str:
    if cell.get("t") == "s":
        value = cell.findtext(tag("v"))
        if value and value.isdigit() and int(value) < len(shared_strings):
            return shared_strings[int(value)]
        return ""
    inline = cell.find(tag("is"))
    if inline is not None:
        return "".join(inline.itertext())
    return cell.findtext(tag("v")) or ""


def worksheet_path(archive: zipfile.ZipFile, sheet_name: str) -> str:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {
        relation.get("Id"): relation.get("Target", "")
        for relation in relationships.findall(f"{{{PACKAGE_REL_NS}}}Relationship")
    }
    for sheet in workbook.findall(f".//{tag('sheet')}"):
        if sheet.get("name") != sheet_name:
            continue
        target = targets.get(sheet.get(f"{{{REL_NS}}}id"), "")
        if target.startswith("/"):
            return target.lstrip("/")
        if target.startswith("xl/"):
            return target
        return str(PurePosixPath("xl") / target)
    raise RuntimeError(f"找不到工作表：{sheet_name}")


def set_inline_text(cell: ET.Element, value: str) -> None:
    for child in list(cell):
        if child.tag in {tag("v"), tag("is")}:
            cell.remove(child)
    cell.set("t", "inlineStr")
    inline = ET.SubElement(cell, tag("is"))
    text = ET.SubElement(inline, tag("t"))
    text.set(f"{{{XML_NS}}}space", "preserve")
    text.text = value


def estimated_row_height(value: str, column_width: int = 72) -> float:
    visual_lines = 0
    for line in value.splitlines() or [""]:
        units = sum(2 if unicodedata.east_asian_width(character) in {"W", "F"} else 1 for character in line)
        visual_lines += max(1, math.ceil(units / column_width))
    return max(54.0, min(409.0, visual_lines * 15.0 + 8.0))


def refresh_workbook(workbook: Path, translations: dict[str, str]) -> int:
    with zipfile.ZipFile(workbook, "r") as source:
        shared_strings = read_shared_strings(source)
        replacements: dict[str, bytes] = {}
        updated = 0

        sheet_path = worksheet_path(source, PAIN_POINT_SHEET)
        worksheet = ET.fromstring(source.read(sheet_path))
        content_column = ""
        header_row = next((row for row in worksheet.iter(tag("row")) if row.get("r") == "1"), None)
        if header_row is not None:
            for cell in header_row.findall(tag("c")):
                if cell_text(cell, shared_strings) == "评价内容":
                    content_column = "".join(character for character in cell.get("r", "") if character.isalpha())
                    break
        if not content_column:
            raise RuntimeError("差评痛点工作表缺少“评价内容”列。")
        for row in worksheet.iter(tag("row")):
            for cell in row.findall(tag("c")):
                coordinate = cell.get("r", "")
                column = "".join(character for character in coordinate if character.isalpha())
                row_number = "".join(character for character in coordinate if character.isdigit())
                if column != content_column or not row_number.isdigit() or row_number == "1":
                    continue
                value = cell_text(cell, shared_strings)
                translation = translations.get(value)
                if not translation:
                    continue
                display_value = f"{value}\n({translation})"
                set_inline_text(cell, display_value)
                current_height = float(row.get("ht", "0") or 0)
                row.set("ht", str(max(current_height, estimated_row_height(display_value))))
                row.set("customHeight", "1")
                updated += 1
        replacements[sheet_path] = ET.tostring(worksheet, encoding="utf-8", xml_declaration=True)

        try:
            frequency_path = worksheet_path(source, PAIN_FREQUENCY_SHEET)
            frequency_sheet = ET.fromstring(source.read(frequency_path))
            minimum_frequency_row = 2
        except RuntimeError:
            try:
                frequency_path = worksheet_path(source, SUMMARY_SHEET)
            except RuntimeError:
                frequency_path = ""
                frequency_sheet = None
                minimum_frequency_row = 0
            else:
                frequency_sheet = ET.fromstring(source.read(frequency_path))
                section_title_row = next(
                    (
                        int(row.get("r", "0") or 0)
                        for row in frequency_sheet.iter(tag("row"))
                        if any(
                            cell.get("r", "").startswith("A")
                            and cell_text(cell, shared_strings) == PAIN_FREQUENCY_SHEET
                            for cell in row.findall(tag("c"))
                        )
                    ),
                    0,
                )
                minimum_frequency_row = section_title_row + 2 if section_title_row else 0
        if frequency_path:
            frequency_header_row = minimum_frequency_row - 1
            header = next(
                (
                    row
                    for row in frequency_sheet.iter(tag("row"))
                    if int(row.get("r", "0") or 0) == frequency_header_row
                ),
                None,
            )
            frequency_columns: dict[str, str] = {}
            if header is not None:
                for cell in header.findall(tag("c")):
                    column = "".join(character for character in cell.get("r", "") if character.isalpha())
                    frequency_columns[cell_text(cell, shared_strings)] = column
            phrase_column = frequency_columns.get("关键词/短语", "A")
            translation_column = frequency_columns.get("中文翻译", "B")
            for row in frequency_sheet.iter(tag("row")):
                if int(row.get("r", "0") or 0) < minimum_frequency_row:
                    continue
                cells = row.findall(tag("c"))
                source_cell = next(
                    (
                        cell
                        for cell in cells
                        if "".join(character for character in cell.get("r", "") if character.isalpha())
                        == phrase_column
                    ),
                    None,
                )
                if source_cell is None:
                    continue
                phrase = cell_text(source_cell, shared_strings)
                translation = translations.get(phrase)
                if not translation:
                    continue
                row_number = "".join(character for character in source_cell.get("r", "") if character.isdigit())
                target_coordinate = f"{translation_column}{row_number}"
                target_cell = next((cell for cell in cells if cell.get("r") == target_coordinate), None)
                if target_cell is None:
                    target_cell = ET.Element(tag("c"), {"r": target_coordinate})
                    if source_cell.get("s"):
                        target_cell.set("s", source_cell.get("s", ""))
                    row.insert(list(row).index(source_cell) + 1, target_cell)
                set_inline_text(target_cell, translation)
                updated += 1
            replacements[frequency_path] = ET.tostring(
                frequency_sheet,
                encoding="utf-8",
                xml_declaration=True,
            )

        temporary = workbook.with_name(f"{workbook.stem}_review_translation_tmp.xlsx")
        with zipfile.ZipFile(temporary, "w", compression=zipfile.ZIP_DEFLATED) as destination:
            for info in source.infolist():
                data = replacements.get(info.filename, source.read(info.filename))
                destination.writestr(info, data)
    os.replace(temporary, workbook)
    return updated


def main() -> int:
    parser = argparse.ArgumentParser(
        description="把 Agent 完成的评论中文翻译回填到差评痛点页，并保留工作簿现有图片。"
    )
    parser.add_argument("--translations", type=Path, required=True)
    parser.add_argument("--workbook", type=Path, required=True)
    parser.add_argument("--cache", type=Path, default=DEFAULT_CACHE_PATH)
    args = parser.parse_args()

    if not args.workbook.exists():
        raise FileNotFoundError(args.workbook)
    translations = load_translation_mapping(args.translations)
    if not translations:
        raise ValueError("翻译文件中没有可回填的翻译。")
    updated = refresh_workbook(args.workbook, translations)
    update_cache(args.cache, translations)
    print(f"updated_cells={updated}")
    print(f"cache_updated={args.cache}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
