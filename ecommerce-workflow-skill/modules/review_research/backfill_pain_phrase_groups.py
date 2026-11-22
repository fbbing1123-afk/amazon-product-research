from __future__ import annotations

import argparse
import json
import os
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


MODULES_DIR = Path(__file__).resolve().parents[1]
if str(MODULES_DIR) not in sys.path:
    sys.path.insert(0, str(MODULES_DIR))

from review_research.backfill_review_translations import (
    cell_text,
    read_shared_strings,
    set_inline_text,
    tag,
    worksheet_path,
)
from review_research.pain_phrase_analysis import (
    SemanticGroup,
    aggregate_semantic_groups,
    normalize_semantic_groups,
    semantic_group_for_phrase,
)


DEFAULT_CACHE_PATH = Path(__file__).resolve().parents[2] / "resources" / "pain_phrase_groups.json"
SUMMARY_SHEET = "评论抓取摘要"
PAIN_POINT_SHEET = "差评痛点"
PAIN_FREQUENCY_SHEET = "差评词频分析"
CELL_COLUMN_PATTERN = re.compile(r"([A-Z]+)")


def load_group_mapping(path: Path) -> dict[str, SemanticGroup]:
    payload = json.loads(path.read_text(encoding="utf-8-sig"))
    if isinstance(payload, dict) and isinstance(payload.get("groups"), dict):
        payload = payload["groups"]
    if not isinstance(payload, dict):
        raise ValueError("分组文件必须是 JSON 对象，或包含 groups 对象。")
    groups: dict[str, SemanticGroup] = {}
    for phrase, value in payload.items():
        if not isinstance(value, dict):
            continue
        english = str(value.get("group_en") or "").strip()
        chinese = str(value.get("group_zh") or "").strip()
        if str(phrase).strip() and english and chinese:
            groups[str(phrase).strip()] = SemanticGroup(english, chinese)
    return groups


def update_cache(cache_path: Path, groups: dict[str, SemanticGroup]) -> None:
    existing: dict[str, object] = {}
    if cache_path.exists():
        payload = json.loads(cache_path.read_text(encoding="utf-8-sig"))
        if isinstance(payload, dict):
            existing = payload
    cached = existing.get("groups")
    if not isinstance(cached, dict):
        cached = {}
    for phrase, group in groups.items():
        cached[phrase] = {"group_en": group.english, "group_zh": group.chinese}
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    cache_path.write_text(
        json.dumps({"groups": cached}, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def row_values(row: ET.Element, shared_strings: list[str]) -> dict[str, str]:
    values: dict[str, str] = {}
    for cell in row.findall(tag("c")):
        match = CELL_COLUMN_PATTERN.match(cell.get("r", ""))
        if match:
            values[match.group(1)] = cell_text(cell, shared_strings)
    return values


def replace_pain_point_groups(
    archive: zipfile.ZipFile,
    shared_strings: list[str],
    groups: dict[str, SemanticGroup],
) -> tuple[str, bytes, list[dict[str, str]], list[list[str]]]:
    sheet_path = worksheet_path(archive, PAIN_POINT_SHEET)
    worksheet = ET.fromstring(archive.read(sheet_path))
    sheet_data = worksheet.find(tag("sheetData"))
    if sheet_data is None:
        raise RuntimeError("差评痛点工作表缺少 sheetData。")
    normalized_groups = normalize_semantic_groups(groups)
    reviews: list[dict[str, str]] = []
    terms_by_review: list[list[str]] = []
    columns: dict[str, str] = {}
    for row in sheet_data.findall(tag("row")):
        row_number = int(row.get("r", "0") or 0)
        values = row_values(row, shared_strings)
        if row_number == 1:
            columns = {header: column for column, header in values.items()}
            required = {
                "竞品ASIN", "实际筛选", "属性/型号",
                "差评词/短语", "归一化痛点（中文）",
            }
            if missing := required - columns.keys():
                raise RuntimeError(f"差评痛点缺少列：{', '.join(sorted(missing))}")
            continue
        if row_number < 2:
            continue
        phrase_column = columns["差评词/短语"]
        variant_column = columns["属性/型号"]
        asin_column = columns["竞品ASIN"]
        source_asin_column = columns.get("评论来源ASIN")
        filter_column = columns["实际筛选"]
        normalized_column = columns["归一化痛点（中文）"]
        phrases = [
            phrase.strip()
            for phrase in values.get(phrase_column, "").splitlines()
            if phrase.strip() and phrase.strip() != "/"
        ]
        reviews.append(
            {
                "asin": values.get(asin_column, ""),
                "source_asin": values.get(source_asin_column, "") if source_asin_column else values.get(asin_column, ""),
                "actual_filter": values.get(filter_column, ""),
                "variant": "" if values.get(variant_column) == "/" else values.get(variant_column, ""),
                "review_text": values.get("评价内容", ""),
            }
        )
        terms_by_review.append(phrases)
        normalized = []
        for phrase in phrases:
            group = semantic_group_for_phrase(
                phrase,
                normalized_groups,
                review_text=values.get("评价内容", ""),
            )
            normalized.append(group.chinese if group else "待归类")
        cells = row.findall(tag("c"))
        coordinate = f"{normalized_column}{row_number}"
        target = next((cell for cell in cells if cell.get("r") == coordinate), None)
        if target is None:
            target = ET.Element(tag("c"), {"r": coordinate})
            source_cell = next(
                (cell for cell in cells if cell.get("r") == f"{phrase_column}{row_number}"),
                None,
            )
            if source_cell is not None and source_cell.get("s"):
                target.set("s", source_cell.get("s", ""))
            insert_at = list(row).index(source_cell) + 1 if source_cell is not None else len(row)
            row.insert(insert_at, target)
        set_inline_text(target, "\n".join(normalized) or "/")
    return (
        sheet_path,
        ET.tostring(worksheet, encoding="utf-8", xml_declaration=True),
        reviews,
        terms_by_review,
    )


def append_cell(row: ET.Element, coordinate: str, value: object, style: str | None = None) -> None:
    cell = ET.SubElement(row, tag("c"), {"r": coordinate})
    if style:
        cell.set("s", style)
    if isinstance(value, (int, float)):
        number = ET.SubElement(cell, tag("v"))
        number.text = str(value)
    else:
        set_inline_text(cell, str(value))


def replace_frequency_rows(
    archive: zipfile.ZipFile,
    shared_strings: list[str],
    rows,
) -> tuple[str, bytes]:
    sheet_path = worksheet_path(archive, SUMMARY_SHEET)
    worksheet = ET.fromstring(archive.read(sheet_path))
    sheet_data = worksheet.find(tag("sheetData"))
    if sheet_data is None:
        raise RuntimeError("评论抓取摘要工作表缺少 sheetData。")

    section_title_row = 0
    for existing_row in sheet_data.findall(tag("row")):
        values = row_values(existing_row, shared_strings)
        if values.get("A") == PAIN_FREQUENCY_SHEET:
            section_title_row = int(existing_row.get("r", "0") or 0)
            break
    if not section_title_row:
        raise RuntimeError("评论抓取摘要中缺少“差评词频分析”区块。")
    frequency_header_row = section_title_row + 1
    header_row = next(
        (
            row
            for row in sheet_data.findall(tag("row"))
            if int(row.get("r", "0") or 0) == frequency_header_row
        ),
        None,
    )
    if header_row is None:
        raise RuntimeError("差评词频分析缺少表头。")
    header_columns = {
        header: column for column, header in row_values(header_row, shared_strings).items()
    }
    required_headers = {
        "竞品ASIN", "实际筛选", "关键词/短语",
        "中文翻译", "评论覆盖数", "覆盖率", "涉及型号",
    }
    if missing := required_headers - header_columns.keys():
        raise RuntimeError(f"差评词频分析缺少列：{', '.join(sorted(missing))}")

    style_by_column: dict[str, str] = {}
    for old_row in sheet_data.findall(tag("row")):
        if int(old_row.get("r", "0") or 0) <= frequency_header_row:
            continue
        for cell in old_row.findall(tag("c")):
            match = CELL_COLUMN_PATTERN.match(cell.get("r", ""))
            if match and cell.get("s"):
                style_by_column.setdefault(match.group(1), cell.get("s", ""))
        sheet_data.remove(old_row)

    for row_number, (asin, source_asin, actual_filter, item) in enumerate(
        rows,
        start=frequency_header_row + 1,
    ):
        row = ET.SubElement(sheet_data, tag("row"), {"r": str(row_number)})
        values = {
            "竞品ASIN": asin,
            "实际筛选": actual_filter,
            "关键词/短语": item.phrase,
            "中文翻译": item.translation,
            "评论覆盖数": item.review_count,
            "覆盖率": item.review_ratio,
            "涉及型号": "\n".join(item.variants) or "/",
        }
        if "评论来源ASIN" in header_columns:
            values["评论来源ASIN"] = source_asin
        for header, value in values.items():
            column = header_columns[header]
            append_cell(row, f"{column}{row_number}", value, style_by_column.get(column))

    last_row = frequency_header_row + len(rows)
    dimension = worksheet.find(tag("dimension"))
    if dimension is not None:
        dimension.set("ref", f"A1:O{last_row}")
    return sheet_path, ET.tostring(worksheet, encoding="utf-8", xml_declaration=True)


def refresh_workbook(workbook: Path, groups: dict[str, SemanticGroup]) -> tuple[int, list[str]]:
    with zipfile.ZipFile(workbook, "r") as source:
        shared_strings = read_shared_strings(source)
        pain_path, pain_xml, reviews, terms_by_review = replace_pain_point_groups(
            source,
            shared_strings,
            groups,
        )
        frequency_rows = []
        missing: list[str] = []
        keys = list(
            dict.fromkeys(
                (
                    review.get("asin", ""),
                    review.get("source_asin", ""),
                    review.get("actual_filter", ""),
                )
                for review in reviews
            )
        )
        for asin, source_asin, actual_filter in keys:
            indexes = [
                index
                for index, review in enumerate(reviews)
                if (
                    review.get("asin", ""),
                    review.get("source_asin", ""),
                    review.get("actual_filter", ""),
                )
                == (asin, source_asin, actual_filter)
            ]
            grouped_rows, current_missing = aggregate_semantic_groups(
                [reviews[index] for index in indexes],
                [terms_by_review[index] for index in indexes],
                custom_groups=groups,
            )
            frequency_rows.extend(
                (asin, source_asin, actual_filter, item) for item in grouped_rows
            )
            missing.extend(current_missing)
        missing = list(dict.fromkeys(missing))
        frequency_path, frequency_xml = replace_frequency_rows(source, shared_strings, frequency_rows)
        replacements = {pain_path: pain_xml, frequency_path: frequency_xml}
        temporary = workbook.with_name(f"{workbook.stem}_pain_group_tmp.xlsx")
        with zipfile.ZipFile(temporary, "w", compression=zipfile.ZIP_DEFLATED) as destination:
            for info in source.infolist():
                destination.writestr(info, replacements.get(info.filename, source.read(info.filename)))
    os.replace(temporary, workbook)
    return len(frequency_rows), missing


def main() -> int:
    parser = argparse.ArgumentParser(
        description="把 Agent 完成的差评短语语义分组写入缓存，并原位刷新差评词频分析。"
    )
    parser.add_argument("--groups", type=Path, required=True)
    parser.add_argument("--workbook", type=Path, required=True)
    parser.add_argument("--cache", type=Path, default=DEFAULT_CACHE_PATH)
    args = parser.parse_args()

    if not args.workbook.exists():
        raise FileNotFoundError(args.workbook)
    supplied = load_group_mapping(args.groups)
    if not supplied:
        raise ValueError("分组文件中没有完整的 group_en/group_zh 映射。")
    cached = load_group_mapping(args.cache) if args.cache.exists() else {}
    cached.update(supplied)
    row_count, missing = refresh_workbook(args.workbook, cached)
    update_cache(args.cache, supplied)
    print(f"frequency_rows={row_count}")
    print(f"remaining_missing_groups={len(missing)}")
    print(f"cache_updated={args.cache}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
