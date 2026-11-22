from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import tempfile
import zipfile
from xml.etree import ElementTree as ET
from pathlib import Path


DEFAULT_CACHE_PATH = Path(__file__).resolve().parents[3] / "resources" / "category_translations.json"
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
XML_NS = "http://www.w3.org/XML/1998/namespace"


def load_translation_mapping(path: Path) -> dict[str, str]:
    data = json.loads(path.read_text(encoding="utf-8-sig"))
    if isinstance(data, dict) and isinstance(data.get("translations"), dict):
        data = data["translations"]
    if not isinstance(data, dict):
        raise ValueError("Translation JSON must be an object or contain a 'translations' object.")

    translations: dict[str, str] = {}
    for english, chinese in data.items():
        english_text = str(english).strip()
        chinese_text = str(chinese).strip()
        if english_text and chinese_text:
            translations[english_text] = chinese_text
    return translations


def update_cache(cache_path: Path, translations: dict[str, str]) -> None:
    existing: dict[str, object] = {}
    if cache_path.exists():
        raw = json.loads(cache_path.read_text(encoding="utf-8-sig"))
        if isinstance(raw, dict):
            existing = raw

    cached = existing.get("translations")
    if not isinstance(cached, dict):
        cached = {}
    cached.update(translations)
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    cache_path.write_text(
        json.dumps({"translations": cached}, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


EXCEL_COM_REFRESH_SCRIPT = r'''
param([string]$PayloadPath)
$ErrorActionPreference = 'Stop'
$payload = Get-Content -LiteralPath $PayloadPath -Raw -Encoding UTF8 | ConvertFrom-Json
$translations = @{}
foreach ($property in $payload.translations.PSObject.Properties) {
    $translations[[string]$property.Name] = [string]$property.Value
}

function Release-ComObject($object) {
    if ($null -ne $object) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($object)
    }
}

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $workbook = $excel.Workbooks.Open([string]$payload.workbook)

    $worksheet = $null
    foreach ($candidate in $workbook.Worksheets) {
        if ([string]$candidate.Name -like '*市场调研') {
            $worksheet = $candidate
            break
        }
    }
    if ($null -eq $worksheet) {
        throw 'Could not find the market research worksheet.'
    }

    $updated = 0
    $usedRange = $worksheet.UsedRange
    $firstRow = [int]$usedRange.Row
    $lastRow = $firstRow + [int]$usedRange.Rows.Count - 1
    $firstColumn = [int]$usedRange.Column
    $lastColumn = $firstColumn + [int]$usedRange.Columns.Count - 1
    for ($row = $firstRow; $row -le $lastRow; $row++) {
        for ($column = $firstColumn; $column -le $lastColumn; $column++) {
            $cell = $worksheet.Cells.Item($row, $column)
            $value = [string]$cell.Value2
            if ($translations.ContainsKey($value)) {
                $cell.Value2 = "$value`n($($translations[$value]))"
                $cell.WrapText = $true
                $updated++
            }
        }
    }
    $workbook.Save()
    Write-Output "updated_cells=$updated"
}
finally {
    if ($null -ne $workbook) {
        $workbook.Close($false)
        Release-ComObject $workbook
    }
    if ($null -ne $excel) {
        $excel.Quit()
        Release-ComObject $excel
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
'''


def tag(name: str) -> str:
    return f"{{{MAIN_NS}}}{name}"


def cell_text(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.get("t")
    if cell_type == "s":
        value = cell.findtext(tag("v"))
        if value and value.isdigit() and int(value) < len(shared_strings):
            return shared_strings[int(value)]
        return ""
    if cell_type == "inlineStr":
        return "".join(cell.find(tag("is")).itertext()) if cell.find(tag("is")) is not None else ""
    return cell.findtext(tag("v")) or ""


def read_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    return ["".join(item.itertext()) for item in root.findall(tag("si"))]


def market_research_sheet_path(archive: zipfile.ZipFile) -> str:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {
        relation.get("Id"): relation.get("Target", "")
        for relation in relationships.findall(f"{{{PACKAGE_REL_NS}}}Relationship")
    }
    for sheet in workbook.findall(f".//{tag('sheet')}"):
        if sheet.get("name", "").endswith("市场调研"):
            target = targets.get(sheet.get(f"{{{REL_NS}}}id"), "")
            if target.startswith("/"):
                return target.lstrip("/")
            return f"xl/{target}"
    raise RuntimeError("Could not find the market research worksheet.")


def refresh_workbook(workbook: Path, translations: dict[str, str]) -> None:
    """Patch only category text in the OOXML package, preserving drawings unchanged."""
    with zipfile.ZipFile(workbook, "r") as source:
        sheet_path = market_research_sheet_path(source)
        shared_strings = read_shared_strings(source)
        worksheet = ET.fromstring(source.read(sheet_path))
        updated = 0
        for cell in worksheet.iter(tag("c")):
            value = cell_text(cell, shared_strings)
            translation = translations.get(value)
            if not translation:
                continue
            for child in list(cell):
                if child.tag in {tag("v"), tag("is")}:
                    cell.remove(child)
            cell.set("t", "inlineStr")
            inline_string = ET.SubElement(cell, tag("is"))
            text = ET.SubElement(inline_string, tag("t"))
            text.set(f"{{{XML_NS}}}space", "preserve")
            text.text = f"{value}\n({translation})"
            updated += 1

        replacement = ET.tostring(worksheet, encoding="utf-8", xml_declaration=True)
        temporary = workbook.with_name(f"{workbook.stem}_category_translation_tmp.xlsx")
        with zipfile.ZipFile(temporary, "w", compression=zipfile.ZIP_DEFLATED) as destination:
            for info in source.infolist():
                data = replacement if info.filename == sheet_path else source.read(info.filename)
                destination.writestr(info, data)
    os.replace(temporary, workbook)
    print(f"updated_cells={updated}")


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Merge agent-approved category translations into the local cache and refresh an existing report without touching images."
    )
    parser.add_argument("--translations", type=Path, required=True, help="JSON object: English category level -> Chinese translation.")
    parser.add_argument("--workbook", type=Path, required=True, help="Completed .xlsx workbook to update in place through Excel COM.")
    parser.add_argument("--cache", type=Path, default=DEFAULT_CACHE_PATH, help="Local category translation cache path.")
    args = parser.parse_args()

    if not args.workbook.exists():
        raise FileNotFoundError(args.workbook)
    translations = load_translation_mapping(args.translations)
    if not translations:
        raise ValueError("No valid translations were provided.")
    update_cache(args.cache, translations)
    refresh_workbook(args.workbook, translations)
    print(f"cache_updated={args.cache}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
