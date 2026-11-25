from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from openpyxl import load_workbook

MODULES_DIR = Path(__file__).resolve().parents[2]
if str(MODULES_DIR) not in sys.path:
    sys.path.insert(0, str(MODULES_DIR))
from shared.excel_inspect import find_source_sheet


def choose_sheet(wb, requested: str | None):
    if requested:
        return next(ws for ws in wb if ws.title.strip() == requested.strip())
    source = find_source_sheet(wb)
    return source if source is not None else wb.worksheets[0]


def clean(value):
    if value is None:
        return None
    return value


def read_records(ws) -> list[dict]:
    rows = ws.iter_rows(values_only=True)
    try:
        headers = next(rows)
    except StopIteration:
        return []
    headers_by_index = {
        index: str(header)
        for index, header in enumerate(headers)
        if header not in (None, "")
    }
    records = []
    for row_values in rows:
        record = {}
        for index, header in headers_by_index.items():
            value = row_values[index] if index < len(row_values) else None
            record[header] = clean(value)
        if any(value not in (None, "") for value in record.values()):
            records.append(record)
    return records


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--sheet")
    args = parser.parse_args()

    wb = load_workbook(args.source, read_only=True, data_only=True)
    try:
        ws = choose_sheet(wb, args.sheet)
        records = read_records(ws)
    finally:
        wb.close()

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps({"source_records": records}, ensure_ascii=False), encoding="utf-8")
    print(args.output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
