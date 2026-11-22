from __future__ import annotations

import json
import re
import sys
import zipfile
from pathlib import Path


XLSX = Path(sys.argv[1] if len(sys.argv) > 1 else "outputs/report_filled.xlsx")


def main() -> int:
    with zipfile.ZipFile(XLSX) as zf:
        xml_text = " ".join(
            zf.read(name).decode("utf-8", "ignore")
            for name in zf.namelist()
            if name.endswith(".xml")
        )
        errors = sorted(set(re.findall(r"#(?:REF!|DIV/0!|VALUE!|NAME\?|N/A)", xml_text)))

    result = {
        "xlsx": str(XLSX.resolve()),
        "size_bytes": XLSX.stat().st_size,
        "formula_errors": errors,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
