from __future__ import annotations

from pathlib import Path


def ensure_dir(path: str | Path) -> Path:
    target = Path(path).expanduser().resolve()
    target.mkdir(parents=True, exist_ok=True)
    return target


def resolve_existing_file(path: str | Path, label: str) -> Path:
    target = Path(path).expanduser().resolve()
    if not target.exists():
        raise FileNotFoundError(f"{label} not found: {target}")
    if not target.is_file():
        raise FileNotFoundError(f"{label} is not a file: {target}")
    return target
