"""Configuration and path rules."""

from __future__ import annotations

import json
from pathlib import Path

EXCEL_SUFFIXES = {".xlsx", ".xlsm"}


def is_output_part(path: Path) -> bool:
    return any(part.endswith("_Img") for part in path.parts)


def is_eligible_workbook(path: Path) -> bool:
    return (
        path.suffix.lower() in EXCEL_SUFFIXES
        and not path.name.startswith(("!", "~$"))
        and not is_output_part(path)
    )


def load_directories(config_path: Path) -> tuple[list[Path], list[Path]]:
    try:
        data = json.loads(config_path.read_text(encoding="utf-8"))
    except FileNotFoundError as exc:
        raise ValueError(
            f"Configuration not found: {config_path}. Copy "
            "excelDirectoryList.example.json to excelDirectoryList.json first."
        ) from exc
    except (json.JSONDecodeError, OSError) as exc:
        raise ValueError(f"Cannot read configuration {config_path}: {exc}") from exc
    raw = data.get("directories") if isinstance(data, dict) else None
    if not isinstance(raw, list) or not all(isinstance(value, str) for value in raw):
        raise ValueError("Configuration must contain a string list named 'directories'.")

    existing: list[Path] = []
    missing: list[Path] = []
    for value in raw:
        path = Path(value).expanduser().resolve()
        (existing if path.is_dir() else missing).append(path)

    # A recursive watch of a parent already covers its descendants.
    unique: list[Path] = []
    for path in sorted(set(existing), key=lambda item: (len(item.parts), str(item))):
        if not any(path == parent or parent in path.parents for parent in unique):
            unique.append(path)
    return unique, sorted(set(missing))


def iter_workbooks(directories: list[Path]):
    seen: set[Path] = set()
    for directory in directories:
        for path in directory.rglob("*"):
            if path.is_file() and is_eligible_workbook(path):
                resolved = path.resolve()
                if resolved not in seen:
                    seen.add(resolved)
                    yield resolved

