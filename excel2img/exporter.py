"""Workbook processing and safe output management."""

from __future__ import annotations

import hashlib
import json
import os
import re
import tempfile
from pathlib import Path
from typing import Callable

from openpyxl import load_workbook

from .detector import workbook_tables
from .paths import is_eligible_workbook
from .renderer import ExcelRenderer, temporary_excel_workbook

MANIFEST = ".excel2img-manifest.json"
_UNSAFE = re.compile(r"[/:\\\x00-\x1f]")


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def sanitize_filename(value: str) -> str:
    cleaned = _UNSAFE.sub("_", value).strip().rstrip(".")
    return cleaned or "Sheet"


def _read_manifest(output_dir: Path) -> dict[str, str]:
    try:
        data = json.loads((output_dir / MANIFEST).read_text(encoding="utf-8"))
        files = data.get("generated", {})
        if not isinstance(files, dict):
            return {}
        return {
            name: digest
            for name, digest in files.items()
            if isinstance(name, str)
            and Path(name).name == name
            and isinstance(digest, str)
            and len(digest) == 64
        }
    except (FileNotFoundError, OSError, json.JSONDecodeError, AttributeError):
        return {}


def _write_manifest(output_dir: Path, generated: dict[str, str]) -> None:
    payload = json.dumps({"generated": generated}, ensure_ascii=False, indent=2) + "\n"
    fd, temporary = tempfile.mkstemp(prefix=".manifest-", dir=output_dir)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as stream:
            stream.write(payload)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, output_dir / MANIFEST)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)


def process_workbook(
    source: Path,
    renderer: ExcelRenderer | None = None,
    log: Callable[[str], None] = print,
) -> list[Path]:
    source = source.resolve()
    if not is_eligible_workbook(source):
        return []
    keep_vba = source.suffix.lower() == ".xlsm"
    workbook = load_workbook(source, read_only=False, data_only=False, keep_vba=keep_vba)
    try:
        tables = workbook_tables(workbook)
    finally:
        workbook.close()

    worksheet_count = len({sheet for sheet, _ in tables})
    log(f"Detected {len(tables)} tables across {worksheet_count} worksheets.")
    output_dir = source.with_name(f"{source.stem}_Img")
    output_dir.mkdir(exist_ok=True)
    old_generated = _read_manifest(output_dir)
    counts: dict[str, int] = {}
    jobs: list[tuple[str, str, str]] = []
    for sheet_name, table in tables:
        counts[sheet_name] = counts.get(sheet_name, 0) + 1
        filename = f"{sanitize_filename(sheet_name)}_{counts[sheet_name]:02d}.png"
        jobs.append((sheet_name, table.coordinate, filename))

    renderer = renderer or ExcelRenderer()
    staged: list[tuple[Path, str]] = []
    try:
        if jobs:
            with temporary_excel_workbook(source) as workbook_name:
                for sheet_name, coordinate, filename in jobs:
                    fd, staging_name = tempfile.mkstemp(
                        prefix=".render-", suffix=".png", dir=output_dir
                    )
                    os.close(fd)
                    staging = Path(staging_name)
                    renderer.render_open_workbook(workbook_name, sheet_name, coordinate, staging)
                    staged.append((staging, filename))
        for staging, filename in staged:
            os.replace(staging, output_dir / filename)
            log(f"Exported: {filename}")
    finally:
        for staging, _ in staged:
            staging.unlink(missing_ok=True)

    current_names = [filename for _, _, filename in jobs]
    for stale in old_generated.keys() - set(current_names):
        candidate = output_dir / stale
        if (
            candidate.parent == output_dir
            and candidate.is_file()
            and _sha256(candidate) == old_generated[stale]
        ):
            candidate.unlink()
            log(f"Removed stale: {stale}")
        elif candidate.is_file():
            log(f"Preserved modified file: {stale}")
    current = {
        filename: _sha256(output_dir / filename)
        for filename in current_names
    }
    _write_manifest(output_dir, current)
    return [output_dir / filename for filename in current_names]

