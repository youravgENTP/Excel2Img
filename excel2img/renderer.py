"""High-fidelity rendering through Microsoft Excel for Mac."""

from __future__ import annotations

import shutil
import subprocess
import tempfile
import time
import uuid
from contextlib import contextmanager
from pathlib import Path
from typing import Iterator

import pymupdf


class RenderError(RuntimeError):
    pass


def _run_osascript(source: str, language: str = "AppleScript") -> str:
    command = ["/usr/bin/osascript"]
    if language != "AppleScript":
        command.extend(["-l", language])
    command.extend(["-e", source])
    completed = subprocess.run(command, text=True, capture_output=True, timeout=90)
    if completed.returncode:
        detail = completed.stderr.strip() or completed.stdout.strip()
        raise RenderError(f"Excel automation failed: {detail}")
    return completed.stdout.strip()


def _as_literal(value: str) -> str:
    return '"' + value.replace("\\", "\\\\").replace('"', '\\"') + '"'


@contextmanager
def temporary_excel_workbook(source: Path) -> Iterator[str]:
    if not Path("/Applications/Microsoft Excel.app").exists():
        raise RenderError("Microsoft Excel was not found in /Applications.")
    with tempfile.TemporaryDirectory(prefix="excel2img-") as directory:
        # Excel refuses to open two workbooks with the same filename, even from
        # different folders. The source is normally already open during a save.
        temp_path = Path(directory) / f"render-{uuid.uuid4().hex}{source.suffix}"
        shutil.copy2(source, temp_path)
        script = f'''
tell application "Microsoft Excel"
    set oldSecurity to automation security
    set oldAlerts to display alerts
    try
        set automation security to msoAutomationSecurityForceDisable
        set display alerts to false
        set wb to open workbook workbook file name {_as_literal(str(temp_path))} update links do not update links read only true
        set automation security to oldSecurity
        set display alerts to oldAlerts
        return name of wb
    on error errorMessage number errorNumber
        set automation security to oldSecurity
        set display alerts to oldAlerts
        error errorMessage number errorNumber
    end try
end tell
'''
        workbook_name = _run_osascript(script)
        try:
            yield workbook_name
        finally:
            close_script = f'''
tell application "Microsoft Excel"
    try
        set wb to workbook {_as_literal(workbook_name)}
        close wb saving no
    end try
end tell
'''
            try:
                _run_osascript(close_script)
            except RenderError:
                pass


def _copy_range(workbook_name: str, sheet_name: str, coordinate: str) -> None:
    script = f'''
tell application "Microsoft Excel"
    set wb to workbook {_as_literal(workbook_name)}
    set ws to worksheet {_as_literal(sheet_name)} of wb
    copy picture (range {_as_literal(coordinate)} of ws) appearance printer format picture
end tell
'''
    _run_osascript(script)
    time.sleep(0.15)


def _write_clipboard_pdf(destination: Path) -> None:
    # JXA provides direct access to the vector PDF pasteboard flavor produced by
    # Excel. TIFF remains a fallback for Excel versions that omit PDF.
    script = f'''
ObjC.import('AppKit');
const pb = $.NSPasteboard.generalPasteboard;
let data = pb.dataForType('com.adobe.pdf');
if (!data) data = pb.dataForType('public.tiff');
if (!data) throw new Error('Excel did not place PDF or TIFF image data on the clipboard');
if (!data.writeToFileAtomically({_as_literal(str(destination))}, true))
    throw new Error('Could not write clipboard image data');
'''
    _run_osascript(script, language="JavaScript")


def _rasterize(source: Path, destination: Path, dpi: int) -> None:
    try:
        document = pymupdf.open(source)
        if document.page_count:
            page = document[0]
            pixmap = page.get_pixmap(
                matrix=pymupdf.Matrix(dpi / 72, dpi / 72), alpha=False
            )
            pixmap.save(destination)
            document.close()
            return
        document.close()
    except (pymupdf.FileDataError, RuntimeError, ValueError):
        pass

    # TIFF fallback: Pillow is included by PyMuPDF in many environments, but is
    # deliberately optional because normal Excel output contains vector PDF.
    try:
        from PIL import Image

        with Image.open(source) as image:
            image.save(destination, format="PNG", dpi=(dpi, dpi))
    except Exception as exc:
        raise RenderError(f"Could not convert Excel clipboard image: {exc}") from exc


class ExcelRenderer:
    def __init__(self, dpi: int = 300):
        self.dpi = dpi

    def render_open_workbook(
        self, workbook_name: str, sheet_name: str, coordinate: str, destination: Path
    ) -> None:
        destination.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(prefix="excel2img-clipboard-") as directory:
            clipboard_data = Path(directory) / "range.data"
            _copy_range(workbook_name, sheet_name, coordinate)
            _write_clipboard_pdf(clipboard_data)
            _rasterize(clipboard_data, destination, self.dpi)

