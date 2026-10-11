"""Non-activating access to the workbook currently being edited in Excel."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from .exporter import process_workbook
from .paths import is_eligible_workbook
from .renderer import RECOVERY_TIMEOUT, RenderError, _as_literal, _run_osascript
from .watcher import wait_until_stable

EXCEL_BUNDLE_ID = "com.microsoft.Excel"


class ActiveWorkbookError(RenderError):
    """The active Excel workbook cannot safely be rendered."""


@dataclass(frozen=True)
class ActiveWorkbook:
    name: str
    path: Path
    saved: bool
    read_only: bool


def frontmost_application() -> tuple[str, str]:
    result = _run_osascript(
        """
ObjC.import('AppKit');
const app = $.NSWorkspace.sharedWorkspace.frontmostApplication;
[ObjC.unwrap(app.bundleIdentifier) || '', ObjC.unwrap(app.localizedName) || ''].join('\u001f');
""",
        language="JavaScript",
        timeout=RECOVERY_TIMEOUT,
        phase="Reading the frontmost application",
    )
    parts = result.split(chr(31), 1)
    if len(parts) != 2:
        raise ActiveWorkbookError(f"macOS returned an unrecognized application: {result!r}")
    return parts[0], parts[1]


def require_excel_frontmost() -> None:
    bundle_id, name = frontmost_application()
    if bundle_id != EXCEL_BUNDLE_ID:
        visible_name = name or bundle_id or "an unknown application"
        raise ActiveWorkbookError(
            f"Microsoft Excel is not frontmost (frontmost: {visible_name}). "
            "Excel2Img did not activate it or process a workbook."
        )


def get_active_workbook() -> ActiveWorkbook:
    result = _run_osascript(
        """
set fieldSeparator to ASCII character 31
tell application "Microsoft Excel"
    if (count of workbooks) is 0 then error "No workbook is open in Microsoft Excel."
    set targetWorkbook to active workbook
    set workbookName to name of targetWorkbook as text
    set workbookPath to path of targetWorkbook as text
    if workbookPath is "" then
        set workbookFullName to ""
    else
        set workbookFullName to full name of targetWorkbook as text
    end if
    set workbookSaved to saved of targetWorkbook as text
    set workbookReadOnly to read only of targetWorkbook as text
end tell
return workbookName & fieldSeparator & workbookPath & fieldSeparator & workbookFullName & fieldSeparator & workbookSaved & fieldSeparator & workbookReadOnly
""",
        timeout=RECOVERY_TIMEOUT,
        phase="Reading Excel's active workbook",
    )
    parts = result.split(chr(31))
    if len(parts) != 5:
        raise ActiveWorkbookError(f"Excel returned unrecognized workbook details: {result!r}")
    name, parent, full_name, saved_text, read_only_text = parts
    if not parent or not full_name or full_name == name:
        raise ActiveWorkbookError(
            f"{name or 'The active workbook'} has never been saved. Save it normally "
            "in Excel before running Excel2Img."
        )
    if saved_text.lower() not in {"true", "false"} or read_only_text.lower() not in {
        "true",
        "false",
    }:
        raise ActiveWorkbookError(f"Excel returned unrecognized workbook details: {result!r}")
    return ActiveWorkbook(
        name=name,
        path=Path(full_name).expanduser().resolve(),
        saved=saved_text.lower() == "true",
        read_only=read_only_text.lower() == "true",
    )


def save_active_workbook(workbook: ActiveWorkbook) -> None:
    if workbook.read_only:
        raise ActiveWorkbookError(
            f"{workbook.name} is read-only and has unsaved changes; it was not rendered."
        )
    expected = _as_literal(str(workbook.path))
    _run_osascript(
        f"""
tell application "Microsoft Excel"
    if (count of workbooks) is 0 then error "No workbook is open in Microsoft Excel."
    set targetWorkbook to active workbook
    if (full name of targetWorkbook as text) is not {expected} then error "The active workbook changed before it could be saved."
    save targetWorkbook
    if not (saved of targetWorkbook) then error "Excel did not finish saving the active workbook."
end tell
""",
        phase=f"Saving {workbook.name}",
    )


def process_active_workbook(log: Callable[[str], None] = print) -> list[Path]:
    require_excel_frontmost()
    workbook = get_active_workbook()
    if not is_eligible_workbook(workbook.path):
        if workbook.name.startswith(("!", "~$")):
            log(f"Skipped by exclusion rule: {workbook.name}")
            return []
        raise ActiveWorkbookError(f"Unsupported active workbook path: {workbook.path}")
    if not workbook.path.is_file():
        raise ActiveWorkbookError(f"The active workbook file does not exist: {workbook.path}")
    if not workbook.saved:
        log(f"Saving modified workbook: {workbook.name}")
        save_active_workbook(workbook)
    if not wait_until_stable(workbook.path):
        raise ActiveWorkbookError(
            f"{workbook.name} did not become stable and readable after saving."
        )
    log(f"Processing active workbook: {workbook.name}")
    return process_workbook(workbook.path, log=log)
