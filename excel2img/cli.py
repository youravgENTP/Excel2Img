from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

from .exporter import process_workbook
from .active import process_active_workbook
from .locking import ActiveRunLock, AlreadyRunningError
from .paths import EXCEL_SUFFIXES, is_eligible_workbook, iter_workbooks, load_directories
from .renderer import (
    AutomationTimeout,
    AutomationUnrecoverableError,
    automation_preflight,
)
from .watcher import watch


def _log(message: str) -> None:
    print(f"[{datetime.now():%H:%M:%S}] {message}", flush=True)


def _error(message: str) -> None:
    print(f"[{datetime.now():%H:%M:%S}] {message}", file=sys.stderr, flush=True)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Export bordered Excel tables as PNG images.")
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--active", action="store_true", help="render Excel's active workbook")
    mode.add_argument("--watch", action="store_true", help="watch configured directories")
    mode.add_argument("--once", action="store_true", help="process current workbooks and exit")
    mode.add_argument("--file", type=Path, help="process only one workbook and exit")
    mode.add_argument("--check", action="store_true", help="test the complete Excel rendering path")
    parser.add_argument(
        "--config", type=Path, default=Path("excelDirectoryList.json"), help=argparse.SUPPRESS
    )
    args = parser.parse_args(argv)

    if args.active:
        try:
            with ActiveRunLock():
                process_active_workbook(_log)
        except AlreadyRunningError as exc:
            _error(f"Skipped: {exc}.")
            return 0
        except Exception as exc:
            _error(f"Failed: {exc}")
            return 1
        _log("Completed.")
        return 0

    if args.check:
        try:
            automation_preflight(_log)
        except Exception as exc:
            _log(f"Excel automation check failed: {exc}")
            return 1
        _log("Excel automation check passed.")
        return 0

    def process(path: Path):
        _log(f"Processing: {path.name}")
        result = process_workbook(path, log=_log)
        _log("Completed.")
        return result

    if args.file is not None:
        workbook = args.file.expanduser().resolve()
        if not workbook.is_file():
            _log(f"Failed: workbook does not exist: {workbook}")
            return 2
        if workbook.name.startswith(("!", "~$")):
            _log(f"Skipped by exclusion rule: {workbook.name}")
            return 0
        if workbook.suffix.lower() not in EXCEL_SUFFIXES or not is_eligible_workbook(workbook):
            _log(f"Failed: unsupported workbook path: {workbook}")
            return 2
        try:
            process(workbook)
        except Exception as exc:
            _log(f"Failed: {workbook.name}: {exc}")
            return 1
        return 0

    if not args.once and not args.watch:
        parser.print_help()
        return 0

    try:
        directories, missing = load_directories(args.config)
    except ValueError as exc:
        parser.error(str(exc))
    for path in missing:
        _log(f"Warning: configured directory does not exist: {path}")
    if not directories:
        _log("No existing configured directories; nothing to do.")
        return 0

    try:
        automation_preflight(_log)
    except Exception as exc:
        _log(f"Excel automation preflight failed: {exc}")
        return 1

    if args.once:
        workbooks = list(iter_workbooks(directories))
        _log(f"Found {len(workbooks)} eligible workbooks.")
        failed = False
        for workbook in workbooks:
            try:
                process(workbook)
            except (AutomationTimeout, AutomationUnrecoverableError) as exc:
                _log(f"Excel automation became unavailable: {exc}")
                _log("Aborting batch to prevent cascading failures. Run: python main.py --check")
                return 1
            except Exception as exc:
                failed = True
                _log(f"Failed: {workbook.name}: {exc}")
        return 1 if failed else 0

    print(f"[Excel2Img] Watching {len(directories)} directories.", flush=True)
    watch(
        directories,
        process,
        _log,
        fatal_exceptions=(AutomationTimeout, AutomationUnrecoverableError),
    )
    return 0

