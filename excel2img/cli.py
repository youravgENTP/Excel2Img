from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path

from .exporter import process_workbook
from .paths import iter_workbooks, load_directories
from .watcher import watch


def _log(message: str) -> None:
    print(f"[{datetime.now():%H:%M:%S}] {message}", flush=True)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Export bordered Excel tables as PNG images.")
    parser.add_argument("--once", action="store_true", help="process current workbooks and exit")
    parser.add_argument(
        "--config", type=Path, default=Path("excelDirectoryList.json"), help=argparse.SUPPRESS
    )
    args = parser.parse_args(argv)
    try:
        directories, missing = load_directories(args.config)
    except ValueError as exc:
        parser.error(str(exc))
    for path in missing:
        _log(f"Warning: configured directory does not exist: {path}")
    if not directories:
        _log("No existing configured directories; nothing to do.")
        return 0

    def process(path: Path):
        _log(f"Processing: {path.name}")
        result = process_workbook(path, log=_log)
        _log("Completed.")
        return result

    if args.once:
        workbooks = list(iter_workbooks(directories))
        _log(f"Found {len(workbooks)} eligible workbooks.")
        failed = False
        for workbook in workbooks:
            try:
                process(workbook)
            except Exception as exc:
                failed = True
                _log(f"Failed: {workbook.name}: {exc}")
        return 1 if failed else 0

    print(f"[Excel2Img] Watching {len(directories)} directories.", flush=True)
    watch(directories, process, _log)
    return 0

