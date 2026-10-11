"""Reliable debounced workbook watching."""

from __future__ import annotations

import threading
import time
import zipfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Callable

from watchdog.events import FileSystemEvent, FileSystemEventHandler
from watchdog.observers import Observer

from .paths import is_eligible_workbook


def wait_until_stable(
    path: Path, interval: float = 0.4, consecutive: int = 2, timeout: float = 20.0
) -> bool:
    deadline = time.monotonic() + timeout
    previous: tuple[int, int] | None = None
    unchanged = 0
    while time.monotonic() < deadline:
        try:
            stat = path.stat()
            signature = (stat.st_size, stat.st_mtime_ns)
            if signature == previous:
                unchanged += 1
            else:
                previous, unchanged = signature, 0
            if unchanged >= consecutive:
                with zipfile.ZipFile(path) as archive:
                    if archive.testzip() is None:
                        return True
        except (FileNotFoundError, PermissionError, OSError, zipfile.BadZipFile):
            unchanged = 0
        time.sleep(interval)
    return False


class DebouncedProcessor:
    def __init__(
        self,
        process: Callable[[Path], object],
        log: Callable[[str], None] = print,
        debounce_seconds: float = 0.8,
    ):
        self.process = process
        self.log = log
        self.debounce_seconds = debounce_seconds
        self._lock = threading.Lock()
        self._states: dict[Path, dict[str, object]] = {}
        # Excel and the macOS pasteboard are process-global resources, so render
        # workbooks serially. Events still queue and coalesce independently.
        self._executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="excel2img")

    def submit(self, path: Path) -> None:
        path = path.resolve()
        if not is_eligible_workbook(path):
            return
        with self._lock:
            state = self._states.setdefault(path, {"generation": 0, "running": False, "timer": None})
            state["generation"] = int(state["generation"]) + 1
            generation = int(state["generation"])
            timer = state.get("timer")
            if isinstance(timer, threading.Timer):
                timer.cancel()
            new_timer = threading.Timer(self.debounce_seconds, self._ready, args=(path, generation))
            new_timer.daemon = True
            state["timer"] = new_timer
            new_timer.start()

    def _ready(self, path: Path, generation: int) -> None:
        with self._lock:
            state = self._states.get(path)
            if not state or generation != state["generation"]:
                return
            state["timer"] = None
            if state["running"]:
                return
            state["running"] = True
        self._executor.submit(self._run, path, generation)

    def _run(self, path: Path, generation: int) -> None:
        try:
            if not wait_until_stable(path):
                raise RuntimeError("file did not become stable and readable")
            self.process(path)
        except Exception as exc:
            self.log(f"Failed: {path.name}: {exc}")
        finally:
            with self._lock:
                state = self._states.get(path)
                if not state:
                    return
                state["running"] = False
                latest = int(state["generation"])
                if latest > generation:
                    timer = threading.Timer(0, self._ready, args=(path, latest))
                    timer.daemon = True
                    state["timer"] = timer
                    timer.start()

    def close(self) -> None:
        with self._lock:
            for state in self._states.values():
                timer = state.get("timer")
                if isinstance(timer, threading.Timer):
                    timer.cancel()
        self._executor.shutdown(wait=True)


class ExcelEventHandler(FileSystemEventHandler):
    def __init__(self, processor: DebouncedProcessor, log: Callable[[str], None] = print):
        self.processor = processor
        self.log = log

    def on_any_event(self, event: FileSystemEvent) -> None:
        if event.is_directory or event.event_type not in {"created", "modified", "moved"}:
            return
        value = event.dest_path if event.event_type == "moved" else event.src_path
        path = Path(value)
        if is_eligible_workbook(path):
            self.log(f"Modified: {path.name}")
            self.processor.submit(path)


def watch(directories: list[Path], process: Callable[[Path], object], log=print) -> None:
    processor = DebouncedProcessor(process, log)
    observer = Observer()
    handler = ExcelEventHandler(processor, log)
    for directory in directories:
        observer.schedule(handler, str(directory), recursive=True)
    observer.start()
    try:
        while observer.is_alive():
            observer.join(1)
    except KeyboardInterrupt:
        log("Stopping.")
    finally:
        observer.stop()
        observer.join()
        processor.close()

