import threading
import time

from excel2img import watcher


def test_debounce_collapses_events_and_latest_event_runs(tmp_path, monkeypatch):
    path = tmp_path / "book.xlsx"
    path.write_bytes(b"placeholder")
    monkeypatch.setattr(watcher, "wait_until_stable", lambda _path: True)
    called = []
    finished = threading.Event()

    def process(value):
        called.append(value)
        finished.set()

    processor = watcher.DebouncedProcessor(process, debounce_seconds=0.03)
    try:
        processor.submit(path)
        processor.submit(path)
        processor.submit(path)
        assert finished.wait(1)
        time.sleep(0.08)
        assert called == [path.resolve()]
    finally:
        processor.close()


def test_fatal_automation_error_stops_queued_processing(tmp_path, monkeypatch):
    class FatalAutomation(Exception):
        pass

    first = tmp_path / "first.xlsx"
    second = tmp_path / "second.xlsx"
    first.write_bytes(b"placeholder")
    second.write_bytes(b"placeholder")
    monkeypatch.setattr(watcher, "wait_until_stable", lambda _path: True)
    called = []
    failed = threading.Event()

    def process(value):
        called.append(value)
        failed.set()
        raise FatalAutomation("Excel unavailable")

    processor = watcher.DebouncedProcessor(
        process,
        debounce_seconds=0.01,
        fatal_exceptions=(FatalAutomation,),
    )
    try:
        processor.submit(first)
        assert failed.wait(1)
        time.sleep(0.05)
        processor.submit(second)
        time.sleep(0.05)
        assert called == [first.resolve()]
    finally:
        processor.close()

