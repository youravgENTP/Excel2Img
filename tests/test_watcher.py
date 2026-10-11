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

