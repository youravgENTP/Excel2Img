"""Cross-process coordination for manual active-workbook renders."""

from __future__ import annotations

import fcntl
import os
import tempfile
from pathlib import Path
from types import TracebackType

ACTIVE_LOCK_PATH = Path(tempfile.gettempdir()) / f"excel2img-active-{os.getuid()}.lock"


class AlreadyRunningError(RuntimeError):
    """Another active-workbook render owns the process lock."""


class ActiveRunLock:
    """A crash-safe advisory lock; the kernel releases it when the process exits."""

    def __init__(self, path: Path = ACTIVE_LOCK_PATH):
        self.path = path
        self._stream = None

    def __enter__(self) -> ActiveRunLock:
        self._stream = self.path.open("a+")
        try:
            fcntl.flock(self._stream.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError as exc:
            self._stream.close()
            self._stream = None
            raise AlreadyRunningError("another --active render is already running") from exc
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        traceback: TracebackType | None,
    ) -> None:
        assert self._stream is not None
        fcntl.flock(self._stream.fileno(), fcntl.LOCK_UN)
        self._stream.close()
        self._stream = None


def active_run_is_locked(path: Path = ACTIVE_LOCK_PATH) -> bool:
    """Return whether another process currently holds the active-render lock."""
    stream = path.open("a+")
    try:
        try:
            fcntl.flock(stream.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError:
            return True
        fcntl.flock(stream.fileno(), fcntl.LOCK_UN)
        return False
    finally:
        stream.close()
