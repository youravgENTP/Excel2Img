from pathlib import Path

import pytest

from excel2img import cli
from excel2img.locking import AlreadyRunningError
from excel2img.renderer import AutomationUnrecoverableError


def test_file_processes_only_requested_workbook(tmp_path, monkeypatch):
    workbook = tmp_path / "한 파일.xlsx"
    workbook.write_bytes(b"file")
    called = []
    monkeypatch.setattr(cli, "process_workbook", lambda path, log: called.append(path) or [])
    assert cli.main(["--file", str(workbook)]) == 0
    assert called == [workbook.resolve()]


def test_file_respects_exclusion_rule(tmp_path, monkeypatch):
    workbook = tmp_path / "!skip.xlsx"
    workbook.write_bytes(b"file")
    called = []
    monkeypatch.setattr(cli, "process_workbook", lambda path, log: called.append(path))
    assert cli.main(["--file", str(workbook)]) == 0
    assert called == []


def test_check_returns_success_or_failure(monkeypatch):
    monkeypatch.setattr(cli, "automation_preflight", lambda _log: None)
    assert cli.main(["--check"]) == 0

    def fail(_log):
        raise RuntimeError("no Excel")

    monkeypatch.setattr(cli, "automation_preflight", fail)
    assert cli.main(["--check"]) == 1


def test_batch_aborts_after_unrecoverable_automation_failure(tmp_path, monkeypatch):
    workbooks = [tmp_path / "one.xlsx", tmp_path / "two.xlsx"]
    for workbook in workbooks:
        workbook.write_bytes(b"file")
    monkeypatch.setattr(cli, "load_directories", lambda _path: ([tmp_path], []))
    monkeypatch.setattr(cli, "iter_workbooks", lambda _paths: iter(workbooks))
    monkeypatch.setattr(cli, "automation_preflight", lambda _log: None)
    called = []

    def fail(path, log):
        called.append(path)
        raise AutomationUnrecoverableError("Excel stuck")

    monkeypatch.setattr(cli, "process_workbook", fail)
    assert cli.main(["--once"]) == 1
    assert called == [workbooks[0]]


def test_default_does_not_start_watcher_or_read_config(monkeypatch):
    monkeypatch.setattr(
        cli, "load_directories", lambda _path: (_ for _ in ()).throw(AssertionError("config read"))
    )
    monkeypatch.setattr(cli, "watch", lambda *_args, **_kwargs: pytest.fail("watch started"))
    assert cli.main([]) == 0


def test_watch_is_explicit_opt_in(tmp_path, monkeypatch):
    called = []
    monkeypatch.setattr(cli, "load_directories", lambda _path: ([tmp_path], []))
    monkeypatch.setattr(cli, "automation_preflight", lambda _log: None)
    monkeypatch.setattr(cli, "watch", lambda directories, *_args, **_kwargs: called.append(directories))
    assert cli.main(["--watch"]) == 0
    assert called == [[tmp_path]]


def test_active_uses_process_lock(monkeypatch, capsys):
    class BusyLock:
        def __enter__(self):
            raise AlreadyRunningError("another --active render is already running")

        def __exit__(self, *_args):
            return None

    monkeypatch.setattr(cli, "ActiveRunLock", BusyLock)
    monkeypatch.setattr(
        cli, "process_active_workbook", lambda _log: pytest.fail("duplicate processed")
    )
    assert cli.main(["--active"]) == 0
    assert "already running" in capsys.readouterr().err


def test_active_failure_is_nonzero_and_written_to_stderr(monkeypatch, capsys):
    monkeypatch.setattr(
        cli,
        "process_active_workbook",
        lambda _log: (_ for _ in ()).throw(RuntimeError("Excel is not frontmost")),
    )
    assert cli.main(["--active"]) == 1
    output = capsys.readouterr()
    assert "Excel is not frontmost" in output.err


def test_active_does_not_start_watcher(monkeypatch):
    monkeypatch.setattr(cli, "process_active_workbook", lambda _log: [])
    monkeypatch.setattr(cli, "watch", lambda *_args, **_kwargs: pytest.fail("watch started"))
    assert cli.main(["--active"]) == 0
