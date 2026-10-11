from pathlib import Path

from excel2img import cli
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
