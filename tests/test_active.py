from pathlib import Path

import pytest

from excel2img import active


def workbook(path: Path, *, saved: bool = True, read_only: bool = False):
    return active.ActiveWorkbook(path.name, path, saved, read_only)


def test_frontmost_application_must_be_excel(monkeypatch):
    monkeypatch.setattr(active, "frontmost_application", lambda: ("com.apple.finder", "Finder"))
    with pytest.raises(active.ActiveWorkbookError, match="not frontmost"):
        active.require_excel_frontmost()


def test_active_workbook_details_include_absolute_path(monkeypatch, tmp_path):
    path = tmp_path / "현재.xlsx"
    separator = chr(31)
    monkeypatch.setattr(
        active,
        "_run_osascript",
        lambda *_args, **_kwargs: separator.join(
            [path.name, str(path.parent), str(path), "false", "true"]
        ),
    )
    result = active.get_active_workbook()
    assert result == active.ActiveWorkbook(path.name, path.resolve(), False, True)


def test_never_saved_workbook_is_rejected(monkeypatch):
    separator = chr(31)
    monkeypatch.setattr(
        active,
        "_run_osascript",
        lambda *_args, **_kwargs: separator.join(["Book1", "", "", "false", "false"]),
    )
    with pytest.raises(active.ActiveWorkbookError, match="never been saved"):
        active.get_active_workbook()


def test_modified_workbook_is_saved_then_processed(monkeypatch, tmp_path):
    path = tmp_path / "book.xlsx"
    path.write_bytes(b"workbook")
    details = workbook(path.resolve(), saved=False)
    calls = []
    monkeypatch.setattr(active, "require_excel_frontmost", lambda: calls.append("frontmost"))
    monkeypatch.setattr(active, "get_active_workbook", lambda: details)
    monkeypatch.setattr(active, "save_active_workbook", lambda value: calls.append(("save", value)))
    monkeypatch.setattr(active, "wait_until_stable", lambda value: calls.append(("stable", value)) or True)
    monkeypatch.setattr(
        active,
        "process_workbook",
        lambda value, log: calls.append(("process", value)) or [Path("image.png")],
    )
    assert active.process_active_workbook(lambda _message: None) == [Path("image.png")]
    assert calls == [
        "frontmost",
        ("save", details),
        ("stable", path.resolve()),
        ("process", path.resolve()),
    ]


def test_unmodified_workbook_is_not_saved(monkeypatch, tmp_path):
    path = tmp_path / "book.xlsx"
    path.write_bytes(b"workbook")
    monkeypatch.setattr(active, "require_excel_frontmost", lambda: None)
    monkeypatch.setattr(active, "get_active_workbook", lambda: workbook(path.resolve()))
    monkeypatch.setattr(
        active,
        "save_active_workbook",
        lambda _value: pytest.fail("an unmodified workbook must not be saved"),
    )
    monkeypatch.setattr(active, "wait_until_stable", lambda _value: True)
    monkeypatch.setattr(active, "process_workbook", lambda value, log: [])
    assert active.process_active_workbook() == []


def test_read_only_modified_workbook_is_not_processed(monkeypatch, tmp_path):
    path = tmp_path / "book.xlsx"
    path.write_bytes(b"workbook")
    monkeypatch.setattr(active, "require_excel_frontmost", lambda: None)
    monkeypatch.setattr(
        active, "get_active_workbook", lambda: workbook(path.resolve(), saved=False, read_only=True)
    )
    monkeypatch.setattr(
        active, "process_workbook", lambda *_args, **_kwargs: pytest.fail("must not process")
    )
    with pytest.raises(active.ActiveWorkbookError, match="read-only"):
        active.process_active_workbook()


def test_save_failure_is_reported_without_processing(monkeypatch, tmp_path):
    path = tmp_path / "book.xlsx"
    path.write_bytes(b"workbook")
    monkeypatch.setattr(active, "require_excel_frontmost", lambda: None)
    monkeypatch.setattr(
        active, "get_active_workbook", lambda: workbook(path.resolve(), saved=False)
    )
    monkeypatch.setattr(
        active,
        "save_active_workbook",
        lambda _value: (_ for _ in ()).throw(active.RenderError("save denied")),
    )
    monkeypatch.setattr(
        active, "process_workbook", lambda *_args, **_kwargs: pytest.fail("must not process")
    )
    with pytest.raises(active.RenderError, match="save denied"):
        active.process_active_workbook()


def test_active_exclusion_rule_skips_without_save(monkeypatch, tmp_path):
    path = tmp_path / "!book.xlsx"
    path.write_bytes(b"workbook")
    monkeypatch.setattr(active, "require_excel_frontmost", lambda: None)
    monkeypatch.setattr(
        active, "get_active_workbook", lambda: workbook(path.resolve(), saved=False)
    )
    monkeypatch.setattr(
        active, "save_active_workbook", lambda _value: pytest.fail("excluded workbook saved")
    )
    assert active.process_active_workbook(lambda _message: None) == []


def test_save_script_has_no_activation_or_navigation(monkeypatch, tmp_path):
    path = (tmp_path / "book.xlsx").resolve()
    captured = {}

    def run(script, **kwargs):
        captured["script"] = script.lower()
        return ""

    monkeypatch.setattr(active, "_run_osascript", run)
    active.save_active_workbook(workbook(path, saved=False))
    script = captured["script"]
    assert "save targetworkbook" in script
    assert "save as" not in script
    assert "activate" not in script
    assert "select" not in script
    assert "goto" not in script
