import subprocess
from pathlib import Path

import pytest

from excel2img import renderer


def prepare(monkeypatch, tmp_path):
    app = tmp_path / "Microsoft Excel.app"
    app.mkdir()
    monkeypatch.setattr(renderer, "EXCEL_APP", app)
    monkeypatch.setattr(renderer, "EXCEL_TEMP_ROOT", tmp_path / "excel-temp")
    monkeypatch.setattr(renderer, "list_open_workbooks", lambda: [])
    source = tmp_path / "book.xlsm"
    source.write_bytes(b"workbook")
    return source


def test_osascript_timeout_names_the_failed_phase(monkeypatch):
    def timeout(*_args, **_kwargs):
        raise subprocess.TimeoutExpired(["osascript"], 3)

    monkeypatch.setattr(renderer.subprocess, "run", timeout)
    with pytest.raises(renderer.AutomationTimeout, match="Opening timed out after 3 seconds"):
        renderer._run_osascript("script", timeout=3, phase="Opening")


def test_open_script_declares_workbook_and_has_no_masking_error_handler(monkeypatch, tmp_path):
    captured = {}

    def run(script, **_kwargs):
        captured["script"] = script
        return "render.xlsx"

    monkeypatch.setattr(renderer, "_run_osascript", run)
    assert renderer._open_workbook(tmp_path / "book.xlsx") == "render.xlsx"
    assert "set openedWorkbookName" in captured["script"]
    assert "set openedWorkbook to open workbook" in captured["script"]
    assert "on error" not in captured["script"]


def test_copy_picture_does_not_activate_or_navigate(monkeypatch):
    captured = {}

    def run(script, **_kwargs):
        captured["script"] = script
        return ""

    monkeypatch.setattr(renderer, "_run_osascript", run)
    monkeypatch.setattr(renderer.time, "sleep", lambda _seconds: None)
    renderer._copy_range("render.xlsx", "Sheet1", "B4:D8")
    assert "copy picture targetRange" in captured["script"]
    assert "activate object" not in captured["script"]
    assert "goto reference" not in captured["script"]


def test_open_failure_restores_observed_security_and_preserves_error(monkeypatch, tmp_path):
    source = prepare(monkeypatch, tmp_path)
    changes = []
    original = renderer.RenderError("original open failure -50")
    monkeypatch.setattr(
        renderer,
        "read_excel_settings",
        lambda: renderer.ExcelSettings("msoAutomationSecurityLow", True),
    )
    monkeypatch.setattr(renderer, "_set_automation_security", changes.append)
    monkeypatch.setattr(renderer, "_open_workbook", lambda *_args, **_kwargs: (_ for _ in ()).throw(original))

    with pytest.raises(renderer.RenderError) as caught:
        with renderer.temporary_excel_workbook(source, log=lambda _: None):
            pass
    assert caught.value is original
    assert changes == ["msoAutomationSecurityForceDisable", "msoAutomationSecurityLow"]


def test_timeout_restores_settings_and_conditionally_cleans_temp(monkeypatch, tmp_path):
    source = prepare(monkeypatch, tmp_path)
    changes = []
    closed = []
    monkeypatch.setattr(
        renderer,
        "read_excel_settings",
        lambda: renderer.ExcelSettings("msoAutomationSecurityByUI", True),
    )
    monkeypatch.setattr(renderer, "_set_automation_security", changes.append)
    monkeypatch.setattr(
        renderer,
        "_open_workbook",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(renderer.AutomationTimeout("open timeout")),
    )
    monkeypatch.setattr(renderer, "check_excel_connection", lambda: "Microsoft Excel")
    monkeypatch.setattr(renderer, "_close_workbook_if_open", lambda name: closed.append(name) or True)

    with pytest.raises(renderer.AutomationTimeout, match="open timeout"):
        with renderer.temporary_excel_workbook(source, log=lambda _: None):
            pass
    assert changes == ["msoAutomationSecurityForceDisable", "msoAutomationSecurityByUI"]
    assert len(closed) == 1
    assert closed[0].startswith("render-") and closed[0].endswith(".xlsm")


def test_restore_failure_reports_primary_as_cause(monkeypatch, tmp_path):
    source = prepare(monkeypatch, tmp_path)
    primary = renderer.RenderError("open failed first")
    calls = 0
    monkeypatch.setattr(
        renderer,
        "read_excel_settings",
        lambda: renderer.ExcelSettings("msoAutomationSecurityLow", True),
    )

    def set_security(_value):
        nonlocal calls
        calls += 1
        if calls == 2:
            raise renderer.AutomationTimeout("restore timeout")

    monkeypatch.setattr(renderer, "_set_automation_security", set_security)
    monkeypatch.setattr(renderer, "_open_workbook", lambda *_a, **_k: (_ for _ in ()).throw(primary))
    with pytest.raises(renderer.AutomationUnrecoverableError, match="open failed first") as caught:
        with renderer.temporary_excel_workbook(source, log=lambda _: None):
            pass
    assert caught.value.__cause__ is primary
    assert "security restoration" in str(caught.value)


def test_interruption_restores_settings_and_checks_for_temp_workbook(monkeypatch, tmp_path):
    source = prepare(monkeypatch, tmp_path)
    changes = []
    closed = []
    monkeypatch.setattr(
        renderer,
        "read_excel_settings",
        lambda: renderer.ExcelSettings("msoAutomationSecurityByUI", True),
    )
    monkeypatch.setattr(renderer, "_set_automation_security", changes.append)
    monkeypatch.setattr(
        renderer,
        "_open_workbook",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(KeyboardInterrupt()),
    )
    monkeypatch.setattr(renderer, "check_excel_connection", lambda: "Microsoft Excel")
    monkeypatch.setattr(renderer, "_close_workbook_if_open", lambda name: closed.append(name) or False)
    with pytest.raises(KeyboardInterrupt):
        with renderer.temporary_excel_workbook(source, log=lambda _: None):
            pass
    assert changes == ["msoAutomationSecurityForceDisable", "msoAutomationSecurityByUI"]
    assert len(closed) == 1


def test_security_change_timeout_still_attempts_exact_restoration(monkeypatch, tmp_path):
    source = prepare(monkeypatch, tmp_path)
    changes = []
    monkeypatch.setattr(
        renderer,
        "read_excel_settings",
        lambda: renderer.ExcelSettings("msoAutomationSecurityLow", True),
    )

    def set_security(value):
        changes.append(value)
        if len(changes) == 1:
            raise renderer.AutomationTimeout("security change timeout")

    monkeypatch.setattr(renderer, "_set_automation_security", set_security)
    monkeypatch.setattr(renderer, "check_excel_connection", lambda: "Microsoft Excel")
    monkeypatch.setattr(renderer, "_close_workbook_if_open", lambda _name: False)
    with pytest.raises(renderer.AutomationTimeout, match="security change timeout"):
        with renderer.temporary_excel_workbook(source, log=lambda _: None):
            pass
    assert changes == ["msoAutomationSecurityForceDisable", "msoAutomationSecurityLow"]


def test_cleanup_closes_only_random_temporary_workbook(monkeypatch, tmp_path):
    source = prepare(monkeypatch, tmp_path).with_suffix(".xlsx")
    source.write_bytes(b"workbook")
    opened = []
    closed = []

    def open_workbook(path, **_kwargs):
        opened.append(path.name)
        return path.name

    monkeypatch.setattr(renderer, "_open_workbook", open_workbook)
    monkeypatch.setattr(renderer, "_close_workbook_if_open", lambda name: closed.append(name) or True)
    with renderer.temporary_excel_workbook(source, log=lambda _: None) as name:
        assert name != source.name
    assert closed == opened
    assert source.name not in closed


def test_render_uses_matching_open_source_without_temp_or_close(monkeypatch, tmp_path):
    source = tmp_path / "자료.xlsx"
    source.write_bytes(b"workbook")
    monkeypatch.setattr(
        renderer,
        "list_open_workbooks",
        lambda: [renderer.OpenWorkbook("자료.xlsx", source)],
    )
    monkeypatch.setattr(
        renderer,
        "temporary_excel_workbook",
        lambda *_a, **_k: (_ for _ in ()).throw(AssertionError("must not open temp")),
    )
    with renderer.render_workbook(source, log=lambda _: None) as workbook_name:
        assert workbook_name == "자료.xlsx"


def test_render_refuses_closed_source_when_user_workbook_is_open(monkeypatch, tmp_path):
    source = tmp_path / "closed.xlsx"
    source.write_bytes(b"workbook")
    user_book = tmp_path / "editing.xlsx"
    monkeypatch.setattr(
        renderer,
        "list_open_workbooks",
        lambda: [renderer.OpenWorkbook("editing.xlsx", user_book)],
    )
    with pytest.raises(renderer.RenderError, match="left the Excel UI unchanged"):
        with renderer.render_workbook(source, log=lambda _: None):
            pass


def test_temporary_open_is_blocked_when_any_workbook_is_open(monkeypatch, tmp_path):
    source = prepare(monkeypatch, tmp_path).with_suffix(".xlsx")
    source.write_bytes(b"workbook")
    monkeypatch.setattr(
        renderer,
        "list_open_workbooks",
        lambda: [renderer.OpenWorkbook("user.xlsx", tmp_path / "user.xlsx")],
    )
    monkeypatch.setattr(
        renderer,
        "_open_workbook",
        lambda *_a, **_k: (_ for _ in ()).throw(AssertionError("must not open")),
    )
    with pytest.raises(renderer.RenderError, match="active editing context"):
        with renderer.temporary_excel_workbook(source, log=lambda _: None):
            pass


def test_preflight_uses_active_cell_without_opening_temp(monkeypatch, tmp_path):
    app = tmp_path / "Microsoft Excel.app"
    app.mkdir()
    monkeypatch.setattr(renderer, "EXCEL_APP", app)
    monkeypatch.setattr(renderer, "check_excel_connection", lambda: "Microsoft Excel")
    monkeypatch.setattr(
        renderer,
        "read_excel_settings",
        lambda: renderer.ExcelSettings("msoAutomationSecurityLow", True),
    )
    monkeypatch.setattr(
        renderer,
        "list_open_workbooks",
        lambda: [renderer.OpenWorkbook("editing.xlsx", tmp_path / "editing.xlsx")],
    )
    monkeypatch.setattr(renderer, "_active_cell_target", lambda: ("editing.xlsx", "Sheet2", "C8"))
    monkeypatch.setattr(
        renderer,
        "temporary_excel_workbook",
        lambda *_a, **_k: (_ for _ in ()).throw(AssertionError("must not open temp")),
    )
    rendered = []

    class FakeRenderer:
        def __init__(self, log):
            pass

        def render_open_workbook(self, workbook_name, sheet_name, coordinate, destination):
            rendered.append((workbook_name, sheet_name, coordinate))
            destination.write_bytes(b"png")

    monkeypatch.setattr(renderer, "ExcelRenderer", FakeRenderer)
    renderer.automation_preflight(log=lambda _: None)
    assert rendered == [("editing.xlsx", "Sheet2", "C8")]

