from contextlib import contextmanager

import pytest
from openpyxl import Workbook

from excel2img import exporter
from excel2img.paths import is_eligible_workbook
from tests.test_detector import outline


class FakeRenderer:
    def render_open_workbook(self, workbook_name, sheet_name, coordinate, destination):
        destination.write_bytes(f"{sheet_name}:{coordinate}".encode())


class FailingRenderer:
    def render_open_workbook(self, workbook_name, sheet_name, coordinate, destination):
        raise RuntimeError("copy failed")


@contextmanager
def fake_workbook(_source, **_kwargs):
    yield "temporary.xlsx"


def make_workbook(path, rectangles, sheet_name="빠른 예습"):
    workbook = Workbook()
    ws = workbook.active
    ws.title = sheet_name
    for rectangle in rectangles:
        outline(ws, *rectangle)
    workbook.save(path)


def test_ignored_workbooks_and_unicode_names(tmp_path):
    assert not is_eligible_workbook(tmp_path / "!무시.xlsx")
    assert not is_eligible_workbook(tmp_path / "~$잠금.xlsx")
    assert is_eligible_workbook(tmp_path / "약물 표 (최종).xlsm")


def test_repeated_saves_replace_and_remove_stale_without_touching_manual_file(tmp_path, monkeypatch):
    source = tmp_path / "약물학.xlsx"
    monkeypatch.setattr(exporter, "temporary_excel_workbook", fake_workbook)
    make_workbook(source, [(1, 1, 2, 2), (5, 1, 6, 2)])
    first = exporter.process_workbook(source, renderer=FakeRenderer(), log=lambda _: None)
    assert [path.name for path in first] == ["빠른 예습_01.png", "빠른 예습_02.png"]
    output = tmp_path / "약물학_Img"
    manual = output / "manual.png"
    manual.write_bytes(b"mine")

    make_workbook(source, [(1, 1, 3, 3)])
    second = exporter.process_workbook(source, renderer=FakeRenderer(), log=lambda _: None)
    assert [path.name for path in second] == ["빠른 예습_01.png"]
    assert not (output / "빠른 예습_02.png").exists()
    assert manual.read_bytes() == b"mine"


def test_modified_generated_file_is_not_deleted_as_stale(tmp_path, monkeypatch):
    source = tmp_path / "notes.xlsx"
    monkeypatch.setattr(exporter, "temporary_excel_workbook", fake_workbook)
    make_workbook(source, [(1, 1, 2, 2)])
    generated = exporter.process_workbook(source, renderer=FakeRenderer(), log=lambda _: None)[0]
    generated.write_bytes(b"user replaced this image")
    make_workbook(source, [])
    exporter.process_workbook(source, renderer=FakeRenderer(), log=lambda _: None)
    assert generated.read_bytes() == b"user replaced this image"


def test_failed_render_removes_staging_file(tmp_path, monkeypatch):
    source = tmp_path / "failure.xlsx"
    monkeypatch.setattr(exporter, "temporary_excel_workbook", fake_workbook)
    make_workbook(source, [(1, 1, 2, 2)])
    with pytest.raises(RuntimeError, match="copy failed"):
        exporter.process_workbook(source, renderer=FailingRenderer(), log=lambda _: None)
    output = tmp_path / "failure_Img"
    assert list(output.glob(".render-*.png")) == []

