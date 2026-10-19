import logging
from datetime import datetime

import pytest

import excel_handler
from case_manager import CaseManager, ValidationError
from excel_handler import ExcelHandler, WorkbookLockedError


def test_workbook_is_created_for_new_folder(handler):
    assert handler.path.exists()


def test_add_first_case(manager):
    case = manager.create_case("Unable to access company portal")
    assert case.case_id == "CASE-0001"
    assert case.explanation == "Unable to access company portal"
    assert case.created_at == "2026-01-15 09:30:00"


def test_sequential_ids(manager):
    ids = [manager.create_case(f"problem {i}").case_id for i in range(3)]
    assert ids == ["CASE-0001", "CASE-0002", "CASE-0003"]


def test_id_generation_continues_after_restart(handler, manager):
    manager.create_case("one")
    manager.create_case("two")
    reopened_handler = ExcelHandler(handler.folder)
    assert reopened_handler.initialize() is False
    reopened = CaseManager(reopened_handler)
    assert reopened.create_case("three").case_id == "CASE-0003"


def test_gaps_use_highest_id_plus_one(manager, seed_rows):
    seed_rows([("CASE-0001", "a", "t"), ("CASE-0003", "b", "t")])
    assert manager.create_case("c").case_id == "CASE-0004"


@pytest.mark.parametrize("bad", ["", "   ", "\n\t "])
def test_empty_explanation_rejected(manager, bad):
    with pytest.raises(ValidationError):
        manager.create_case(bad)
    assert manager.count_cases() == 0


def test_search_existing_case_is_case_insensitive(manager):
    manager.create_case("Application crashes during login")
    for query in ("CASE-0001", "case-0001", " Case-1 "):
        found = manager.get_case(query)
        assert found is not None
        assert found.explanation == "Application crashes during login"


def test_search_nonexistent_case(manager):
    manager.create_case("one")
    assert manager.get_case("CASE-0099") is None


def test_search_with_invalid_id_format(manager):
    with pytest.raises(ValidationError):
        manager.get_case("hello")


def test_count_cases(manager):
    assert manager.count_cases() == 0
    manager.create_case("a")
    manager.create_case("b")
    assert manager.count_cases() == 2


def test_list_cases(manager):
    manager.create_case("a")
    manager.create_case("b")
    assert [c.explanation for c in manager.list_cases()] == ["a", "b"]


def test_records_preserved_after_reload(handler, manager):
    manager.create_case("Ünïcödé record")
    fresh = CaseManager(ExcelHandler(handler.folder))
    cases = fresh.list_cases()
    assert len(cases) == 1
    assert cases[0].explanation == "Ünïcödé record"


def test_malformed_ids_are_skipped_with_warning(manager, seed_rows, caplog):
    seed_rows(
        [
            ("CASE-0001", "valid", "t"),
            ("garbage", "bad id", "t"),
            ("CASE-12", "too few digits", "t"),
            (None, "no id", "t"),
            ("CASE-0007", "valid too", "t"),
        ]
    )
    with caplog.at_level(logging.WARNING, logger="case_manager"):
        cases = manager.list_cases()
    assert [c.case_id for c in cases] == ["CASE-0001", "CASE-0007"]
    assert len([r for r in caplog.records if "malformed" in r.getMessage()]) == 3
    assert manager.next_case_id() == "CASE-0008"


def test_failed_save_is_not_reported_as_success(manager, monkeypatch):
    def deny(*_args, **_kwargs):
        raise PermissionError("locked")

    monkeypatch.setattr(excel_handler.os, "replace", deny)
    with pytest.raises(WorkbookLockedError):
        manager.create_case("will not be saved")
    monkeypatch.undo()
    assert manager.count_cases() == 0
    assert manager.create_case("retry works").case_id == "CASE-0001"


def test_illegal_control_characters_are_removed(manager):
    case = manager.create_case("bad\x00char\x07s")
    assert case.explanation == "badchars"


def test_very_long_explanation_is_rejected(manager):
    with pytest.raises(ValidationError):
        manager.create_case("x" * 40000)
