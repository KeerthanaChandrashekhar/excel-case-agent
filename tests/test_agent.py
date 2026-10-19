import pytest

import excel_handler
from agent import CaseAgent, parse_command
from case_manager import CaseManager


@pytest.mark.parametrize(
    "text, name, argument",
    [
        (
            "Create a case for a customer who cannot access the application.",
            "create",
            "A customer who cannot access the application.",
        ),
        (
            "create case: Payment deducted twice from customer account.",
            "create",
            "Payment deducted twice from customer account.",
        ),
        ("create case", "create", ""),
        ("Show all cases", "list", ""),
        ("List cases", "list", ""),
        ("Display existing cases", "list", ""),
        ("Find CASE-0003.", "search", "CASE-0003"),
        ("find case-3", "search", "CASE-0003"),
        ("How many cases have been created?", "count", ""),
        ("Help", "help", ""),
        ("make me a sandwich", "unknown", ""),
        ("   ", "empty", ""),
    ],
)
def test_parse_command(text, name, argument):
    command = parse_command(text)
    assert (command.name, command.argument) == (name, argument)


def test_no_folder_selected_asks_for_one():
    agent = CaseAgent()
    response = agent.handle("list cases")
    assert response.is_error and "Select Folder" in response.text
    assert not agent.handle("help").is_error


def test_unknown_command_shows_help():
    response = CaseAgent().handle("do a barrel roll")
    assert "don't recognise" in response.text and "create case" in response.text


def test_empty_explanation_is_rejected(manager):
    response = CaseAgent(manager).handle("create case:")
    assert response.is_error and "empty" in response.text
    assert manager.count_cases() == 0


def test_end_to_end(tmp_path):
    agent = CaseAgent()
    first = agent.select_folder(tmp_path)
    assert agent.is_ready and "Created a new" in first.text

    assert "CASE-0001" in agent.handle("create case: Payment deducted twice").text
    assert "CASE-0002" in agent.handle("Create a case for a user who cannot log in").text
    assert "There are 2 cases" in agent.handle("how many cases?").text
    listing = agent.handle("show all cases").text
    assert "CASE-0001" in listing and "CASE-0002" in listing
    assert "Payment deducted twice" in agent.handle("find case-0001").text
    assert "No case found" in agent.handle("find CASE-0042").text

    restarted = CaseAgent()
    assert "Loaded existing" in restarted.select_folder(tmp_path).text
    assert "CASE-0003" in restarted.handle("create case: third one").text


def test_failed_save_never_claims_success(manager, monkeypatch):
    def deny(*_args, **_kwargs):
        raise PermissionError("locked")

    monkeypatch.setattr(excel_handler.os, "replace", deny)
    response = CaseAgent(manager).handle("create case: this should fail")
    assert response.is_error
    assert "Case created" not in response.text
    monkeypatch.undo()
    assert manager.count_cases() == 0


def test_invalid_workbook_blocks_folder_selection(tmp_path):
    (tmp_path / "cases.xlsx").write_bytes(b"junk")
    agent = CaseAgent()
    response = agent.select_folder(tmp_path)
    assert response.is_error and not agent.is_ready
    assert (tmp_path / "cases.xlsx").read_bytes() == b"junk"
