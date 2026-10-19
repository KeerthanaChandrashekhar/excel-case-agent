"""LLM controller tests use a scripted fake client: no network, no API key."""

import json
from types import SimpleNamespace

import pytest

import excel_handler
import llm_agent
from agent import CaseAgent
from llm_agent import LLMController, LLMUnavailableError


def text(t):
    return SimpleNamespace(type="text", text=t)


def tool(name, **args):
    return SimpleNamespace(type="tool_use", id=f"id_{name}", name=name, input=args)


def reply(*blocks, stop="tool_use"):
    return SimpleNamespace(stop_reason=stop, content=list(blocks))


class FakeClient:
    """Returns scripted responses in order and records every request."""

    def __init__(self, *responses):
        self._responses = list(responses)
        self.calls = []
        self.messages = self

    def create(self, **kwargs):
        self.calls.append(json.loads(json.dumps(kwargs["messages"], default=lambda o: "<obj>")))
        item = self._responses.pop(0)
        if isinstance(item, Exception):
            raise item
        return item


def last_tool_result(client):
    return client.calls[-1][-1]["content"][0]


def make_agent(manager, *responses):
    client = FakeClient(*responses)
    return CaseAgent(manager, llm=LLMController(client, model="test-model")), client


def test_llm_creates_case_through_tool(manager):
    agent, client = make_agent(
        manager,
        reply(tool("create_case", explanation="Customer cannot access the application")),
        reply(text("Created CASE-0001."), stop="end_turn"),
    )
    response = agent.handle("my customer can't get into the app, please log it")
    assert response.text == "Created CASE-0001."
    cases = manager.list_cases()
    assert [(c.case_id, c.explanation) for c in cases] == [
        ("CASE-0001", "Customer cannot access the application")
    ]
    result = last_tool_result(client)
    assert result["is_error"] is False
    assert json.loads(result["content"])["case_id"] == "CASE-0001"


def test_llm_list_find_and_count(manager):
    manager.create_case("first")
    manager.create_case("second")
    agent, client = make_agent(
        manager,
        reply(tool("count_cases")),
        reply(tool("find_case", case_id="case-2")),
        reply(tool("list_cases")),
        reply(text("done"), stop="end_turn"),
    )
    assert agent.handle("tell me everything").text == "done"
    assert json.loads(client.calls[1][-1]["content"][0]["content"]) == {"count": 2}
    found = json.loads(client.calls[2][-1]["content"][0]["content"])
    assert found["found"] and found["explanation"] == "second"
    listing = json.loads(client.calls[3][-1]["content"][0]["content"])
    assert listing["count"] == 2


def test_tool_error_is_returned_to_model_and_nothing_saved(manager):
    agent, client = make_agent(
        manager,
        reply(tool("create_case", explanation="   ")),
        reply(text("That was empty, nothing saved."), stop="end_turn"),
    )
    agent.handle("create an empty case")
    result = last_tool_result(client)
    assert result["is_error"] is True
    assert manager.count_cases() == 0


def test_failed_save_reaches_model_as_error(manager, monkeypatch):
    def deny(*_a, **_k):
        raise PermissionError("locked")

    monkeypatch.setattr(excel_handler.os, "replace", deny)
    agent, client = make_agent(
        manager,
        reply(tool("create_case", explanation="will fail")),
        reply(text("Could not save."), stop="end_turn"),
    )
    agent.handle("create case will fail")
    result = last_tool_result(client)
    assert result["is_error"] is True
    assert json.loads(result["content"])["saved"] is False
    monkeypatch.undo()
    assert manager.count_cases() == 0


def test_unknown_tool_is_an_error(manager):
    agent, client = make_agent(
        manager, reply(tool("delete_everything")), reply(text("no"), stop="end_turn")
    )
    agent.handle("delete everything")
    assert last_tool_result(client)["is_error"] is True


def test_api_failure_falls_back_to_rule_parser(manager):
    agent, _ = make_agent(manager, RuntimeError("network down"))
    response = agent.handle("create case: Payment deducted twice")
    assert "could not be reached" in response.text
    assert "CASE-0001" in response.text
    assert manager.count_cases() == 1


def test_too_many_steps_raises(manager):
    client = FakeClient(*[reply(tool("count_cases")) for _ in range(10)])
    with pytest.raises(LLMUnavailableError):
        LLMController(client, max_steps=3).run("loop forever", manager)


def test_llm_needs_a_folder_first():
    agent = CaseAgent(llm=LLMController(FakeClient()))
    response = agent.handle("create a case for printer failure")
    assert response.is_error and "Select Folder" in response.text


def test_mode_label_is_honest(manager, monkeypatch):
    assert "Rule-based" in CaseAgent(manager).mode_label
    assert "test-model" in make_agent(manager)[0].mode_label
    monkeypatch.delenv(llm_agent.API_KEY_ENV_VAR, raising=False)
    assert LLMController.from_environment() is None
