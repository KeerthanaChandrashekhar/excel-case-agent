"""Optional LLM controller: the model decides which Excel case action to run.

The model never touches Excel directly and never invents Case IDs. It can only
call four tools (create/list/find/count). Each tool runs through CaseManager,
so IDs, validation, safe saving and error handling stay in plain Python code.
Enabled only when ANTHROPIC_API_KEY is set; otherwise the app uses the
rule-based parser in agent.py.
"""

from __future__ import annotations

import json
import logging
import os
from typing import Any

from case_manager import CaseManager, CaseManagerError
from excel_handler import ExcelHandlerError

logger = logging.getLogger(__name__)

DEFAULT_MODEL = "claude-sonnet-5-5"
MODEL_ENV_VAR = "CASE_AGENT_MODEL"
API_KEY_ENV_VAR = "ANTHROPIC_API_KEY"
MAX_STEPS = 6

SYSTEM_PROMPT = (
    "You are a case management agent. Cases live in an Excel workbook that you "
    "control only through the provided tools: create_case, list_cases, find_case "
    "and count_cases. Rules: (1) Never invent Case IDs, timestamps or case data; "
    "use only what tool results return. (2) Only say a case was created if "
    "create_case returned success. If a tool returns an error, tell the user "
    "plainly what went wrong and that nothing was saved. (3) For create requests, "
    "pass the problem description as the explanation, keeping the user's meaning. "
    "(4) If the request is not about creating, listing, finding or counting cases, "
    "briefly say what you can do. (5) Keep replies short and clear."
)

TOOLS: list[dict[str, Any]] = [
    {
        "name": "create_case",
        "description": "Create a new case with the given explanation and save it to Excel. "
        "The Case ID and timestamp are generated automatically.",
        "input_schema": {
            "type": "object",
            "properties": {
                "explanation": {"type": "string", "description": "What the case is about."}
            },
            "required": ["explanation"],
        },
    },
    {
        "name": "list_cases",
        "description": "List every case (Case ID, explanation, creation time).",
        "input_schema": {"type": "object", "properties": {}},
    },
    {
        "name": "find_case",
        "description": "Look up one case by Case ID, e.g. CASE-0003.",
        "input_schema": {
            "type": "object",
            "properties": {"case_id": {"type": "string", "description": "e.g. CASE-0003"}},
            "required": ["case_id"],
        },
    },
    {
        "name": "count_cases",
        "description": "Return the total number of cases.",
        "input_schema": {"type": "object", "properties": {}},
    },
]


class LLMUnavailableError(Exception):
    """The model could not be reached or did not finish; use the fallback."""


class LLMController:
    """Runs a tool-use loop between the model and the CaseManager."""

    def __init__(self, client: Any, model: str = DEFAULT_MODEL, max_steps: int = MAX_STEPS) -> None:
        self._client = client
        self.model = model
        self._max_steps = max_steps

    @classmethod
    def from_environment(cls) -> "LLMController | None":
        """Build a controller if an API key and the SDK are available, else None."""
        api_key = os.environ.get(API_KEY_ENV_VAR, "").strip()
        if not api_key:
            return None
        try:
            import anthropic
        except ImportError:
            logger.warning("anthropic package not installed; using rule-based agent")
            return None
        model = os.environ.get(MODEL_ENV_VAR, "").strip() or DEFAULT_MODEL
        return cls(anthropic.Anthropic(api_key=api_key), model=model)

    def run(self, user_text: str, manager: CaseManager) -> str:
        """Let the model handle one user message and return its final reply."""
        messages: list[dict[str, Any]] = [{"role": "user", "content": user_text}]
        for _ in range(self._max_steps):
            try:
                response = self._client.messages.create(
                    model=self.model,
                    max_tokens=1024,
                    system=SYSTEM_PROMPT,
                    tools=TOOLS,
                    messages=messages,
                )
            except Exception as exc:
                logger.exception("LLM request failed")
                raise LLMUnavailableError(f"{type(exc).__name__}: {exc}") from exc

            if response.stop_reason != "tool_use":
                text = "".join(b.text for b in response.content if b.type == "text").strip()
                if not text:
                    raise LLMUnavailableError("The model returned an empty reply.")
                return text

            messages.append({"role": "assistant", "content": response.content})
            results = []
            for block in response.content:
                if block.type != "tool_use":
                    continue
                content, is_error = self._execute_tool(block.name, block.input, manager)
                results.append(
                    {
                        "type": "tool_result",
                        "tool_use_id": block.id,
                        "content": content,
                        "is_error": is_error,
                    }
                )
            messages.append({"role": "user", "content": results})
        raise LLMUnavailableError("The model took too many steps without finishing.")

    @staticmethod
    def _execute_tool(name: str, args: dict[str, Any], manager: CaseManager) -> tuple[str, bool]:
        """Run one tool. Returns (JSON text, is_error). Never raises."""
        logger.info("LLM tool call: %s", name)
        try:
            if name == "create_case":
                case = manager.create_case(str(args.get("explanation", "")))
                return json.dumps({"saved": True, **case.__dict__}), False
            if name == "list_cases":
                cases = manager.list_cases()
                return json.dumps({"count": len(cases), "cases": [c.__dict__ for c in cases]}), False
            if name == "find_case":
                case = manager.get_case(str(args.get("case_id", "")))
                if case is None:
                    return json.dumps({"found": False}), False
                return json.dumps({"found": True, **case.__dict__}), False
            if name == "count_cases":
                return json.dumps({"count": manager.count_cases()}), False
            return json.dumps({"error": f"Unknown tool '{name}'."}), True
        except (CaseManagerError, ExcelHandlerError) as exc:
            logger.error("Tool %s failed: %s", name, exc)
            return json.dumps({"error": str(exc), "saved": False}), True
        except Exception as exc:
            logger.exception("Unexpected tool failure in %s", name)
            return json.dumps({"error": f"Unexpected error: {exc}", "saved": False}), True
