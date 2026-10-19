"""Agent: routes messages to the LLM controller when available, otherwise to a
rule-based command parser. Both paths use the same CaseManager."""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass
from pathlib import Path

import config
from case_manager import CaseManager, CaseManagerError, CaseRecord
from excel_handler import ExcelHandler, ExcelHandlerError
from llm_agent import LLMController, LLMUnavailableError

logger = logging.getLogger(__name__)

WELCOME_TEXT = (
    "Welcome! I create and look up support cases stored in an Excel workbook.\n"
    "Click 'Select Folder' to choose where cases.xlsx lives, then type a command.\n"
    "Type 'help' to see what I can do."
)

HELP_TEXT = (
    "I understand these commands:\n"
    "  Create a case:   create case: Payment deducted twice from customer account.\n"
    "                   Create a case for a customer who cannot access the application.\n"
    "  List all cases:  show all cases  |  list cases  |  display existing cases\n"
    "  Find by ID:      find CASE-0003\n"
    "  Count cases:     how many cases have been created?\n"
    "  Help:            help"
)

NO_FOLDER_TEXT = (
    "Please select a folder first. Click 'Select Folder' so I know where "
    "to create or find cases.xlsx."
)

UNKNOWN_TEXT = "Sorry, I don't recognise that command.\n\n" + HELP_TEXT

HELP_RE = re.compile(r"^(help|\?|commands|what can you do)\s*[?!.]*$", re.IGNORECASE)
CREATE_RE = re.compile(
    r"^(?:please\s+)?(?:create|add|open|log|new|raise)\s+(?:an?\s+)?(?:new\s+)?case\b"
    r"(?:\s*[:\-\u2013\u2014]\s*|\s+(?:(?:for|about|regarding|concerning|of|on)\b\s*)?|\s*$)"
    r"(?P<exp>.*)$",
    re.IGNORECASE | re.DOTALL,
)
ID_SEARCH_RE = re.compile(r"\bcase[-_ ]?(\d+)\b", re.IGNORECASE)
CASE_WORD_RE = re.compile(r"\bcases?\b", re.IGNORECASE)
COUNT_RE = re.compile(r"\b(how many|count|number of|total)\b", re.IGNORECASE)
LIST_RE = re.compile(r"\b(list|show|display|view|see|all)\b", re.IGNORECASE)


@dataclass(frozen=True)
class Command:
    """A parsed user command."""

    name: str  # help | create | search | count | list | unknown | empty
    argument: str = ""


@dataclass(frozen=True)
class AgentResponse:
    """Text to show the user, plus whether it describes a problem."""

    text: str
    is_error: bool = False


def parse_command(text: str) -> Command:
    """Turn raw user text into a Command (pure function, easy to test)."""
    stripped = (text or "").strip()
    if not stripped:
        return Command("empty")
    if HELP_RE.match(stripped):
        return Command("help")

    create = CREATE_RE.match(stripped)
    if create:
        explanation = create.group("exp").strip()
        if explanation:
            explanation = explanation[0].upper() + explanation[1:]
        return Command("create", explanation)

    found_id = ID_SEARCH_RE.search(stripped)
    if found_id:
        return Command("search", f"{config.CASE_ID_PREFIX}{int(found_id.group(1)):04d}")

    has_case_word = bool(CASE_WORD_RE.search(stripped))
    if has_case_word and COUNT_RE.search(stripped):
        return Command("count")
    if (has_case_word and LIST_RE.search(stripped)) or stripped.lower() in ("list", "cases"):
        return Command("list")
    return Command("unknown")


def _format_case(case: CaseRecord) -> str:
    return f"{case.case_id}  (created {case.created_at})\n    {case.explanation}"


class CaseAgent:
    """Routes commands to the CaseManager and returns friendly responses.

    If an LLMController is supplied, the model decides which action to run;
    if it is missing or unavailable, the deterministic parser is used.
    """

    def __init__(
        self, manager: CaseManager | None = None, llm: LLMController | None = None
    ) -> None:
        self._manager = manager
        self._llm = llm

    @property
    def is_ready(self) -> bool:
        return self._manager is not None

    @property
    def mode_label(self) -> str:
        """Honest description of how commands are interpreted."""
        if self._llm is not None:
            return f"AI agent ({self._llm.model})"
        return "Rule-based (no API key set)"

    def select_folder(self, folder: str | Path) -> AgentResponse:
        """Create or load cases.xlsx in the folder and prepare the manager."""
        self._manager = None
        try:
            handler = ExcelHandler(folder)
            created = handler.initialize()
            manager = CaseManager(handler)
            count = manager.count_cases()
        except (ExcelHandlerError, CaseManagerError) as exc:
            logger.error("Folder selection failed for %s: %s", folder, exc)
            return AgentResponse(f"Error: {exc}", is_error=True)
        self._manager = manager
        if created:
            detail = f"Created a new {config.WORKBOOK_NAME} (no cases yet)."
        else:
            detail = f"Loaded existing {config.WORKBOOK_NAME} with {_plural(count)}."
        return AgentResponse(f"Folder: {handler.folder}\n{detail} The case file is ready.")

    def handle(self, text: str) -> AgentResponse:
        """Interpret one user message. Never raises."""
        command = parse_command(text)
        logger.info("Command: %s", command.name)
        try:
            return self._dispatch(command, text)
        except (ExcelHandlerError, CaseManagerError) as exc:
            logger.error("Command %s failed: %s", command.name, exc)
            return AgentResponse(f"Error: {exc}", is_error=True)
        except Exception:  # keep the app alive no matter what
            logger.exception("Unexpected error while handling %s", command.name)
            return AgentResponse(
                "Error: something unexpected went wrong. Nothing was reported as saved. "
                "See the log file for details.",
                is_error=True,
            )

    # ----------------------------------------------------------------- routing
    def _dispatch(self, command: Command, text: str) -> AgentResponse:
        if command.name == "help":
            return AgentResponse(HELP_TEXT)
        if command.name == "empty":
            return AgentResponse("Please type a command. Type 'help' to see examples.")
        if self._manager is None:
            if command.name == "unknown" and self._llm is None:
                return AgentResponse(UNKNOWN_TEXT, is_error=True)
            return AgentResponse(NO_FOLDER_TEXT, is_error=True)

        if self._llm is not None:
            try:
                return AgentResponse(self._llm.run(text.strip(), self._manager))
            except LLMUnavailableError as exc:
                logger.error("LLM unavailable, using rule-based parser: %s", exc)
                fallback = self._run_rules(command)
                note = (
                    "(The AI model could not be reached, so the built-in command "
                    f"parser was used. Reason: {exc})\n\n"
                )
                return AgentResponse(note + fallback.text, fallback.is_error)
        return self._run_rules(command)

    def _run_rules(self, command: Command) -> AgentResponse:
        """Deterministic command handling (also the LLM fallback)."""
        assert self._manager is not None
        if command.name == "unknown":
            return AgentResponse(UNKNOWN_TEXT, is_error=True)
        if command.name == "create":
            case = self._manager.create_case(command.argument)
            return AgentResponse(
                "Case created and saved to Excel.\n"
                f"Case ID: {case.case_id}\n"
                f"Explanation: {case.explanation}\n"
                f"Created at: {case.created_at}"
            )
        if command.name == "search":
            case = self._manager.get_case(command.argument)
            if case is None:
                return AgentResponse(f"No case found with ID {command.argument}.")
            return AgentResponse(
                f"Case ID: {case.case_id}\n"
                f"Explanation: {case.explanation}\n"
                f"Created at: {case.created_at}"
            )
        if command.name == "count":
            count = self._manager.count_cases()
            verb = "is" if count == 1 else "are"
            return AgentResponse(f"There {verb} {_plural(count)} in the workbook.")
        if command.name == "list":
            cases = self._manager.list_cases()
            if not cases:
                return AgentResponse("There are no cases yet. Try: create case: <explanation>")
            lines = "\n".join(_format_case(c) for c in cases)
            return AgentResponse(f"{_plural(len(cases))}:\n\n{lines}")
        return AgentResponse(UNKNOWN_TEXT, is_error=True)


def _plural(count: int) -> str:
    return f"{count} case" + ("" if count == 1 else "s")
