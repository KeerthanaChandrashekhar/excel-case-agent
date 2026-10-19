"""Business logic for cases: validation, ID generation, search, list, count."""

from __future__ import annotations

import logging
from dataclasses import dataclass
from datetime import datetime
from typing import Callable

from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

import config
from excel_handler import ExcelHandler

logger = logging.getLogger(__name__)


class CaseManagerError(Exception):
    """Base class for case-level errors."""


class ValidationError(CaseManagerError):
    """User input was invalid (empty explanation, bad Case ID, ...)."""


class DuplicateCaseIdError(CaseManagerError):
    """The generated Case ID already exists in the workbook."""


@dataclass(frozen=True)
class CaseRecord:
    """A single case as stored in the workbook."""

    case_id: str
    explanation: str
    created_at: str

    @property
    def number(self) -> int:
        """Numeric suffix of the Case ID (CASE-0042 -> 42)."""
        match = config.CASE_ID_PATTERN.match(self.case_id)
        return int(match.group(1)) if match else 0


def parse_case_number(raw: str) -> int | None:
    """Return the numeric suffix of a stored Case ID, or None if malformed."""
    match = config.CASE_ID_PATTERN.match((raw or "").strip())
    return int(match.group(1)) if match else None


def format_case_id(number: int) -> str:
    """Build a Case ID such as CASE-0004 from a number."""
    return f"{config.CASE_ID_PREFIX}{number:0{config.CASE_ID_MIN_DIGITS}d}"


class CaseManager:
    """Creates and retrieves cases. The workbook is always the source of truth."""

    def __init__(
        self,
        handler: ExcelHandler,
        clock: Callable[[], datetime] = datetime.now,
    ) -> None:
        self._handler = handler
        self._clock = clock

    @property
    def workbook_path(self):
        return self._handler.path

    # ----------------------------------------------------------------- reading
    def load_cases(self) -> list[CaseRecord]:
        """Read all valid cases from Excel, skipping malformed IDs safely."""
        cases: list[CaseRecord] = []
        seen: set[int] = set()
        for row in self._handler.read_rows():
            number = parse_case_number(row.case_id)
            if number is None:
                logger.warning(
                    "Skipping row %d: malformed Case ID %r", row.row_number, row.case_id
                )
                continue
            if number in seen:
                logger.warning(
                    "Row %d repeats Case ID %s", row.row_number, row.case_id.strip()
                )
            seen.add(number)
            cases.append(
                CaseRecord(row.case_id.strip().upper(), row.explanation, row.created_at)
            )
        return cases

    def list_cases(self) -> list[CaseRecord]:
        return self.load_cases()

    def count_cases(self) -> int:
        return len(self.load_cases())

    def get_case(self, case_id: str) -> CaseRecord | None:
        """Find a case by ID (case-insensitive). Returns None if not found."""
        match = config.CASE_LOOKUP_PATTERN.match((case_id or "").strip())
        if not match:
            raise ValidationError(
                f"'{case_id}' is not a valid Case ID. Expected something like CASE-0001."
            )
        wanted = int(match.group(1))
        for case in self.load_cases():
            if case.number == wanted:
                return case
        return None

    # ---------------------------------------------------------------- creating
    @staticmethod
    def validate_explanation(explanation: str) -> str:
        """Return a cleaned explanation or raise ValidationError."""
        if not isinstance(explanation, str):
            raise ValidationError("The explanation must be text.")
        cleaned = ILLEGAL_CHARACTERS_RE.sub("", explanation).strip()
        if not cleaned:
            raise ValidationError(
                "The explanation is empty. Describe the problem, for example: "
                "create case: Payment deducted twice."
            )
        if len(cleaned) > config.MAX_EXPLANATION_LENGTH:
            raise ValidationError(
                f"The explanation is too long ({len(cleaned)} characters). "
                f"The limit is {config.MAX_EXPLANATION_LENGTH}."
            )
        return cleaned

    def next_case_id(self, cases: list[CaseRecord] | None = None) -> str:
        """Highest existing numeric suffix + 1 (gaps are never reused)."""
        existing = self.load_cases() if cases is None else cases
        return format_case_id(max((c.number for c in existing), default=0) + 1)

    def create_case(self, explanation: str) -> CaseRecord:
        """Validate, generate an ID, save to Excel, and return the new record.

        Raises if the save fails, so callers can never report a false success.
        """
        text = self.validate_explanation(explanation)
        cases = self.load_cases()
        case_id = self.next_case_id(cases)
        if any(c.case_id == case_id for c in cases):
            raise DuplicateCaseIdError(f"{case_id} already exists; nothing was saved.")
        created = self._clock().replace(microsecond=0)
        self._handler.append_record(case_id, text, created)
        return CaseRecord(case_id, text, created.strftime(config.TIMESTAMP_FORMAT))
