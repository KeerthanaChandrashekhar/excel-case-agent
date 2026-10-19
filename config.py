"""Application-wide constants for the Excel Case Management Agent."""

from __future__ import annotations

import re

APP_TITLE = "Excel Case Management Agent"
APP_DESCRIPTION = (
    "Create and retrieve support cases stored in an Excel workbook "
    "using plain-English commands."
)

# Workbook layout
WORKBOOK_NAME = "cases.xlsx"
WORKSHEET_NAME = "Cases"
HEADERS = ("Case ID", "Explanation", "Created At")
COLUMN_WIDTHS = {"A": 14, "B": 70, "C": 22}
HEADER_FILL_COLOR = "1F4E78"
HEADER_FONT_COLOR = "FFFFFF"

# Timestamps: Python format for display/reading, Excel format for the cell.
TIMESTAMP_FORMAT = "%Y-%m-%d %H:%M:%S"
EXCEL_TIMESTAMP_FORMAT = "yyyy-mm-dd hh:mm:ss"

# Case IDs
CASE_ID_PREFIX = "CASE-"
CASE_ID_MIN_DIGITS = 4
# Strict pattern for IDs stored in the workbook (CASE-0001, CASE-12345, ...).
CASE_ID_PATTERN = re.compile(r"^CASE-(\d{4,})$", re.IGNORECASE)
# Lenient pattern for IDs typed by the user (CASE-3, case-0003, CASE0003).
CASE_LOOKUP_PATTERN = re.compile(r"^CASE-?(\d+)$", re.IGNORECASE)

# Excel cells hold at most 32,767 characters; stay safely below that.
MAX_EXPLANATION_LENGTH = 32000

# Safe-save helpers
BACKUP_SUFFIX = ".bak"
TEMP_SUFFIX = ".tmp"

# Logging
LOG_DIR_NAME = "logs"
LOG_FILE_NAME = "app.log"
