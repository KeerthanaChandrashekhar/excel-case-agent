"""Shared fixtures. Every test uses its own temporary folder."""

from datetime import datetime

import pytest
from openpyxl import load_workbook

import config
from case_manager import CaseManager
from excel_handler import ExcelHandler


@pytest.fixture
def handler(tmp_path):
    h = ExcelHandler(tmp_path)
    h.initialize()
    return h


@pytest.fixture
def manager(handler):
    return CaseManager(handler, clock=lambda: datetime(2026, 1, 15, 9, 30, 0, 123456))


@pytest.fixture
def seed_rows(handler):
    """Write raw rows straight into the workbook (simulates manual edits)."""

    def _seed(rows):
        wb = load_workbook(handler.path)
        ws = wb[config.WORKSHEET_NAME]
        for row in rows:
            ws.append(list(row))
        wb.save(handler.path)
        wb.close()

    return _seed
