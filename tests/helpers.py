"""Shared test helpers for xlsxturbo integration tests."""

from __future__ import annotations

import os
import tempfile
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from openpyxl import load_workbook
    from openpyxl.workbook.workbook import Workbook
    from openpyxl.worksheet.worksheet import Worksheet

    HAS_OPENPYXL = True
else:
    try:
        from openpyxl import load_workbook

        HAS_OPENPYXL = True
    except ImportError:
        load_workbook = None
        HAS_OPENPYXL = False

__all__ = ["HAS_OPENPYXL", "active_ws", "get_temp_path", "load_workbook"]


def get_temp_path() -> str:
    """Return a temporary ``.xlsx`` file path with its handle closed.

    The handle is closed immediately so Windows allows the file to be
    reopened and rewritten by the library under test.
    """
    fd, path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    return path


def active_ws(wb: Workbook) -> Worksheet:
    """Return the active worksheet of ``wb``, asserting one exists.

    openpyxl types :attr:`Workbook.active` as ``Worksheet | None``; in these
    tests a freshly written workbook always has an active sheet, so this helper
    narrows the type for callers.
    """
    ws = wb.active
    assert ws is not None, "workbook has no active worksheet"
    return ws
