"""Shared test helpers for xlsxturbo integration tests."""

import os
import tempfile

import pandas as pd
import polars as pl
import pytest
import xlsxturbo

try:
    from openpyxl import load_workbook

    HAS_OPENPYXL = True
except ImportError:
    load_workbook = None
    HAS_OPENPYXL = False


def get_temp_path():
    """Get a temporary file path that's closed for Windows compatibility."""
    fd, path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    return path
