"""Tests for the formula_columns feature."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest
import xlsxturbo

from tests.helpers import HAS_OPENPYXL, active_ws, get_temp_path, load_workbook

pytestmark = pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required for content verification")


class TestFormulaColumns:
    """Tests for formula_columns feature (v0.9.0)."""

    def test_basic_formula(self) -> None:
        """Formula column appended after data columns."""
        df = pd.DataFrame({"price": [100, 200], "quantity": [5, 3]})
        path = get_temp_path()
        try:
            _rows, cols = xlsxturbo.df_to_xlsx(
                df, path, formula_columns={"Total": "=A{row}*B{row}"}
            )
            assert cols == 3  # price, quantity, Total
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                # Header should be "Total"
                assert ws["C1"].value == "Total"
                # Data rows should have formulas (openpyxl shows them as =formula)
                assert ws["C2"].value == "=A2*B2"
                assert ws["C3"].value == "=A3*B3"
                wb.close()
        finally:
            Path(path).unlink()

    def test_multiple_formula_columns(self) -> None:
        """Multiple formula columns in order."""
        df = pd.DataFrame({"price": [100], "qty": [5], "tax": [0.1]})
        path = get_temp_path()
        try:
            _rows, cols = xlsxturbo.df_to_xlsx(
                df,
                path,
                formula_columns={
                    "Subtotal": "=A{row}*B{row}",
                    "TaxAmt": "=D{row}*C{row}",
                },
            )
            assert cols == 5  # price, qty, tax, Subtotal, TaxAmt
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["D1"].value == "Subtotal"
                assert ws["E1"].value == "TaxAmt"
                assert ws["D2"].value == "=A2*B2"
                assert ws["E2"].value == "=D2*C2"
                wb.close()
        finally:
            Path(path).unlink()

    def test_formula_row_placeholder(self) -> None:
        """The {row} placeholder is correctly replaced per row."""
        df = pd.DataFrame({"A": [10, 20, 30]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df, path, formula_columns={"Double": "=A{row}*2"}
            )
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["B2"].value == "=A2*2"
                assert ws["B3"].value == "=A3*2"
                assert ws["B4"].value == "=A4*2"
                wb.close()
        finally:
            Path(path).unlink()

    def test_formula_columns_empty_dataframe(self) -> None:
        """An empty DataFrame writes cleanly, skipping the formula column when there are no data rows."""
        df = pd.DataFrame({"A": []})
        path = get_temp_path()
        try:
            _rows, cols = xlsxturbo.df_to_xlsx(
                df, path, formula_columns={"Total": "=A{row}*2"}
            )
            # No data rows -> the formula column is not appended.
            assert cols == 1
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["A1"].value == "A"  # header still written
                assert ws["B1"].value is None  # no formula column emitted
                wb.close()
        finally:
            Path(path).unlink()

    def test_formula_with_dfs_to_xlsx(self) -> None:
        """Formula columns work in multi-sheet mode."""
        df = pd.DataFrame({"A": [1, 2]})
        path = get_temp_path()
        try:
            xlsxturbo.dfs_to_xlsx(
                [(df, "Sheet1", {"formula_columns": {"Sum": "=A{row}+10"}})],
                path,
            )
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb["Sheet1"]
                assert ws["B1"].value == "Sum"
                assert ws["B2"].value == "=A2+10"
                wb.close()
        finally:
            Path(path).unlink()


class TestFormulaColumnsHeaderFalse:
    """Regression tests for formula_columns with header=False (v0.10.5 fix)."""

    def test_formula_columns_header_false(self) -> None:
        """Formula columns work correctly when header=False."""
        df = pd.DataFrame({"A": [10, 20], "B": [1, 2]})
        path = get_temp_path()
        try:
            _rows, cols = xlsxturbo.df_to_xlsx(
                df, path,
                header=False,
                formula_columns={"Sum": "=A{row}+B{row}"},
            )
            assert cols == 3  # 2 data + 1 formula
            wb = load_workbook(path)
            ws = active_ws(wb)
            # Row 1 should have data, not headers
            assert ws["A1"].value == 10
            # Formula column should be in C (0-indexed col 2)
            assert ws["C1"].value is not None
            wb.close()
        finally:
            Path(path).unlink()

    def test_formula_columns_header_true(self) -> None:
        """Formula columns still work correctly when header=True."""
        df = pd.DataFrame({"A": [10, 20], "B": [1, 2]})
        path = get_temp_path()
        try:
            _rows, cols = xlsxturbo.df_to_xlsx(
                df, path,
                header=True,
                formula_columns={"Sum": "=A{row}+B{row}"},
            )
            assert cols == 3
            wb = load_workbook(path)
            ws = active_ws(wb)
            # Row 1 should have headers
            assert ws["A1"].value == "A"
            assert ws["C1"].value == "Sum"
            # Row 2 should have data
            assert ws["A2"].value == 10
            wb.close()
        finally:
            Path(path).unlink()
