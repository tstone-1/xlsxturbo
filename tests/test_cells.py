"""Tests for arbitrary cell writes and per-cell formatting."""

from __future__ import annotations

import pandas as pd
import polars as pl
import pytest
import xlsxturbo

from tests.helpers import HAS_OPENPYXL, active_ws, load_workbook

pytestmark = pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required for content verification")


class TestCells:
    """Tests for arbitrary cell writes (v0.11.0)."""

    def test_simple_string_cell(self, tmp_xlsx: str) -> None:
        """Write a string to a specific cell."""
        df = pd.DataFrame({"a": [1, 2]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, cells={"C1": "hello"})
        wb = load_workbook(tmp_xlsx)
        assert active_ws(wb)["C1"].value == "hello"
        wb.close()

    def test_numeric_cells(self, tmp_xlsx: str) -> None:
        """Write int and float to cells."""
        df = pd.DataFrame({"a": [1]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, cells={"B5": 42, "C5": 3.14})
        wb = load_workbook(tmp_xlsx)
        assert active_ws(wb)["B5"].value == 42
        assert abs(active_ws(wb)["C5"].value - 3.14) < 0.001
        wb.close()

    def test_bool_cell(self, tmp_xlsx: str) -> None:
        """Write a boolean to a cell."""
        df = pd.DataFrame({"a": [1]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, cells={"B2": True})
        wb = load_workbook(tmp_xlsx)
        assert active_ws(wb)["B2"].value is True
        wb.close()

    def test_cell_with_num_format(self, tmp_xlsx: str) -> None:
        """Dict-style cell with num_format preserves text format."""
        df = pd.DataFrame({"a": [1]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, cells={
            "D6": {"value": "934728173849", "num_format": "@"}
        })
        wb = load_workbook(tmp_xlsx)
        cell = active_ws(wb)["D6"]
        assert cell.value == "934728173849"
        assert cell.number_format == "@"
        wb.close()

    def test_cell_overwrites_dataframe_data(self, tmp_xlsx: str) -> None:
        """Cells parameter overwrites existing DataFrame values."""
        df = pd.DataFrame({"a": ["original"]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, cells={"A2": "overwritten"})
        wb = load_workbook(tmp_xlsx)
        assert active_ws(wb)["A2"].value == "overwritten"
        wb.close()

    def test_cell_dict_missing_value_key(self, tmp_xlsx: str) -> None:
        """Dict-style cell without 'value' key raises ValueError."""
        df = pd.DataFrame({"a": [1]})
        with pytest.raises(ValueError, match="missing 'value' key"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, cells={"A1": {"num_format": "@"}})

    def test_cell_dict_unknown_key_raises(self, tmp_xlsx: str) -> None:
        """A stray key like 'bold' (a format concern, not a cell key) is rejected, not silently dropped."""
        df = pd.DataFrame({"a": [1]})
        with pytest.raises(ValueError, match="unknown option 'bold'"):
            # Intentionally invalid: 'bold' is not a recognized cells dict key.
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, cells={"B1": {"value": "x", "bold": True}})  # type: ignore[typeddict-unknown-key]

    def test_invalid_cell_ref(self, tmp_xlsx: str) -> None:
        """Invalid cell reference raises ValueError."""
        df = pd.DataFrame({"a": [1]})
        with pytest.raises(ValueError):  # noqa: PT011
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, cells={"ZZZZ1": "x"})

    def test_cells_per_sheet_override(self, tmp_xlsx: str) -> None:
        """Per-sheet cells override global cells in dfs_to_xlsx."""
        df1 = pd.DataFrame({"a": [1]})
        df2 = pd.DataFrame({"b": [2]})
        xlsxturbo.dfs_to_xlsx(
            [(df1, "S1"), (df2, "S2", {"cells": {"C1": "per-sheet"}})],
            tmp_xlsx, cells={"C1": "global"})
        wb = load_workbook(tmp_xlsx)
        assert wb["S1"]["C1"].value == "global"
        assert wb["S2"]["C1"].value == "per-sheet"
        wb.close()

    def test_num_format_wrong_type_raises(self, tmp_xlsx: str) -> None:
        """Non-string num_format raises TypeError."""
        df = pd.DataFrame({"a": [1]})
        with pytest.raises(TypeError):
            # num_format must be str; int is intentionally invalid to trigger TypeError
            xlsxturbo.df_to_xlsx(df, tmp_xlsx,
                cells={"A1": {"value": "x", "num_format": 123}})  # type: ignore[dict-item]

    def test_cells_with_polars(self, tmp_xlsx: str) -> None:
        """Cells work with polars DataFrames."""
        df = pl.DataFrame({"a": [1, 2]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, cells={"C1": "extra"})
        wb = load_workbook(tmp_xlsx)
        assert active_ws(wb)["C1"].value == "extra"
        wb.close()


class TestCellsPerSheet:
    """Tests for cells with per-sheet SheetOptions in dfs_to_xlsx (item 6)."""

    def test_cells_per_sheet_3tuple(self, tmp_xlsx: str) -> None:
        """Per-sheet cells override via 3-tuple SheetOptions."""
        df1 = pd.DataFrame({"a": [1, 2]})
        df2 = pd.DataFrame({"b": [3, 4]})
        xlsxturbo.dfs_to_xlsx([
            (df1, "Sheet1", {"cells": {"C1": "note1", "C2": 100}}),
            (df2, "Sheet2", {"cells": {"C1": "note2"}}),
        ], tmp_xlsx)
        wb = load_workbook(tmp_xlsx)
        assert wb["Sheet1"]["C1"].value == "note1"
        assert wb["Sheet1"]["C2"].value == 100
        assert wb["Sheet2"]["C1"].value == "note2"
        assert wb["Sheet2"]["C2"].value is None
        wb.close()

    def test_cells_per_sheet_overrides_global(self, tmp_xlsx: str) -> None:
        """Per-sheet cells replace (not merge with) global cells."""
        df1 = pd.DataFrame({"a": [1]})
        df2 = pd.DataFrame({"a": [2]})
        xlsxturbo.dfs_to_xlsx([
            (df1, "S1"),
            (df2, "S2", {"cells": {"D1": "override"}}),
        ], tmp_xlsx, cells={"C1": "global"})
        wb = load_workbook(tmp_xlsx)
        # S1 gets global cells
        assert wb["S1"]["C1"].value == "global"
        # S2 gets per-sheet cells (override replaces global)
        assert wb["S2"]["D1"].value == "override"
        assert wb["S2"]["C1"].value is None
        wb.close()

    def test_cells_per_sheet_with_num_format(self, tmp_xlsx: str) -> None:
        """Per-sheet cells with num_format via SheetOptions."""
        df = pd.DataFrame({"a": [1]})
        xlsxturbo.dfs_to_xlsx([
            (df, "S1", {"cells": {
                "C1": {"value": "12345", "num_format": "@"}
            }}),
        ], tmp_xlsx)
        wb = load_workbook(tmp_xlsx)
        cell = wb["S1"]["C1"]
        assert str(cell.value) == "12345"
        assert cell.number_format == "@"
        wb.close()

    def test_cells_per_sheet_with_alignment(self, tmp_xlsx: str) -> None:
        """Per-sheet cells with alignment via SheetOptions."""
        df = pd.DataFrame({"a": [1]})
        xlsxturbo.dfs_to_xlsx([
            (df, "S1", {"cells": {
                "C1": {"value": "centered", "align_horizontal": "center"}
            }}),
        ], tmp_xlsx)
        wb = load_workbook(tmp_xlsx)
        cell = wb["S1"]["C1"]
        assert cell.value == "centered"
        assert cell.alignment.horizontal == "center"
        wb.close()


class TestCellsFormatting:
    """Tests for cells with formatting options beyond num_format (item 7)."""

    def test_cells_with_horizontal_alignment(self, tmp_xlsx: str) -> None:
        """Write cells with align_horizontal."""
        df = pd.DataFrame({"a": [1]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, cells={
            "C1": {"value": "right-aligned", "align_horizontal": "right"}
        })
        wb = load_workbook(tmp_xlsx)
        cell = active_ws(wb)["C1"]
        assert cell.value == "right-aligned"
        assert cell.alignment.horizontal == "right"
        wb.close()

    def test_cells_with_vertical_alignment(self, tmp_xlsx: str) -> None:
        """Write cells with align_vertical."""
        df = pd.DataFrame({"a": [1]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, cells={
            "C1": {"value": "top", "align_vertical": "top"}
        })
        wb = load_workbook(tmp_xlsx)
        cell = active_ws(wb)["C1"]
        assert cell.value == "top"
        assert cell.alignment.vertical == "top"
        wb.close()

    def test_cells_with_wrap_text(self, tmp_xlsx: str) -> None:
        """Write cells with wrap_text."""
        df = pd.DataFrame({"a": [1]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, cells={
            "C1": {"value": "long text here", "wrap_text": True}
        })
        wb = load_workbook(tmp_xlsx)
        cell = active_ws(wb)["C1"]
        assert cell.value == "long text here"
        assert cell.alignment.wrapText is True
        wb.close()

    def test_cells_with_combined_formatting(self, tmp_xlsx: str) -> None:
        """Write cells with num_format + alignment + wrap_text together."""
        df = pd.DataFrame({"a": [1]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, cells={
            "C1": {
                "value": 0.15,
                "num_format": "0.00%",
                "align_horizontal": "center",
                "align_vertical": "top",
                "wrap_text": True
            }
        })
        wb = load_workbook(tmp_xlsx)
        cell = active_ws(wb)["C1"]
        assert cell.number_format == "0.00%"
        assert cell.alignment.horizontal == "center"
        assert cell.alignment.vertical == "top"
        assert cell.alignment.wrapText is True
        wb.close()

    def test_cells_formatting_with_polars(self, tmp_xlsx: str) -> None:
        """Cells formatting works with polars DataFrames."""
        df = pl.DataFrame({"a": [1]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, cells={
            "C1": {"value": "test", "align_horizontal": "center"}
        })
        wb = load_workbook(tmp_xlsx)
        cell = active_ws(wb)["C1"]
        assert cell.value == "test"
        assert cell.alignment.horizontal == "center"
        wb.close()
