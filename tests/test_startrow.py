"""Tests for the startrow option: placing the frame below free rows.

``startrow`` moves the frame, not the sheet. Everything derived from the frame's
position (header, data, table, freeze panes, formula columns, conditional formats,
validations) must move with it, and everything addressed by an absolute cell
reference (``cells``, ``merged_ranges``, ``row_heights``, charts) must not.
"""

from __future__ import annotations

import warnings
from pathlib import Path

import pandas as pd
import pytest
import xlsxturbo

from tests.helpers import HAS_OPENPYXL, active_ws, load_workbook

pytestmark = pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required for content verification")


def _frame() -> pd.DataFrame:
    """Two columns, three rows."""
    return pd.DataFrame({"Name": ["Alice", "Bob", "Carol"], "Score": [10, 50, 90]})


def _column(ws: object, letter: str, last_row: int) -> list[object]:
    """The values of one column from row 1 to ``last_row``."""
    return [ws[f"{letter}{row}"].value for row in range(1, last_row + 1)]  # type: ignore[index]


class TestPlacement:
    """Where the header and the data land."""

    def test_header_and_data_move_down(self, tmp_xlsx: str) -> None:
        """startrow=2 leaves rows 1-2 empty, header on row 3, data from row 4."""
        assert xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, startrow=2) == (4, 2)
        ws = active_ws(load_workbook(tmp_xlsx))
        assert _column(ws, "A", 6) == [None, None, "Name", "Alice", "Bob", "Carol"]
        assert _column(ws, "B", 6) == [None, None, "Score", 10, 50, 90]

    def test_without_header_the_data_starts_at_startrow(self, tmp_xlsx: str) -> None:
        """header=False puts the first data row on the startrow row itself."""
        assert xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, startrow=1, header=False) == (3, 2)
        ws = active_ws(load_workbook(tmp_xlsx))
        assert _column(ws, "A", 4) == [None, "Alice", "Bob", "Carol"]

    @pytest.mark.parametrize("value", [0, None])
    def test_zero_and_none_are_the_default(self, tmp_xlsx: str, value: int | None) -> None:
        """0 and an explicit None both write from A1."""
        xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, startrow=value)  # type: ignore[arg-type]
        assert active_ws(load_workbook(tmp_xlsx))["A1"].value == "Name"

    def test_polars_frame(self, tmp_xlsx: str) -> None:
        """The offset applies to a polars frame the same way."""
        pl = pytest.importorskip("polars")
        xlsxturbo.df_to_xlsx(pl.from_pandas(_frame()), tmp_xlsx, startrow=1)
        ws = active_ws(load_workbook(tmp_xlsx))
        assert _column(ws, "A", 3) == [None, "Name", "Alice"]

    def test_title_above_the_table(self, tmp_xlsx: str) -> None:
        """The use case: a merged title and a note above a styled table."""
        xlsxturbo.df_to_xlsx(
            _frame(),
            tmp_xlsx,
            startrow=2,
            table_style="Medium2",
            merged_ranges=[("A1:B1", "Q3 report", {"bold": True})],
            cells={"A2": "Scores by person"},
        )
        ws = active_ws(load_workbook(tmp_xlsx))
        assert _column(ws, "A", 4) == ["Q3 report", "Scores by person", "Name", "Alice"]
        assert "A1:B1" in {str(r) for r in ws.merged_cells.ranges}


class TestFeaturesMoveWithTheFrame:
    """Everything positioned relative to the frame follows the offset."""

    def test_table_covers_the_moved_frame(self, tmp_xlsx: str) -> None:
        """The table starts at the header row, not at A1."""
        xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, startrow=2, table_style="Medium2", table_name="Scores")
        ws = active_ws(load_workbook(tmp_xlsx))
        assert ws.tables["Scores"].ref == "A3:B6"

    def test_freeze_panes_below_the_header(self, tmp_xlsx: str) -> None:
        """The split sits under the moved header, so the rows above it freeze too."""
        xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, startrow=2, freeze_panes=True)
        assert active_ws(load_workbook(tmp_xlsx)).freeze_panes == "A4"

    def test_formula_columns(self, tmp_xlsx: str) -> None:
        """The formula header joins the moved header, and {row} names the real sheet row."""
        xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, startrow=2, formula_columns={"Double": "=B{row}*2"})
        ws = active_ws(load_workbook(tmp_xlsx))
        assert _column(ws, "C", 6) == [None, None, "Double", "=B4*2", "=B5*2", "=B6*2"]

    def test_formula_columns_without_header(self, tmp_xlsx: str) -> None:
        """With no header, the first formula lands on the startrow row."""
        xlsxturbo.df_to_xlsx(
            _frame(), tmp_xlsx, startrow=1, header=False, formula_columns={"Double": "=B{row}*2"}
        )
        ws = active_ws(load_workbook(tmp_xlsx))
        assert _column(ws, "C", 4) == [None, "=B2*2", "=B3*2", "=B4*2"]

    def test_conditional_format_range(self, tmp_xlsx: str) -> None:
        """A conditional format covers the moved data rows only."""
        xlsxturbo.df_to_xlsx(
            _frame(), tmp_xlsx, startrow=2, conditional_formats={"Score": {"type": "data_bar"}}
        )
        ws = active_ws(load_workbook(tmp_xlsx))
        assert [str(rng.sqref) for rng in ws.conditional_formatting] == ["B4:B6"]

    def test_validation_range(self, tmp_xlsx: str) -> None:
        """A validation covers the moved data rows only."""
        xlsxturbo.df_to_xlsx(
            _frame(),
            tmp_xlsx,
            startrow=2,
            validations={"Score": {"type": "whole_number", "min": 0, "max": 100}},
        )
        ws = active_ws(load_workbook(tmp_xlsx))
        assert [str(v.sqref) for v in ws.data_validations.dataValidation] == ["B4:B6"]


class TestAbsoluteReferencesStay:
    """Options addressed by a cell reference or row index do not move."""

    def test_row_heights_are_sheet_rows(self, tmp_xlsx: str) -> None:
        """row_heights={0: 30} sets the first sheet row, not the header row."""
        xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, startrow=2, row_heights={0: 30})
        ws = active_ws(load_workbook(tmp_xlsx))
        assert ws.row_dimensions[1].height == 30
        assert ws.row_dimensions[3].height is None


class TestMultipleSheets:
    """dfs_to_xlsx takes startrow as a default and as a per-sheet override."""

    def test_default_and_override(self, tmp_xlsx: str) -> None:
        """The workbook default applies unless a sheet sets its own."""
        result = xlsxturbo.dfs_to_xlsx(
            [(_frame(), "Default"), (_frame(), "Override", {"startrow": 0})],
            tmp_xlsx,
            startrow=3,
        )
        assert result == [(4, 2), (4, 2)]
        wb = load_workbook(tmp_xlsx)
        assert wb["Default"]["A4"].value == "Name"
        assert wb["Default"]["A1"].value is None
        assert wb["Override"]["A1"].value == "Name"


class TestConstantMemory:
    """The offset is applied while rows are written, so streaming mode keeps it."""

    def test_applies_without_a_warning(self, tmp_xlsx: str) -> None:
        """No RuntimeWarning names startrow, and the data lands below the offset."""
        with warnings.catch_warnings():
            warnings.simplefilter("error")
            xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, startrow=2, constant_memory=True)
        ws = active_ws(load_workbook(tmp_xlsx))
        assert _column(ws, "A", 4) == [None, None, "Name", "Alice"]


class TestValidation:
    """Bad values are refused inside the exception hierarchy, before any file exists."""

    @pytest.mark.parametrize("value", [True, 1.5, "2", [2]])
    def test_wrong_type(self, tmp_path: Path, value: object) -> None:
        """A bool, float, string or list is a ConfigurationTypeError naming startrow."""
        out = tmp_path / "out.xlsx"
        with pytest.raises(xlsxturbo.ConfigurationTypeError, match="startrow"):
            xlsxturbo.df_to_xlsx(_frame(), out, startrow=value)  # type: ignore[arg-type]
        assert not out.exists()

    @pytest.mark.parametrize("value", [-1, 1_048_576])
    def test_out_of_range(self, tmp_path: Path, value: int) -> None:
        """A negative row or one past Excel's last row is a ConfigurationError."""
        out = tmp_path / "out.xlsx"
        with pytest.raises(xlsxturbo.ConfigurationError, match="between 0 and 1048575"):
            xlsxturbo.df_to_xlsx(_frame(), out, startrow=value)
        assert not out.exists()

    def test_per_sheet_wrong_type(self, tmp_xlsx: str) -> None:
        """The per-sheet option is validated the same way."""
        with pytest.raises(xlsxturbo.ConfigurationTypeError, match="sheet option 'startrow'"):
            xlsxturbo.dfs_to_xlsx([(_frame(), "S", {"startrow": True})], tmp_xlsx)

    def test_frame_that_does_not_fit_below_the_offset(self, tmp_path: Path) -> None:
        """Header plus three rows need four rows; three left is refused before writing."""
        out = tmp_path / "out.xlsx"
        with pytest.raises(xlsxturbo.ConfigurationError, match="startrow=1048573 leaves room for 3 rows"):
            xlsxturbo.df_to_xlsx(_frame(), out, startrow=1_048_573)
        assert not out.exists()

    def test_frame_that_exactly_fits(self, tmp_xlsx: str) -> None:
        """The same frame one row higher fills the grid to its last row."""
        assert xlsxturbo.df_to_xlsx(_frame(), tmp_xlsx, startrow=1_048_572) == (4, 2)
