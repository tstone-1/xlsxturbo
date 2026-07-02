"""Tests for formatting features: column widths, tables, headers, borders, alignment, rich text."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pandas as pd
import polars as pl
import pytest
import xlsxturbo

from tests.helpers import HAS_OPENPYXL, active_ws, load_workbook

pytestmark = pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required for content verification")


class TestColumnWidthCap:
    """Tests for column_widths={'_all': value} feature."""

    def test_all_columns_capped(self, tmp_xlsx: str) -> None:
        """Set width for all columns via the '_all' key."""
        df = pd.DataFrame({"A": ["x" * 100], "B": ["y" * 100], "C": ["z" * 100]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_widths={"_all": 20})
        assert Path(tmp_xlsx).exists()
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        for col in ["A", "B", "C"]:
            assert ws.column_dimensions[col].width <= 21
        wb.close()

    def test_specific_overrides_all(self, tmp_xlsx: str) -> None:
        """Override '_all' with a specific column width."""
        df = pd.DataFrame({"A": ["x"], "B": ["y"], "C": ["z"]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_widths={0: 30, "_all": 10})
        assert Path(tmp_xlsx).exists()
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws.column_dimensions["A"].width > 25
        assert ws.column_dimensions["B"].width <= 11
        wb.close()

    def test_all_with_autofit(self, tmp_xlsx: str) -> None:
        """Use '_all' as a cap when combined with autofit."""
        df = pd.DataFrame({"Short": ["x"], "VeryLongColumnName": ["y" * 100]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, autofit=True, column_widths={"_all": 25})
        assert Path(tmp_xlsx).exists()
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        # Long column should be capped at ~25
        assert ws.column_dimensions["B"].width <= 26
        wb.close()

    def test_dfs_to_xlsx_per_sheet_all(self, tmp_xlsx: str) -> None:
        """Apply per-sheet '_all' override in dfs_to_xlsx."""
        df1 = pd.DataFrame({"A": ["x" * 50]})
        df2 = pd.DataFrame({"B": ["y" * 50]})
        xlsxturbo.dfs_to_xlsx(
            [
                (df1, "Sheet1", {"column_widths": {"_all": 20}}),
                (df2, "Sheet2", {"column_widths": {"_all": 40}}),
            ],
            tmp_xlsx,
        )
        assert Path(tmp_xlsx).exists()
        wb = load_workbook(tmp_xlsx)
        # Sheet1 should have narrower column than Sheet2
        w1 = wb["Sheet1"].column_dimensions["A"].width or 0
        w2 = wb["Sheet2"].column_dimensions["A"].width or 0
        assert w1 <= 21, f"Sheet1 col A width {w1} should be <= 21"
        assert w2 > 21, f"Sheet2 col A width {w2} should be > 21"
        wb.close()

    def test_autofit_with_all_cap(self, tmp_xlsx: str) -> None:
        """Cap autofit widths with _all instead of overriding."""
        df = pd.DataFrame({
            "Short": ["ab", "cd"],
            "VeryLong": ["A" * 100, "B" * 80],
            "Medium": ["hello world", "test data"],
        })
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, autofit=True, column_widths={"_all": 25})
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        w_short = ws.column_dimensions["A"].width or 0
        w_long = ws.column_dimensions["B"].width or 0
        w_med = ws.column_dimensions["C"].width or 0
        # Short column should be narrow (content-fitted, not inflated to 25)
        assert w_short < 15, f"Short col width {w_short} should be < 15"
        # VeryLong column should be capped at ~25
        assert w_long <= 26, f"VeryLong col width {w_long} should be <= 26"
        # Medium column should be content-fitted (< 25)
        assert w_med < 20, f"Medium col width {w_med} should be < 20"
        wb.close()

    def test_autofit_with_all_cap_polars(self, tmp_xlsx: str) -> None:
        """Apply autofit + _all cap with polars DataFrames."""
        df = pl.DataFrame({"Short": ["x"], "Long": ["A" * 80]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, autofit=True, column_widths={"_all": 20})
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        w_short = ws.column_dimensions["A"].width or 0
        w_long = ws.column_dimensions["B"].width or 0
        assert w_short < 15, f"Short col width {w_short} should be < 15"
        assert w_long <= 21, f"Long col width {w_long} should be <= 21"
        wb.close()

    def test_empty_column_widths_suppresses_autofit_per_sheet(self, tmp_xlsx: str) -> None:
        """A per-sheet column_widths={} combined with autofit=True suppresses autofit for that sheet.

        This is a pre-existing interplay, not a new behavior change: the mere
        presence of a `column_widths` dict (even an empty one, with no '_all'
        key) selects the plain apply_column_widths branch instead of calling
        `worksheet.autofit()`, so an empty per-sheet override silently turns
        autofit off for that sheet.
        """
        df = pd.DataFrame({"VeryLongColumnName": ["y" * 100]})
        xlsxturbo.dfs_to_xlsx(
            [(df, "S1", {"column_widths": {}})],
            tmp_xlsx,
            autofit=True,
        )
        wb = load_workbook(tmp_xlsx)
        width = wb["S1"].column_dimensions["A"].width
        # Default Excel width (~8.43), not autofit-expanded to fit the long content.
        assert width is None or width < 15
        wb.close()

    def test_column_widths_negative_key_raises(self, tmp_xlsx: str) -> None:
        """A negative column_widths key raises with a clear message."""
        df = pd.DataFrame({"A": [1], "B": [2]})
        with pytest.raises(ValueError, match="must be a non-negative column index"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_widths={-1: 20.0})

    def test_column_widths_key_exceeds_max_column_raises(self, tmp_xlsx: str) -> None:
        """A column_widths key beyond Excel's maximum column index (16383) raises."""
        df = pd.DataFrame({"A": [1], "B": [2]})
        with pytest.raises(ValueError, match="exceeds Excel's maximum column index"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_widths={16384: 20.0})

    def test_column_widths_non_integer_key_raises(self, tmp_xlsx: str) -> None:
        """A non-integer, non-'_all' column_widths key raises TypeError."""
        df = pd.DataFrame({"A": [1], "B": [2]})
        with pytest.raises(TypeError, match="must be an integer column index or the string '_all'"):
            # Intentionally invalid: a float key is neither an int index nor '_all'.
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_widths={1.5: 20.0})  # type: ignore[dict-item]

    def test_column_widths_key_beyond_data_range_is_applied(self, tmp_xlsx: str) -> None:
        """A column_widths key beyond the DataFrame's column count is now applied, not ignored."""
        df = pd.DataFrame({"A": [1], "B": [2]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_widths={5: 30.0})
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        # Column index 5 (0-based) is column F, beyond the 2-column data range.
        assert ws.column_dimensions["F"].width is not None
        assert ws.column_dimensions["F"].width > 25
        wb.close()

    def test_column_widths_string_key_exceeds_max_column_raises(self, tmp_xlsx: str) -> None:
        """A string column_widths key must pass the same max-column-index validation as an int key."""
        df = pd.DataFrame({"A": [1], "B": [2]})
        with pytest.raises(ValueError, match="exceeds Excel's maximum column index"):
            # Intentionally invalid: "20000" is beyond Excel's max column index (16383).
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_widths={"20000": 5.0})  # type: ignore[dict-item]

    def test_column_widths_non_numeric_string_key_raises(self, tmp_xlsx: str) -> None:
        """A non-numeric, non-'_all' string column_widths key raises, not silently ignored."""
        df = pd.DataFrame({"A": [1], "B": [2]})
        with pytest.raises(TypeError, match="must be an integer column index or the string '_all'"):
            # Intentionally invalid: "banana" is neither a numeric string nor '_all'.
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_widths={"banana": 5.0})  # type: ignore[dict-item]

    def test_column_widths_negative_string_key_raises(self, tmp_xlsx: str) -> None:
        """A negative string column_widths key must pass the same validation as an int key."""
        df = pd.DataFrame({"A": [1], "B": [2]})
        with pytest.raises(ValueError, match="must be a non-negative column index"):
            # Intentionally invalid: "-5" is a negative column index.
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_widths={"-5": 5.0})  # type: ignore[dict-item]

    def test_column_widths_valid_string_key_beyond_data_range_is_applied(self, tmp_xlsx: str) -> None:
        """A valid numeric-string column_widths key beyond the DataFrame's column count is applied."""
        df = pd.DataFrame({"A": [1], "B": [2]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_widths={"5": 30.0})  # type: ignore[dict-item]
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        # Column index 5 (0-based) is column F, beyond the 2-column data range.
        assert ws.column_dimensions["F"].width is not None
        assert ws.column_dimensions["F"].width > 25
        wb.close()


class TestTableName:
    """Tests for table_name parameter."""

    def test_explicit_table_name(self, tmp_xlsx: str) -> None:
        """Apply an explicit table name."""
        df = pd.DataFrame({"A": [1, 2, 3]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, table_style="Medium2", table_name="MyTable")
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        tables = list(ws.tables.keys())
        assert "MyTable" in tables
        wb.close()

    def test_table_name_sanitization(self, tmp_xlsx: str) -> None:
        """Sanitize invalid characters in table names."""
        df = pd.DataFrame({"A": [1]})
        xlsxturbo.df_to_xlsx(
            df, tmp_xlsx, table_style="Medium2", table_name="My Table-2024!"
        )
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        tables = list(ws.tables.keys())
        assert len(tables) == 1
        assert "_" in tables[0]  # Some characters replaced with underscore
        wb.close()

    def test_table_name_starts_with_digit(self, tmp_xlsx: str) -> None:
        """Prefix table names starting with a digit with an underscore."""
        df = pd.DataFrame({"A": [1]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, table_style="Medium2", table_name="123Data")
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        tables = list(ws.tables.keys())
        assert tables[0].startswith("_")
        wb.close()

    def test_per_sheet_table_names(self, tmp_xlsx: str) -> None:
        """Apply different table names per sheet."""
        df1 = pd.DataFrame({"A": [1]})
        df2 = pd.DataFrame({"B": [2]})
        xlsxturbo.dfs_to_xlsx(
            [
                (df1, "Sheet1", {"table_style": "Medium2", "table_name": "Table1"}),
                (df2, "Sheet2", {"table_style": "Medium2", "table_name": "Table2"}),
            ],
            tmp_xlsx,
        )
        wb = load_workbook(tmp_xlsx)
        assert "Table1" in wb["Sheet1"].tables
        assert "Table2" in wb["Sheet2"].tables
        wb.close()

    def test_no_table_name_without_table_style(self, tmp_xlsx: str) -> None:
        """Ignore table_name when table_style is None."""
        df = pd.DataFrame({"A": [1]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, table_name="Ignored")
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert len(ws.tables) == 0
        wb.close()


class TestHeaderFormat:
    """Tests for header_format parameter."""

    def test_bold_header(self, tmp_xlsx: str) -> None:
        """Apply a bold header."""
        df = pd.DataFrame({"A": [1], "B": [2]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, header_format={"bold": True})
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A1"].font.bold is True
        assert ws["B1"].font.bold is True
        wb.close()

    def test_header_background_color(self, tmp_xlsx: str) -> None:
        """Apply a background color to the header."""
        df = pd.DataFrame({"A": [1]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, header_format={"bg_color": "#4F81BD"})
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        # openpyxl uses ARGB format
        assert ws["A1"].fill.fgColor.rgb == "FF4F81BD"
        wb.close()

    def test_header_format_hex_color_with_sign_raises(self, tmp_xlsx: str) -> None:
        """A hex color with a sign character (e.g. '#+12345') is rejected, not parsed as a number."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(ValueError, match="Invalid hex color"):
            # Intentionally invalid: '+' is not a hex digit, but
            # u32::from_str_radix would otherwise accept it as a sign.
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, header_format={"bg_color": "#+12345"})

    def test_header_font_color(self, tmp_xlsx: str) -> None:
        """Apply a font color to the header."""
        df = pd.DataFrame({"A": [1]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, header_format={"font_color": "white"})
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A1"].font.color.rgb == "FFFFFFFF"
        wb.close()

    def test_combined_header_format(self, tmp_xlsx: str) -> None:
        """Combine multiple header format options."""
        df = pd.DataFrame({"A": [1]})
        xlsxturbo.df_to_xlsx(
            df,
            tmp_xlsx,
            header_format={"bold": True, "bg_color": "#4F81BD", "font_color": "white"},
        )
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A1"].font.bold is True
        assert ws["A1"].fill.fgColor.rgb == "FF4F81BD"
        assert ws["A1"].font.color.rgb == "FFFFFFFF"
        wb.close()

    def test_header_format_no_header(self, tmp_xlsx: str) -> None:
        """Ignore header_format when header=False."""
        df = pd.DataFrame({"A": [1]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, header=False, header_format={"bold": True})
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        # First row should be data, not header
        assert ws["A1"].value == 1
        wb.close()

    def test_per_sheet_header_format(self, tmp_xlsx: str) -> None:
        """Apply different header formats per sheet."""
        df1 = pd.DataFrame({"A": [1]})
        df2 = pd.DataFrame({"B": [2]})
        xlsxturbo.dfs_to_xlsx(
            [
                (df1, "Sheet1", {"header_format": {"bold": True}}),
                (df2, "Sheet2", {"header_format": {"bg_color": "#FF0000"}}),
            ],
            tmp_xlsx,
        )
        wb = load_workbook(tmp_xlsx)
        assert wb["Sheet1"]["A1"].font.bold is True
        assert wb["Sheet2"]["A1"].fill.fgColor.rgb == "FFFF0000"
        wb.close()


class TestRichText:
    """Tests for rich text feature (v0.10.0)."""

    def test_rich_text_bold(self, tmp_xlsx: str) -> None:
        """Apply rich text with bold formatting."""
        df = pd.DataFrame({"A": [1]})
        xlsxturbo.df_to_xlsx(
            df, tmp_xlsx, rich_text={"A1": [("Bold", {"bold": True}), " normal"]}
        )
        # openpyxl flattens rich text, so inspect the runs in sharedStrings.xml:
        # the two segments must be distinct runs and the first must be bold.
        with zipfile.ZipFile(tmp_xlsx) as zf:
            shared = zf.read("xl/sharedStrings.xml").decode("utf-8")
            assert "Bold" in shared
            assert "normal" in shared
            assert "<b/>" in shared

    def test_rich_text_mixed_formats(self, tmp_xlsx: str) -> None:
        """Apply rich text with multiple format segments."""
        df = pd.DataFrame({"A": [1]})
        xlsxturbo.df_to_xlsx(
            df,
            tmp_xlsx,
            rich_text={
                "B2": [
                    ("Red text", {"font_color": "red"}),
                    " and ",
                    ("blue text", {"font_color": "blue", "italic": True}),
                ]
            },
        )
        with zipfile.ZipFile(tmp_xlsx) as zf:
            shared = zf.read("xl/sharedStrings.xml").decode("utf-8")
            assert "Red text" in shared
            assert "blue text" in shared
            # Red (FF0000) and blue (0000FF) runs and an italic run must be present.
            assert "FF0000" in shared.upper()
            assert "0000FF" in shared.upper()
            assert "<i/>" in shared

    def test_rich_text_plain_segments(self, tmp_xlsx: str) -> None:
        """Apply rich text with plain string segments."""
        df = pd.DataFrame({"A": [1]})
        xlsxturbo.df_to_xlsx(
            df, tmp_xlsx, rich_text={"A3": ["Plain text ", ("bold", {"bold": True}), " more plain"]}
        )
        assert Path(tmp_xlsx).exists()


class TestRowHeights:
    """Tests for row_heights parameter (v0.4.0)."""

    def test_basic_row_heights(self, tmp_xlsx: str) -> None:
        """Set specific row heights."""
        df = pd.DataFrame({"A": [1, 2, 3]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, row_heights={0: 30, 2: 40})
        assert Path(tmp_xlsx).exists()
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        # openpyxl is 1-indexed; Excel may round heights slightly
        assert abs(ws.row_dimensions[1].height - 30) < 1
        assert abs(ws.row_dimensions[3].height - 40) < 1
        # Rows without explicit height should not have customHeight
        assert ws.row_dimensions[2].customHeight is False
        wb.close()

    def test_row_heights_with_dfs_to_xlsx(self, tmp_xlsx: str) -> None:
        """Apply row heights per-sheet."""
        df = pd.DataFrame({"A": [1, 2]})
        xlsxturbo.dfs_to_xlsx(
            [(df, "Sheet1", {"row_heights": {0: 25}})],
            tmp_xlsx,
        )
        assert Path(tmp_xlsx).exists()
        wb = load_workbook(tmp_xlsx)
        ws = wb["Sheet1"]
        assert abs(ws.row_dimensions[1].height - 25) < 1
        wb.close()

    def test_row_heights_ignored_in_constant_memory(self, tmp_xlsx: str) -> None:
        """Silently ignore row heights in constant memory mode."""
        df = pd.DataFrame({"A": [1]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, constant_memory=True, row_heights={0: 50})
        assert Path(tmp_xlsx).exists()
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        # Row height should NOT be set (constant memory ignores it)
        # Default height is ~15, so it should not be 50
        assert ws.row_dimensions[1].height != 50 or ws.row_dimensions[1].height is None
        wb.close()


class TestBorderStyles:
    """Tests for per-side border styles (v0.12.0)."""

    def test_border_bool_backward_compat(self, tmp_xlsx: str) -> None:
        """Apply a thin border on all sides with border=True."""
        df = pd.DataFrame({"A": [1, 2], "B": [3, 4]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_formats={
            "A": {"border": True}
        })
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        cell = ws["A2"]
        assert cell.border.left.style == "thin"
        assert cell.border.right.style == "thin"
        assert cell.border.top.style == "thin"
        assert cell.border.bottom.style == "thin"
        wb.close()

    def test_border_string_all_sides(self, tmp_xlsx: str) -> None:
        """Apply a thick border on all 4 sides with border='thick'."""
        df = pd.DataFrame({"A": [1, 2], "B": [3, 4]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_formats={
            "A": {"border": "thick"}
        })
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        cell = ws["A2"]
        assert cell.border.left.style == "thick"
        assert cell.border.right.style == "thick"
        assert cell.border.top.style == "thick"
        assert cell.border.bottom.style == "thick"
        wb.close()

    def test_border_right_only(self, tmp_xlsx: str) -> None:
        """Apply a thick right border only with border_right='thick'."""
        df = pd.DataFrame({"A": [1, 2], "B": [3, 4]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_formats={
            "A": {"border_right": "thick"}
        })
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        cell = ws["A2"]
        assert cell.border.right.style == "thick"
        assert cell.border.left.style is None
        assert cell.border.top.style is None
        assert cell.border.bottom.style is None
        wb.close()

    def test_border_left_and_right(self, tmp_xlsx: str) -> None:
        """Apply mixed per-side borders."""
        df = pd.DataFrame({"A": [1], "B": [2]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_formats={
            "A": {"border_left": "medium", "border_right": "thick"}
        })
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        cell = ws["A2"]
        assert cell.border.left.style == "medium"
        assert cell.border.right.style == "thick"
        assert cell.border.top.style is None
        assert cell.border.bottom.style is None
        wb.close()

    def test_border_all_four_sides_individually(self, tmp_xlsx: str) -> None:
        """Set all four sides independently."""
        df = pd.DataFrame({"A": [1]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_formats={
            "A": {
                "border_left": "thin",
                "border_right": "thick",
                "border_top": "medium",
                "border_bottom": "dashed",
            }
        })
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        cell = ws["A2"]
        assert cell.border.left.style == "thin"
        assert cell.border.right.style == "thick"
        assert cell.border.top.style == "medium"
        assert cell.border.bottom.style == "dashed"
        wb.close()

    def test_border_string_thin(self, tmp_xlsx: str) -> None:
        """Treat border='thin' as equivalent to border=True."""
        df = pd.DataFrame({"A": [1]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_formats={
            "A": {"border": "thin"}
        })
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        cell = ws["A2"]
        assert cell.border.left.style == "thin"
        assert cell.border.right.style == "thin"
        wb.close()

    def test_border_with_wildcard_pattern(self, tmp_xlsx: str) -> None:
        """Apply per-side borders with wildcard column matching."""
        df = pd.DataFrame({"price_usd": [10], "price_eur": [9], "name": ["x"]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_formats={
            "price_*": {"border_right": "thick"}
        })
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A2"].border.right.style == "thick"
        assert ws["B2"].border.right.style == "thick"
        assert ws["C2"].border.right.style is None
        wb.close()

    def test_border_color(self, tmp_xlsx: str) -> None:
        """Set color for all borders via border_color."""
        df = pd.DataFrame({"A": [1]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_formats={
            "A": {"border": "thin", "border_color": "red"}
        })
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        cell = ws["A2"]
        assert cell.border.left.style == "thin"
        assert cell.border.left.color is not None
        assert cell.border.left.color.rgb.endswith("FF0000")
        wb.close()

    def test_invalid_border_style_raises(self, tmp_xlsx: str) -> None:
        """Raise ValueError for an invalid border style string."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(ValueError, match="Unknown border style"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_formats={
                "A": {"border": "invalid_style"}
            })

    def test_border_with_other_formats(self, tmp_xlsx: str) -> None:
        """Combine border styles with other column format options."""
        df = pd.DataFrame({"A": [1.5]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_formats={
            "A": {"border_right": "thick", "bold": True, "num_format": "0.00"}
        })
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        cell = ws["A2"]
        assert cell.border.right.style == "thick"
        assert cell.font.bold
        assert cell.number_format == "0.00"
        wb.close()

    def test_border_per_side_overrides_all(self, tmp_xlsx: str) -> None:
        """Override all-sides border with a per-side border for that side."""
        df = pd.DataFrame({"A": [1]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_formats={
            "A": {"border": "thin", "border_right": "thick"}
        })
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        cell = ws["A2"]
        assert cell.border.left.style == "thin"
        assert cell.border.right.style == "thick"
        assert cell.border.top.style == "thin"
        assert cell.border.bottom.style == "thin"
        wb.close()

    def test_border_medium_style(self, tmp_xlsx: str) -> None:
        """Apply the medium border style."""
        df = pd.DataFrame({"A": [1]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_formats={
            "A": {"border": "medium"}
        })
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A2"].border.left.style == "medium"
        wb.close()

    def test_border_double_style(self, tmp_xlsx: str) -> None:
        """Apply the double border style."""
        df = pd.DataFrame({"A": [1]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_formats={
            "A": {"border_bottom": "double"}
        })
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A2"].border.bottom.style == "double"
        wb.close()

    def test_border_with_polars(self, tmp_xlsx: str) -> None:
        """Apply border styles with polars DataFrames."""
        df = pl.DataFrame({"A": [1, 2]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_formats={
            "A": {"border_right": "thick"}
        })
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A2"].border.right.style == "thick"
        wb.close()

    def test_border_right_bool_treated_as_thin(self, tmp_xlsx: str) -> None:
        """Apply a thin right border for border_right=True (bool)."""
        df = pd.DataFrame({"A": [1]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_formats={
            "A": {"border_right": True}
        })
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        cell = ws["A2"]
        assert cell.border.right.style == "thin"
        assert cell.border.left.style is None
        wb.close()

    def test_header_format_border(self, tmp_xlsx: str) -> None:
        """Support border styles in header_format."""
        df = pd.DataFrame({"A": [1, 2], "B": [3, 4]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, header_format={
            "bold": True, "border": "thick"
        })
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        header = ws["A1"]
        assert header.border.left.style == "thick"
        assert header.border.right.style == "thick"
        wb.close()

    def test_header_format_border_right_only(self, tmp_xlsx: str) -> None:
        """Support per-side borders in header_format."""
        df = pd.DataFrame({"A": [1], "B": [2]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, header_format={
            "border_bottom": "medium"
        })
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        header = ws["A1"]
        assert header.border.bottom.style == "medium"
        assert header.border.top.style is None
        wb.close()

    def test_border_dfs_to_xlsx_per_sheet(self, tmp_xlsx: str) -> None:
        """Apply per-side borders with dfs_to_xlsx per-sheet overrides."""
        df1 = pd.DataFrame({"A": [1]})
        df2 = pd.DataFrame({"A": [2]})
        xlsxturbo.dfs_to_xlsx([
            (df1, "S1"),
            (df2, "S2", {"column_formats": {"A": {"border_right": "thick"}}})
        ], tmp_xlsx, column_formats={"A": {"border": "thin"}})
        wb = load_workbook(tmp_xlsx)
        assert wb["S1"]["A2"].border.left.style == "thin"
        assert wb["S2"]["A2"].border.right.style == "thick"
        wb.close()


class TestTextAlignment:
    """Tests for text alignment (v0.12.0)."""

    def test_column_format_horizontal_center(self, tmp_xlsx: str) -> None:
        """Center cell text with align_horizontal='center'."""
        df = pd.DataFrame({"A": ["hello", "world"]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_formats={
            "A": {"align_horizontal": "center"}
        })
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A2"].alignment.horizontal == "center"
        wb.close()

    def test_column_format_horizontal_right(self, tmp_xlsx: str) -> None:
        """Right-align cell text with align_horizontal='right'."""
        df = pd.DataFrame({"A": ["hello"]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_formats={
            "A": {"align_horizontal": "right"}
        })
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A2"].alignment.horizontal == "right"
        wb.close()

    def test_column_format_vertical_top(self, tmp_xlsx: str) -> None:
        """Top-align cell text with align_vertical='top'."""
        df = pd.DataFrame({"A": ["hello"]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_formats={
            "A": {"align_vertical": "top"}
        })
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A2"].alignment.vertical == "top"
        wb.close()

    def test_column_format_vertical_center(self, tmp_xlsx: str) -> None:
        """Vertically center cell text with align_vertical='center'."""
        df = pd.DataFrame({"A": ["hello"]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_formats={
            "A": {"align_vertical": "center"}
        })
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A2"].alignment.vertical == "center"
        wb.close()

    def test_column_format_wrap_text(self, tmp_xlsx: str) -> None:
        """Enable text wrapping with wrap_text=True."""
        df = pd.DataFrame({"A": ["hello world this is a long text"]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_formats={
            "A": {"wrap_text": True}
        })
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A2"].alignment.wrapText is True
        wb.close()

    def test_column_format_combined_alignment(self, tmp_xlsx: str) -> None:
        """Combine horizontal, vertical, and wrap_text alignment."""
        df = pd.DataFrame({"A": ["hello"]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_formats={
            "A": {"align_horizontal": "center", "align_vertical": "top", "wrap_text": True}
        })
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        cell = ws["A2"]
        assert cell.alignment.horizontal == "center"
        assert cell.alignment.vertical == "top"
        assert cell.alignment.wrapText is True
        wb.close()

    def test_header_format_alignment(self, tmp_xlsx: str) -> None:
        """Support alignment in header_format."""
        df = pd.DataFrame({"A": [1], "B": [2]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, header_format={
            "bold": True, "align_horizontal": "center", "wrap_text": True
        })
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        header = ws["A1"]
        assert header.alignment.horizontal == "center"
        assert header.alignment.wrapText is True
        wb.close()

    def test_merged_range_alignment(self, tmp_xlsx: str) -> None:
        """Support alignment in merged_ranges format."""
        df = pd.DataFrame({"A": [1], "B": [2], "C": [3]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, merged_ranges=[
            ("A1:C1", "Title", {"bold": True, "align_horizontal": "left"})
        ])
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A1"].alignment.horizontal == "left"
        wb.close()

    def test_merged_range_default_center(self, tmp_xlsx: str) -> None:
        """Auto-center merged_ranges without explicit format."""
        df = pd.DataFrame({"A": [1], "B": [2]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, merged_ranges=[
            ("A1:B1", "Title")
        ])
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A1"].alignment.horizontal == "center"
        wb.close()

    def test_cells_alignment(self, tmp_xlsx: str) -> None:
        """Support alignment options in the cells parameter."""
        df = pd.DataFrame({"A": [1]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, cells={
            "C1": {"value": "Header", "align_horizontal": "center", "wrap_text": True}
        })
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        cell = ws["C1"]
        assert cell.value == "Header"
        assert cell.alignment.horizontal == "center"
        assert cell.alignment.wrapText is True
        wb.close()

    def test_alignment_with_wildcard(self, tmp_xlsx: str) -> None:
        """Apply alignment with wildcard column patterns."""
        df = pd.DataFrame({"desc_a": ["x"], "desc_b": ["y"], "num": [1]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_formats={
            "desc_*": {"align_horizontal": "left", "wrap_text": True}
        })
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A2"].alignment.horizontal == "left"
        assert ws["A2"].alignment.wrapText is True
        assert ws["B2"].alignment.horizontal == "left"
        assert ws["C2"].alignment.horizontal is None
        wb.close()

    def test_alignment_with_border(self, tmp_xlsx: str) -> None:
        """Combine alignment with border styles."""
        df = pd.DataFrame({"A": ["hello"]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_formats={
            "A": {"align_horizontal": "center", "border": "thin"}
        })
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        cell = ws["A2"]
        assert cell.alignment.horizontal == "center"
        assert cell.border.left.style == "thin"
        wb.close()

    def test_invalid_horizontal_alignment_raises(self, tmp_xlsx: str) -> None:
        """Raise ValueError for an invalid horizontal alignment."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(ValueError, match="Unknown horizontal alignment"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_formats={
                "A": {"align_horizontal": "middle"}
            })

    def test_invalid_vertical_alignment_raises(self, tmp_xlsx: str) -> None:
        """Raise ValueError for an invalid vertical alignment."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(ValueError, match="Unknown vertical alignment"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_formats={
                "A": {"align_vertical": "left"}
            })

    def test_alignment_with_polars(self, tmp_xlsx: str) -> None:
        """Apply alignment with polars DataFrames."""
        df = pl.DataFrame({"A": ["hello"]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_formats={
            "A": {"align_horizontal": "center"}
        })
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A2"].alignment.horizontal == "center"
        wb.close()
