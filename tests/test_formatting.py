"""Tests for formatting features: column widths, tables, headers, borders, alignment, rich text."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pandas as pd
import polars as pl
import pytest
import xlsxturbo

from tests.helpers import HAS_OPENPYXL, active_ws, get_temp_path, load_workbook

pytestmark = pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required for content verification")


class TestColumnWidthCap:
    """Tests for column_widths={'_all': value} feature."""

    def test_all_columns_capped(self) -> None:
        """Set width for all columns via the '_all' key."""
        df = pd.DataFrame({"A": ["x" * 100], "B": ["y" * 100], "C": ["z" * 100]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_widths={"_all": 20})
            assert Path(path).exists()
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                for col in ["A", "B", "C"]:
                    assert ws.column_dimensions[col].width <= 21
                wb.close()
        finally:
            Path(path).unlink()

    def test_specific_overrides_all(self) -> None:
        """Override '_all' with a specific column width."""
        df = pd.DataFrame({"A": ["x"], "B": ["y"], "C": ["z"]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_widths={0: 30, "_all": 10})
            assert Path(path).exists()
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws.column_dimensions["A"].width > 25
                assert ws.column_dimensions["B"].width <= 11
                wb.close()
        finally:
            Path(path).unlink()

    def test_all_with_autofit(self) -> None:
        """Use '_all' as a cap when combined with autofit."""
        df = pd.DataFrame({"Short": ["x"], "VeryLongColumnName": ["y" * 100]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, autofit=True, column_widths={"_all": 25})
            assert Path(path).exists()
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                # Long column should be capped at ~25
                assert ws.column_dimensions["B"].width <= 26
                wb.close()
        finally:
            Path(path).unlink()

    def test_dfs_to_xlsx_per_sheet_all(self) -> None:
        """Apply per-sheet '_all' override in dfs_to_xlsx."""
        df1 = pd.DataFrame({"A": ["x" * 50]})
        df2 = pd.DataFrame({"B": ["y" * 50]})
        path = get_temp_path()
        try:
            xlsxturbo.dfs_to_xlsx(
                [
                    (df1, "Sheet1", {"column_widths": {"_all": 20}}),
                    (df2, "Sheet2", {"column_widths": {"_all": 40}}),
                ],
                path,
            )
            assert Path(path).exists()
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                # Sheet1 should have narrower column than Sheet2
                w1 = wb["Sheet1"].column_dimensions["A"].width or 0
                w2 = wb["Sheet2"].column_dimensions["A"].width or 0
                assert w1 <= 21, f"Sheet1 col A width {w1} should be <= 21"
                assert w2 > 21, f"Sheet2 col A width {w2} should be > 21"
                wb.close()
        finally:
            Path(path).unlink()

    def test_autofit_with_all_cap(self) -> None:
        """Cap autofit widths with _all instead of overriding."""
        df = pd.DataFrame({
            "Short": ["ab", "cd"],
            "VeryLong": ["A" * 100, "B" * 80],
            "Medium": ["hello world", "test data"],
        })
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, autofit=True, column_widths={"_all": 25})
            if HAS_OPENPYXL:
                wb = load_workbook(path)
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
        finally:
            Path(path).unlink()

    def test_autofit_with_all_cap_polars(self) -> None:
        """Apply autofit + _all cap with polars DataFrames."""
        df = pl.DataFrame({"Short": ["x"], "Long": ["A" * 80]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, autofit=True, column_widths={"_all": 20})
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                w_short = ws.column_dimensions["A"].width or 0
                w_long = ws.column_dimensions["B"].width or 0
                assert w_short < 15, f"Short col width {w_short} should be < 15"
                assert w_long <= 21, f"Long col width {w_long} should be <= 21"
                wb.close()
        finally:
            Path(path).unlink()


class TestTableName:
    """Tests for table_name parameter."""

    def test_explicit_table_name(self) -> None:
        """Apply an explicit table name."""
        df = pd.DataFrame({"A": [1, 2, 3]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, table_style="Medium2", table_name="MyTable")
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                tables = list(ws.tables.keys())
                assert "MyTable" in tables
                wb.close()
        finally:
            Path(path).unlink()

    def test_table_name_sanitization(self) -> None:
        """Sanitize invalid characters in table names."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df, path, table_style="Medium2", table_name="My Table-2024!"
            )
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                tables = list(ws.tables.keys())
                assert len(tables) == 1
                assert "_" in tables[0]  # Some characters replaced with underscore
                wb.close()
        finally:
            Path(path).unlink()

    def test_table_name_starts_with_digit(self) -> None:
        """Prefix table names starting with a digit with an underscore."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, table_style="Medium2", table_name="123Data")
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                tables = list(ws.tables.keys())
                assert tables[0].startswith("_")
                wb.close()
        finally:
            Path(path).unlink()

    def test_per_sheet_table_names(self) -> None:
        """Apply different table names per sheet."""
        df1 = pd.DataFrame({"A": [1]})
        df2 = pd.DataFrame({"B": [2]})
        path = get_temp_path()
        try:
            xlsxturbo.dfs_to_xlsx(
                [
                    (df1, "Sheet1", {"table_style": "Medium2", "table_name": "Table1"}),
                    (df2, "Sheet2", {"table_style": "Medium2", "table_name": "Table2"}),
                ],
                path,
            )
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                assert "Table1" in wb["Sheet1"].tables
                assert "Table2" in wb["Sheet2"].tables
                wb.close()
        finally:
            Path(path).unlink()

    def test_no_table_name_without_table_style(self) -> None:
        """Ignore table_name when table_style is None."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, table_name="Ignored")
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert len(ws.tables) == 0
                wb.close()
        finally:
            Path(path).unlink()


class TestHeaderFormat:
    """Tests for header_format parameter."""

    def test_bold_header(self) -> None:
        """Apply a bold header."""
        df = pd.DataFrame({"A": [1], "B": [2]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, header_format={"bold": True})
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["A1"].font.bold is True
                assert ws["B1"].font.bold is True
                wb.close()
        finally:
            Path(path).unlink()

    def test_header_background_color(self) -> None:
        """Apply a background color to the header."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, header_format={"bg_color": "#4F81BD"})
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                # openpyxl uses ARGB format
                assert ws["A1"].fill.fgColor.rgb == "FF4F81BD"
                wb.close()
        finally:
            Path(path).unlink()

    def test_header_font_color(self) -> None:
        """Apply a font color to the header."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, header_format={"font_color": "white"})
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["A1"].font.color.rgb == "FFFFFFFF"
                wb.close()
        finally:
            Path(path).unlink()

    def test_combined_header_format(self) -> None:
        """Combine multiple header format options."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df,
                path,
                header_format={"bold": True, "bg_color": "#4F81BD", "font_color": "white"},
            )
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["A1"].font.bold is True
                assert ws["A1"].fill.fgColor.rgb == "FF4F81BD"
                assert ws["A1"].font.color.rgb == "FFFFFFFF"
                wb.close()
        finally:
            Path(path).unlink()

    def test_header_format_no_header(self) -> None:
        """Ignore header_format when header=False."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, header=False, header_format={"bold": True})
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                # First row should be data, not header
                assert ws["A1"].value == 1
                wb.close()
        finally:
            Path(path).unlink()

    def test_per_sheet_header_format(self) -> None:
        """Apply different header formats per sheet."""
        df1 = pd.DataFrame({"A": [1]})
        df2 = pd.DataFrame({"B": [2]})
        path = get_temp_path()
        try:
            xlsxturbo.dfs_to_xlsx(
                [
                    (df1, "Sheet1", {"header_format": {"bold": True}}),
                    (df2, "Sheet2", {"header_format": {"bg_color": "#FF0000"}}),
                ],
                path,
            )
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                assert wb["Sheet1"]["A1"].font.bold is True
                assert wb["Sheet2"]["A1"].fill.fgColor.rgb == "FFFF0000"
                wb.close()
        finally:
            Path(path).unlink()


class TestRichText:
    """Tests for rich text feature (v0.10.0)."""

    def test_rich_text_bold(self) -> None:
        """Apply rich text with bold formatting."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df, path, rich_text={"A1": [("Bold", {"bold": True}), " normal"]}
            )
            # openpyxl flattens rich text, so inspect the runs in sharedStrings.xml:
            # the two segments must be distinct runs and the first must be bold.
            with zipfile.ZipFile(path) as zf:
                shared = zf.read("xl/sharedStrings.xml").decode("utf-8")
                assert "Bold" in shared
                assert "normal" in shared
                assert "<b/>" in shared
        finally:
            Path(path).unlink()

    def test_rich_text_mixed_formats(self) -> None:
        """Apply rich text with multiple format segments."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df,
                path,
                rich_text={
                    "B2": [
                        ("Red text", {"font_color": "red"}),
                        " and ",
                        ("blue text", {"font_color": "blue", "italic": True}),
                    ]
                },
            )
            with zipfile.ZipFile(path) as zf:
                shared = zf.read("xl/sharedStrings.xml").decode("utf-8")
                assert "Red text" in shared
                assert "blue text" in shared
                # Red (FF0000) and blue (0000FF) runs and an italic run must be present.
                assert "FF0000" in shared.upper()
                assert "0000FF" in shared.upper()
                assert "<i/>" in shared
        finally:
            Path(path).unlink()

    def test_rich_text_plain_segments(self) -> None:
        """Apply rich text with plain string segments."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df, path, rich_text={"A3": ["Plain text ", ("bold", {"bold": True}), " more plain"]}
            )
            assert Path(path).exists()
        finally:
            Path(path).unlink()


class TestRowHeights:
    """Tests for row_heights parameter (v0.4.0)."""

    def test_basic_row_heights(self) -> None:
        """Set specific row heights."""
        df = pd.DataFrame({"A": [1, 2, 3]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, row_heights={0: 30, 2: 40})
            assert Path(path).exists()
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                # openpyxl is 1-indexed; Excel may round heights slightly
                assert abs(ws.row_dimensions[1].height - 30) < 1
                assert abs(ws.row_dimensions[3].height - 40) < 1
                # Rows without explicit height should not have customHeight
                assert ws.row_dimensions[2].customHeight is False
                wb.close()
        finally:
            Path(path).unlink()

    def test_row_heights_with_dfs_to_xlsx(self) -> None:
        """Apply row heights per-sheet."""
        df = pd.DataFrame({"A": [1, 2]})
        path = get_temp_path()
        try:
            xlsxturbo.dfs_to_xlsx(
                [(df, "Sheet1", {"row_heights": {0: 25}})],
                path,
            )
            assert Path(path).exists()
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb["Sheet1"]
                assert abs(ws.row_dimensions[1].height - 25) < 1
                wb.close()
        finally:
            Path(path).unlink()

    def test_row_heights_ignored_in_constant_memory(self) -> None:
        """Silently ignore row heights in constant memory mode."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, constant_memory=True, row_heights={0: 50})
            assert Path(path).exists()
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                # Row height should NOT be set (constant memory ignores it)
                # Default height is ~15, so it should not be 50
                assert ws.row_dimensions[1].height != 50 or ws.row_dimensions[1].height is None
                wb.close()
        finally:
            Path(path).unlink()


class TestBorderStyles:
    """Tests for per-side border styles (v0.12.0)."""

    def test_border_bool_backward_compat(self) -> None:
        """Apply a thin border on all sides with border=True."""
        df = pd.DataFrame({"A": [1, 2], "B": [3, 4]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"border": True}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                cell = ws["A2"]
                assert cell.border.left.style == "thin"
                assert cell.border.right.style == "thin"
                assert cell.border.top.style == "thin"
                assert cell.border.bottom.style == "thin"
                wb.close()
        finally:
            Path(path).unlink()

    def test_border_string_all_sides(self) -> None:
        """Apply a thick border on all 4 sides with border='thick'."""
        df = pd.DataFrame({"A": [1, 2], "B": [3, 4]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"border": "thick"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                cell = ws["A2"]
                assert cell.border.left.style == "thick"
                assert cell.border.right.style == "thick"
                assert cell.border.top.style == "thick"
                assert cell.border.bottom.style == "thick"
                wb.close()
        finally:
            Path(path).unlink()

    def test_border_right_only(self) -> None:
        """Apply a thick right border only with border_right='thick'."""
        df = pd.DataFrame({"A": [1, 2], "B": [3, 4]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"border_right": "thick"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                cell = ws["A2"]
                assert cell.border.right.style == "thick"
                assert cell.border.left.style is None
                assert cell.border.top.style is None
                assert cell.border.bottom.style is None
                wb.close()
        finally:
            Path(path).unlink()

    def test_border_left_and_right(self) -> None:
        """Apply mixed per-side borders."""
        df = pd.DataFrame({"A": [1], "B": [2]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"border_left": "medium", "border_right": "thick"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                cell = ws["A2"]
                assert cell.border.left.style == "medium"
                assert cell.border.right.style == "thick"
                assert cell.border.top.style is None
                assert cell.border.bottom.style is None
                wb.close()
        finally:
            Path(path).unlink()

    def test_border_all_four_sides_individually(self) -> None:
        """Set all four sides independently."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {
                    "border_left": "thin",
                    "border_right": "thick",
                    "border_top": "medium",
                    "border_bottom": "dashed",
                }
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                cell = ws["A2"]
                assert cell.border.left.style == "thin"
                assert cell.border.right.style == "thick"
                assert cell.border.top.style == "medium"
                assert cell.border.bottom.style == "dashed"
                wb.close()
        finally:
            Path(path).unlink()

    def test_border_string_thin(self) -> None:
        """Treat border='thin' as equivalent to border=True."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"border": "thin"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                cell = ws["A2"]
                assert cell.border.left.style == "thin"
                assert cell.border.right.style == "thin"
                wb.close()
        finally:
            Path(path).unlink()

    def test_border_with_wildcard_pattern(self) -> None:
        """Apply per-side borders with wildcard column matching."""
        df = pd.DataFrame({"price_usd": [10], "price_eur": [9], "name": ["x"]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "price_*": {"border_right": "thick"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["A2"].border.right.style == "thick"
                assert ws["B2"].border.right.style == "thick"
                assert ws["C2"].border.right.style is None
                wb.close()
        finally:
            Path(path).unlink()

    def test_border_color(self) -> None:
        """Set color for all borders via border_color."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"border": "thin", "border_color": "red"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                cell = ws["A2"]
                assert cell.border.left.style == "thin"
                assert cell.border.left.color is not None
                assert cell.border.left.color.rgb.endswith("FF0000")
                wb.close()
        finally:
            Path(path).unlink()

    def test_invalid_border_style_raises(self) -> None:
        """Raise ValueError for an invalid border style string."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            with pytest.raises(ValueError, match="Unknown border style"):
                xlsxturbo.df_to_xlsx(df, path, column_formats={
                    "A": {"border": "invalid_style"}
                })
        finally:
            Path(path).unlink(missing_ok=True)

    def test_border_with_other_formats(self) -> None:
        """Combine border styles with other column format options."""
        df = pd.DataFrame({"A": [1.5]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"border_right": "thick", "bold": True, "num_format": "0.00"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                cell = ws["A2"]
                assert cell.border.right.style == "thick"
                assert cell.font.bold
                assert cell.number_format == "0.00"
                wb.close()
        finally:
            Path(path).unlink()

    def test_border_per_side_overrides_all(self) -> None:
        """Override all-sides border with a per-side border for that side."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"border": "thin", "border_right": "thick"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                cell = ws["A2"]
                assert cell.border.left.style == "thin"
                assert cell.border.right.style == "thick"
                assert cell.border.top.style == "thin"
                assert cell.border.bottom.style == "thin"
                wb.close()
        finally:
            Path(path).unlink()

    def test_border_medium_style(self) -> None:
        """Apply the medium border style."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"border": "medium"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["A2"].border.left.style == "medium"
                wb.close()
        finally:
            Path(path).unlink()

    def test_border_double_style(self) -> None:
        """Apply the double border style."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"border_bottom": "double"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["A2"].border.bottom.style == "double"
                wb.close()
        finally:
            Path(path).unlink()

    def test_border_with_polars(self) -> None:
        """Apply border styles with polars DataFrames."""
        df = pl.DataFrame({"A": [1, 2]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"border_right": "thick"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["A2"].border.right.style == "thick"
                wb.close()
        finally:
            Path(path).unlink()

    def test_border_right_bool_treated_as_thin(self) -> None:
        """Apply a thin right border for border_right=True (bool)."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"border_right": True}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                cell = ws["A2"]
                assert cell.border.right.style == "thin"
                assert cell.border.left.style is None
                wb.close()
        finally:
            Path(path).unlink()

    def test_header_format_border(self) -> None:
        """Support border styles in header_format."""
        df = pd.DataFrame({"A": [1, 2], "B": [3, 4]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, header_format={
                "bold": True, "border": "thick"
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                header = ws["A1"]
                assert header.border.left.style == "thick"
                assert header.border.right.style == "thick"
                wb.close()
        finally:
            Path(path).unlink()

    def test_header_format_border_right_only(self) -> None:
        """Support per-side borders in header_format."""
        df = pd.DataFrame({"A": [1], "B": [2]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, header_format={
                "border_bottom": "medium"
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                header = ws["A1"]
                assert header.border.bottom.style == "medium"
                assert header.border.top.style is None
                wb.close()
        finally:
            Path(path).unlink()

    def test_border_dfs_to_xlsx_per_sheet(self) -> None:
        """Apply per-side borders with dfs_to_xlsx per-sheet overrides."""
        df1 = pd.DataFrame({"A": [1]})
        df2 = pd.DataFrame({"A": [2]})
        path = get_temp_path()
        try:
            xlsxturbo.dfs_to_xlsx([
                (df1, "S1"),
                (df2, "S2", {"column_formats": {"A": {"border_right": "thick"}}})
            ], path, column_formats={"A": {"border": "thin"}})
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                assert wb["S1"]["A2"].border.left.style == "thin"
                assert wb["S2"]["A2"].border.right.style == "thick"
                wb.close()
        finally:
            Path(path).unlink()


class TestTextAlignment:
    """Tests for text alignment (v0.12.0)."""

    def test_column_format_horizontal_center(self) -> None:
        """Center cell text with align_horizontal='center'."""
        df = pd.DataFrame({"A": ["hello", "world"]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"align_horizontal": "center"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["A2"].alignment.horizontal == "center"
                wb.close()
        finally:
            Path(path).unlink()

    def test_column_format_horizontal_right(self) -> None:
        """Right-align cell text with align_horizontal='right'."""
        df = pd.DataFrame({"A": ["hello"]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"align_horizontal": "right"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["A2"].alignment.horizontal == "right"
                wb.close()
        finally:
            Path(path).unlink()

    def test_column_format_vertical_top(self) -> None:
        """Top-align cell text with align_vertical='top'."""
        df = pd.DataFrame({"A": ["hello"]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"align_vertical": "top"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["A2"].alignment.vertical == "top"
                wb.close()
        finally:
            Path(path).unlink()

    def test_column_format_vertical_center(self) -> None:
        """Vertically center cell text with align_vertical='center'."""
        df = pd.DataFrame({"A": ["hello"]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"align_vertical": "center"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["A2"].alignment.vertical == "center"
                wb.close()
        finally:
            Path(path).unlink()

    def test_column_format_wrap_text(self) -> None:
        """Enable text wrapping with wrap_text=True."""
        df = pd.DataFrame({"A": ["hello world this is a long text"]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"wrap_text": True}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["A2"].alignment.wrapText is True
                wb.close()
        finally:
            Path(path).unlink()

    def test_column_format_combined_alignment(self) -> None:
        """Combine horizontal, vertical, and wrap_text alignment."""
        df = pd.DataFrame({"A": ["hello"]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"align_horizontal": "center", "align_vertical": "top", "wrap_text": True}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                cell = ws["A2"]
                assert cell.alignment.horizontal == "center"
                assert cell.alignment.vertical == "top"
                assert cell.alignment.wrapText is True
                wb.close()
        finally:
            Path(path).unlink()

    def test_header_format_alignment(self) -> None:
        """Support alignment in header_format."""
        df = pd.DataFrame({"A": [1], "B": [2]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, header_format={
                "bold": True, "align_horizontal": "center", "wrap_text": True
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                header = ws["A1"]
                assert header.alignment.horizontal == "center"
                assert header.alignment.wrapText is True
                wb.close()
        finally:
            Path(path).unlink()

    def test_merged_range_alignment(self) -> None:
        """Support alignment in merged_ranges format."""
        df = pd.DataFrame({"A": [1], "B": [2], "C": [3]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, merged_ranges=[
                ("A1:C1", "Title", {"bold": True, "align_horizontal": "left"})
            ])
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["A1"].alignment.horizontal == "left"
                wb.close()
        finally:
            Path(path).unlink()

    def test_merged_range_default_center(self) -> None:
        """Auto-center merged_ranges without explicit format."""
        df = pd.DataFrame({"A": [1], "B": [2]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, merged_ranges=[
                ("A1:B1", "Title")
            ])
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["A1"].alignment.horizontal == "center"
                wb.close()
        finally:
            Path(path).unlink()

    def test_cells_alignment(self) -> None:
        """Support alignment options in the cells parameter."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, cells={
                "C1": {"value": "Header", "align_horizontal": "center", "wrap_text": True}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                cell = ws["C1"]
                assert cell.value == "Header"
                assert cell.alignment.horizontal == "center"
                assert cell.alignment.wrapText is True
                wb.close()
        finally:
            Path(path).unlink()

    def test_alignment_with_wildcard(self) -> None:
        """Apply alignment with wildcard column patterns."""
        df = pd.DataFrame({"desc_a": ["x"], "desc_b": ["y"], "num": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "desc_*": {"align_horizontal": "left", "wrap_text": True}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["A2"].alignment.horizontal == "left"
                assert ws["A2"].alignment.wrapText is True
                assert ws["B2"].alignment.horizontal == "left"
                assert ws["C2"].alignment.horizontal is None
                wb.close()
        finally:
            Path(path).unlink()

    def test_alignment_with_border(self) -> None:
        """Combine alignment with border styles."""
        df = pd.DataFrame({"A": ["hello"]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"align_horizontal": "center", "border": "thin"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                cell = ws["A2"]
                assert cell.alignment.horizontal == "center"
                assert cell.border.left.style == "thin"
                wb.close()
        finally:
            Path(path).unlink()

    def test_invalid_horizontal_alignment_raises(self) -> None:
        """Raise ValueError for an invalid horizontal alignment."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            with pytest.raises(ValueError, match="Unknown horizontal alignment"):
                xlsxturbo.df_to_xlsx(df, path, column_formats={
                    "A": {"align_horizontal": "middle"}
                })
        finally:
            Path(path).unlink(missing_ok=True)

    def test_invalid_vertical_alignment_raises(self) -> None:
        """Raise ValueError for an invalid vertical alignment."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            with pytest.raises(ValueError, match="Unknown vertical alignment"):
                xlsxturbo.df_to_xlsx(df, path, column_formats={
                    "A": {"align_vertical": "left"}
                })
        finally:
            Path(path).unlink(missing_ok=True)

    def test_alignment_with_polars(self) -> None:
        """Apply alignment with polars DataFrames."""
        df = pl.DataFrame({"A": ["hello"]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"align_horizontal": "center"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["A2"].alignment.horizontal == "center"
                wb.close()
        finally:
            Path(path).unlink()
