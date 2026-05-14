from tests.helpers import HAS_OPENPYXL, get_temp_path, load_workbook, os, pd, pl, pytest, xlsxturbo


pytestmark = pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required for content verification")


class TestColumnWidthCap:
    """Tests for column_widths={'_all': value} feature"""

    def test_all_columns_capped(self):
        """'_all' key sets width for all columns"""
        df = pd.DataFrame({"A": ["x" * 100], "B": ["y" * 100], "C": ["z" * 100]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_widths={"_all": 20})
            assert os.path.exists(path)
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                for col in ["A", "B", "C"]:
                    assert ws.column_dimensions[col].width <= 21
                wb.close()
        finally:
            os.unlink(path)

    def test_specific_overrides_all(self):
        """Specific column width overrides '_all'"""
        df = pd.DataFrame({"A": ["x"], "B": ["y"], "C": ["z"]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_widths={0: 30, "_all": 10})
            assert os.path.exists(path)
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                assert ws.column_dimensions["A"].width > 25
                assert ws.column_dimensions["B"].width <= 11
                wb.close()
        finally:
            os.unlink(path)

    def test_all_with_autofit(self):
        """'_all' acts as cap when combined with autofit"""
        df = pd.DataFrame({"Short": ["x"], "VeryLongColumnName": ["y" * 100]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, autofit=True, column_widths={"_all": 25})
            assert os.path.exists(path)
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                # Long column should be capped at ~25
                assert ws.column_dimensions["B"].width <= 26
                wb.close()
        finally:
            os.unlink(path)

    def test_dfs_to_xlsx_per_sheet_all(self):
        """Per-sheet '_all' override in dfs_to_xlsx"""
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
            assert os.path.exists(path)
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                # Sheet1 should have narrower column than Sheet2
                w1 = wb["Sheet1"].column_dimensions["A"].width or 0
                w2 = wb["Sheet2"].column_dimensions["A"].width or 0
                assert w1 <= 21, f"Sheet1 col A width {w1} should be <= 21"
                assert w2 > 21, f"Sheet2 col A width {w2} should be > 21"
                wb.close()
        finally:
            os.unlink(path)


    def test_autofit_with_all_cap(self):
        """autofit=True + _all caps autofit widths instead of overriding"""
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
                ws = wb.active
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
            os.unlink(path)

    def test_autofit_with_all_cap_polars(self):
        """autofit + _all cap works with polars DataFrames"""
        df = pl.DataFrame({"Short": ["x"], "Long": ["A" * 80]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, autofit=True, column_widths={"_all": 20})
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                w_short = ws.column_dimensions["A"].width or 0
                w_long = ws.column_dimensions["B"].width or 0
                assert w_short < 15, f"Short col width {w_short} should be < 15"
                assert w_long <= 21, f"Long col width {w_long} should be <= 21"
                wb.close()
        finally:
            os.unlink(path)

class TestTableName:
    """Tests for table_name parameter"""

    def test_explicit_table_name(self):
        """Explicit table name is applied"""
        df = pd.DataFrame({"A": [1, 2, 3]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, table_style="Medium2", table_name="MyTable")
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                tables = list(ws.tables.keys())
                assert "MyTable" in tables
                wb.close()
        finally:
            os.unlink(path)

    def test_table_name_sanitization(self):
        """Invalid characters are sanitized"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df, path, table_style="Medium2", table_name="My Table-2024!"
            )
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                tables = list(ws.tables.keys())
                assert len(tables) == 1
                assert "_" in tables[0]  # Some characters replaced with underscore
                wb.close()
        finally:
            os.unlink(path)

    def test_table_name_starts_with_digit(self):
        """Table names starting with digit get underscore prefix"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, table_style="Medium2", table_name="123Data")
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                tables = list(ws.tables.keys())
                assert tables[0].startswith("_")
                wb.close()
        finally:
            os.unlink(path)

    def test_per_sheet_table_names(self):
        """Different table names per sheet"""
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
            os.unlink(path)

    def test_no_table_name_without_table_style(self):
        """table_name is ignored if table_style is None"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, table_name="Ignored")
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                assert len(ws.tables) == 0
                wb.close()
        finally:
            os.unlink(path)

class TestHeaderFormat:
    """Tests for header_format parameter"""

    def test_bold_header(self):
        """Bold header is applied"""
        df = pd.DataFrame({"A": [1], "B": [2]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, header_format={"bold": True})
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                assert ws["A1"].font.bold == True
                assert ws["B1"].font.bold == True
                wb.close()
        finally:
            os.unlink(path)

    def test_header_background_color(self):
        """Background color is applied to header"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, header_format={"bg_color": "#4F81BD"})
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                # openpyxl uses ARGB format
                assert ws["A1"].fill.fgColor.rgb == "FF4F81BD"
                wb.close()
        finally:
            os.unlink(path)

    def test_header_font_color(self):
        """Font color is applied to header"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, header_format={"font_color": "white"})
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                assert ws["A1"].font.color.rgb == "FFFFFFFF"
                wb.close()
        finally:
            os.unlink(path)

    def test_combined_header_format(self):
        """Multiple format options combined"""
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
                ws = wb.active
                assert ws["A1"].font.bold == True
                assert ws["A1"].fill.fgColor.rgb == "FF4F81BD"
                assert ws["A1"].font.color.rgb == "FFFFFFFF"
                wb.close()
        finally:
            os.unlink(path)

    def test_header_format_no_header(self):
        """header_format ignored when header=False"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, header=False, header_format={"bold": True})
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                # First row should be data, not header
                assert ws["A1"].value == 1
                wb.close()
        finally:
            os.unlink(path)

    def test_per_sheet_header_format(self):
        """Different header formats per sheet"""
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
                assert wb["Sheet1"]["A1"].font.bold == True
                assert wb["Sheet2"]["A1"].fill.fgColor.rgb == "FFFF0000"
                wb.close()
        finally:
            os.unlink(path)

class TestRichText:
    """Tests for rich text feature (v0.10.0)"""

    def test_rich_text_bold(self):
        """Rich text with bold formatting"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df, path, rich_text={"A1": [("Bold", {"bold": True}), " normal"]}
            )
            assert os.path.exists(path)
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                # openpyxl may have issues reading rich text, just verify file exists
                wb.close()
        finally:
            os.unlink(path)

    def test_rich_text_mixed_formats(self):
        """Rich text with multiple format segments"""
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
            assert os.path.exists(path)
        finally:
            os.unlink(path)

    def test_rich_text_plain_segments(self):
        """Rich text with plain string segments"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(
                df, path, rich_text={"A3": ["Plain text ", ("bold", {"bold": True}), " more plain"]}
            )
            assert os.path.exists(path)
        finally:
            os.unlink(path)

class TestRowHeights:
    """Tests for row_heights parameter (v0.4.0)"""

    def test_basic_row_heights(self):
        """Set specific row heights"""
        df = pd.DataFrame({"A": [1, 2, 3]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, row_heights={0: 30, 2: 40})
            assert os.path.exists(path)
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                # openpyxl is 1-indexed; Excel may round heights slightly
                assert abs(ws.row_dimensions[1].height - 30) < 1
                assert abs(ws.row_dimensions[3].height - 40) < 1
                # Rows without explicit height should not have customHeight
                assert ws.row_dimensions[2].customHeight is False
                wb.close()
        finally:
            os.unlink(path)

    def test_row_heights_with_dfs_to_xlsx(self):
        """Row heights work per-sheet"""
        df = pd.DataFrame({"A": [1, 2]})
        path = get_temp_path()
        try:
            xlsxturbo.dfs_to_xlsx(
                [(df, "Sheet1", {"row_heights": {0: 25}})],
                path,
            )
            assert os.path.exists(path)
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb["Sheet1"]
                assert abs(ws.row_dimensions[1].height - 25) < 1
                wb.close()
        finally:
            os.unlink(path)

    def test_row_heights_ignored_in_constant_memory(self):
        """Row heights silently ignored in constant memory mode"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, constant_memory=True, row_heights={0: 50})
            assert os.path.exists(path)
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                # Row height should NOT be set (constant memory ignores it)
                # Default height is ~15, so it should not be 50
                assert ws.row_dimensions[1].height != 50 or ws.row_dimensions[1].height is None
                wb.close()
        finally:
            os.unlink(path)

class TestBorderStyles:
    """Tests for per-side border styles (v0.12.0)"""

    def test_border_bool_backward_compat(self):
        """border=True still applies thin border on all sides"""
        df = pd.DataFrame({"A": [1, 2], "B": [3, 4]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"border": True}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                cell = ws["A2"]
                assert cell.border.left.style == "thin"
                assert cell.border.right.style == "thin"
                assert cell.border.top.style == "thin"
                assert cell.border.bottom.style == "thin"
                wb.close()
        finally:
            os.unlink(path)

    def test_border_string_all_sides(self):
        """border='thick' applies thick border on all 4 sides"""
        df = pd.DataFrame({"A": [1, 2], "B": [3, 4]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"border": "thick"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                cell = ws["A2"]
                assert cell.border.left.style == "thick"
                assert cell.border.right.style == "thick"
                assert cell.border.top.style == "thick"
                assert cell.border.bottom.style == "thick"
                wb.close()
        finally:
            os.unlink(path)

    def test_border_right_only(self):
        """border_right='thick' applies thick right border only"""
        df = pd.DataFrame({"A": [1, 2], "B": [3, 4]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"border_right": "thick"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                cell = ws["A2"]
                assert cell.border.right.style == "thick"
                assert cell.border.left.style is None
                assert cell.border.top.style is None
                assert cell.border.bottom.style is None
                wb.close()
        finally:
            os.unlink(path)

    def test_border_left_and_right(self):
        """Mixed per-side borders"""
        df = pd.DataFrame({"A": [1], "B": [2]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"border_left": "medium", "border_right": "thick"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                cell = ws["A2"]
                assert cell.border.left.style == "medium"
                assert cell.border.right.style == "thick"
                assert cell.border.top.style is None
                assert cell.border.bottom.style is None
                wb.close()
        finally:
            os.unlink(path)

    def test_border_all_four_sides_individually(self):
        """All four sides set independently"""
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
                ws = wb.active
                cell = ws["A2"]
                assert cell.border.left.style == "thin"
                assert cell.border.right.style == "thick"
                assert cell.border.top.style == "medium"
                assert cell.border.bottom.style == "dashed"
                wb.close()
        finally:
            os.unlink(path)

    def test_border_string_thin(self):
        """border='thin' is equivalent to border=True"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"border": "thin"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                cell = ws["A2"]
                assert cell.border.left.style == "thin"
                assert cell.border.right.style == "thin"
                wb.close()
        finally:
            os.unlink(path)

    def test_border_with_wildcard_pattern(self):
        """Per-side borders work with wildcard column matching"""
        df = pd.DataFrame({"price_usd": [10], "price_eur": [9], "name": ["x"]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "price_*": {"border_right": "thick"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                assert ws["A2"].border.right.style == "thick"
                assert ws["B2"].border.right.style == "thick"
                assert ws["C2"].border.right.style is None
                wb.close()
        finally:
            os.unlink(path)

    def test_border_color(self):
        """border_color sets color for all borders"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"border": "thin", "border_color": "red"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                cell = ws["A2"]
                assert cell.border.left.style == "thin"
                assert cell.border.left.color is not None
                assert cell.border.left.color.rgb.endswith("FF0000")
                wb.close()
        finally:
            os.unlink(path)

    def test_invalid_border_style_raises(self):
        """Invalid border style string raises ValueError"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            import pytest
            with pytest.raises(ValueError, match="Unknown border style"):
                xlsxturbo.df_to_xlsx(df, path, column_formats={
                    "A": {"border": "invalid_style"}
                })
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_border_with_other_formats(self):
        """Border styles combine with other column format options"""
        df = pd.DataFrame({"A": [1.5]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"border_right": "thick", "bold": True, "num_format": "0.00"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                cell = ws["A2"]
                assert cell.border.right.style == "thick"
                assert cell.font.bold
                assert cell.number_format == "0.00"
                wb.close()
        finally:
            os.unlink(path)

    def test_border_per_side_overrides_all(self):
        """Per-side border overrides all-sides border for that side"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"border": "thin", "border_right": "thick"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                cell = ws["A2"]
                assert cell.border.left.style == "thin"
                assert cell.border.right.style == "thick"
                assert cell.border.top.style == "thin"
                assert cell.border.bottom.style == "thin"
                wb.close()
        finally:
            os.unlink(path)

    def test_border_medium_style(self):
        """Medium border style works"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"border": "medium"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                assert ws["A2"].border.left.style == "medium"
                wb.close()
        finally:
            os.unlink(path)

    def test_border_double_style(self):
        """Double border style works"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"border_bottom": "double"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                assert ws["A2"].border.bottom.style == "double"
                wb.close()
        finally:
            os.unlink(path)

    def test_border_with_polars(self):
        """Border styles work with polars DataFrames"""
        df = pl.DataFrame({"A": [1, 2]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"border_right": "thick"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                assert ws["A2"].border.right.style == "thick"
                wb.close()
        finally:
            os.unlink(path)

    def test_border_right_bool_treated_as_thin(self):
        """border_right=True (bool) applies thin right border"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"border_right": True}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                cell = ws["A2"]
                assert cell.border.right.style == "thin"
                assert cell.border.left.style is None
                wb.close()
        finally:
            os.unlink(path)

    def test_header_format_border(self):
        """header_format supports border styles"""
        df = pd.DataFrame({"A": [1, 2], "B": [3, 4]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, header_format={
                "bold": True, "border": "thick"
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                header = ws["A1"]
                assert header.border.left.style == "thick"
                assert header.border.right.style == "thick"
                wb.close()
        finally:
            os.unlink(path)

    def test_header_format_border_right_only(self):
        """header_format supports per-side borders"""
        df = pd.DataFrame({"A": [1], "B": [2]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, header_format={
                "border_bottom": "medium"
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                header = ws["A1"]
                assert header.border.bottom.style == "medium"
                assert header.border.top.style is None
                wb.close()
        finally:
            os.unlink(path)

    def test_border_dfs_to_xlsx_per_sheet(self):
        """Per-side borders work with dfs_to_xlsx per-sheet overrides"""
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
            os.unlink(path)

class TestTextAlignment:
    """Tests for text alignment (v0.12.0)"""

    def test_column_format_horizontal_center(self):
        """align_horizontal='center' centers cell text"""
        df = pd.DataFrame({"A": ["hello", "world"]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"align_horizontal": "center"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                assert ws["A2"].alignment.horizontal == "center"
                wb.close()
        finally:
            os.unlink(path)

    def test_column_format_horizontal_right(self):
        """align_horizontal='right' right-aligns cell text"""
        df = pd.DataFrame({"A": ["hello"]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"align_horizontal": "right"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                assert ws["A2"].alignment.horizontal == "right"
                wb.close()
        finally:
            os.unlink(path)

    def test_column_format_vertical_top(self):
        """align_vertical='top' top-aligns cell text"""
        df = pd.DataFrame({"A": ["hello"]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"align_vertical": "top"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                assert ws["A2"].alignment.vertical == "top"
                wb.close()
        finally:
            os.unlink(path)

    def test_column_format_vertical_center(self):
        """align_vertical='center' vertically centers cell text"""
        df = pd.DataFrame({"A": ["hello"]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"align_vertical": "center"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                assert ws["A2"].alignment.vertical == "center"
                wb.close()
        finally:
            os.unlink(path)

    def test_column_format_wrap_text(self):
        """wrap_text=True enables text wrapping"""
        df = pd.DataFrame({"A": ["hello world this is a long text"]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"wrap_text": True}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                assert ws["A2"].alignment.wrapText is True
                wb.close()
        finally:
            os.unlink(path)

    def test_column_format_combined_alignment(self):
        """Horizontal + vertical + wrap_text together"""
        df = pd.DataFrame({"A": ["hello"]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"align_horizontal": "center", "align_vertical": "top", "wrap_text": True}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                cell = ws["A2"]
                assert cell.alignment.horizontal == "center"
                assert cell.alignment.vertical == "top"
                assert cell.alignment.wrapText is True
                wb.close()
        finally:
            os.unlink(path)

    def test_header_format_alignment(self):
        """header_format supports alignment"""
        df = pd.DataFrame({"A": [1], "B": [2]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, header_format={
                "bold": True, "align_horizontal": "center", "wrap_text": True
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                header = ws["A1"]
                assert header.alignment.horizontal == "center"
                assert header.alignment.wrapText is True
                wb.close()
        finally:
            os.unlink(path)

    def test_merged_range_alignment(self):
        """merged_ranges format supports alignment"""
        df = pd.DataFrame({"A": [1], "B": [2], "C": [3]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, merged_ranges=[
                ("A1:C1", "Title", {"bold": True, "align_horizontal": "left"})
            ])
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                assert ws["A1"].alignment.horizontal == "left"
                wb.close()
        finally:
            os.unlink(path)

    def test_merged_range_default_center(self):
        """merged_ranges without format still auto-center"""
        df = pd.DataFrame({"A": [1], "B": [2]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, merged_ranges=[
                ("A1:B1", "Title")
            ])
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                assert ws["A1"].alignment.horizontal == "center"
                wb.close()
        finally:
            os.unlink(path)

    def test_cells_alignment(self):
        """cells parameter supports alignment options"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, cells={
                "C1": {"value": "Header", "align_horizontal": "center", "wrap_text": True}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                cell = ws["C1"]
                assert cell.value == "Header"
                assert cell.alignment.horizontal == "center"
                assert cell.alignment.wrapText is True
                wb.close()
        finally:
            os.unlink(path)

    def test_alignment_with_wildcard(self):
        """Alignment works with wildcard column patterns"""
        df = pd.DataFrame({"desc_a": ["x"], "desc_b": ["y"], "num": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "desc_*": {"align_horizontal": "left", "wrap_text": True}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                assert ws["A2"].alignment.horizontal == "left"
                assert ws["A2"].alignment.wrapText is True
                assert ws["B2"].alignment.horizontal == "left"
                assert ws["C2"].alignment.horizontal is None
                wb.close()
        finally:
            os.unlink(path)

    def test_alignment_with_border(self):
        """Alignment combines with border styles"""
        df = pd.DataFrame({"A": ["hello"]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"align_horizontal": "center", "border": "thin"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                cell = ws["A2"]
                assert cell.alignment.horizontal == "center"
                assert cell.border.left.style == "thin"
                wb.close()
        finally:
            os.unlink(path)

    def test_invalid_horizontal_alignment_raises(self):
        """Invalid horizontal alignment raises ValueError"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            import pytest
            with pytest.raises(ValueError, match="Unknown horizontal alignment"):
                xlsxturbo.df_to_xlsx(df, path, column_formats={
                    "A": {"align_horizontal": "middle"}
                })
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_invalid_vertical_alignment_raises(self):
        """Invalid vertical alignment raises ValueError"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            import pytest
            with pytest.raises(ValueError, match="Unknown vertical alignment"):
                xlsxturbo.df_to_xlsx(df, path, column_formats={
                    "A": {"align_vertical": "left"}
                })
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_alignment_with_polars(self):
        """Alignment works with polars DataFrames"""
        df = pl.DataFrame({"A": ["hello"]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_formats={
                "A": {"align_horizontal": "center"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                assert ws["A2"].alignment.horizontal == "center"
                wb.close()
        finally:
            os.unlink(path)
