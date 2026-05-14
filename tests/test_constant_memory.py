from tests.helpers import HAS_OPENPYXL, get_temp_path, load_workbook, os, pd, pl, pytest, xlsxturbo


pytestmark = pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required for content verification")


class TestConstantMemoryMode:
    """Tests for constant_memory mode (v0.4.0)"""

    def test_basic_constant_memory(self):
        """File is created in constant memory mode"""
        df = pd.DataFrame({"A": list(range(100)), "B": list(range(100, 200))})
        path = get_temp_path()
        try:
            rows, cols = xlsxturbo.df_to_xlsx(df, path, constant_memory=True)
            assert rows > 0
            assert cols == 2
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                assert ws["A1"].value == "A"  # header
                assert ws["A2"].value == 0  # first data row
                assert ws["B2"].value == 100
                wb.close()
        finally:
            os.unlink(path)

    def test_constant_memory_silently_disables_features(self):
        """Features are silently disabled in constant memory mode (no crash)"""
        df = pd.DataFrame({"Score": [1, 2, 3]})
        path = get_temp_path()
        try:
            rows, cols = xlsxturbo.df_to_xlsx(
                df,
                path,
                constant_memory=True,
                # All these should be silently ignored:
                table_style="Medium9",
                freeze_panes=True,
                autofit=True,
                row_heights={0: 30},
                conditional_formats={"Score": {"type": "data_bar", "bar_color": "#638EC6"}},
                formula_columns={"Double": "=A{row}*2"},
                merged_ranges=[("A1:A1", "Merge")],
                hyperlinks=[("A1", "https://example.com", "Link")],
                comments={"A1": "Comment"},
                validations={"Score": {"type": "whole_number", "min": 0, "max": 100}},
                rich_text={"B1": [("Bold", {"bold": True})]},
            )
            assert rows > 0
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                # Table should NOT be created
                assert len(ws.tables) == 0
                # Data should still be written
                assert ws["A1"].value == "Score"
                assert ws["A2"].value == 1
                wb.close()
        finally:
            os.unlink(path)

    def test_constant_memory_with_column_widths(self):
        """Column widths still work in constant memory mode"""
        df = pd.DataFrame({"A": [1], "B": [2]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, constant_memory=True, column_widths={0: 25})
            assert os.path.exists(path)
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                assert ws.column_dimensions["A"].width > 20
                wb.close()
        finally:
            os.unlink(path)

class TestConstantMemoryWarning:
    """Tests for constant_memory warning emission"""

    def test_constant_memory_warns_on_incompatible_options(self):
        """constant_memory=True with incompatible options emits RuntimeWarning"""
        import warnings

        df = pd.DataFrame({"A": [1, 2]})
        path = get_temp_path()
        try:
            with warnings.catch_warnings(record=True) as w:
                warnings.simplefilter("always")
                xlsxturbo.df_to_xlsx(
                    df, path,
                    constant_memory=True,
                    table_style="Medium2",
                    freeze_panes=True,
                )
                assert len(w) == 1
                assert issubclass(w[0].category, RuntimeWarning)
                assert "table_style" in str(w[0].message)
                assert "freeze_panes" in str(w[0].message)
        finally:
            os.unlink(path)

    def test_constant_memory_no_warning_when_clean(self):
        """constant_memory=True without incompatible options emits no warning"""
        import warnings

        df = pd.DataFrame({"A": [1, 2]})
        path = get_temp_path()
        try:
            with warnings.catch_warnings(record=True) as w:
                warnings.simplefilter("always")
                xlsxturbo.df_to_xlsx(df, path, constant_memory=True)
                assert len(w) == 0
        finally:
            os.unlink(path)
