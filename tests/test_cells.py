from tests.helpers import HAS_OPENPYXL, get_temp_path, load_workbook, os, pd, pl, pytest, xlsxturbo


pytestmark = pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required for content verification")


class TestCells:
    """Tests for arbitrary cell writes (v0.11.0)"""

    def test_simple_string_cell(self):
        """Write a string to a specific cell"""
        df = pd.DataFrame({"a": [1, 2]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, cells={"C1": "hello"})
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                assert wb.active["C1"].value == "hello"
                wb.close()
        finally:
            os.unlink(path)

    def test_numeric_cells(self):
        """Write int and float to cells"""
        df = pd.DataFrame({"a": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, cells={"B5": 42, "C5": 3.14})
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                assert wb.active["B5"].value == 42
                assert abs(wb.active["C5"].value - 3.14) < 0.001
                wb.close()
        finally:
            os.unlink(path)

    def test_bool_cell(self):
        """Write boolean to cell"""
        df = pd.DataFrame({"a": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, cells={"B2": True})
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                assert wb.active["B2"].value is True
                wb.close()
        finally:
            os.unlink(path)

    def test_cell_with_num_format(self):
        """Dict-style cell with num_format preserves text format"""
        df = pd.DataFrame({"a": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, cells={
                "D6": {"value": "934728173849", "num_format": "@"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                cell = wb.active["D6"]
                assert cell.value == "934728173849"
                assert cell.number_format == "@"
                wb.close()
        finally:
            os.unlink(path)

    def test_cell_overwrites_dataframe_data(self):
        """cells parameter overwrites existing DataFrame values"""
        df = pd.DataFrame({"a": ["original"]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, cells={"A2": "overwritten"})
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                assert wb.active["A2"].value == "overwritten"
                wb.close()
        finally:
            os.unlink(path)

    def test_cell_dict_missing_value_key(self):
        """Dict-style cell without 'value' key raises ValueError"""
        df = pd.DataFrame({"a": [1]})
        path = get_temp_path()
        try:
            import pytest
            with pytest.raises(ValueError, match="missing 'value' key"):
                xlsxturbo.df_to_xlsx(df, path, cells={"A1": {"num_format": "@"}})
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_invalid_cell_ref(self):
        """Invalid cell reference raises ValueError"""
        df = pd.DataFrame({"a": [1]})
        path = get_temp_path()
        try:
            import pytest
            with pytest.raises(ValueError):
                xlsxturbo.df_to_xlsx(df, path, cells={"ZZZZ1": "x"})
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_cells_per_sheet_override(self):
        """Per-sheet cells override global cells in dfs_to_xlsx"""
        df1 = pd.DataFrame({"a": [1]})
        df2 = pd.DataFrame({"b": [2]})
        path = get_temp_path()
        try:
            xlsxturbo.dfs_to_xlsx(
                [(df1, "S1"), (df2, "S2", {"cells": {"C1": "per-sheet"}})],
                path, cells={"C1": "global"})
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                assert wb["S1"]["C1"].value == "global"
                assert wb["S2"]["C1"].value == "per-sheet"
                wb.close()
        finally:
            os.unlink(path)

    def test_cells_constant_memory_warns(self):
        """cells with constant_memory emits a warning"""
        df = pd.DataFrame({"a": [1, 2]})
        path = get_temp_path()
        try:
            import warnings
            with warnings.catch_warnings(record=True) as w:
                warnings.simplefilter("always")
                xlsxturbo.df_to_xlsx(df, path, constant_memory=True,
                    cells={"C1": "test"})
                assert len(w) == 1
                assert "cells" in str(w[0].message)
        finally:
            os.unlink(path)

    def test_num_format_wrong_type_raises(self):
        """Non-string num_format raises TypeError"""
        df = pd.DataFrame({"a": [1]})
        path = get_temp_path()
        try:
            import pytest
            with pytest.raises(TypeError):
                xlsxturbo.df_to_xlsx(df, path,
                    cells={"A1": {"value": "x", "num_format": 123}})
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_cells_with_polars(self):
        """cells work with polars DataFrames"""
        df = pl.DataFrame({"a": [1, 2]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, cells={"C1": "extra"})
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                assert wb.active["C1"].value == "extra"
                wb.close()
        finally:
            os.unlink(path)

class TestCellsPerSheet:
    """Tests for cells with per-sheet SheetOptions in dfs_to_xlsx (item 6)"""

    def test_cells_per_sheet_3tuple(self):
        """Per-sheet cells override via 3-tuple SheetOptions"""
        df1 = pd.DataFrame({"a": [1, 2]})
        df2 = pd.DataFrame({"b": [3, 4]})
        path = get_temp_path()
        try:
            xlsxturbo.dfs_to_xlsx([
                (df1, "Sheet1", {"cells": {"C1": "note1", "C2": 100}}),
                (df2, "Sheet2", {"cells": {"C1": "note2"}}),
            ], path)
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                assert wb["Sheet1"]["C1"].value == "note1"
                assert wb["Sheet1"]["C2"].value == 100
                assert wb["Sheet2"]["C1"].value == "note2"
                assert wb["Sheet2"]["C2"].value is None
                wb.close()
        finally:
            os.unlink(path)

    def test_cells_per_sheet_overrides_global(self):
        """Per-sheet cells replace (not merge with) global cells"""
        df1 = pd.DataFrame({"a": [1]})
        df2 = pd.DataFrame({"a": [2]})
        path = get_temp_path()
        try:
            xlsxturbo.dfs_to_xlsx([
                (df1, "S1"),
                (df2, "S2", {"cells": {"D1": "override"}}),
            ], path, cells={"C1": "global"})
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                # S1 gets global cells
                assert wb["S1"]["C1"].value == "global"
                # S2 gets per-sheet cells (override replaces global)
                assert wb["S2"]["D1"].value == "override"
                assert wb["S2"]["C1"].value is None
                wb.close()
        finally:
            os.unlink(path)

    def test_cells_per_sheet_with_num_format(self):
        """Per-sheet cells with num_format via SheetOptions"""
        df = pd.DataFrame({"a": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.dfs_to_xlsx([
                (df, "S1", {"cells": {
                    "C1": {"value": "12345", "num_format": "@"}
                }}),
            ], path)
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                cell = wb["S1"]["C1"]
                assert str(cell.value) == "12345"
                assert cell.number_format == "@"
                wb.close()
        finally:
            os.unlink(path)

    def test_cells_per_sheet_with_alignment(self):
        """Per-sheet cells with alignment via SheetOptions"""
        df = pd.DataFrame({"a": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.dfs_to_xlsx([
                (df, "S1", {"cells": {
                    "C1": {"value": "centered", "align_horizontal": "center"}
                }}),
            ], path)
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                cell = wb["S1"]["C1"]
                assert cell.value == "centered"
                assert cell.alignment.horizontal == "center"
                wb.close()
        finally:
            os.unlink(path)

class TestCellsFormatting:
    """Tests for cells with formatting options beyond num_format (item 7)"""

    def test_cells_with_horizontal_alignment(self):
        """cells with align_horizontal"""
        df = pd.DataFrame({"a": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, cells={
                "C1": {"value": "right-aligned", "align_horizontal": "right"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                cell = wb.active["C1"]
                assert cell.value == "right-aligned"
                assert cell.alignment.horizontal == "right"
                wb.close()
        finally:
            os.unlink(path)

    def test_cells_with_vertical_alignment(self):
        """cells with align_vertical"""
        df = pd.DataFrame({"a": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, cells={
                "C1": {"value": "top", "align_vertical": "top"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                cell = wb.active["C1"]
                assert cell.value == "top"
                assert cell.alignment.vertical == "top"
                wb.close()
        finally:
            os.unlink(path)

    def test_cells_with_wrap_text(self):
        """cells with wrap_text"""
        df = pd.DataFrame({"a": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, cells={
                "C1": {"value": "long text here", "wrap_text": True}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                cell = wb.active["C1"]
                assert cell.value == "long text here"
                assert cell.alignment.wrapText is True
                wb.close()
        finally:
            os.unlink(path)

    def test_cells_with_combined_formatting(self):
        """cells with num_format + alignment + wrap_text together"""
        df = pd.DataFrame({"a": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, cells={
                "C1": {
                    "value": 0.15,
                    "num_format": "0.00%",
                    "align_horizontal": "center",
                    "align_vertical": "top",
                    "wrap_text": True
                }
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                cell = wb.active["C1"]
                assert cell.number_format == "0.00%"
                assert cell.alignment.horizontal == "center"
                assert cell.alignment.vertical == "top"
                assert cell.alignment.wrapText is True
                wb.close()
        finally:
            os.unlink(path)

    def test_cells_formatting_with_polars(self):
        """cells formatting works with polars DataFrames"""
        df = pl.DataFrame({"a": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, cells={
                "C1": {"value": "test", "align_horizontal": "center"}
            })
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                cell = wb.active["C1"]
                assert cell.value == "test"
                assert cell.alignment.horizontal == "center"
                wb.close()
        finally:
            os.unlink(path)
