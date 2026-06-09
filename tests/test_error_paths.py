from tests.helpers import HAS_OPENPYXL, get_temp_path, load_workbook, os, pd, pl, pytest, xlsxturbo


pytestmark = pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required for content verification")


class TestErrorPaths:
    """Tests for error handling (v0.10.0)"""

    def test_nonexistent_image_file_raises_error(self):
        """Non-existent image file raises clear error"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, images={"B1": "/nonexistent/path/to/image.png"})
            assert False, "Should have raised an error"
        except ValueError as e:
            assert "Failed to load image" in str(e) or "image" in str(e).lower()
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_validation_list_exceeds_255_chars_raises_error(self):
        """Validation list exceeding 255 chars raises clear error"""
        df = pd.DataFrame({"Status": ["A"]})
        path = get_temp_path()
        # Create values that exceed 255 chars total
        long_values = ["A" * 100, "B" * 100, "C" * 100]  # 300+ chars
        try:
            xlsxturbo.df_to_xlsx(
                df, path, validations={"Status": {"type": "list", "values": long_values}}
            )
            assert False, "Should have raised an error"
        except ValueError as e:
            assert "255" in str(e) and "character" in str(e).lower()
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_invalid_validation_config_raises_error(self):
        """Invalid validation config (not a dict) raises clear error"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, validations={"A": "not_a_dict"})
            assert False, "Should have raised an error"
        except TypeError as e:
            assert "expected dict" in str(e).lower()
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_validation_min_wrong_type_raises_error(self):
        """Validation min/max values reject wrong types instead of defaulting"""
        df = pd.DataFrame({"Score": [85]})
        path = get_temp_path()
        try:
            with pytest.raises(ValueError, match="validations\\['Score'\\].*min"):
                xlsxturbo.df_to_xlsx(
                    df,
                    path,
                    validations={"Score": {"type": "whole_number", "min": "zero", "max": 100}},
                )
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_validation_unknown_option_raises_error(self):
        """Validation configs reject unknown keys instead of ignoring typos"""
        df = pd.DataFrame({"Score": [85]})
        path = get_temp_path()
        try:
            with pytest.raises(ValueError, match="unknown option 'minimum'"):
                xlsxturbo.df_to_xlsx(
                    df,
                    path,
                    validations={"Score": {"type": "whole_number", "minimum": 0, "max": 100}},
                )
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_column_format_value_must_be_dict(self):
        """Column format entries reject non-dict values"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            with pytest.raises(TypeError, match="column_formats.*expected dict"):
                xlsxturbo.df_to_xlsx(df, path, column_formats={"A": "bold"})
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_invalid_rich_text_segment_raises_error(self):
        """Invalid rich_text segment (not string or tuple) raises clear error"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, rich_text={"A1": [123]})  # int is invalid
            assert False, "Should have raised an error"
        except TypeError as e:
            assert "segment" in str(e).lower() and ("string" in str(e).lower() or "tuple" in str(e).lower())
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_rich_text_format_must_be_dict(self):
        """Rich text tuple formats reject non-dict values"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            with pytest.raises(TypeError, match="rich_text.*format must be a dict"):
                xlsxturbo.df_to_xlsx(df, path, rich_text={"A1": [("Bold", "not_a_dict")]})
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_wrong_type_column_widths_raises_error(self):
        """Passing a list instead of dict for column_widths raises TypeError"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_widths=[10, 20])
            assert False, "Should have raised TypeError"
        except TypeError as e:
            assert "expected dict" in str(e).lower()
            assert "column_widths" in str(e)
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_wrong_type_header_format_raises_error(self):
        """Passing a string instead of dict for header_format raises TypeError"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, header_format="bold")
            assert False, "Should have raised TypeError"
        except TypeError as e:
            assert "expected dict" in str(e).lower()
            assert "header_format" in str(e)
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_wrong_type_merged_ranges_raises_error(self):
        """Passing a dict instead of list for merged_ranges raises TypeError"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, merged_ranges={"A1:B1": "Title"})
            assert False, "Should have raised TypeError"
        except TypeError as e:
            assert "expected list" in str(e).lower()
            assert "merged_ranges" in str(e)
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_merged_range_format_must_be_dict(self):
        """Merged range format entries reject non-dict values"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            with pytest.raises(TypeError, match="merged_ranges.*format must be a dict"):
                xlsxturbo.df_to_xlsx(df, path, merged_ranges=[("A1:B1", "Title", "bold")])
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_wrong_type_hyperlinks_raises_error(self):
        """Passing a dict instead of list for hyperlinks raises TypeError"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, hyperlinks={"A1": "https://example.com"})
            assert False, "Should have raised TypeError"
        except TypeError as e:
            assert "expected list" in str(e).lower()
            assert "hyperlinks" in str(e)
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_invalid_rich_text_not_list_raises_error(self):
        """Invalid rich_text value (not a list) raises clear error"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, rich_text={"A1": "not_a_list"})
            assert False, "Should have raised an error"
        except TypeError as e:
            assert "expected list" in str(e).lower()
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_dfs_to_xlsx_per_sheet_invalid_dict_option_raises(self):
        """Per-sheet dict options reject wrong container types"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            with pytest.raises(TypeError, match="validations.*dict"):
                xlsxturbo.dfs_to_xlsx(
                    [(df, "Sheet1", {"validations": "not_a_dict"})], path
                )
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_dfs_to_xlsx_per_sheet_invalid_list_option_raises(self):
        """Per-sheet list options reject wrong container types"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            with pytest.raises(TypeError, match="merged_ranges.*list"):
                xlsxturbo.dfs_to_xlsx(
                    [(df, "Sheet1", {"merged_ranges": {"A1:B1": "Title"}})], path
                )
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_dfs_to_xlsx_per_sheet_invalid_cells_option_raises(self):
        """Per-sheet cells option rejects wrong container types"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            with pytest.raises(TypeError, match="cells.*dict"):
                xlsxturbo.dfs_to_xlsx(
                    [(df, "Sheet1", {"cells": "not_a_dict"})], path
                )
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_dfs_to_xlsx_per_sheet_options_must_be_dict(self):
        """Third sheet tuple item must be an options dict or None"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            with pytest.raises(TypeError, match="Sheet options must be a dict"):
                xlsxturbo.dfs_to_xlsx([(df, "Sheet1", "not_a_dict")], path)
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_dfs_to_xlsx_per_sheet_unknown_option_raises(self):
        """Per-sheet option typos are rejected with the valid-key list."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            with pytest.raises(ValueError, match="Unknown sheet option 'tabel_style'"):
                xlsxturbo.dfs_to_xlsx([(df, "Sheet1", {"tabel_style": "Medium2"})], path)
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_cells_wrap_text_wrong_type_raises(self):
        """Invalid cells wrap_text type raises a clear TypeError"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            with pytest.raises(TypeError, match="wrap_text.*bool"):
                xlsxturbo.df_to_xlsx(
                    df, path, cells={"B1": {"value": "x", "wrap_text": "yes"}}
                )
        finally:
            if os.path.exists(path):
                os.unlink(path)
