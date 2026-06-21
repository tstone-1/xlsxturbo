"""Tests for error handling and input validation in xlsxturbo."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest
import xlsxturbo

from tests.helpers import HAS_OPENPYXL, get_temp_path

pytestmark = pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required for content verification")


class TestErrorPaths:
    """Tests for error handling (v0.10.0)."""

    def test_nonexistent_image_file_raises_error(self) -> None:
        """Non-existent image file raises clear error."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            with pytest.raises(ValueError, match=r"(?i)image") as exc_info:
                xlsxturbo.df_to_xlsx(df, path, images={"B1": "/nonexistent/path/to/image.png"})
            message = str(exc_info.value)
            assert "Failed to load image" in message or "image" in message.lower()
        finally:
            Path(path).unlink(missing_ok=True)

    def test_validation_list_exceeds_255_chars_raises_error(self) -> None:
        """Validation list exceeding 255 chars raises clear error."""
        df = pd.DataFrame({"Status": ["A"]})
        path = get_temp_path()
        # Create values that exceed 255 chars total
        long_values = ["A" * 100, "B" * 100, "C" * 100]  # 300+ chars
        try:
            with pytest.raises(ValueError, match="255") as exc_info:
                xlsxturbo.df_to_xlsx(
                    df, path, validations={"Status": {"type": "list", "values": long_values}}
                )
            message = str(exc_info.value)
            assert "255" in message
            assert "character" in message.lower()
        finally:
            Path(path).unlink(missing_ok=True)

    def test_invalid_validation_config_raises_error(self) -> None:
        """Invalid validation config (not a dict) raises clear error."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            with pytest.raises(TypeError, match="expected dict"):
                xlsxturbo.df_to_xlsx(df, path, validations={"A": "not_a_dict"})  # type: ignore[arg-type]  # intentionally invalid value type
        finally:
            Path(path).unlink(missing_ok=True)

    def test_validation_min_wrong_type_raises_error(self) -> None:
        """Validation min/max values reject wrong types instead of defaulting."""
        df = pd.DataFrame({"Score": [85]})
        path = get_temp_path()
        try:
            with pytest.raises(ValueError, match=r"validations\['Score'\].*min"):
                xlsxturbo.df_to_xlsx(
                    df,
                    path,
                    # min must be numeric; "zero" is intentionally the wrong type.
                    validations={"Score": {"type": "whole_number", "min": "zero", "max": 100}},  # type: ignore[arg-type]
                )
        finally:
            Path(path).unlink(missing_ok=True)

    def test_validation_unknown_option_raises_error(self) -> None:
        """Validation configs reject unknown keys instead of ignoring typos."""
        df = pd.DataFrame({"Score": [85]})
        path = get_temp_path()
        try:
            with pytest.raises(ValueError, match="unknown option 'minimum'"):
                xlsxturbo.df_to_xlsx(
                    df,
                    path,
                    # "minimum" is an intentionally unknown key for ValidationOptions.
                    validations={"Score": {"type": "whole_number", "minimum": 0, "max": 100}},  # type: ignore[arg-type]
                )
        finally:
            Path(path).unlink(missing_ok=True)

    def test_column_format_value_must_be_dict(self) -> None:
        """Column format entries reject non-dict values."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            with pytest.raises(TypeError, match=r"column_formats.*expected dict"):
                xlsxturbo.df_to_xlsx(df, path, column_formats={"A": "bold"})  # type: ignore[arg-type]  # value must be a dict
        finally:
            Path(path).unlink(missing_ok=True)

    def test_invalid_rich_text_segment_raises_error(self) -> None:
        """Invalid rich_text segment (not string or tuple) raises clear error."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            with pytest.raises(TypeError) as exc_info:
                xlsxturbo.df_to_xlsx(df, path, rich_text={"A1": [123]})  # type: ignore[arg-type]  # int segment is intentionally invalid
            message = str(exc_info.value).lower()
            assert "segment" in message
            assert "string" in message or "tuple" in message
        finally:
            Path(path).unlink(missing_ok=True)

    def test_rich_text_format_must_be_dict(self) -> None:
        """Rich text tuple formats reject non-dict values."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            with pytest.raises(TypeError, match=r"rich_text.*format must be a dict"):
                xlsxturbo.df_to_xlsx(df, path, rich_text={"A1": [("Bold", "not_a_dict")]})  # type: ignore[arg-type]  # format must be a dict
        finally:
            Path(path).unlink(missing_ok=True)

    def test_wrong_type_column_widths_raises_error(self) -> None:
        """Passing a list instead of dict for column_widths raises TypeError."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            with pytest.raises(TypeError) as exc_info:
                xlsxturbo.df_to_xlsx(df, path, column_widths=[10, 20])  # type: ignore[arg-type]  # must be a dict, not a list
            message = str(exc_info.value)
            assert "expected dict" in message.lower()
            assert "column_widths" in message
        finally:
            Path(path).unlink(missing_ok=True)

    def test_wrong_type_header_format_raises_error(self) -> None:
        """Passing a string instead of dict for header_format raises TypeError."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            with pytest.raises(TypeError) as exc_info:
                xlsxturbo.df_to_xlsx(df, path, header_format="bold")  # type: ignore[arg-type]  # must be a dict, not a str
            message = str(exc_info.value)
            assert "expected dict" in message.lower()
            assert "header_format" in message
        finally:
            Path(path).unlink(missing_ok=True)

    def test_wrong_type_merged_ranges_raises_error(self) -> None:
        """Passing a dict instead of list for merged_ranges raises TypeError."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            with pytest.raises(TypeError) as exc_info:
                xlsxturbo.df_to_xlsx(df, path, merged_ranges={"A1:B1": "Title"})  # type: ignore[arg-type]  # must be a list, not a dict
            message = str(exc_info.value)
            assert "expected list" in message.lower()
            assert "merged_ranges" in message
        finally:
            Path(path).unlink(missing_ok=True)

    def test_merged_range_format_must_be_dict(self) -> None:
        """Merged range format entries reject non-dict values."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            with pytest.raises(TypeError, match=r"merged_ranges.*format must be a dict"):
                xlsxturbo.df_to_xlsx(df, path, merged_ranges=[("A1:B1", "Title", "bold")])  # type: ignore[arg-type]  # tuple format must be a dict
        finally:
            Path(path).unlink(missing_ok=True)

    def test_wrong_type_hyperlinks_raises_error(self) -> None:
        """Passing a dict instead of list for hyperlinks raises TypeError."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            with pytest.raises(TypeError) as exc_info:
                xlsxturbo.df_to_xlsx(df, path, hyperlinks={"A1": "https://example.com"})  # type: ignore[arg-type]  # must be a list, not a dict
            message = str(exc_info.value)
            assert "expected list" in message.lower()
            assert "hyperlinks" in message
        finally:
            Path(path).unlink(missing_ok=True)

    def test_invalid_rich_text_not_list_raises_error(self) -> None:
        """Invalid rich_text value (not a list) raises clear error."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            with pytest.raises(TypeError, match="expected list"):
                xlsxturbo.df_to_xlsx(df, path, rich_text={"A1": "not_a_list"})  # type: ignore[arg-type]  # value must be a list
        finally:
            Path(path).unlink(missing_ok=True)

    def test_dfs_to_xlsx_per_sheet_invalid_dict_option_raises(self) -> None:
        """Per-sheet dict options reject wrong container types."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            # validations must be a dict; passing a str is intentionally invalid.
            bad_sheets = [(df, "Sheet1", {"validations": "not_a_dict"})]
            with pytest.raises(TypeError, match=r"validations.*dict"):
                xlsxturbo.dfs_to_xlsx(bad_sheets, path)  # type: ignore[arg-type]
        finally:
            Path(path).unlink(missing_ok=True)

    def test_dfs_to_xlsx_per_sheet_invalid_list_option_raises(self) -> None:
        """Per-sheet list options reject wrong container types."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            # merged_ranges must be a list; passing a dict is intentionally invalid.
            bad_sheets = [(df, "Sheet1", {"merged_ranges": {"A1:B1": "Title"}})]
            with pytest.raises(TypeError, match=r"merged_ranges.*list"):
                xlsxturbo.dfs_to_xlsx(bad_sheets, path)  # type: ignore[arg-type]
        finally:
            Path(path).unlink(missing_ok=True)

    def test_dfs_to_xlsx_per_sheet_invalid_cells_option_raises(self) -> None:
        """Per-sheet cells option rejects wrong container types."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            # cells must be a dict; passing a str is intentionally invalid.
            bad_sheets = [(df, "Sheet1", {"cells": "not_a_dict"})]
            with pytest.raises(TypeError, match=r"cells.*dict"):
                xlsxturbo.dfs_to_xlsx(bad_sheets, path)  # type: ignore[arg-type]
        finally:
            Path(path).unlink(missing_ok=True)

    def test_dfs_to_xlsx_per_sheet_options_must_be_dict(self) -> None:
        """Third sheet tuple item must be an options dict or None."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            # the third tuple item must be a dict; a str is intentionally invalid.
            bad_sheets = [(df, "Sheet1", "not_a_dict")]
            with pytest.raises(TypeError, match="Sheet options must be a dict"):
                xlsxturbo.dfs_to_xlsx(bad_sheets, path)  # type: ignore[arg-type]
        finally:
            Path(path).unlink(missing_ok=True)

    def test_dfs_to_xlsx_per_sheet_unknown_option_raises(self) -> None:
        """Per-sheet option typos are rejected with the valid-key list."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            # "tabel_style" is an intentionally unknown SheetOptions key.
            bad_sheets = [(df, "Sheet1", {"tabel_style": "Medium2"})]
            with pytest.raises(ValueError, match="Unknown sheet option 'tabel_style'"):
                xlsxturbo.dfs_to_xlsx(bad_sheets, path)  # type: ignore[arg-type]
        finally:
            Path(path).unlink(missing_ok=True)

    def test_cells_wrap_text_wrong_type_raises(self) -> None:
        """Invalid cells wrap_text type raises a clear TypeError."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            with pytest.raises(TypeError, match=r"wrap_text.*bool"):
                # wrap_text must be a bool; "yes" is intentionally the wrong type.
                xlsxturbo.df_to_xlsx(df, path, cells={"B1": {"value": "x", "wrap_text": "yes"}})  # type: ignore[arg-type]
        finally:
            Path(path).unlink(missing_ok=True)
