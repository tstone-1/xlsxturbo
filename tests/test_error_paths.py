"""Tests for error handling and input validation in xlsxturbo."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest
import xlsxturbo

from tests.helpers import HAS_OPENPYXL, load_workbook

pytestmark = pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required for content verification")


class TestErrorPaths:
    """Tests for error handling (v0.10.0)."""

    def test_nonexistent_image_file_raises_error(self, tmp_xlsx: str) -> None:
        """Non-existent image file raises clear error."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(ValueError, match=r"(?i)image") as exc_info:
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, images={"B1": "/nonexistent/path/to/image.png"})
        message = str(exc_info.value)
        assert "Failed to load image" in message or "image" in message.lower()

    def test_validation_list_exceeds_255_chars_raises_error(self, tmp_xlsx: str) -> None:
        """Validation list exceeding 255 chars raises clear error."""
        df = pd.DataFrame({"Status": ["A"]})
        # Create values that exceed 255 chars total
        long_values = ["A" * 100, "B" * 100, "C" * 100]  # 300+ chars
        with pytest.raises(ValueError, match="255") as exc_info:
            xlsxturbo.df_to_xlsx(
                df, tmp_xlsx, validations={"Status": {"type": "list", "values": long_values}}
            )
        message = str(exc_info.value)
        assert "255" in message
        assert "character" in message.lower()

    def test_invalid_validation_config_raises_error(self, tmp_xlsx: str) -> None:
        """Invalid validation config (not a dict) raises clear error."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(TypeError, match="expected dict"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, validations={"A": "not_a_dict"})  # type: ignore[arg-type]  # intentionally invalid value type

    def test_validation_min_wrong_type_raises_error(self, tmp_xlsx: str) -> None:
        """Validation min/max values reject wrong types instead of defaulting."""
        df = pd.DataFrame({"Score": [85]})
        with pytest.raises(ValueError, match=r"validations\['Score'\].*min"):
            xlsxturbo.df_to_xlsx(
                df,
                tmp_xlsx,
                # min must be numeric; "zero" is intentionally the wrong type.
                validations={"Score": {"type": "whole_number", "min": "zero", "max": 100}},  # type: ignore[arg-type]
            )

    def test_validation_unknown_option_raises_error(self, tmp_xlsx: str) -> None:
        """Validation configs reject unknown keys instead of ignoring typos."""
        df = pd.DataFrame({"Score": [85]})
        with pytest.raises(ValueError, match="unknown option 'minimum'"):
            xlsxturbo.df_to_xlsx(
                df,
                tmp_xlsx,
                # "minimum" is an intentionally unknown key for ValidationOptions.
                validations={"Score": {"type": "whole_number", "minimum": 0, "max": 100}},  # type: ignore[arg-type]
            )

    def test_column_format_value_must_be_dict(self, tmp_xlsx: str) -> None:
        """Column format entries reject non-dict values."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(TypeError, match=r"column_formats.*expected dict"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_formats={"A": "bold"})  # type: ignore[arg-type]  # value must be a dict

    def test_column_format_pattern_must_match(self, tmp_xlsx: str) -> None:
        """Column format patterns that match no columns raise an error."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(ValueError, match=r"column_formats.*Missing.*matched no columns"):
            xlsxturbo.df_to_xlsx(
                df,
                tmp_xlsx,
                column_formats={"Missing": {"bold": True}},
            )

    def test_invalid_rich_text_segment_raises_error(self, tmp_xlsx: str) -> None:
        """Invalid rich_text segment (not string or tuple) raises clear error."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(TypeError) as exc_info:
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, rich_text={"A1": [123]})  # type: ignore[arg-type]  # int segment is intentionally invalid
        message = str(exc_info.value).lower()
        assert "segment" in message
        assert "string" in message or "tuple" in message

    def test_rich_text_format_must_be_dict(self, tmp_xlsx: str) -> None:
        """Rich text tuple formats reject non-dict values."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(TypeError, match=r"rich_text.*format must be a dict"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, rich_text={"A1": [("Bold", "not_a_dict")]})  # type: ignore[arg-type]  # format must be a dict

    def test_wrong_type_column_widths_raises_error(self, tmp_xlsx: str) -> None:
        """Passing a list instead of dict for column_widths raises TypeError."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(TypeError) as exc_info:
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_widths=[10, 20])  # type: ignore[arg-type]  # must be a dict, not a list
        message = str(exc_info.value)
        assert "expected dict" in message.lower()
        assert "column_widths" in message

    def test_wrong_type_header_format_raises_error(self, tmp_xlsx: str) -> None:
        """Passing a string instead of dict for header_format raises TypeError."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(TypeError) as exc_info:
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, header_format="bold")  # type: ignore[arg-type]  # must be a dict, not a str
        message = str(exc_info.value)
        assert "expected dict" in message.lower()
        assert "header_format" in message

    def test_wrong_type_merged_ranges_raises_error(self, tmp_xlsx: str) -> None:
        """Passing a dict instead of list for merged_ranges raises TypeError."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(TypeError) as exc_info:
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, merged_ranges={"A1:B1": "Title"})  # type: ignore[arg-type]  # must be a list, not a dict
        message = str(exc_info.value)
        assert "expected list" in message.lower()
        assert "merged_ranges" in message

    def test_merged_range_format_must_be_dict(self, tmp_xlsx: str) -> None:
        """Merged range format entries reject non-dict values."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(TypeError, match=r"merged_ranges.*format must be a dict"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, merged_ranges=[("A1:B1", "Title", "bold")])  # type: ignore[arg-type]  # tuple format must be a dict

    def test_merged_range_tuple_rejects_extra_items(self, tmp_xlsx: str) -> None:
        """Merged-range tuples require their documented exact arity."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(ValueError, match="exactly 2 or 3 elements"):
            xlsxturbo.df_to_xlsx(
                df,
                tmp_xlsx,
                merged_ranges=[("A1:B1", "Title", {"bold": True}, "ignored")],  # type: ignore[list-item]
            )

    def test_wrong_type_hyperlinks_raises_error(self, tmp_xlsx: str) -> None:
        """Passing a dict instead of list for hyperlinks raises TypeError."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(TypeError) as exc_info:
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, hyperlinks={"A1": "https://example.com"})  # type: ignore[arg-type]  # must be a list, not a dict
        message = str(exc_info.value)
        assert "expected list" in message.lower()
        assert "hyperlinks" in message

    def test_hyperlink_tuple_rejects_extra_items(self, tmp_xlsx: str) -> None:
        """Hyperlink tuples require their documented exact arity."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(ValueError, match="exactly 2 or 3 elements"):
            xlsxturbo.df_to_xlsx(
                df,
                tmp_xlsx,
                hyperlinks=[("A1", "https://example.com", "Example", "ignored")],  # type: ignore[list-item]
            )

    def test_invalid_rich_text_not_list_raises_error(self, tmp_xlsx: str) -> None:
        """Invalid rich_text value (not a list) raises clear error."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(TypeError, match="expected list"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, rich_text={"A1": "not_a_list"})  # type: ignore[arg-type]  # value must be a list

    def test_rich_text_tuple_rejects_extra_items(self, tmp_xlsx: str) -> None:
        """Rich-text segment tuples require exactly text and format."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(ValueError, match="tuple must have exactly 2 elements"):
            xlsxturbo.df_to_xlsx(
                df,
                tmp_xlsx,
                rich_text={"A1": [("Bold", {"bold": True}, "ignored")]},  # type: ignore[list-item]
            )

    def test_dfs_to_xlsx_per_sheet_invalid_dict_option_raises(self, tmp_xlsx: str) -> None:
        """Per-sheet dict options reject wrong container types."""
        df = pd.DataFrame({"A": [1]})
        # validations must be a dict; passing a str is intentionally invalid.
        bad_sheets = [(df, "Sheet1", {"validations": "not_a_dict"})]
        with pytest.raises(TypeError, match=r"validations.*dict"):
            xlsxturbo.dfs_to_xlsx(bad_sheets, tmp_xlsx)  # type: ignore[arg-type]

    def test_dfs_to_xlsx_per_sheet_invalid_list_option_raises(self, tmp_xlsx: str) -> None:
        """Per-sheet list options reject wrong container types."""
        df = pd.DataFrame({"A": [1]})
        # merged_ranges must be a list; passing a dict is intentionally invalid.
        bad_sheets = [(df, "Sheet1", {"merged_ranges": {"A1:B1": "Title"}})]
        with pytest.raises(TypeError, match=r"merged_ranges.*list"):
            xlsxturbo.dfs_to_xlsx(bad_sheets, tmp_xlsx)  # type: ignore[arg-type]

    def test_dfs_to_xlsx_per_sheet_invalid_cells_option_raises(self, tmp_xlsx: str) -> None:
        """Per-sheet cells option rejects wrong container types."""
        df = pd.DataFrame({"A": [1]})
        # cells must be a dict; passing a str is intentionally invalid.
        bad_sheets = [(df, "Sheet1", {"cells": "not_a_dict"})]
        with pytest.raises(TypeError, match=r"cells.*dict"):
            xlsxturbo.dfs_to_xlsx(bad_sheets, tmp_xlsx)  # type: ignore[arg-type]

    def test_dfs_to_xlsx_per_sheet_options_must_be_dict(self, tmp_xlsx: str) -> None:
        """Third sheet tuple item must be an options dict or None."""
        df = pd.DataFrame({"A": [1]})
        # the third tuple item must be a dict; a str is intentionally invalid.
        bad_sheets = [(df, "Sheet1", "not_a_dict")]
        with pytest.raises(TypeError, match="Sheet options must be a dict"):
            xlsxturbo.dfs_to_xlsx(bad_sheets, tmp_xlsx)  # type: ignore[arg-type]

    def test_dfs_to_xlsx_sheet_tuple_rejects_extra_items(self, tmp_xlsx: str) -> None:
        """Sheet tuples require their documented two- or three-item shape."""
        df = pd.DataFrame({"A": [1]})
        bad_sheets = [(df, "Sheet1", {}, "ignored")]
        with pytest.raises(ValueError, match="exactly 2 or 3 elements"):
            xlsxturbo.dfs_to_xlsx(bad_sheets, tmp_xlsx)  # type: ignore[arg-type]

    def test_dfs_to_xlsx_per_sheet_unknown_option_raises(self, tmp_xlsx: str) -> None:
        """Per-sheet option typos are rejected with the valid-key list."""
        df = pd.DataFrame({"A": [1]})
        # "tabel_style" is an intentionally unknown SheetOptions key.
        bad_sheets = [(df, "Sheet1", {"tabel_style": "Medium2"})]
        with pytest.raises(ValueError, match="Unknown sheet option 'tabel_style'"):
            xlsxturbo.dfs_to_xlsx(bad_sheets, tmp_xlsx)  # type: ignore[arg-type]

    def test_dfs_to_xlsx_per_sheet_header_wrong_type_raises(self, tmp_xlsx: str) -> None:
        """A wrong-typed per-sheet scalar option names the option and the received type.

        Exercises the per-sheet `header` option, which is extracted from the
        options dict via the `extract_scalar!` macro in `src/extract.rs`
        (distinct from the top-level `header` kwarg below, which pyo3 types
        and converts directly). The macro must propagate a context-rich error
        naming both `header` and the offending Python type instead of the
        default pyo3 conversion error.
        """
        df = pd.DataFrame({"A": [1]})
        # "yes" is intentionally the wrong type for the bool-typed 'header' option.
        bad_sheets = [(df, "Sheet1", {"header": "yes"})]
        with pytest.raises(TypeError, match=r"header.*bool.*str"):
            xlsxturbo.dfs_to_xlsx(bad_sheets, tmp_xlsx)  # type: ignore[arg-type]

    def test_df_to_xlsx_header_wrong_type_raises(self, tmp_xlsx: str) -> None:
        """A wrong-typed top-level `header` kwarg still raises a clear TypeError.

        Unlike the per-sheet option above, the top-level `header` kwarg is
        typed directly in the pyo3 function signature, so pyo3's own argument
        conversion rejects it before any of our Rust code runs; the message
        does not carry our `<option>: ...` context phrasing, but it must
        still identify both the expected and the offending type.
        """
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(TypeError) as exc_info:
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, header="yes")  # type: ignore[arg-type]
        message = str(exc_info.value).lower()
        assert "bool" in message
        assert "str" in message

    def test_cells_wrap_text_wrong_type_raises(self, tmp_xlsx: str) -> None:
        """Invalid cells wrap_text type raises a clear TypeError."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(TypeError, match=r"wrap_text.*bool"):
            # wrap_text must be a bool; "yes" is intentionally the wrong type.
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, cells={"B1": {"value": "x", "wrap_text": "yes"}})  # type: ignore[arg-type]

    def test_dfs_to_xlsx_write_error_names_the_failing_sheet(self, tmp_xlsx: str) -> None:
        """A per-sheet write-phase error (raised inside write_sheet_data) names its sheet.

        `column_formats` validation runs inside `write_sheet_data` (via
        `build_column_formats`), not during the up-front option extraction,
        so its error is a good probe for the `sheet '<name>': ` prefix that
        `dfs_to_xlsx` adds around each sheet's write result.
        """
        df1 = pd.DataFrame({"A": [1]})
        df2 = pd.DataFrame({"B": [2]})
        sheets = [
            (df1, "Sheet1"),
            # "Missing" matches no column on Sheet2, which is intentionally invalid.
            (df2, "Sheet2", {"column_formats": {"Missing": {"bold": True}}}),
        ]
        with pytest.raises(ValueError, match=r"sheet 'Sheet2'.*column_formats.*Missing"):
            xlsxturbo.dfs_to_xlsx(sheets, tmp_xlsx)

    def test_dfs_to_xlsx_empty_dataframes_same_table_name_succeeds(self, tmp_xlsx: str) -> None:
        """Two empty DataFrames sharing a table_name/table_style do not false-positive as a conflict.

        A table is only actually created when `row_count > 0` (the same gate
        `apply_worksheet_features` uses), so the duplicate-table-name
        pre-check must skip empty sheets too, matching that behavior.
        """
        df1 = pd.DataFrame({"A": pd.Series([], dtype="int64")})
        df2 = pd.DataFrame({"B": pd.Series([], dtype="int64")})
        sheets = [
            (df1, "Sheet1", {"table_style": "Medium2", "table_name": "SharedTable"}),
            (df2, "Sheet2", {"table_style": "Medium2", "table_name": "SharedTable"}),
        ]
        result = xlsxturbo.dfs_to_xlsx(sheets, tmp_xlsx)  # type: ignore[arg-type]
        assert len(result) == 2
        wb = load_workbook(tmp_xlsx)
        # Neither sheet has any data rows, so neither gets an actual Excel table.
        assert len(wb["Sheet1"].tables) == 0
        assert len(wb["Sheet2"].tables) == 0
        wb.close()

    def test_dfs_to_xlsx_empty_sheets_list_raises(self, tmp_xlsx: str) -> None:
        """dfs_to_xlsx rejects an empty sheets list instead of silently writing a blank workbook."""
        with pytest.raises(ValueError, match="at least one sheet"):
            xlsxturbo.dfs_to_xlsx([], tmp_xlsx)

    def test_df_to_xlsx_save_error_includes_output_path(self, tmp_path: Path) -> None:
        """A save failure (e.g. a nonexistent output directory) names the output path."""
        df = pd.DataFrame({"A": [1]})
        bad_path = str(tmp_path / "does_not_exist_dir" / "out.xlsx")
        with pytest.raises(ValueError, match="Failed to save") as exc_info:
            xlsxturbo.df_to_xlsx(df, bad_path)
        message = str(exc_info.value)
        assert bad_path in message

    def test_bytes_fspath_output_path_raises_clear_message(self) -> None:
        """A path-like object whose __fspath__ returns bytes is rejected with a clear message.

        os.PathLike permits bytes, but xlsxturbo only accepts str (or a
        path-like returning str); the error should say so explicitly rather
        than failing with an opaque extraction error.
        """

        class BytesPath:
            """A minimal os.PathLike whose __fspath__ intentionally returns bytes."""

            def __fspath__(self) -> bytes:
                """Return a bytes path, which xlsxturbo does not support.

                Returns:
                    The (unsupported) bytes path.
                """
                return b"/tmp/bytes_output.xlsx"

        df = pd.DataFrame({"A": [1]})
        with pytest.raises(TypeError, match="bytes paths are not supported"):
            xlsxturbo.df_to_xlsx(df, BytesPath())  # type: ignore[arg-type]
