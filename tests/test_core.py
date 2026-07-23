"""Tests for core df_to_xlsx, csv_to_xlsx, and related behavior."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path

import numpy as np
import pandas as pd
import polars as pl
import pytest
import xlsxturbo

from tests.helpers import HAS_OPENPYXL, active_ws, load_workbook

pytestmark = pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required for content verification")


class TestBackwardCompatibility:
    """Ensure existing functionality still works."""

    def test_old_column_widths_still_works(self, tmp_xlsx: str) -> None:
        """Integer key column_widths still works."""
        df = pd.DataFrame({"A": [1], "B": [2]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_widths={0: 20, 1: 30})
        assert Path(tmp_xlsx).exists()
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws.column_dimensions["A"].width > 15
        assert ws.column_dimensions["B"].width > 25
        wb.close()

    def test_table_style_without_table_name(self, tmp_xlsx: str) -> None:
        """table_style works without table_name (existing behavior)."""
        df = pd.DataFrame({"A": [1, 2, 3]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, table_style="Medium9")
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert len(ws.tables) == 1
        wb.close()


class TestPolarsSupport:
    """Ensure all features work with polars DataFrames."""

    def test_polars_column_width_cap(self, tmp_xlsx: str) -> None:
        """Polars column width cap is respected."""
        df = pl.DataFrame({"A": ["x" * 100], "B": ["y" * 100]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, column_widths={"_all": 20})
        assert Path(tmp_xlsx).exists()
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws.column_dimensions["A"].width <= 21
        wb.close()

    def test_polars_datetime_fractional_seconds(self, tmp_xlsx: str) -> None:
        """Polars datetime columns write as Excel datetimes.

        Polars `iter_rows` yields Python `datetime.datetime` objects, so this
        hits the "datetime" branch of write_py_value_with_format -- a distinct
        path from the pandas datetime64/Timestamp branches.
        """
        from datetime import datetime

        df = pl.DataFrame({"t": [datetime(2024, 1, 1, 12, 34, 56, 789000)]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx)
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A2"].value == datetime(2024, 1, 1, 12, 34, 56, 789000)
        wb.close()

    def test_polars_table_name(self, tmp_xlsx: str) -> None:
        """Polars DataFrame with a custom table name creates that table."""
        df = pl.DataFrame({"A": [1, 2, 3]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, table_style="Medium2", table_name="PolarsTable")
        assert Path(tmp_xlsx).exists()
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert "PolarsTable" in ws.tables
        wb.close()

    def test_polars_header_format(self, tmp_xlsx: str) -> None:
        """Polars DataFrame honors header_format."""
        df = pl.DataFrame({"A": [1], "B": [2]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, header_format={"bold": True})
        assert Path(tmp_xlsx).exists()
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A1"].font.bold is True
        wb.close()


class TestBooleanDtype:
    """Tests for pure-bool-dtype columns (write.rs write_py_value_with_format).

    A pandas DataFrame whose column is entirely bool-typed yields `np.bool_`
    scalars from `df.values` (a plain object-dtype/mixed column yields real
    Python `bool` instead). Before the numpy-bool branch was added, those
    scalars fell through to the numpy-scalar-int fallback via `__index__`
    and were written as the numbers 0/1 rather than Excel booleans.
    """

    def test_pure_bool_dtype_pandas_dataframe_writes_excel_booleans(self, tmp_xlsx: str) -> None:
        """A pandas DataFrame with a pure bool dtype column writes real Excel booleans."""
        df = pd.DataFrame({"flag": [True, False]})
        assert df["flag"].dtype == bool
        xlsxturbo.df_to_xlsx(df, tmp_xlsx)
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A2"].data_type == "b"
        assert ws["A2"].value is True
        assert ws["A3"].data_type == "b"
        assert ws["A3"].value is False
        wb.close()

    def test_polars_boolean_column_writes_excel_booleans(self, tmp_xlsx: str) -> None:
        """A polars DataFrame with a Boolean column writes real Excel booleans."""
        df = pl.DataFrame({"flag": [True, False]})
        assert df["flag"].dtype == pl.Boolean
        xlsxturbo.df_to_xlsx(df, tmp_xlsx)
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A2"].data_type == "b"
        assert ws["A2"].value is True
        assert ws["A3"].data_type == "b"
        assert ws["A3"].value is False
        wb.close()

    def test_mixed_dtype_dataframe_bool_column_still_writes_booleans(self, tmp_xlsx: str) -> None:
        """A mixed-dtype DataFrame's bool column still writes real Excel booleans.

        Regression guard for the pre-existing (already-working) path: mixed
        dtypes keep object-dtype columns holding real Python `bool` values,
        which the `PyBool` cast handles directly, distinct from the
        `np.bool_` scalar path pure-bool-dtype columns exercise above.
        """
        df = pd.DataFrame({"flag": [True, False], "count": [1, 2]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx)
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A2"].data_type == "b"
        assert ws["A2"].value is True
        assert ws["A3"].data_type == "b"
        assert ws["A3"].value is False
        wb.close()


class TestEdgeCases:
    """Tests for edge cases and error handling."""

    def test_empty_dataframe(self, tmp_xlsx: str) -> None:
        """Empty DataFrame writes successfully."""
        df = pd.DataFrame({"A": [], "B": []})
        rows, cols = xlsxturbo.df_to_xlsx(df, tmp_xlsx)
        assert rows == 1  # Just header
        assert cols == 2
        assert Path(tmp_xlsx).exists()

    def test_pathlib_output_path(self, tmp_xlsx: str) -> None:
        """pathlib.Path is accepted for output_path."""
        df = pd.DataFrame({"A": [1]})
        output_path = Path(tmp_xlsx)
        rows, cols = xlsxturbo.df_to_xlsx(df, output_path)
        assert rows == 2
        assert cols == 1
        assert output_path.exists()

    def test_empty_dataframe_with_table_style(self, tmp_xlsx: str) -> None:
        """Empty DataFrame with table_style writes without creating table."""
        df = pd.DataFrame({"A": [], "B": []})
        _rows, _cols = xlsxturbo.df_to_xlsx(df, tmp_xlsx, table_style="Medium2")
        assert Path(tmp_xlsx).exists()
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        # No table should be created for empty DataFrame
        assert len(ws.tables) == 0
        wb.close()

    def test_invalid_table_style_raises_error(self, tmp_xlsx: str) -> None:
        """Invalid table_style raises ValueError."""
        df = pd.DataFrame({"A": [1, 2]})
        with pytest.raises(ValueError, match="Unknown table_style") as exc_info:
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, table_style="InvalidStyle")
        message = str(exc_info.value)
        assert "Unknown table_style" in message
        assert "InvalidStyle" in message

    def test_invalid_hex_color_raises_error(self, tmp_xlsx: str) -> None:
        """Invalid hex color format raises ValueError."""
        df = pd.DataFrame({"A": [1]})
        with pytest.raises(ValueError, match="expected 6 characters"):
            xlsxturbo.df_to_xlsx(df, tmp_xlsx, header_format={"bg_color": "#FF"})

    def test_column_formats_order_preserved(self, tmp_xlsx: str) -> None:
        """Column format patterns are matched in order (first match wins)."""
        df = pd.DataFrame({"price_usd": [1.0], "price_eur": [2.0], "other": [3.0]})
        # The more specific pattern should be listed first to take priority
        xlsxturbo.df_to_xlsx(
            df,
            tmp_xlsx,
            column_formats={
                "price_usd": {"bg_color": "#FF0000"},  # Specific - should match first
                "price_*": {"bg_color": "#0000FF"},  # General - should match price_eur
            },
        )
        assert Path(tmp_xlsx).exists()
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        # price_usd should be red (specific match)
        assert ws["A2"].fill.fgColor.rgb == "FFFF0000"
        # price_eur should be blue (wildcard match)
        assert ws["B2"].fill.fgColor.rgb == "FF0000FF"
        # other should have no background
        wb.close()

    def test_empty_dataframe_no_header(self, tmp_xlsx: str) -> None:
        """Empty DataFrame with header=False."""
        df = pd.DataFrame({"A": [], "B": []})
        rows, cols = xlsxturbo.df_to_xlsx(df, tmp_xlsx, header=False)
        assert rows == 0
        assert cols == 2
        assert Path(tmp_xlsx).exists()

    def test_table_style_skipped_when_header_false(self, tmp_xlsx: str) -> None:
        """table_style is skipped when header=False (tables require headers)."""
        df = pd.DataFrame({"A": [1, 2]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx, header=False, table_style="Medium2")
        assert Path(tmp_xlsx).exists()
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A1"].value == 1  # data in row 1, no header
        wb.close()

    def test_dfs_per_sheet_header_false_with_global_table_style(self, tmp_xlsx: str) -> None:
        """Per-sheet header=False skips table even with global table_style."""
        df1 = pd.DataFrame({"A": [1]})
        df2 = pd.DataFrame({"B": [2]})
        xlsxturbo.dfs_to_xlsx([
            (df1, "WithHeader"),
            (df2, "NoHeader", {"header": False}),
        ], tmp_xlsx, table_style="Medium2")
        assert Path(tmp_xlsx).exists()
        wb = load_workbook(tmp_xlsx)
        # WithHeader sheet should have table
        assert wb["WithHeader"]["A1"].value == "A"
        # NoHeader sheet should have data in row 1
        assert wb["NoHeader"]["A1"].value == 2
        wb.close()


class TestDateOrder:
    """Tests for date_order parameter in csv_to_xlsx."""

    def test_date_order_us_parses_mdy(self, tmp_xlsx_factory: Callable[..., str]) -> None:
        """US date order parses 01-02-2024 as January 2."""
        import csv
        from datetime import datetime

        # Create CSV with ambiguous date
        csv_path = tmp_xlsx_factory(".csv")
        xlsx_path = tmp_xlsx_factory()
        with Path(csv_path).open("w", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["date", "value"])
            writer.writerow(["01-02-2024", "100"])

        xlsxturbo.csv_to_xlsx(csv_path, xlsx_path, date_order="us")

        wb = load_workbook(xlsx_path)
        ws = active_ws(wb)
        cell_value = ws["A2"].value
        # openpyxl returns datetime objects for Excel dates
        assert isinstance(cell_value, datetime), f"Expected datetime, got {type(cell_value)}"
        # US format: 01-02-2024 = January 2, 2024
        assert cell_value.month == 1, f"Expected January (1), got month {cell_value.month}"
        assert cell_value.day == 2, f"Expected day 2, got day {cell_value.day}"
        wb.close()

    def test_date_order_eu_parses_dmy(self, tmp_xlsx_factory: Callable[..., str]) -> None:
        """European date order parses 01-02-2024 as February 1."""
        import csv
        from datetime import datetime

        csv_path = tmp_xlsx_factory(".csv")
        xlsx_path = tmp_xlsx_factory()
        with Path(csv_path).open("w", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["date", "value"])
            writer.writerow(["01-02-2024", "100"])

        xlsxturbo.csv_to_xlsx(csv_path, xlsx_path, date_order="eu")

        wb = load_workbook(xlsx_path)
        ws = active_ws(wb)
        cell_value = ws["A2"].value
        assert isinstance(cell_value, datetime), f"Expected datetime, got {type(cell_value)}"
        # EU format: 01-02-2024 = February 1, 2024
        assert cell_value.month == 2, f"Expected February (2), got month {cell_value.month}"
        assert cell_value.day == 1, f"Expected day 1, got day {cell_value.day}"
        wb.close()

    def test_date_order_produces_different_results(self, tmp_xlsx_factory: Callable[..., str]) -> None:
        """US and EU date orders produce different Excel dates for ambiguous input."""
        import csv

        csv_path = tmp_xlsx_factory(".csv")
        xlsx_us = tmp_xlsx_factory()
        xlsx_eu = tmp_xlsx_factory()
        with Path(csv_path).open("w", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["date"])
            writer.writerow(["03-04-2024"])  # Mar 4 (US) vs Apr 3 (EU)

        xlsxturbo.csv_to_xlsx(csv_path, xlsx_us, date_order="us")
        xlsxturbo.csv_to_xlsx(csv_path, xlsx_eu, date_order="eu")

        wb_us = load_workbook(xlsx_us)
        wb_eu = load_workbook(xlsx_eu)
        us_value = active_ws(wb_us)["A2"].value
        eu_value = active_ws(wb_eu)["A2"].value
        wb_us.close()
        wb_eu.close()

        # Values should be different dates
        assert us_value != eu_value, "US and EU should produce different dates"
        # US: March 4, EU: April 3 (30 days difference)
        diff = abs((us_value - eu_value).days)
        assert diff == 30, f"Expected 30 day difference, got {diff}"

    def test_invalid_date_order_raises(self, tmp_xlsx_factory: Callable[..., str]) -> None:
        """Invalid date_order raises ValueError."""
        import csv

        csv_path = tmp_xlsx_factory(".csv")
        xlsx_path = tmp_xlsx_factory()
        with Path(csv_path).open("w", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["a"])
            writer.writerow(["1"])

        with pytest.raises(ValueError, match=r"(?i)invalid") as exc_info:
            xlsxturbo.csv_to_xlsx(csv_path, xlsx_path, date_order="invalid")  # type: ignore[arg-type]  # invalid on purpose
        message = str(exc_info.value)
        assert "invalid" in message.lower()
        # The valid-values list must include the 'european' alias for 'dmy'/'eu'.
        assert "european" in message.lower()


class TestCsvConversion:
    """Tests for csv_to_xlsx function."""

    def test_basic_csv(self, tmp_xlsx_factory: Callable[..., str]) -> None:
        """Basic CSV with header and data rows."""
        import csv

        csv_path = tmp_xlsx_factory(".csv")
        xlsx_path = tmp_xlsx_factory()
        with Path(csv_path).open("w", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["name", "age", "score"])
            writer.writerow(["Alice", "30", "95.5"])
            writer.writerow(["Bob", "25", "88.0"])

        rows, cols = xlsxturbo.csv_to_xlsx(csv_path, xlsx_path)
        assert rows == 3
        assert cols == 3
        wb = load_workbook(xlsx_path)
        ws = active_ws(wb)
        assert ws["A1"].value == "name"
        assert ws["B2"].value == 30  # should be detected as integer
        assert ws["C2"].value == 95.5  # should be detected as float
        wb.close()

    def test_csv_type_detection(self, tmp_xlsx_factory: Callable[..., str]) -> None:
        """CSV type detection: int, float, bool, date, string."""
        import csv

        csv_path = tmp_xlsx_factory(".csv")
        xlsx_path = tmp_xlsx_factory()
        with Path(csv_path).open("w", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["int", "float", "bool", "date", "text"])
            writer.writerow(["42", "3.14", "true", "2024-01-15", "hello"])

        xlsxturbo.csv_to_xlsx(csv_path, xlsx_path)
        wb = load_workbook(xlsx_path)
        ws = active_ws(wb)
        assert ws["A2"].value == 42
        assert abs(ws["B2"].value - 3.14) < 0.001
        assert ws["C2"].value is True
        # Date should be a datetime object in openpyxl
        from datetime import datetime

        assert isinstance(ws["D2"].value, datetime)
        assert ws["E2"].value == "hello"
        wb.close()

    def test_csv_datetime_fractional_seconds(self, tmp_xlsx_factory: Callable[..., str]) -> None:
        """CSV datetime values preserve fractional seconds in Excel serials."""
        import csv
        from datetime import datetime

        csv_path = tmp_xlsx_factory(".csv")
        xlsx_path = tmp_xlsx_factory()
        with Path(csv_path).open("w", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["timestamp"])
            writer.writerow(["2024-01-01T12:34:56.789"])

        xlsxturbo.csv_to_xlsx(csv_path, xlsx_path)
        wb = load_workbook(xlsx_path)
        ws = active_ws(wb)
        value = ws["A2"].value
        assert isinstance(value, datetime)
        assert value == datetime(2024, 1, 1, 12, 34, 56, 789000)
        wb.close()

    def test_csv_special_values(self, tmp_xlsx_factory: Callable[..., str]) -> None:
        """CSV with NaN, Inf, empty cells."""
        import csv

        csv_path = tmp_xlsx_factory(".csv")
        xlsx_path = tmp_xlsx_factory()
        with Path(csv_path).open("w", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["a", "b", "c"])
            writer.writerow(["NaN", "Inf", ""])
            writer.writerow(["nan", "-Inf", "   "])

        rows, _cols = xlsxturbo.csv_to_xlsx(csv_path, xlsx_path)
        assert rows == 3
        wb = load_workbook(xlsx_path)
        ws = active_ws(wb)
        # NaN/Inf/empty all become empty cells (write_cell -> CellValue::Empty
        # writes an empty string, which openpyxl reads back as None or "").
        for ref in ("A2", "B2", "C2", "A3", "B3", "C3"):
            assert ws[ref].value in (None, ""), f"{ref} should be empty"
        wb.close()

    def test_csv_string_cells_preserve_surrounding_whitespace(self, tmp_xlsx_factory: Callable[..., str]) -> None:
        """String cells keep leading/trailing whitespace; type detection still trims to classify.

        Quoted fields are used so the padding survives however csv.writer
        would otherwise decide to quote plain fields. A whitespace-padded
        string stays padded; a whitespace-padded number is still detected and
        written as a number (detection trims a private copy, but the original
        untrimmed value is what a genuine string falls back to).
        """
        import csv

        csv_path = tmp_xlsx_factory(".csv")
        xlsx_path = tmp_xlsx_factory()
        with Path(csv_path).open("w", newline="") as f:
            writer = csv.writer(f, quoting=csv.QUOTE_ALL)
            writer.writerow(["text", "number"])
            writer.writerow([" padded ", " 123 "])

        xlsxturbo.csv_to_xlsx(csv_path, xlsx_path)
        wb = load_workbook(xlsx_path)
        ws = active_ws(wb)
        assert ws["A2"].value == " padded "
        assert ws["B2"].value == 123
        wb.close()

    def test_csv_int_min_writes_as_string_without_overflow(self, tmp_xlsx_factory: Callable[..., str]) -> None:
        """i64::MIN via the CSV path (write_cell) is written as text, no overflow.

        Complements test_i64_min_writes_as_string_without_overflow, which covers
        the distinct DataFrame write_int path.
        """
        import csv as _csv

        int_min = -9223372036854775808  # i64::MIN
        csv_path = tmp_xlsx_factory(".csv")
        xlsx_path = tmp_xlsx_factory()
        with Path(csv_path).open("w", newline="") as f:
            w = _csv.writer(f)
            w.writerow(["n"])
            w.writerow([str(int_min)])
        xlsxturbo.csv_to_xlsx(csv_path, xlsx_path)
        wb = load_workbook(xlsx_path)
        ws = active_ws(wb)
        assert ws["A2"].value == str(int_min)
        wb.close()

    def test_csv_parallel(self, tmp_xlsx_factory: Callable[..., str]) -> None:
        """CSV parallel mode produces same output."""
        import csv

        csv_path = tmp_xlsx_factory(".csv")
        xlsx_seq = tmp_xlsx_factory()
        xlsx_par = tmp_xlsx_factory()
        with Path(csv_path).open("w", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["num", "text"])
            for i in range(100):
                writer.writerow([str(i), f"row_{i}"])

        rows_s, cols_s = xlsxturbo.csv_to_xlsx(csv_path, xlsx_seq, parallel=False)
        rows_p, cols_p = xlsxturbo.csv_to_xlsx(csv_path, xlsx_par, parallel=True)
        assert rows_s == rows_p
        assert cols_s == cols_p
        wb_s = load_workbook(xlsx_seq)
        wb_p = load_workbook(xlsx_par)
        ws_s = active_ws(wb_s)
        ws_p = active_ws(wb_p)
        # Spot check some cells match
        for row in [1, 2, 50, 101]:
            assert ws_s[f"A{row}"].value == ws_p[f"A{row}"].value
            assert ws_s[f"B{row}"].value == ws_p[f"B{row}"].value
        wb_s.close()
        wb_p.close()

    def test_csv_with_sheet_name(self, tmp_xlsx_factory: Callable[..., str]) -> None:
        """CSV conversion with custom sheet name."""
        import csv

        csv_path = tmp_xlsx_factory(".csv")
        xlsx_path = tmp_xlsx_factory()
        with Path(csv_path).open("w", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["a"])
            writer.writerow(["1"])

        xlsxturbo.csv_to_xlsx(csv_path, xlsx_path, sheet_name="MySheet")
        wb = load_workbook(xlsx_path)
        assert "MySheet" in wb.sheetnames
        wb.close()


class TestUnicodeAndSpecialData:
    """Tests for Unicode, mixed types, nulls, and CSV edge cases."""

    def test_unicode_column_names_and_data(self, tmp_xlsx: str) -> None:
        """Unicode characters in column names and cell data."""
        df = pd.DataFrame({
            "价格": [100, 200],       # Chinese: "price"
            "Straße": ["Berlin", "München"],  # German: street, Munich
            "名前": ["太郎", "花子"],           # Japanese names
        })
        rows, cols = xlsxturbo.df_to_xlsx(df, tmp_xlsx)
        assert rows == 3  # header + 2 data rows
        assert cols == 3
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A1"].value == "价格"
        assert ws["B1"].value == "Straße"
        assert ws["C1"].value == "名前"
        assert ws["B2"].value == "Berlin"
        assert ws["C2"].value == "太郎"
        wb.close()

    def test_emoji_in_data(self, tmp_xlsx: str) -> None:
        """Emoji characters in cell values."""
        df = pd.DataFrame({
            "status": ["done", "pending"],
            "icon": ["\U0001f680", "\U0001f525"],
        })
        rows, _cols = xlsxturbo.df_to_xlsx(df, tmp_xlsx)
        assert rows == 3
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["B2"].value == "\U0001f680"
        assert ws["B3"].value == "\U0001f525"
        wb.close()

    def test_mixed_type_column(self, tmp_xlsx: str) -> None:
        """Column with mixed int and string values (pandas object dtype)."""
        df = pd.DataFrame({"mixed": [1, "two", 3, "four", 5.5]})
        rows, _cols = xlsxturbo.df_to_xlsx(df, tmp_xlsx)
        assert rows == 6  # header + 5 rows
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A2"].value == 1
        assert ws["A3"].value == "two"
        assert ws["A4"].value == 3
        assert ws["A5"].value == "four"
        assert ws["A6"].value == 5.5
        wb.close()

    def test_none_and_nat_values(self, tmp_xlsx: str) -> None:
        """None, NaT, and pd.NA values write as empty cells."""
        df = pd.DataFrame({
            "a": [1, None, 3],
            "b": pd.array([10, pd.NA, 30], dtype="Int64"),
            "c": pd.to_datetime(["2024-01-01", "NaT", "2024-03-01"]),
        })
        rows, _cols = xlsxturbo.df_to_xlsx(df, tmp_xlsx)
        assert rows == 4  # header + 3 rows
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        # None/NA cells should be empty
        assert ws["A3"].value is None or ws["A3"].value == ""
        assert ws["B3"].value is None or ws["B3"].value == ""
        assert ws["C2"].value.year == 2024
        assert ws["C2"].value.month == 1
        assert ws["C2"].value.day == 1
        assert ws["C3"].value is None or ws["C3"].value == ""
        assert ws["C4"].value.year == 2024
        assert ws["C4"].value.month == 3
        assert ws["C4"].value.day == 1
        wb.close()

    def test_pandas_datetime64_preserves_datetime_and_fractional_seconds(self, tmp_xlsx: str) -> None:
        """Pandas datetime64[ns] columns write as datetimes, not strings."""
        from datetime import datetime

        df = pd.DataFrame({
            "timestamp": pd.to_datetime([
                "2024-01-01 12:34:56.789",
                "NaT",
            ])
        })
        rows, cols = xlsxturbo.df_to_xlsx(df, tmp_xlsx)
        assert rows == 3
        assert cols == 1
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A2"].value == datetime(2024, 1, 1, 12, 34, 56, 789000)
        assert ws["A3"].value is None or ws["A3"].value == ""
        wb.close()

    def test_non_ns_datetime64_out_of_range_writes_correct_date(self, tmp_xlsx: str) -> None:
        """datetime64[us] dates outside ns range must not wrap around."""
        from datetime import datetime

        df = pd.DataFrame({
            "timestamp": np.array(["3000-01-01T00:00:00"], dtype="datetime64[us]")
        })
        rows, cols = xlsxturbo.df_to_xlsx(df, tmp_xlsx)
        assert rows == 2
        assert cols == 1
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A2"].value == datetime(3000, 1, 1, 0, 0, 0)
        wb.close()

    def test_python_int_beyond_i64_writes_as_string(self, tmp_xlsx: str) -> None:
        """Oversized Python ints should not fall through to rounded f64."""
        value = 2**63 + 1025
        df = pd.DataFrame({"big": [value]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx)
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A2"].value == str(value)
        wb.close()

    def test_i64_min_writes_as_string_without_overflow(self, tmp_xlsx: str) -> None:
        """The signed minimum value must use the precision-preserving fallback."""
        value = np.iinfo(np.int64).min
        df = pd.DataFrame({"min": np.array([value], dtype=np.int64)})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx)
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A2"].value == str(value)
        wb.close()

    def test_dataframe_pre_1900_datetime_writes_as_string(self, tmp_xlsx: str) -> None:
        """DataFrame datetime paths match CSV behavior for Excel-unsupported dates."""
        from datetime import datetime

        df = pd.DataFrame({
            "old": pd.Series([datetime(1850, 1, 1, 12, 0, 0)], dtype=object)
        })
        xlsxturbo.df_to_xlsx(df, tmp_xlsx)
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A2"].value == "1850-01-01 12:00:00"
        wb.close()

    def test_object_timestamp_fractional_seconds(self, tmp_xlsx: str) -> None:
        """Object-dtype pandas Timestamps go through the attribute branch.

        Object-dtype Timestamps preserve fractional seconds (microsecond*1000 +
        nanosecond fold). This is a distinct code path from the datetime64[ns]
        numpy-scalar branch: an object-dtype column yields the Python Timestamp
        object, exercising the `.microsecond`/`.nanosecond` extraction in
        write_py_value_with_format. (True sub-microsecond precision cannot survive
        Excel's f64 serial at real-world date magnitudes, so millisecond precision
        is the meaningful assertion here.)
        """
        from datetime import datetime

        ts = pd.Timestamp("2024-01-01 12:34:56.789")
        df = pd.DataFrame({"t": pd.Series([ts], dtype=object)})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx)
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A2"].value == datetime(2024, 1, 1, 12, 34, 56, 789000)
        wb.close()

    def test_timezone_aware_datetime_writes_wall_clock(self, tmp_xlsx: str) -> None:
        """Timezone-aware datetimes are written as their local wall-clock value.

        The UTC offset is intentionally dropped (Excel has no timezone concept).
        Characterization test pinning the documented contract: 12:00 US/Eastern
        is stored as 12:00, NOT converted to its 17:00 UTC equivalent.
        """
        from datetime import datetime

        df = pd.DataFrame(
            {"t": pd.to_datetime(["2024-01-01 12:00:00"]).tz_localize("US/Eastern")}
        )
        xlsxturbo.df_to_xlsx(df, tmp_xlsx)
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A2"].value == datetime(2024, 1, 1, 12, 0, 0)
        wb.close()

    def test_all_none_column(self, tmp_xlsx: str) -> None:
        """Column with all None values."""
        df = pd.DataFrame({"empty": [None, None, None]})
        rows, cols = xlsxturbo.df_to_xlsx(df, tmp_xlsx)
        assert rows == 4
        assert cols == 1

    def test_large_integers_written_as_strings(self, tmp_xlsx: str) -> None:
        """Integers > 2^53 should be written as strings to prevent precision loss."""
        large_int = 9007199254740993  # 2^53 + 1
        df = pd.DataFrame({"id": [large_int, 42]})
        rows, _cols = xlsxturbo.df_to_xlsx(df, tmp_xlsx)
        assert rows == 3
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        # Large int should be written as string to preserve precision
        assert str(ws["A2"].value) == str(large_int)
        # Normal int should be a number
        assert ws["A3"].value == 42
        wb.close()

    def test_csv_with_bom(self, tmp_xlsx_factory: Callable[..., str]) -> None:
        """CSV file with UTF-8 BOM."""
        csv_path = tmp_xlsx_factory(".csv")
        xlsx_path = tmp_xlsx_factory()
        with Path(csv_path).open("w", encoding="utf-8-sig") as f:
            f.write("name,value\nAlice,1\nBob,2\n")
        rows, cols = xlsxturbo.csv_to_xlsx(csv_path, xlsx_path)
        assert rows == 3  # header + 2 data rows
        assert cols == 2

    def test_csv_with_crlf(self, tmp_xlsx_factory: Callable[..., str]) -> None:
        """CSV file with Windows CRLF line endings."""
        csv_path = tmp_xlsx_factory(".csv")
        xlsx_path = tmp_xlsx_factory()
        with Path(csv_path).open("wb") as f:
            f.write(b"a,b\r\n1,2\r\n3,4\r\n")
        rows, cols = xlsxturbo.csv_to_xlsx(csv_path, xlsx_path)
        assert rows == 3
        assert cols == 2

    def test_csv_quoted_fields_with_delimiters(self, tmp_xlsx_factory: Callable[..., str]) -> None:
        """CSV with quoted fields containing commas and newlines."""
        csv_path = tmp_xlsx_factory(".csv")
        xlsx_path = tmp_xlsx_factory()
        with Path(csv_path).open("w", encoding="utf-8") as f:
            f.write('name,address\n"Smith, John","123 Main St"\n"Doe, Jane","456 Oak Ave"\n')
        rows, cols = xlsxturbo.csv_to_xlsx(csv_path, xlsx_path)
        assert rows == 3
        assert cols == 2
        wb = load_workbook(xlsx_path)
        ws = active_ws(wb)
        assert ws["A2"].value == "Smith, John"
        assert ws["B2"].value == "123 Main St"
        wb.close()

    def test_polars_unicode(self, tmp_xlsx: str) -> None:
        """Unicode data through Polars DataFrames."""
        df = pl.DataFrame({
            "city": ["Tökyö", "Zürich", "São Paulo"],
            "pop": [14000000, 420000, 12300000],
        })
        rows, _cols = xlsxturbo.df_to_xlsx(df, tmp_xlsx)
        assert rows == 4
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A2"].value == "Tökyö"
        wb.close()


class TestCsvErrorPaths:
    """Tests for CSV conversion error handling."""

    def test_csv_nonexistent_input_raises_error(self, tmp_xlsx: str) -> None:
        """csv_to_xlsx with nonexistent input file raises ValueError with path info."""
        with pytest.raises(ValueError, match="Failed to open"):
            xlsxturbo.csv_to_xlsx("/nonexistent/file.csv", tmp_xlsx)


class TestPreEpochDates:
    """Tests for dates in and before Excel's 1900 leap-year-bug window.

    Excel's serial date numbering assumes a phantom 1900-02-29 that never
    existed, so every date from 1900-01-01 through 1900-02-28 (serials that
    would land below 61) cannot be represented as a correct serial number.
    Those dates fall back to being written as plain strings; 1900-03-01
    (serial 61) is the first date Excel can represent correctly.
    """

    def test_pre_epoch_date_csv_becomes_string(self, tmp_xlsx_factory: Callable[..., str]) -> None:
        """CSV dates before 1900 are written as strings, not invalid serial numbers."""
        import csv

        csv_path = tmp_xlsx_factory(".csv")
        xlsx_path = tmp_xlsx_factory()
        with Path(csv_path).open("w", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["date", "value"])
            writer.writerow(["1899-01-01", "old"])
            writer.writerow(["2024-01-15", "new"])
        xlsxturbo.csv_to_xlsx(csv_path, xlsx_path)
        wb = load_workbook(xlsx_path)
        ws = active_ws(wb)
        # Pre-epoch date should be a string
        assert ws["A2"].value == "1899-01-01"
        wb.close()

    def test_1900_leap_year_bug_window_csv_becomes_string(self, tmp_xlsx_factory: Callable[..., str]) -> None:
        """CSV dates inside the 1900 leap-year-bug window (Jan/Feb 1900) become strings."""
        import csv

        csv_path = tmp_xlsx_factory(".csv")
        xlsx_path = tmp_xlsx_factory()
        with Path(csv_path).open("w", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["date"])
            writer.writerow(["1900-01-15"])
        xlsxturbo.csv_to_xlsx(csv_path, xlsx_path)
        wb = load_workbook(xlsx_path)
        ws = active_ws(wb)
        assert ws["A2"].value == "1900-01-15"
        wb.close()

    def test_first_real_date_csv_becomes_date(self, tmp_xlsx_factory: Callable[..., str]) -> None:
        """CSV date 1900-03-01 (Excel serial 61) is the first date written as a real date."""
        import csv
        from datetime import datetime

        csv_path = tmp_xlsx_factory(".csv")
        xlsx_path = tmp_xlsx_factory()
        with Path(csv_path).open("w", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["date"])
            writer.writerow(["1900-03-01"])
        xlsxturbo.csv_to_xlsx(csv_path, xlsx_path)
        wb = load_workbook(xlsx_path)
        ws = active_ws(wb)
        value = ws["A2"].value
        assert isinstance(value, datetime)
        assert (value.year, value.month, value.day) == (1900, 3, 1)
        wb.close()

    def test_1900_leap_year_bug_window_dataframe_becomes_string(self, tmp_xlsx: str) -> None:
        """A datetime.date inside the 1900 leap-year-bug window is written as a string."""
        import datetime

        df = pd.DataFrame({"d": [datetime.date(1900, 1, 15)]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx)
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        assert ws["A2"].value == "1900-01-15"
        wb.close()

    def test_first_real_date_dataframe_becomes_date(self, tmp_xlsx: str) -> None:
        """datetime.date(1900, 3, 1) (Excel serial 61) is written as a real date."""
        import datetime

        df = pd.DataFrame({"d": [datetime.date(1900, 3, 1)]})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx)
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        value = ws["A2"].value
        assert value is not None
        assert (value.year, value.month, value.day) == (1900, 3, 1)
        wb.close()

    def test_1900_leap_year_bug_window_datetime64_becomes_string(self, tmp_xlsx: str) -> None:
        """A datetime64[ns] column value inside the 1900 leap-year-bug window is written as a string.

        Regression coverage for the datetime64 guard at src/write.rs:232-240,
        which is separate from the datetime.date path exercised above. The
        fallback string is numpy's own str() of the datetime64 scalar (a full
        timestamp), not the plain "YYYY-MM-DD" used by the datetime.date path.
        """
        df = pd.DataFrame({"d": pd.to_datetime(["1900-01-15"])})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx)
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        value = ws["A2"].value
        assert isinstance(value, str)
        assert value.startswith("1900-01-15")
        wb.close()

    def test_first_real_date_datetime64_becomes_date(self, tmp_xlsx: str) -> None:
        """A datetime64[ns] column value of 1900-03-01 (Excel serial 61) is written as a real date."""
        df = pd.DataFrame({"d": pd.to_datetime(["1900-03-01"])})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx)
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        value = ws["A2"].value
        assert value is not None
        assert (value.year, value.month, value.day) == (1900, 3, 1)
        wb.close()


class TestDatetimeDateSubclasses:
    """Tests for datetime.datetime/datetime.date subclasses in object-dtype columns.

    The Rust writer now performs a typed `PyDateTime`/`PyDate` cast (an
    isinstance-style check) instead of dispatching on the exact class name, so
    subclasses of datetime/date (e.g. pendulum.DateTime, freezegun's
    FakeDatetime, or a user-defined subclass) are written as real Excel
    datetimes/dates rather than falling through to a plain str() cell.
    """

    def test_datetime_subclass_written_as_datetime(self, tmp_xlsx: str) -> None:
        """A datetime.datetime subclass instance is written as a real datetime, not str()."""
        import datetime

        class SubDT(datetime.datetime):
            """A trivial datetime.datetime subclass used to exercise the typed cast path."""

        df = pd.DataFrame({"d": pd.array([SubDT(2024, 6, 15, 10, 30, 0)], dtype=object)})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx)
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        value = ws["A2"].value
        assert isinstance(value, datetime.datetime)
        assert (value.year, value.month, value.day, value.hour, value.minute) == (
            2024,
            6,
            15,
            10,
            30,
        )
        wb.close()

    def test_date_subclass_written_as_date(self, tmp_xlsx: str) -> None:
        """A datetime.date subclass instance is written as a real date, not str()."""
        import datetime

        class SubDate(datetime.date):
            """A trivial datetime.date subclass used to exercise the typed cast path."""

        df = pd.DataFrame({"d": pd.array([SubDate(2024, 6, 15)], dtype=object)})
        xlsxturbo.df_to_xlsx(df, tmp_xlsx)
        wb = load_workbook(tmp_xlsx)
        ws = active_ws(wb)
        value = ws["A2"].value
        assert value is not None
        assert not isinstance(value, str)
        assert (value.year, value.month, value.day) == (2024, 6, 15)
        wb.close()
