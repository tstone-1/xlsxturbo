"""Tests for core df_to_xlsx, csv_to_xlsx, and related behavior."""

from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd
import polars as pl
import pytest
import xlsxturbo

from tests.helpers import HAS_OPENPYXL, active_ws, get_temp_path, load_workbook

pytestmark = pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required for content verification")


class TestBackwardCompatibility:
    """Ensure existing functionality still works."""

    def test_old_column_widths_still_works(self) -> None:
        """Integer key column_widths still works."""
        df = pd.DataFrame({"A": [1], "B": [2]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_widths={0: 20, 1: 30})
            assert Path(path).exists()
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws.column_dimensions["A"].width > 15
                assert ws.column_dimensions["B"].width > 25
                wb.close()
        finally:
            Path(path).unlink()

    def test_table_style_without_table_name(self) -> None:
        """table_style works without table_name (existing behavior)."""
        df = pd.DataFrame({"A": [1, 2, 3]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, table_style="Medium9")
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert len(ws.tables) == 1
                wb.close()
        finally:
            Path(path).unlink()


class TestPolarsSupport:
    """Ensure all features work with polars DataFrames."""

    def test_polars_column_width_cap(self) -> None:
        """Polars column width cap is respected."""
        df = pl.DataFrame({"A": ["x" * 100], "B": ["y" * 100]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_widths={"_all": 20})
            assert Path(path).exists()
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws.column_dimensions["A"].width <= 21
                wb.close()
        finally:
            Path(path).unlink()

    def test_polars_datetime_fractional_seconds(self) -> None:
        """Polars datetime columns write as Excel datetimes.

        Polars `iter_rows` yields Python `datetime.datetime` objects, so this
        hits the "datetime" branch of write_py_value_with_format -- a distinct
        path from the pandas datetime64/Timestamp branches.
        """
        from datetime import datetime

        df = pl.DataFrame({"t": [datetime(2024, 1, 1, 12, 34, 56, 789000)]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path)
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["A2"].value == datetime(2024, 1, 1, 12, 34, 56, 789000)
                wb.close()
        finally:
            Path(path).unlink()

    def test_polars_table_name(self) -> None:
        """Polars DataFrame with a custom table name creates that table."""
        df = pl.DataFrame({"A": [1, 2, 3]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, table_style="Medium2", table_name="PolarsTable")
            assert Path(path).exists()
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert "PolarsTable" in ws.tables
                wb.close()
        finally:
            Path(path).unlink()

    def test_polars_header_format(self) -> None:
        """Polars DataFrame honors header_format."""
        df = pl.DataFrame({"A": [1], "B": [2]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, header_format={"bold": True})
            assert Path(path).exists()
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["A1"].font.bold is True
                wb.close()
        finally:
            Path(path).unlink()


class TestEdgeCases:
    """Tests for edge cases and error handling."""

    def test_empty_dataframe(self) -> None:
        """Empty DataFrame writes successfully."""
        df = pd.DataFrame({"A": [], "B": []})
        path = get_temp_path()
        try:
            rows, cols = xlsxturbo.df_to_xlsx(df, path)
            assert rows == 1  # Just header
            assert cols == 2
            assert Path(path).exists()
        finally:
            Path(path).unlink()

    def test_pathlib_output_path(self) -> None:
        """pathlib.Path is accepted for output_path."""
        df = pd.DataFrame({"A": [1]})
        path = Path(get_temp_path())
        try:
            rows, cols = xlsxturbo.df_to_xlsx(df, path)
            assert rows == 2
            assert cols == 1
            assert path.exists()
        finally:
            path.unlink()

    def test_empty_dataframe_with_table_style(self) -> None:
        """Empty DataFrame with table_style writes without creating table."""
        df = pd.DataFrame({"A": [], "B": []})
        path = get_temp_path()
        try:
            _rows, _cols = xlsxturbo.df_to_xlsx(df, path, table_style="Medium2")
            assert Path(path).exists()
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                # No table should be created for empty DataFrame
                assert len(ws.tables) == 0
                wb.close()
        finally:
            Path(path).unlink()

    def test_invalid_table_style_raises_error(self) -> None:
        """Invalid table_style raises ValueError."""
        df = pd.DataFrame({"A": [1, 2]})
        path = get_temp_path()
        try:
            with pytest.raises(ValueError, match="Unknown table_style") as exc_info:
                xlsxturbo.df_to_xlsx(df, path, table_style="InvalidStyle")
            message = str(exc_info.value)
            assert "Unknown table_style" in message
            assert "InvalidStyle" in message
        finally:
            Path(path).unlink(missing_ok=True)

    def test_invalid_hex_color_raises_error(self) -> None:
        """Invalid hex color format raises ValueError."""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            with pytest.raises(ValueError, match="expected 6 characters"):
                xlsxturbo.df_to_xlsx(df, path, header_format={"bg_color": "#FF"})
        finally:
            Path(path).unlink(missing_ok=True)

    def test_column_formats_order_preserved(self) -> None:
        """Column format patterns are matched in order (first match wins)."""
        df = pd.DataFrame({"price_usd": [1.0], "price_eur": [2.0], "other": [3.0]})
        path = get_temp_path()
        try:
            # The more specific pattern should be listed first to take priority
            xlsxturbo.df_to_xlsx(
                df,
                path,
                column_formats={
                    "price_usd": {"bg_color": "#FF0000"},  # Specific - should match first
                    "price_*": {"bg_color": "#0000FF"},  # General - should match price_eur
                },
            )
            assert Path(path).exists()
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                # price_usd should be red (specific match)
                assert ws["A2"].fill.fgColor.rgb == "FFFF0000"
                # price_eur should be blue (wildcard match)
                assert ws["B2"].fill.fgColor.rgb == "FF0000FF"
                # other should have no background
                wb.close()
        finally:
            Path(path).unlink()

    def test_empty_dataframe_no_header(self) -> None:
        """Empty DataFrame with header=False."""
        df = pd.DataFrame({"A": [], "B": []})
        path = get_temp_path()
        try:
            rows, cols = xlsxturbo.df_to_xlsx(df, path, header=False)
            assert rows == 0
            assert cols == 2
            assert Path(path).exists()
        finally:
            Path(path).unlink()

    def test_table_style_skipped_when_header_false(self) -> None:
        """table_style is skipped when header=False (tables require headers)."""
        df = pd.DataFrame({"A": [1, 2]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, header=False, table_style="Medium2")
            assert Path(path).exists()
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["A1"].value == 1  # data in row 1, no header
                wb.close()
        finally:
            Path(path).unlink()

    def test_dfs_per_sheet_header_false_with_global_table_style(self) -> None:
        """Per-sheet header=False skips table even with global table_style."""
        df1 = pd.DataFrame({"A": [1]})
        df2 = pd.DataFrame({"B": [2]})
        path = get_temp_path()
        try:
            xlsxturbo.dfs_to_xlsx([
                (df1, "WithHeader"),
                (df2, "NoHeader", {"header": False}),
            ], path, table_style="Medium2")
            assert Path(path).exists()
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                # WithHeader sheet should have table
                assert wb["WithHeader"]["A1"].value == "A"
                # NoHeader sheet should have data in row 1
                assert wb["NoHeader"]["A1"].value == 2
                wb.close()
        finally:
            Path(path).unlink()


class TestDateOrder:
    """Tests for date_order parameter in csv_to_xlsx."""

    def test_date_order_us_parses_mdy(self) -> None:
        """US date order parses 01-02-2024 as January 2."""
        import csv
        from datetime import datetime

        # Create CSV with ambiguous date
        csv_path = get_temp_path().replace(".xlsx", ".csv")
        xlsx_path = get_temp_path()
        try:
            with Path(csv_path).open("w", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["date", "value"])
                writer.writerow(["01-02-2024", "100"])

            xlsxturbo.csv_to_xlsx(csv_path, xlsx_path, date_order="us")

            if HAS_OPENPYXL:
                wb = load_workbook(xlsx_path)
                ws = active_ws(wb)
                cell_value = ws["A2"].value
                # openpyxl returns datetime objects for Excel dates
                assert isinstance(cell_value, datetime), f"Expected datetime, got {type(cell_value)}"
                # US format: 01-02-2024 = January 2, 2024
                assert cell_value.month == 1, f"Expected January (1), got month {cell_value.month}"
                assert cell_value.day == 2, f"Expected day 2, got day {cell_value.day}"
                wb.close()
        finally:
            Path(csv_path).unlink(missing_ok=True)
            Path(xlsx_path).unlink(missing_ok=True)

    def test_date_order_eu_parses_dmy(self) -> None:
        """European date order parses 01-02-2024 as February 1."""
        import csv
        from datetime import datetime

        csv_path = get_temp_path().replace(".xlsx", ".csv")
        xlsx_path = get_temp_path()
        try:
            with Path(csv_path).open("w", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["date", "value"])
                writer.writerow(["01-02-2024", "100"])

            xlsxturbo.csv_to_xlsx(csv_path, xlsx_path, date_order="eu")

            if HAS_OPENPYXL:
                wb = load_workbook(xlsx_path)
                ws = active_ws(wb)
                cell_value = ws["A2"].value
                assert isinstance(cell_value, datetime), f"Expected datetime, got {type(cell_value)}"
                # EU format: 01-02-2024 = February 1, 2024
                assert cell_value.month == 2, f"Expected February (2), got month {cell_value.month}"
                assert cell_value.day == 1, f"Expected day 1, got day {cell_value.day}"
                wb.close()
        finally:
            Path(csv_path).unlink(missing_ok=True)
            Path(xlsx_path).unlink(missing_ok=True)

    def test_date_order_produces_different_results(self) -> None:
        """US and EU date orders produce different Excel dates for ambiguous input."""
        import csv

        csv_path = get_temp_path().replace(".xlsx", ".csv")
        xlsx_us = get_temp_path()
        xlsx_eu = get_temp_path()
        try:
            with Path(csv_path).open("w", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["date"])
                writer.writerow(["03-04-2024"])  # Mar 4 (US) vs Apr 3 (EU)

            xlsxturbo.csv_to_xlsx(csv_path, xlsx_us, date_order="us")
            xlsxturbo.csv_to_xlsx(csv_path, xlsx_eu, date_order="eu")

            if HAS_OPENPYXL:
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
        finally:
            Path(csv_path).unlink(missing_ok=True)
            Path(xlsx_us).unlink(missing_ok=True)
            Path(xlsx_eu).unlink(missing_ok=True)

    def test_invalid_date_order_raises(self) -> None:
        """Invalid date_order raises ValueError."""
        import csv

        csv_path = get_temp_path().replace(".xlsx", ".csv")
        xlsx_path = get_temp_path()
        try:
            with Path(csv_path).open("w", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["a"])
                writer.writerow(["1"])

            with pytest.raises(ValueError, match=r"(?i)invalid") as exc_info:
                xlsxturbo.csv_to_xlsx(csv_path, xlsx_path, date_order="invalid")  # type: ignore[arg-type]  # invalid on purpose
            message = str(exc_info.value)
            assert "invalid" in message.lower()
        finally:
            Path(csv_path).unlink(missing_ok=True)
            Path(xlsx_path).unlink(missing_ok=True)


class TestCsvConversion:
    """Tests for csv_to_xlsx function."""

    def test_basic_csv(self) -> None:
        """Basic CSV with header and data rows."""
        import csv

        csv_path = get_temp_path().replace(".xlsx", ".csv")
        xlsx_path = get_temp_path()
        try:
            with Path(csv_path).open("w", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["name", "age", "score"])
                writer.writerow(["Alice", "30", "95.5"])
                writer.writerow(["Bob", "25", "88.0"])

            rows, cols = xlsxturbo.csv_to_xlsx(csv_path, xlsx_path)
            assert rows == 3
            assert cols == 3
            if HAS_OPENPYXL:
                wb = load_workbook(xlsx_path)
                ws = active_ws(wb)
                assert ws["A1"].value == "name"
                assert ws["B2"].value == 30  # should be detected as integer
                assert ws["C2"].value == 95.5  # should be detected as float
                wb.close()
        finally:
            Path(csv_path).unlink(missing_ok=True)
            Path(xlsx_path).unlink(missing_ok=True)

    def test_csv_type_detection(self) -> None:
        """CSV type detection: int, float, bool, date, string."""
        import csv

        csv_path = get_temp_path().replace(".xlsx", ".csv")
        xlsx_path = get_temp_path()
        try:
            with Path(csv_path).open("w", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["int", "float", "bool", "date", "text"])
                writer.writerow(["42", "3.14", "true", "2024-01-15", "hello"])

            xlsxturbo.csv_to_xlsx(csv_path, xlsx_path)
            if HAS_OPENPYXL:
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
        finally:
            Path(csv_path).unlink(missing_ok=True)
            Path(xlsx_path).unlink(missing_ok=True)

    def test_csv_datetime_fractional_seconds(self) -> None:
        """CSV datetime values preserve fractional seconds in Excel serials."""
        import csv
        from datetime import datetime

        csv_path = get_temp_path().replace(".xlsx", ".csv")
        xlsx_path = get_temp_path()
        try:
            with Path(csv_path).open("w", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["timestamp"])
                writer.writerow(["2024-01-01T12:34:56.789"])

            xlsxturbo.csv_to_xlsx(csv_path, xlsx_path)
            if HAS_OPENPYXL:
                wb = load_workbook(xlsx_path)
                ws = active_ws(wb)
                value = ws["A2"].value
                assert isinstance(value, datetime)
                assert value == datetime(2024, 1, 1, 12, 34, 56, 789000)
                wb.close()
        finally:
            Path(csv_path).unlink(missing_ok=True)
            Path(xlsx_path).unlink(missing_ok=True)

    def test_csv_special_values(self) -> None:
        """CSV with NaN, Inf, empty cells."""
        import csv

        csv_path = get_temp_path().replace(".xlsx", ".csv")
        xlsx_path = get_temp_path()
        try:
            with Path(csv_path).open("w", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["a", "b", "c"])
                writer.writerow(["NaN", "Inf", ""])
                writer.writerow(["nan", "-Inf", "   "])

            rows, _cols = xlsxturbo.csv_to_xlsx(csv_path, xlsx_path)
            assert rows == 3
            if HAS_OPENPYXL:
                wb = load_workbook(xlsx_path)
                ws = active_ws(wb)
                # NaN/Inf/empty all become empty cells (write_cell -> CellValue::Empty
                # writes an empty string, which openpyxl reads back as None or "").
                for ref in ("A2", "B2", "C2", "A3", "B3", "C3"):
                    assert ws[ref].value in (None, ""), f"{ref} should be empty"
                wb.close()
        finally:
            Path(csv_path).unlink(missing_ok=True)
            Path(xlsx_path).unlink(missing_ok=True)

    def test_csv_int_min_writes_as_string_without_overflow(self) -> None:
        """i64::MIN via the CSV path (write_cell) is written as text, no overflow.

        Complements test_i64_min_writes_as_string_without_overflow, which covers
        the distinct DataFrame write_int path.
        """
        import csv as _csv

        int_min = -9223372036854775808  # i64::MIN
        csv_path = get_temp_path().replace(".xlsx", ".csv")
        xlsx_path = get_temp_path()
        try:
            with Path(csv_path).open("w", newline="") as f:
                w = _csv.writer(f)
                w.writerow(["n"])
                w.writerow([str(int_min)])
            xlsxturbo.csv_to_xlsx(csv_path, xlsx_path)
            if HAS_OPENPYXL:
                wb = load_workbook(xlsx_path)
                ws = active_ws(wb)
                assert ws["A2"].value == str(int_min)
                wb.close()
        finally:
            for p in (csv_path, xlsx_path):
                Path(p).unlink(missing_ok=True)

    def test_csv_parallel(self) -> None:
        """CSV parallel mode produces same output."""
        import csv

        csv_path = get_temp_path().replace(".xlsx", ".csv")
        xlsx_seq = get_temp_path()
        xlsx_par = get_temp_path()
        try:
            with Path(csv_path).open("w", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["num", "text"])
                for i in range(100):
                    writer.writerow([str(i), f"row_{i}"])

            rows_s, cols_s = xlsxturbo.csv_to_xlsx(csv_path, xlsx_seq, parallel=False)
            rows_p, cols_p = xlsxturbo.csv_to_xlsx(csv_path, xlsx_par, parallel=True)
            assert rows_s == rows_p
            assert cols_s == cols_p
            if HAS_OPENPYXL:
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
        finally:
            Path(csv_path).unlink(missing_ok=True)
            Path(xlsx_seq).unlink(missing_ok=True)
            Path(xlsx_par).unlink(missing_ok=True)

    def test_csv_with_sheet_name(self) -> None:
        """CSV conversion with custom sheet name."""
        import csv

        csv_path = get_temp_path().replace(".xlsx", ".csv")
        xlsx_path = get_temp_path()
        try:
            with Path(csv_path).open("w", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["a"])
                writer.writerow(["1"])

            xlsxturbo.csv_to_xlsx(csv_path, xlsx_path, sheet_name="MySheet")
            if HAS_OPENPYXL:
                wb = load_workbook(xlsx_path)
                assert "MySheet" in wb.sheetnames
                wb.close()
        finally:
            Path(csv_path).unlink(missing_ok=True)
            Path(xlsx_path).unlink(missing_ok=True)


class TestUnicodeAndSpecialData:
    """Tests for Unicode, mixed types, nulls, and CSV edge cases."""

    def test_unicode_column_names_and_data(self) -> None:
        """Unicode characters in column names and cell data."""
        df = pd.DataFrame({
            "\u4ef7\u683c": [100, 200],       # Chinese: "price"
            "Stra\u00dfe": ["Berlin", "M\u00fcnchen"],  # German: street, Munich
            "\u540d\u524d": ["\u592a\u90ce", "\u82b1\u5b50"],           # Japanese names
        })
        path = get_temp_path()
        try:
            rows, cols = xlsxturbo.df_to_xlsx(df, path)
            assert rows == 3  # header + 2 data rows
            assert cols == 3
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["A1"].value == "\u4ef7\u683c"
                assert ws["B1"].value == "Stra\u00dfe"
                assert ws["C1"].value == "\u540d\u524d"
                assert ws["B2"].value == "Berlin"
                assert ws["C2"].value == "\u592a\u90ce"
                wb.close()
        finally:
            Path(path).unlink()

    def test_emoji_in_data(self) -> None:
        """Emoji characters in cell values."""
        df = pd.DataFrame({
            "status": ["done", "pending"],
            "icon": ["\U0001f680", "\U0001f525"],
        })
        path = get_temp_path()
        try:
            rows, _cols = xlsxturbo.df_to_xlsx(df, path)
            assert rows == 3
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["B2"].value == "\U0001f680"
                assert ws["B3"].value == "\U0001f525"
                wb.close()
        finally:
            Path(path).unlink()

    def test_mixed_type_column(self) -> None:
        """Column with mixed int and string values (pandas object dtype)."""
        df = pd.DataFrame({"mixed": [1, "two", 3, "four", 5.5]})
        path = get_temp_path()
        try:
            rows, _cols = xlsxturbo.df_to_xlsx(df, path)
            assert rows == 6  # header + 5 rows
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["A2"].value == 1
                assert ws["A3"].value == "two"
                assert ws["A4"].value == 3
                assert ws["A5"].value == "four"
                assert ws["A6"].value == 5.5
                wb.close()
        finally:
            Path(path).unlink()

    def test_none_and_nat_values(self) -> None:
        """None, NaT, and pd.NA values write as empty cells."""
        df = pd.DataFrame({
            "a": [1, None, 3],
            "b": pd.array([10, pd.NA, 30], dtype="Int64"),
            "c": pd.to_datetime(["2024-01-01", "NaT", "2024-03-01"]),
        })
        path = get_temp_path()
        try:
            rows, _cols = xlsxturbo.df_to_xlsx(df, path)
            assert rows == 4  # header + 3 rows
            if HAS_OPENPYXL:
                wb = load_workbook(path)
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
        finally:
            Path(path).unlink()

    def test_pandas_datetime64_preserves_datetime_and_fractional_seconds(self) -> None:
        """Pandas datetime64[ns] columns write as datetimes, not strings."""
        from datetime import datetime

        df = pd.DataFrame({
            "timestamp": pd.to_datetime([
                "2024-01-01 12:34:56.789",
                "NaT",
            ])
        })
        path = get_temp_path()
        try:
            rows, cols = xlsxturbo.df_to_xlsx(df, path)
            assert rows == 3
            assert cols == 1
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["A2"].value == datetime(2024, 1, 1, 12, 34, 56, 789000)
                assert ws["A3"].value is None or ws["A3"].value == ""
                wb.close()
        finally:
            Path(path).unlink()

    def test_non_ns_datetime64_out_of_range_writes_correct_date(self) -> None:
        """datetime64[us] dates outside ns range must not wrap around."""
        from datetime import datetime

        df = pd.DataFrame({
            "timestamp": np.array(["3000-01-01T00:00:00"], dtype="datetime64[us]")
        })
        path = get_temp_path()
        try:
            rows, cols = xlsxturbo.df_to_xlsx(df, path)
            assert rows == 2
            assert cols == 1
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["A2"].value == datetime(3000, 1, 1, 0, 0, 0)
                wb.close()
        finally:
            Path(path).unlink()

    def test_python_int_beyond_i64_writes_as_string(self) -> None:
        """Oversized Python ints should not fall through to rounded f64."""
        value = 2**63 + 1025
        df = pd.DataFrame({"big": [value]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path)
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["A2"].value == str(value)
                wb.close()
        finally:
            Path(path).unlink()

    def test_i64_min_writes_as_string_without_overflow(self) -> None:
        """The signed minimum value must use the precision-preserving fallback."""
        value = np.iinfo(np.int64).min
        df = pd.DataFrame({"min": np.array([value], dtype=np.int64)})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path)
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["A2"].value == str(value)
                wb.close()
        finally:
            Path(path).unlink()

    def test_dataframe_pre_1900_datetime_writes_as_string(self) -> None:
        """DataFrame datetime paths match CSV behavior for Excel-unsupported dates."""
        from datetime import datetime

        df = pd.DataFrame({
            "old": pd.Series([datetime(1850, 1, 1, 12, 0, 0)], dtype=object)
        })
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path)
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["A2"].value == "1850-01-01 12:00:00"
                wb.close()
        finally:
            Path(path).unlink()

    def test_object_timestamp_fractional_seconds(self) -> None:
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
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path)
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["A2"].value == datetime(2024, 1, 1, 12, 34, 56, 789000)
                wb.close()
        finally:
            Path(path).unlink()

    def test_timezone_aware_datetime_writes_wall_clock(self) -> None:
        """Timezone-aware datetimes are written as their local wall-clock value.

        The UTC offset is intentionally dropped (Excel has no timezone concept).
        Characterization test pinning the documented contract: 12:00 US/Eastern
        is stored as 12:00, NOT converted to its 17:00 UTC equivalent.
        """
        from datetime import datetime

        df = pd.DataFrame(
            {"t": pd.to_datetime(["2024-01-01 12:00:00"]).tz_localize("US/Eastern")}
        )
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path)
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["A2"].value == datetime(2024, 1, 1, 12, 0, 0)
                wb.close()
        finally:
            Path(path).unlink()

    def test_all_none_column(self) -> None:
        """Column with all None values."""
        df = pd.DataFrame({"empty": [None, None, None]})
        path = get_temp_path()
        try:
            rows, cols = xlsxturbo.df_to_xlsx(df, path)
            assert rows == 4
            assert cols == 1
        finally:
            Path(path).unlink()

    def test_large_integers_written_as_strings(self) -> None:
        """Integers > 2^53 should be written as strings to prevent precision loss."""
        large_int = 9007199254740993  # 2^53 + 1
        df = pd.DataFrame({"id": [large_int, 42]})
        path = get_temp_path()
        try:
            rows, _cols = xlsxturbo.df_to_xlsx(df, path)
            assert rows == 3
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                # Large int should be written as string to preserve precision
                assert str(ws["A2"].value) == str(large_int)
                # Normal int should be a number
                assert ws["A3"].value == 42
                wb.close()
        finally:
            Path(path).unlink()

    def test_csv_with_bom(self) -> None:
        """CSV file with UTF-8 BOM."""
        csv_path = get_temp_path().replace(".xlsx", ".csv")
        xlsx_path = get_temp_path()
        try:
            with Path(csv_path).open("w", encoding="utf-8-sig") as f:
                f.write("name,value\nAlice,1\nBob,2\n")
            rows, cols = xlsxturbo.csv_to_xlsx(csv_path, xlsx_path)
            assert rows == 3  # header + 2 data rows
            assert cols == 2
        finally:
            Path(xlsx_path).unlink()
            Path(csv_path).unlink(missing_ok=True)

    def test_csv_with_crlf(self) -> None:
        """CSV file with Windows CRLF line endings."""
        csv_path = get_temp_path().replace(".xlsx", ".csv")
        xlsx_path = get_temp_path()
        try:
            with Path(csv_path).open("wb") as f:
                f.write(b"a,b\r\n1,2\r\n3,4\r\n")
            rows, cols = xlsxturbo.csv_to_xlsx(csv_path, xlsx_path)
            assert rows == 3
            assert cols == 2
        finally:
            Path(xlsx_path).unlink()
            Path(csv_path).unlink(missing_ok=True)

    def test_csv_quoted_fields_with_delimiters(self) -> None:
        """CSV with quoted fields containing commas and newlines."""
        csv_path = get_temp_path().replace(".xlsx", ".csv")
        xlsx_path = get_temp_path()
        try:
            with Path(csv_path).open("w", encoding="utf-8") as f:
                f.write('name,address\n"Smith, John","123 Main St"\n"Doe, Jane","456 Oak Ave"\n')
            rows, cols = xlsxturbo.csv_to_xlsx(csv_path, xlsx_path)
            assert rows == 3
            assert cols == 2
            if HAS_OPENPYXL:
                wb = load_workbook(xlsx_path)
                ws = active_ws(wb)
                assert ws["A2"].value == "Smith, John"
                assert ws["B2"].value == "123 Main St"
                wb.close()
        finally:
            Path(xlsx_path).unlink()
            Path(csv_path).unlink(missing_ok=True)

    def test_polars_unicode(self) -> None:
        """Unicode data through Polars DataFrames."""
        df = pl.DataFrame({
            "city": ["T\u00f6ky\u00f6", "Z\u00fcrich", "S\u00e3o Paulo"],
            "pop": [14000000, 420000, 12300000],
        })
        path = get_temp_path()
        try:
            rows, _cols = xlsxturbo.df_to_xlsx(df, path)
            assert rows == 4
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = active_ws(wb)
                assert ws["A2"].value == "T\u00f6ky\u00f6"
                wb.close()
        finally:
            Path(path).unlink()


class TestCsvErrorPaths:
    """Tests for CSV conversion error handling."""

    def test_csv_nonexistent_input_raises_error(self) -> None:
        """csv_to_xlsx with nonexistent input file raises ValueError with path info."""
        path = get_temp_path()
        try:
            with pytest.raises(ValueError, match="Failed to open"):
                xlsxturbo.csv_to_xlsx("/nonexistent/file.csv", path)
        finally:
            Path(path).unlink(missing_ok=True)


class TestPreEpochDates:
    """Tests for dates before Excel epoch (1900-01-01)."""

    def test_pre_epoch_date_csv_becomes_string(self) -> None:
        """CSV dates before 1900 are written as strings, not invalid serial numbers."""
        import csv

        csv_path = get_temp_path().replace(".xlsx", ".csv")
        xlsx_path = get_temp_path()
        try:
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
        finally:
            Path(csv_path).unlink(missing_ok=True)
            Path(xlsx_path).unlink()
