from tests.helpers import (
    HAS_OPENPYXL,
    get_temp_path,
    load_workbook,
    os,
    pd,
    pl,
    pytest,
    tempfile,
    xlsxturbo,
)


pytestmark = pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required for content verification")


class TestBackwardCompatibility:
    """Ensure existing functionality still works"""

    def test_old_column_widths_still_works(self):
        """Integer key column_widths still works"""
        df = pd.DataFrame({"A": [1], "B": [2]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_widths={0: 20, 1: 30})
            assert os.path.exists(path)
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                assert ws.column_dimensions["A"].width > 15
                assert ws.column_dimensions["B"].width > 25
                wb.close()
        finally:
            os.unlink(path)

    def test_table_style_without_table_name(self):
        """table_style works without table_name (existing behavior)"""
        df = pd.DataFrame({"A": [1, 2, 3]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, table_style="Medium9")
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                assert len(ws.tables) == 1
                wb.close()
        finally:
            os.unlink(path)

class TestPolarsSupport:
    """Ensure all features work with polars DataFrames"""

    def test_polars_column_width_cap(self):
        df = pl.DataFrame({"A": ["x" * 100], "B": ["y" * 100]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, column_widths={"_all": 20})
            assert os.path.exists(path)
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                assert ws.column_dimensions["A"].width <= 21
                wb.close()
        finally:
            os.unlink(path)

    def test_polars_table_name(self):
        df = pl.DataFrame({"A": [1, 2, 3]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, table_style="Medium2", table_name="PolarsTable")
            assert os.path.exists(path)
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                assert "PolarsTable" in ws.tables
                wb.close()
        finally:
            os.unlink(path)

    def test_polars_header_format(self):
        df = pl.DataFrame({"A": [1], "B": [2]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, header_format={"bold": True})
            assert os.path.exists(path)
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                assert ws["A1"].font.bold == True
                wb.close()
        finally:
            os.unlink(path)

class TestEdgeCases:
    """Tests for edge cases and error handling"""

    def test_empty_dataframe(self):
        """Empty DataFrame writes successfully"""
        df = pd.DataFrame({"A": [], "B": []})
        path = get_temp_path()
        try:
            rows, cols = xlsxturbo.df_to_xlsx(df, path)
            assert rows == 1  # Just header
            assert cols == 2
            assert os.path.exists(path)
        finally:
            os.unlink(path)

    def test_empty_dataframe_with_table_style(self):
        """Empty DataFrame with table_style writes without creating table"""
        df = pd.DataFrame({"A": [], "B": []})
        path = get_temp_path()
        try:
            rows, cols = xlsxturbo.df_to_xlsx(df, path, table_style="Medium2")
            assert os.path.exists(path)
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                # No table should be created for empty DataFrame
                assert len(ws.tables) == 0
                wb.close()
        finally:
            os.unlink(path)

    def test_invalid_table_style_raises_error(self):
        """Invalid table_style raises ValueError"""
        df = pd.DataFrame({"A": [1, 2]})
        path = get_temp_path()
        try:
            try:
                xlsxturbo.df_to_xlsx(df, path, table_style="InvalidStyle")
                assert False, "Expected ValueError for invalid table_style"
            except ValueError as e:
                assert "Unknown table_style" in str(e)
                assert "InvalidStyle" in str(e)
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_invalid_hex_color_raises_error(self):
        """Invalid hex color format raises ValueError"""
        df = pd.DataFrame({"A": [1]})
        path = get_temp_path()
        try:
            try:
                xlsxturbo.df_to_xlsx(df, path, header_format={"bg_color": "#FF"})
                assert False, "Expected ValueError for invalid hex color"
            except ValueError as e:
                assert "expected 6 characters" in str(e)
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_column_formats_order_preserved(self):
        """Column format patterns are matched in order (first match wins)"""
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
            assert os.path.exists(path)
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                # price_usd should be red (specific match)
                assert ws["A2"].fill.fgColor.rgb == "FFFF0000"
                # price_eur should be blue (wildcard match)
                assert ws["B2"].fill.fgColor.rgb == "FF0000FF"
                # other should have no background
                wb.close()
        finally:
            os.unlink(path)

    def test_empty_dataframe_no_header(self):
        """Empty DataFrame with header=False"""
        df = pd.DataFrame({"A": [], "B": []})
        path = get_temp_path()
        try:
            rows, cols = xlsxturbo.df_to_xlsx(df, path, header=False)
            assert rows == 0
            assert cols == 2
            assert os.path.exists(path)
        finally:
            os.unlink(path)

    def test_table_style_skipped_when_header_false(self):
        """table_style is skipped when header=False (tables require headers)"""
        df = pd.DataFrame({"A": [1, 2]})
        path = get_temp_path()
        try:
            xlsxturbo.df_to_xlsx(df, path, header=False, table_style="Medium2")
            assert os.path.exists(path)
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                assert ws["A1"].value == 1  # data in row 1, no header
                wb.close()
        finally:
            os.unlink(path)

    def test_dfs_per_sheet_header_false_with_global_table_style(self):
        """Per-sheet header=False skips table even with global table_style"""
        df1 = pd.DataFrame({"A": [1]})
        df2 = pd.DataFrame({"B": [2]})
        path = get_temp_path()
        try:
            xlsxturbo.dfs_to_xlsx([
                (df1, "WithHeader"),
                (df2, "NoHeader", {"header": False}),
            ], path, table_style="Medium2")
            assert os.path.exists(path)
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                # WithHeader sheet should have table
                assert wb["WithHeader"]["A1"].value == "A"
                # NoHeader sheet should have data in row 1
                assert wb["NoHeader"]["A1"].value == 2
                wb.close()
        finally:
            os.unlink(path)

class TestDateOrder:
    """Tests for date_order parameter in csv_to_xlsx"""

    def test_date_order_us_parses_mdy(self):
        """US date order parses 01-02-2024 as January 2"""
        import csv
        from datetime import datetime

        # Create CSV with ambiguous date
        csv_path = get_temp_path().replace(".xlsx", ".csv")
        xlsx_path = get_temp_path()
        try:
            with open(csv_path, "w", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["date", "value"])
                writer.writerow(["01-02-2024", "100"])

            xlsxturbo.csv_to_xlsx(csv_path, xlsx_path, date_order="us")

            if HAS_OPENPYXL:
                wb = load_workbook(xlsx_path)
                ws = wb.active
                cell_value = ws["A2"].value
                # openpyxl returns datetime objects for Excel dates
                assert isinstance(cell_value, datetime), f"Expected datetime, got {type(cell_value)}"
                # US format: 01-02-2024 = January 2, 2024
                assert cell_value.month == 1, f"Expected January (1), got month {cell_value.month}"
                assert cell_value.day == 2, f"Expected day 2, got day {cell_value.day}"
                wb.close()
        finally:
            if os.path.exists(csv_path):
                os.unlink(csv_path)
            if os.path.exists(xlsx_path):
                os.unlink(xlsx_path)

    def test_date_order_eu_parses_dmy(self):
        """European date order parses 01-02-2024 as February 1"""
        import csv
        from datetime import datetime

        csv_path = get_temp_path().replace(".xlsx", ".csv")
        xlsx_path = get_temp_path()
        try:
            with open(csv_path, "w", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["date", "value"])
                writer.writerow(["01-02-2024", "100"])

            xlsxturbo.csv_to_xlsx(csv_path, xlsx_path, date_order="eu")

            if HAS_OPENPYXL:
                wb = load_workbook(xlsx_path)
                ws = wb.active
                cell_value = ws["A2"].value
                assert isinstance(cell_value, datetime), f"Expected datetime, got {type(cell_value)}"
                # EU format: 01-02-2024 = February 1, 2024
                assert cell_value.month == 2, f"Expected February (2), got month {cell_value.month}"
                assert cell_value.day == 1, f"Expected day 1, got day {cell_value.day}"
                wb.close()
        finally:
            if os.path.exists(csv_path):
                os.unlink(csv_path)
            if os.path.exists(xlsx_path):
                os.unlink(xlsx_path)

    def test_date_order_produces_different_results(self):
        """US and EU date orders produce different Excel dates for ambiguous input"""
        import csv
        from datetime import datetime

        csv_path = get_temp_path().replace(".xlsx", ".csv")
        xlsx_us = get_temp_path()
        xlsx_eu = get_temp_path()
        try:
            with open(csv_path, "w", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["date"])
                writer.writerow(["03-04-2024"])  # Mar 4 (US) vs Apr 3 (EU)

            xlsxturbo.csv_to_xlsx(csv_path, xlsx_us, date_order="us")
            xlsxturbo.csv_to_xlsx(csv_path, xlsx_eu, date_order="eu")

            if HAS_OPENPYXL:
                wb_us = load_workbook(xlsx_us)
                wb_eu = load_workbook(xlsx_eu)
                us_value = wb_us.active["A2"].value
                eu_value = wb_eu.active["A2"].value
                wb_us.close()
                wb_eu.close()

                # Values should be different dates
                assert us_value != eu_value, "US and EU should produce different dates"
                # US: March 4, EU: April 3 (30 days difference)
                diff = abs((us_value - eu_value).days)
                assert diff == 30, f"Expected 30 day difference, got {diff}"
        finally:
            if os.path.exists(csv_path):
                os.unlink(csv_path)
            if os.path.exists(xlsx_us):
                os.unlink(xlsx_us)
            if os.path.exists(xlsx_eu):
                os.unlink(xlsx_eu)

    def test_invalid_date_order_raises(self):
        """Invalid date_order raises ValueError"""
        import csv

        csv_path = get_temp_path().replace(".xlsx", ".csv")
        xlsx_path = get_temp_path()
        try:
            with open(csv_path, "w", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["a"])
                writer.writerow(["1"])

            try:
                xlsxturbo.csv_to_xlsx(csv_path, xlsx_path, date_order="invalid")
                assert False, "Expected ValueError"
            except ValueError as e:
                assert "invalid" in str(e).lower()
        finally:
            if os.path.exists(csv_path):
                os.unlink(csv_path)
            if os.path.exists(xlsx_path):
                os.unlink(xlsx_path)

class TestCsvConversion:
    """Tests for csv_to_xlsx function"""

    def test_basic_csv(self):
        """Basic CSV with header and data rows"""
        import csv

        csv_path = get_temp_path().replace(".xlsx", ".csv")
        xlsx_path = get_temp_path()
        try:
            with open(csv_path, "w", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["name", "age", "score"])
                writer.writerow(["Alice", "30", "95.5"])
                writer.writerow(["Bob", "25", "88.0"])

            rows, cols = xlsxturbo.csv_to_xlsx(csv_path, xlsx_path)
            assert rows == 3
            assert cols == 3
            if HAS_OPENPYXL:
                wb = load_workbook(xlsx_path)
                ws = wb.active
                assert ws["A1"].value == "name"
                assert ws["B2"].value == 30  # should be detected as integer
                assert ws["C2"].value == 95.5  # should be detected as float
                wb.close()
        finally:
            if os.path.exists(csv_path):
                os.unlink(csv_path)
            if os.path.exists(xlsx_path):
                os.unlink(xlsx_path)

    def test_csv_type_detection(self):
        """CSV type detection: int, float, bool, date, string"""
        import csv

        csv_path = get_temp_path().replace(".xlsx", ".csv")
        xlsx_path = get_temp_path()
        try:
            with open(csv_path, "w", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["int", "float", "bool", "date", "text"])
                writer.writerow(["42", "3.14", "true", "2024-01-15", "hello"])

            xlsxturbo.csv_to_xlsx(csv_path, xlsx_path)
            if HAS_OPENPYXL:
                wb = load_workbook(xlsx_path)
                ws = wb.active
                assert ws["A2"].value == 42
                assert abs(ws["B2"].value - 3.14) < 0.001
                assert ws["C2"].value is True
                # Date should be a datetime object in openpyxl
                from datetime import datetime

                assert isinstance(ws["D2"].value, datetime)
                assert ws["E2"].value == "hello"
                wb.close()
        finally:
            if os.path.exists(csv_path):
                os.unlink(csv_path)
            if os.path.exists(xlsx_path):
                os.unlink(xlsx_path)

    def test_csv_datetime_fractional_seconds(self):
        """CSV datetime values preserve fractional seconds in Excel serials"""
        import csv
        from datetime import datetime

        csv_path = get_temp_path().replace(".xlsx", ".csv")
        xlsx_path = get_temp_path()
        try:
            with open(csv_path, "w", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["timestamp"])
                writer.writerow(["2024-01-01T12:34:56.789"])

            xlsxturbo.csv_to_xlsx(csv_path, xlsx_path)
            if HAS_OPENPYXL:
                wb = load_workbook(xlsx_path)
                ws = wb.active
                value = ws["A2"].value
                assert isinstance(value, datetime)
                assert value == datetime(2024, 1, 1, 12, 34, 56, 789000)
                wb.close()
        finally:
            if os.path.exists(csv_path):
                os.unlink(csv_path)
            if os.path.exists(xlsx_path):
                os.unlink(xlsx_path)


    def test_csv_special_values(self):
        """CSV with NaN, Inf, empty cells"""
        import csv

        csv_path = get_temp_path().replace(".xlsx", ".csv")
        xlsx_path = get_temp_path()
        try:
            with open(csv_path, "w", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["a", "b", "c"])
                writer.writerow(["NaN", "Inf", ""])
                writer.writerow(["nan", "-Inf", "   "])

            rows, cols = xlsxturbo.csv_to_xlsx(csv_path, xlsx_path)
            assert rows == 3
            if HAS_OPENPYXL:
                wb = load_workbook(xlsx_path)
                ws = wb.active
                # NaN/Inf/empty should become empty strings or None
                # (written as empty string in write_cell for CellValue::Empty)
                wb.close()
        finally:
            if os.path.exists(csv_path):
                os.unlink(csv_path)
            if os.path.exists(xlsx_path):
                os.unlink(xlsx_path)

    def test_csv_parallel(self):
        """CSV parallel mode produces same output"""
        import csv

        csv_path = get_temp_path().replace(".xlsx", ".csv")
        xlsx_seq = get_temp_path()
        xlsx_par = get_temp_path()
        try:
            with open(csv_path, "w", newline="") as f:
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
                ws_s = wb_s.active
                ws_p = wb_p.active
                # Spot check some cells match
                for row in [1, 2, 50, 101]:
                    assert ws_s[f"A{row}"].value == ws_p[f"A{row}"].value
                    assert ws_s[f"B{row}"].value == ws_p[f"B{row}"].value
                wb_s.close()
                wb_p.close()
        finally:
            if os.path.exists(csv_path):
                os.unlink(csv_path)
            if os.path.exists(xlsx_seq):
                os.unlink(xlsx_seq)
            if os.path.exists(xlsx_par):
                os.unlink(xlsx_par)

    def test_csv_with_sheet_name(self):
        """CSV conversion with custom sheet name"""
        import csv

        csv_path = get_temp_path().replace(".xlsx", ".csv")
        xlsx_path = get_temp_path()
        try:
            with open(csv_path, "w", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["a"])
                writer.writerow(["1"])

            xlsxturbo.csv_to_xlsx(csv_path, xlsx_path, sheet_name="MySheet")
            if HAS_OPENPYXL:
                wb = load_workbook(xlsx_path)
                assert "MySheet" in wb.sheetnames
                wb.close()
        finally:
            if os.path.exists(csv_path):
                os.unlink(csv_path)
            if os.path.exists(xlsx_path):
                os.unlink(xlsx_path)

class TestUnicodeAndSpecialData:
    """Tests for Unicode, mixed types, nulls, and CSV edge cases."""

    def test_unicode_column_names_and_data(self):
        """Unicode characters in column names and cell data"""
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
                ws = wb.active
                assert ws["A1"].value == "\u4ef7\u683c"
                assert ws["B1"].value == "Stra\u00dfe"
                assert ws["C1"].value == "\u540d\u524d"
                assert ws["B2"].value == "Berlin"
                assert ws["C2"].value == "\u592a\u90ce"
                wb.close()
        finally:
            os.unlink(path)

    def test_emoji_in_data(self):
        """Emoji characters in cell values"""
        df = pd.DataFrame({
            "status": ["done", "pending"],
            "icon": ["\U0001f680", "\U0001f525"],
        })
        path = get_temp_path()
        try:
            rows, cols = xlsxturbo.df_to_xlsx(df, path)
            assert rows == 3
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                assert ws["B2"].value == "\U0001f680"
                assert ws["B3"].value == "\U0001f525"
                wb.close()
        finally:
            os.unlink(path)

    def test_mixed_type_column(self):
        """Column with mixed int and string values (pandas object dtype)"""
        df = pd.DataFrame({"mixed": [1, "two", 3, "four", 5.5]})
        path = get_temp_path()
        try:
            rows, cols = xlsxturbo.df_to_xlsx(df, path)
            assert rows == 6  # header + 5 rows
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                assert ws["A2"].value == 1
                assert ws["A3"].value == "two"
                assert ws["A4"].value == 3
                assert ws["A5"].value == "four"
                assert ws["A6"].value == 5.5
                wb.close()
        finally:
            os.unlink(path)

    def test_none_and_nat_values(self):
        """None, NaT, and pd.NA values write as empty cells"""
        df = pd.DataFrame({
            "a": [1, None, 3],
            "b": pd.array([10, pd.NA, 30], dtype="Int64"),
            "c": pd.to_datetime(["2024-01-01", pd.NaT, "2024-03-01"]),
        })
        path = get_temp_path()
        try:
            rows, cols = xlsxturbo.df_to_xlsx(df, path)
            assert rows == 4  # header + 3 rows
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
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
            os.unlink(path)

    def test_pandas_datetime64_preserves_datetime_and_fractional_seconds(self):
        """pandas datetime64[ns] columns write as datetimes, not strings"""
        from datetime import datetime

        df = pd.DataFrame({
            "timestamp": pd.to_datetime([
                "2024-01-01 12:34:56.789",
                pd.NaT,
            ])
        })
        path = get_temp_path()
        try:
            rows, cols = xlsxturbo.df_to_xlsx(df, path)
            assert rows == 3
            assert cols == 1
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                assert ws["A2"].value == datetime(2024, 1, 1, 12, 34, 56, 789000)
                assert ws["A3"].value is None or ws["A3"].value == ""
                wb.close()
        finally:
            os.unlink(path)

    def test_all_none_column(self):
        """Column with all None values"""
        df = pd.DataFrame({"empty": [None, None, None]})
        path = get_temp_path()
        try:
            rows, cols = xlsxturbo.df_to_xlsx(df, path)
            assert rows == 4
            assert cols == 1
        finally:
            os.unlink(path)

    def test_large_integers_written_as_strings(self):
        """Integers > 2^53 should be written as strings to prevent precision loss"""
        large_int = 9007199254740993  # 2^53 + 1
        df = pd.DataFrame({"id": [large_int, 42]})
        path = get_temp_path()
        try:
            rows, cols = xlsxturbo.df_to_xlsx(df, path)
            assert rows == 3
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                # Large int should be written as string to preserve precision
                assert str(ws["A2"].value) == str(large_int)
                # Normal int should be a number
                assert ws["A3"].value == 42
                wb.close()
        finally:
            os.unlink(path)

    def test_csv_with_bom(self):
        """CSV file with UTF-8 BOM"""
        csv_path = tempfile.mktemp(suffix=".csv")
        xlsx_path = get_temp_path()
        try:
            with open(csv_path, "w", encoding="utf-8-sig") as f:
                f.write("name,value\nAlice,1\nBob,2\n")
            rows, cols = xlsxturbo.csv_to_xlsx(csv_path, xlsx_path)
            assert rows == 3  # header + 2 data rows
            assert cols == 2
        finally:
            os.unlink(xlsx_path)
            if os.path.exists(csv_path):
                os.unlink(csv_path)

    def test_csv_with_crlf(self):
        """CSV file with Windows CRLF line endings"""
        csv_path = tempfile.mktemp(suffix=".csv")
        xlsx_path = get_temp_path()
        try:
            with open(csv_path, "wb") as f:
                f.write(b"a,b\r\n1,2\r\n3,4\r\n")
            rows, cols = xlsxturbo.csv_to_xlsx(csv_path, xlsx_path)
            assert rows == 3
            assert cols == 2
        finally:
            os.unlink(xlsx_path)
            if os.path.exists(csv_path):
                os.unlink(csv_path)

    def test_csv_quoted_fields_with_delimiters(self):
        """CSV with quoted fields containing commas and newlines"""
        csv_path = tempfile.mktemp(suffix=".csv")
        xlsx_path = get_temp_path()
        try:
            with open(csv_path, "w", encoding="utf-8") as f:
                f.write('name,address\n"Smith, John","123 Main St"\n"Doe, Jane","456 Oak Ave"\n')
            rows, cols = xlsxturbo.csv_to_xlsx(csv_path, xlsx_path)
            assert rows == 3
            assert cols == 2
            if HAS_OPENPYXL:
                wb = load_workbook(xlsx_path)
                ws = wb.active
                assert ws["A2"].value == "Smith, John"
                assert ws["B2"].value == "123 Main St"
                wb.close()
        finally:
            os.unlink(xlsx_path)
            if os.path.exists(csv_path):
                os.unlink(csv_path)

    def test_polars_unicode(self):
        """Unicode data through Polars DataFrames"""
        df = pl.DataFrame({
            "city": ["T\u00f6ky\u00f6", "Z\u00fcrich", "S\u00e3o Paulo"],
            "pop": [14000000, 420000, 12300000],
        })
        path = get_temp_path()
        try:
            rows, cols = xlsxturbo.df_to_xlsx(df, path)
            assert rows == 4
            if HAS_OPENPYXL:
                wb = load_workbook(path)
                ws = wb.active
                assert ws["A2"].value == "T\u00f6ky\u00f6"
                wb.close()
        finally:
            os.unlink(path)

class TestCsvErrorPaths:
    """Tests for CSV conversion error handling"""

    def test_csv_nonexistent_input_raises_error(self):
        """csv_to_xlsx with nonexistent input file raises ValueError with path info"""
        path = get_temp_path()
        try:
            with pytest.raises(ValueError, match="Failed to open"):
                xlsxturbo.csv_to_xlsx("/nonexistent/file.csv", path)
        finally:
            if os.path.exists(path):
                os.unlink(path)

class TestPreEpochDates:
    """Tests for dates before Excel epoch (1900-01-01)"""

    def test_pre_epoch_date_csv_becomes_string(self):
        """CSV dates before 1900 are written as strings, not invalid serial numbers"""
        import csv
        import tempfile

        csv_path = tempfile.mktemp(suffix=".csv")
        xlsx_path = get_temp_path()
        try:
            with open(csv_path, "w", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["date", "value"])
                writer.writerow(["1899-01-01", "old"])
                writer.writerow(["2024-01-15", "new"])
            xlsxturbo.csv_to_xlsx(csv_path, xlsx_path)
            wb = load_workbook(xlsx_path)
            ws = wb.active
            # Pre-epoch date should be a string
            assert ws["A2"].value == "1899-01-01"
            wb.close()
        finally:
            if os.path.exists(csv_path):
                os.unlink(csv_path)
            os.unlink(xlsx_path)


if __name__ == "__main__":
    import sys

    # Run simple tests without pytest
    test_classes = [
        TestColumnWidthCap,
        TestTableName,
        TestHeaderFormat,
        TestBackwardCompatibility,
        TestPolarsSupport,
        TestAllFeaturesCombined,
        TestEdgeCases,
        TestDateOrder,
        TestFormulaColumns,
        TestMergedRanges,
        TestHyperlinks,
        TestCsvConversion,
        TestComments,
        TestValidations,
        TestRichText,
        TestImages,
        TestV10AllFeatures,
        TestErrorPaths,
        TestConditionalFormatting,
        TestConstantMemoryMode,
        TestRowHeights,
        TestUnicodeAndSpecialData,
        TestDefinedNames,
        TestCells,
        TestCellsPerSheet,
        TestCellsFormatting,
        TestBorderStyles,
        TestTextAlignment,
        TestCellConditionalFormat,
        TestCsvErrorPaths,
        TestConstantMemoryWarning,
        TestDefinedNamesVerification,
        TestFormulaColumnsHeaderFalse,
        TestPreEpochDates,
    ]

    failed = 0
    passed = 0

    for test_class in test_classes:
        instance = test_class()
        for method_name in dir(instance):
            if method_name.startswith("test_"):
                try:
                    getattr(instance, method_name)()
                    print(f"[PASS] {test_class.__name__}.{method_name}")
                    passed += 1
                except Exception as e:
                    print(f"[FAIL] {test_class.__name__}.{method_name}: {e}")
                    failed += 1

    print(f"\n{passed} passed, {failed} failed")
    sys.exit(1 if failed else 0)
