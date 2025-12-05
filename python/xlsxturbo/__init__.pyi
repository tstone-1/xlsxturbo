"""Type stubs for xlsxturbo"""

def csv_to_xlsx(
    input_path: str,
    output_path: str,
    sheet_name: str = "Sheet1",
    parallel: bool = False,
) -> tuple[int, int]:
    """
    Convert a CSV file to XLSX format with automatic type detection.

    Args:
        input_path: Path to the input CSV file
        output_path: Path for the output XLSX file
        sheet_name: Name of the worksheet (default: "Sheet1")
        parallel: Use multi-core parallel processing (default: False).
                  Faster for large files (100K+ rows) but uses more memory.

    Returns:
        Tuple of (rows, columns) written to the Excel file

    Raises:
        ValueError: If the conversion fails

    Example:
        >>> import xlsxturbo
        >>> rows, cols = xlsxturbo.csv_to_xlsx("data.csv", "output.xlsx")
        >>> # For large files, use parallel processing:
        >>> rows, cols = xlsxturbo.csv_to_xlsx("big.csv", "out.xlsx", parallel=True)
    """
    ...

def df_to_xlsx(
    df: object,
    output_path: str,
    sheet_name: str = "Sheet1",
    header: bool = True,
    autofit: bool = False,
    table_style: str | None = None,
    freeze_panes: bool = False,
) -> tuple[int, int]:
    """
    Convert a pandas or polars DataFrame to XLSX format.

    This function writes a DataFrame directly to an Excel XLSX file,
    preserving data types without intermediate CSV conversion.

    Args:
        df: pandas DataFrame or polars DataFrame to export
        output_path: Path for the output XLSX file
        sheet_name: Name of the worksheet (default: "Sheet1")
        header: Include column names as header row (default: True)
        autofit: Automatically adjust column widths to fit content (default: False)
        table_style: Apply Excel table formatting with this style name (default: None).
            Styles: "Light1"-"Light21", "Medium1"-"Medium28", "Dark1"-"Dark11", "None".
            Tables include autofilter dropdowns and banded rows.
        freeze_panes: Freeze the header row for easier scrolling (default: False)

    Returns:
        Tuple of (rows, columns) written to the Excel file

    Raises:
        ValueError: If the conversion fails

    Example:
        >>> import xlsxturbo
        >>> import pandas as pd
        >>> df = pd.DataFrame({'name': ['Alice', 'Bob'], 'age': [30, 25]})
        >>> rows, cols = xlsxturbo.df_to_xlsx(df, "output.xlsx")
        >>> print(f"Wrote {rows} rows and {cols} columns")
        >>> # With table formatting and auto-width columns:
        >>> xlsxturbo.df_to_xlsx(df, "styled.xlsx", table_style="Medium9", autofit=True, freeze_panes=True)
    """
    ...

def dfs_to_xlsx(
    sheets: list[tuple[object, str]],
    output_path: str,
    header: bool = True,
    autofit: bool = False,
    table_style: str | None = None,
    freeze_panes: bool = False,
) -> list[tuple[int, int]]:
    """
    Write multiple DataFrames to separate sheets in a single workbook.

    This is a convenience function that writes multiple DataFrames to
    separate sheets in one workbook, which is more efficient than
    calling df_to_xlsx multiple times.

    Args:
        sheets: List of (DataFrame, sheet_name) tuples
        output_path: Path for the output XLSX file
        header: Include column names as header row (default: True)
        autofit: Automatically adjust column widths to fit content (default: False)
        table_style: Apply Excel table formatting with this style name (default: None).
            Styles: "Light1"-"Light21", "Medium1"-"Medium28", "Dark1"-"Dark11", "None".
            Tables include autofilter dropdowns and banded rows.
        freeze_panes: Freeze the header row for easier scrolling (default: False)

    Returns:
        List of (rows, columns) tuples for each sheet

    Raises:
        ValueError: If the conversion fails

    Example:
        >>> import xlsxturbo
        >>> import pandas as pd
        >>> df1 = pd.DataFrame({'a': [1, 2]})
        >>> df2 = pd.DataFrame({'b': [3, 4]})
        >>> xlsxturbo.dfs_to_xlsx([(df1, "Sheet1"), (df2, "Sheet2")], "out.xlsx")
        >>> # With styling applied to all sheets:
        >>> xlsxturbo.dfs_to_xlsx([(df1, "Sales"), (df2, "Regions")], "report.xlsx",
        ...                       table_style="Medium9", autofit=True, freeze_panes=True)
    """
    ...

def version() -> str:
    """Get the version of the xlsxturbo library."""
    ...

__version__: str
