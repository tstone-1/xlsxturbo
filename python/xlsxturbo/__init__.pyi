"""Type stubs for xlsxturbo"""

from typing import TypedDict

class SheetOptions(TypedDict, total=False):
    """Per-sheet options for dfs_to_xlsx. All fields are optional."""
    header: bool
    autofit: bool
    table_style: str | None
    freeze_panes: bool
    column_widths: dict[int, float] | None
    row_heights: dict[int, float] | None

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
    column_widths: dict[int, float] | None = None,
    row_heights: dict[int, float] | None = None,
    constant_memory: bool = False,
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
        column_widths: Dict mapping column index (0-based) to width in characters.
            Example: {0: 25, 1: 15} sets column A to 25 and B to 15.
        row_heights: Dict mapping row index (0-based) to height in points.
            Example: {0: 22, 5: 30} sets row 1 to 22pt and row 6 to 30pt.
        constant_memory: Use streaming mode for minimal RAM usage (default: False).
            Ideal for very large files (millions of rows). Note: Disables
            table_style, freeze_panes, row_heights, and autofit.

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
        >>> # For very large files with minimal memory:
        >>> xlsxturbo.df_to_xlsx(big_df, "big.xlsx", constant_memory=True)
    """
    ...

def dfs_to_xlsx(
    sheets: list[tuple[object, str] | tuple[object, str, SheetOptions]],
    output_path: str,
    header: bool = True,
    autofit: bool = False,
    table_style: str | None = None,
    freeze_panes: bool = False,
    column_widths: dict[int, float] | None = None,
    row_heights: dict[int, float] | None = None,
    constant_memory: bool = False,
) -> list[tuple[int, int]]:
    """
    Write multiple DataFrames to separate sheets in a single workbook.

    This function writes multiple DataFrames to separate sheets in one workbook,
    which is more efficient than calling df_to_xlsx multiple times.

    Supports per-sheet options: each sheet can override global defaults by
    providing a 3-tuple (df, sheet_name, options_dict) instead of 2-tuple.

    Args:
        sheets: List of tuples. Each tuple can be:
            - (DataFrame, sheet_name) - uses global defaults
            - (DataFrame, sheet_name, options_dict) - per-sheet overrides
            Options dict keys: header, autofit, table_style, freeze_panes,
            column_widths, row_heights
        output_path: Path for the output XLSX file
        header: Include column names as header row (default: True)
        autofit: Automatically adjust column widths to fit content (default: False)
        table_style: Apply Excel table formatting with this style name (default: None).
            Styles: "Light1"-"Light21", "Medium1"-"Medium28", "Dark1"-"Dark11", "None".
            Tables include autofilter dropdowns and banded rows.
        freeze_panes: Freeze the header row for easier scrolling (default: False)
        column_widths: Dict mapping column index (0-based) to width in characters.
            Applied to all sheets unless overridden. Example: {0: 25, 1: 15}.
        row_heights: Dict mapping row index (0-based) to height in points.
            Applied to all sheets unless overridden. Example: {0: 22, 5: 30}.
        constant_memory: Use streaming mode for minimal RAM usage (default: False).
            Note: Disables table_style, freeze_panes, row_heights, and autofit.

    Returns:
        List of (rows, columns) tuples for each sheet

    Raises:
        ValueError: If the conversion fails

    Example:
        >>> import xlsxturbo
        >>> import pandas as pd
        >>> df1 = pd.DataFrame({'a': [1, 2]})
        >>> df2 = pd.DataFrame({'b': [3, 4]})
        >>> # Old API still works:
        >>> xlsxturbo.dfs_to_xlsx([(df1, "Sheet1"), (df2, "Sheet2")], "out.xlsx")
        >>> # With per-sheet options (header=False for one sheet):
        >>> xlsxturbo.dfs_to_xlsx([
        ...     (df1, "Data", {"header": True, "table_style": "Medium2"}),
        ...     (df2, "Instructions", {"header": False})
        ... ], "report.xlsx", autofit=True)
    """
    ...

def version() -> str:
    """Get the version of the xlsxturbo library."""
    ...

__version__: str
