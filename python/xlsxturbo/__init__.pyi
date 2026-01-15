"""Type stubs for xlsxturbo"""

from typing import Literal, TypedDict

DateOrder = Literal["auto", "mdy", "us", "dmy", "eu", "european"]

class HeaderFormat(TypedDict, total=False):
    """Header cell formatting options. All fields are optional."""
    bold: bool
    italic: bool
    font_color: str  # '#RRGGBB' or named color (white, black, red, blue, etc.)
    bg_color: str    # '#RRGGBB' or named color
    font_size: float
    underline: bool

class ColumnFormat(TypedDict, total=False):
    """Column cell formatting options. All fields are optional."""
    bold: bool
    italic: bool
    font_color: str  # '#RRGGBB' or named color (white, black, red, blue, etc.)
    bg_color: str    # '#RRGGBB' or named color
    font_size: float
    underline: bool
    num_format: str  # Excel number format string, e.g. '0.00', '#,##0', '0.00%'
    border: bool     # Add thin border around cells

class SheetOptions(TypedDict, total=False):
    """Per-sheet options for dfs_to_xlsx. All fields are optional."""
    header: bool
    autofit: bool
    table_style: str | None
    freeze_panes: bool
    column_widths: dict[int | str, float] | None  # Keys: int index or '_all'
    row_heights: dict[int, float] | None
    table_name: str | None
    header_format: HeaderFormat | None
    column_formats: dict[str, ColumnFormat] | None  # Pattern -> format. Patterns: 'prefix*', '*suffix', '*contains*', exact

def csv_to_xlsx(
    input_path: str,
    output_path: str,
    sheet_name: str = "Sheet1",
    parallel: bool = False,
    date_order: DateOrder = "auto",
) -> tuple[int, int]:
    """
    Convert a CSV file to XLSX format with automatic type detection.

    Args:
        input_path: Path to the input CSV file
        output_path: Path for the output XLSX file
        sheet_name: Name of the worksheet (default: "Sheet1")
        parallel: Use multi-core parallel processing (default: False).
                  Faster for large files (100K+ rows) but uses more memory.
        date_order: Date parsing order for ambiguous dates like "01-02-2024".
            "auto" - ISO first, then European (DMY), then US (MDY)
            "mdy" or "us" - US format: 01-02-2024 = January 2nd
            "dmy" or "eu" - European format: 01-02-2024 = February 1st

    Returns:
        Tuple of (rows, columns) written to the Excel file

    Raises:
        ValueError: If the conversion fails
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
    column_widths: dict[int | str, float] | None = None,
    row_heights: dict[int, float] | None = None,
    constant_memory: bool = False,
    table_name: str | None = None,
    header_format: HeaderFormat | None = None,
    column_formats: dict[str, ColumnFormat] | None = None,
) -> tuple[int, int]:
    """
    Convert a pandas or polars DataFrame to XLSX format.

    Args:
        df: pandas DataFrame or polars DataFrame to export
        output_path: Path for the output XLSX file
        sheet_name: Name of the worksheet (default: "Sheet1")
        header: Include column names as header row (default: True)
        autofit: Automatically adjust column widths to fit content (default: False)
        table_style: Apply Excel table formatting (default: None).
            Styles: "Light1"-"Light21", "Medium1"-"Medium28", "Dark1"-"Dark11", "None".
        freeze_panes: Freeze the header row for easier scrolling (default: False)
        column_widths: Dict mapping column index to width. Use '_all' to cap all columns.
        row_heights: Dict mapping row index to height in points.
        constant_memory: Use streaming mode for minimal RAM usage (default: False).
        table_name: Custom name for the Excel table (requires table_style).
        header_format: Dict of header cell formatting options.
        column_formats: Dict mapping column name patterns to format options.
            Patterns: 'prefix*', '*suffix', '*contains*', or exact match.
            First matching pattern wins (order preserved).
    """
    ...

def dfs_to_xlsx(
    sheets: list[tuple[object, str] | tuple[object, str, SheetOptions]],
    output_path: str,
    header: bool = True,
    autofit: bool = False,
    table_style: str | None = None,
    freeze_panes: bool = False,
    column_widths: dict[int | str, float] | None = None,
    row_heights: dict[int, float] | None = None,
    constant_memory: bool = False,
    table_name: str | None = None,
    header_format: HeaderFormat | None = None,
    column_formats: dict[str, ColumnFormat] | None = None,
) -> list[tuple[int, int]]:
    """
    Write multiple DataFrames to separate sheets in a single workbook.

    Args:
        sheets: List of (DataFrame, sheet_name) or (DataFrame, sheet_name, options) tuples.
        output_path: Path for the output XLSX file
        header: Include column names as header row (default: True)
        autofit: Automatically adjust column widths (default: False)
        table_style: Apply Excel table formatting (default: None).
        freeze_panes: Freeze the header row (default: False)
        column_widths: Dict mapping column index to width. Use '_all' to cap all columns.
        row_heights: Dict mapping row index to height in points.
        constant_memory: Use streaming mode (default: False).
        table_name: Custom name for Excel tables (requires table_style).
        header_format: Dict of header cell formatting options.
        column_formats: Dict mapping column name patterns to format options.
            Patterns: 'prefix*', '*suffix', '*contains*', or exact match.
    """
    ...

def version() -> str:
    """Get the version of the xlsxturbo library."""
    ...

__version__: str
