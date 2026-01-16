"""Type stubs for xlsxturbo"""

from typing import Literal, TypedDict

DateOrder = Literal["auto", "mdy", "us", "dmy", "eu", "european"]
ValidationType = Literal["list", "whole_number", "decimal", "text_length"]

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

class ConditionalFormat(TypedDict, total=False):
    """Conditional formatting options for a column. 'type' is required.

    Supported types:
    - '2_color_scale': Gradient from min_color to max_color
    - '3_color_scale': Gradient with min_color, mid_color, max_color
    - 'data_bar': In-cell bar chart
    - 'icon_set': Traffic lights, arrows, or other icons
    """
    type: str  # Required: '2_color_scale', '3_color_scale', 'data_bar', 'icon_set'
    # For color scales:
    min_color: str   # '#RRGGBB' or named color for minimum value
    mid_color: str   # '#RRGGBB' or named color for midpoint (3_color_scale only)
    max_color: str   # '#RRGGBB' or named color for maximum value
    # For data bars:
    bar_color: str   # '#RRGGBB' or named color for the bar fill
    border_color: str  # '#RRGGBB' or named color for bar border
    solid: bool      # True for solid fill, False for gradient (default)
    direction: str   # 'left_to_right', 'right_to_left', or 'context' (default)
    # For icon sets:
    icon_type: str   # '3_arrows', '3_traffic_lights', '3_flags', '4_arrows', '5_arrows', etc. (see README for full list)
    reverse: bool    # Reverse icon order
    icons_only: bool # Show only icons, hide values

class CommentOptions(TypedDict, total=False):
    """Options for cell comments/notes.

    Note: 'text' is required at runtime but TypedDict doesn't enforce this.
    """
    text: str    # The comment text (required at runtime)
    author: str  # Author name for the comment

class ValidationOptions(TypedDict, total=False):
    """Data validation options for a column. 'type' is required.

    Supported types:
    - 'list': Dropdown with specified values
    - 'whole_number': Integer between min and max
    - 'decimal': Decimal number between min and max
    - 'text_length': Text length between min and max
    """
    type: ValidationType  # Required: validation type
    values: list[str]  # For 'list' type: dropdown options
    min: int | float   # For number/text_length: minimum value
    max: int | float   # For number/text_length: maximum value
    input_title: str   # Title for input prompt
    input_message: str # Message for input prompt
    error_title: str   # Title for error message
    error_message: str # Message for error message

class RichTextFormat(TypedDict, total=False):
    """Format options for a rich text segment."""
    bold: bool
    italic: bool
    font_color: str  # '#RRGGBB' or named color
    bg_color: str    # '#RRGGBB' or named color
    font_size: float
    underline: bool

class ImageOptions(TypedDict, total=False):
    """Options for embedding images.

    Note: 'path' is required at runtime but TypedDict doesn't enforce this.
    """
    path: str            # Path to image file - PNG, JPEG, GIF, BMP (required at runtime)
    scale_width: float   # Scale factor for width (1.0 = original)
    scale_height: float  # Scale factor for height (1.0 = original)
    alt_text: str        # Alternative text for accessibility

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
    conditional_formats: dict[str, ConditionalFormat] | None  # Column name/pattern -> conditional format config
    formula_columns: dict[str, str] | None  # Column name -> Excel formula template with {row} placeholder
    merged_ranges: list[tuple[str, str] | tuple[str, str, HeaderFormat]] | None  # (range, text) or (range, text, format)
    hyperlinks: list[tuple[str, str] | tuple[str, str, str]] | None  # (cell, url) or (cell, url, display_text)
    comments: dict[str, str | CommentOptions] | None  # Cell ref -> comment text or options
    validations: dict[str, ValidationOptions] | None  # Column name/pattern -> validation options
    rich_text: dict[str, list[tuple[str, RichTextFormat] | str]] | None  # Cell ref -> list of (text, format) or plain text
    images: dict[str, str | ImageOptions] | None  # Cell ref -> image path or options

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
    table_name: str | None = None,
    header_format: HeaderFormat | None = None,
    row_heights: dict[int, float] | None = None,
    constant_memory: bool = False,
    column_formats: dict[str, ColumnFormat] | None = None,
    conditional_formats: dict[str, ConditionalFormat] | None = None,
    formula_columns: dict[str, str] | None = None,
    merged_ranges: list[tuple[str, str] | tuple[str, str, HeaderFormat]] | None = None,
    hyperlinks: list[tuple[str, str] | tuple[str, str, str]] | None = None,
    comments: dict[str, str | CommentOptions] | None = None,
    validations: dict[str, ValidationOptions] | None = None,
    rich_text: dict[str, list[tuple[str, RichTextFormat] | str]] | None = None,
    images: dict[str, str | ImageOptions] | None = None,
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
        conditional_formats: Dict mapping column names to conditional format configs.
            Supported types: '2_color_scale', '3_color_scale', 'data_bar', 'icon_set'.
            Example: {'score': {'type': '2_color_scale', 'min_color': '#FF0000', 'max_color': '#00FF00'}}
        formula_columns: Dict mapping new column names to Excel formula templates.
            Use {row} placeholder for the current row number (1-based Excel row).
            Example: {'Total': '=A{row}+B{row}', 'Percentage': '=C{row}/D{row}*100'}
        merged_ranges: List of (range, text) or (range, text, format) tuples to merge cells.
            Range uses Excel notation (e.g., 'A1:D1'). Format uses HeaderFormat options.
            Example: [('A1:B1', 'Title'), ('C1:D1', 'Subtitle', {'bold': True})]
        hyperlinks: List of (cell, url) or (cell, url, display_text) tuples to add clickable links.
            Cell uses Excel notation (e.g., 'A1'). Display text is optional.
            Example: [('A2', 'https://example.com'), ('B2', 'https://google.com', 'Google')]
        comments: Dict mapping cell refs to comment text or CommentOptions.
            Example: {'A1': 'Simple note'} or {'A1': {'text': 'Note', 'author': 'John'}}
        validations: Dict mapping column name/pattern to data validation config.
            Types: 'list' (dropdown), 'whole_number', 'decimal', 'text_length'.
            Example: {'Status': {'type': 'list', 'values': ['Open', 'Closed']}}
        rich_text: Dict mapping cell refs to list of (text, format) tuples or plain strings.
            Example: {'A1': [('Bold', {'bold': True}), ' normal text']}
        images: Dict mapping cell refs to image path or ImageOptions.
            Example: {'B5': 'logo.png'} or {'B5': {'path': 'logo.png', 'scale_width': 0.5}}
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
    table_name: str | None = None,
    header_format: HeaderFormat | None = None,
    row_heights: dict[int, float] | None = None,
    constant_memory: bool = False,
    column_formats: dict[str, ColumnFormat] | None = None,
    conditional_formats: dict[str, ConditionalFormat] | None = None,
    formula_columns: dict[str, str] | None = None,
    merged_ranges: list[tuple[str, str] | tuple[str, str, HeaderFormat]] | None = None,
    hyperlinks: list[tuple[str, str] | tuple[str, str, str]] | None = None,
    comments: dict[str, str | CommentOptions] | None = None,
    validations: dict[str, ValidationOptions] | None = None,
    rich_text: dict[str, list[tuple[str, RichTextFormat] | str]] | None = None,
    images: dict[str, str | ImageOptions] | None = None,
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
        conditional_formats: Dict mapping column names to conditional format configs.
            Supported types: '2_color_scale', '3_color_scale', 'data_bar', 'icon_set'.
        formula_columns: Dict mapping new column names to Excel formula templates.
            Use {row} placeholder for the current row number (1-based Excel row).
        merged_ranges: List of (range, text) or (range, text, format) tuples to merge cells.
            Range uses Excel notation (e.g., 'A1:D1'). Format uses HeaderFormat options.
        hyperlinks: List of (cell, url) or (cell, url, display_text) tuples to add clickable links.
            Cell uses Excel notation (e.g., 'A1'). Display text is optional.
        comments: Dict mapping cell refs to comment text or CommentOptions.
        validations: Dict mapping column name/pattern to data validation config.
        rich_text: Dict mapping cell refs to list of (text, format) tuples or plain strings.
        images: Dict mapping cell refs to image path or ImageOptions.
    """
    ...

def version() -> str:
    """Get the version of the xlsxturbo library."""
    ...

__version__: str
