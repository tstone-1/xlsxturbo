"""Type stubs for the xlsxturbo compiled (Rust) extension module."""

from os import PathLike
from typing import Literal, TypedDict

PathArg = str | PathLike[str]

DateOrder = Literal["auto", "mdy", "us", "dmy", "eu", "european"]
ValidationType = Literal[
    "list",
    "whole_number",
    "whole",
    "integer",
    "decimal",
    "number",
    "text_length",
    "textlength",
    "length",
]

class HeaderFormat(TypedDict, total=False):
    """Header cell formatting options. All fields are optional."""

    bold: bool
    italic: bool
    font_color: str  # '#RRGGBB' or named color (white, black, red, blue, etc.)
    bg_color: str  # '#RRGGBB' or named color
    font_size: float
    underline: bool
    border: bool | str  # True = thin all sides, str = named style all sides
    border_left: bool | str  # True = thin, or named style (thin, medium, thick, dashed, dotted, double, hair, etc.)
    border_right: bool | str  # True = thin, or named style for right side only
    border_top: bool | str  # True = thin, or named style for top side only
    border_bottom: bool | str  # True = thin, or named style for bottom side only
    border_color: str  # Color for all borders. Requires a border to be set for a visible effect
    align_horizontal: str  # 'left', 'center', 'right', 'fill', 'justify', 'center_across', 'distributed'
    align_vertical: str  # 'top', 'center', 'bottom', 'justify', 'distributed'
    wrap_text: bool  # Enable text wrapping within cell

class ColumnFormat(TypedDict, total=False):
    """Column cell formatting options. All fields are optional."""

    bold: bool
    italic: bool
    font_color: str  # '#RRGGBB' or named color (white, black, red, blue, etc.)
    bg_color: str  # '#RRGGBB' or named color
    font_size: float
    underline: bool
    num_format: str  # Excel number format string, e.g. '0.00', '#,##0', '0.00%'
    border: bool | str  # True = thin all sides (backward compat), str = named style all sides
    border_left: bool | str  # True = thin, or named style (thin, medium, thick, dashed, dotted, double, hair, etc.)
    border_right: bool | str  # True = thin, or named style for right side only
    border_top: bool | str  # True = thin, or named style for top side only
    border_bottom: bool | str  # True = thin, or named style for bottom side only
    border_color: str  # Color for all borders. Requires a border to be set for a visible effect
    align_horizontal: str  # 'left', 'center', 'right', 'fill', 'justify', 'center_across', 'distributed'
    align_vertical: str  # 'top', 'center', 'bottom', 'justify', 'distributed'
    wrap_text: bool  # Enable text wrapping within cell

class ConditionalFormat(TypedDict, total=False):
    """Conditional formatting options for a column. 'type' is required.

    Supported types:
    - '2_color_scale': Gradient from min_color to max_color
    - '3_color_scale': Gradient with min_color, mid_color, max_color
    - 'data_bar': In-cell bar chart
    - 'icon_set': Traffic lights, arrows, or other icons
    - 'cell': Rule-based formatting (highlight cells matching a condition)

    For 'cell' type, use 'criteria' to specify the condition and 'format' for styling.
    Multiple rules on one column: pass a list of ConditionalFormat dicts instead of a single dict.
    """

    type: str  # Required: '2_color_scale', '3_color_scale', 'data_bar', 'icon_set', 'cell'
    # For color scales:
    min_color: str  # '#RRGGBB' or named color for minimum value
    mid_color: str  # '#RRGGBB' or named color for midpoint (3_color_scale only)
    max_color: str  # '#RRGGBB' or named color for maximum value
    # For data bars:
    bar_color: str  # '#RRGGBB' or named color for the bar fill
    border_color: str  # '#RRGGBB' or named color for bar border
    solid: bool  # True for solid fill, False for gradient (default)
    direction: str  # 'left_to_right', 'right_to_left', or 'context' (default)
    # For icon sets:
    icon_type: str  # '3_arrows', '3_traffic_lights', '3_flags', '4_arrows', '5_arrows', etc. (see README)
    reverse: bool  # Reverse icon order
    icons_only: bool  # Show only icons, hide values
    # For cell rules (type='cell'):
    criteria: str  # 'equal_to', 'not_equal_to', 'greater_than', 'less_than', 'between', 'containing', etc.
    value: str | int | float  # Target value for comparison criteria
    min_value: int | float  # Min value for 'between'/'not_between' criteria
    max_value: int | float  # Max value for 'between'/'not_between' criteria
    format: ColumnFormat  # Format to apply when condition is met (bg_color, font_color, bold, etc.)

class CommentOptions(TypedDict, total=False):
    """Options for cell comments/notes.

    Note: 'text' is required at runtime but TypedDict doesn't enforce this.
    """

    text: str  # The comment text (required at runtime)
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
    min: int | float  # For number/text_length: minimum value (defaults to type minimum if omitted)
    max: int | float  # For number/text_length: maximum value (defaults to type maximum if omitted)
    input_title: str  # Title for input prompt
    input_message: str  # Message for input prompt
    error_title: str  # Title for error message
    error_message: str  # Message for error message

class RichTextFormat(TypedDict, total=False):
    """Format options for a rich text segment."""

    bold: bool
    italic: bool
    font_color: str  # '#RRGGBB' or named color
    bg_color: str  # '#RRGGBB' or named color
    font_size: float
    underline: bool

class ImageOptions(TypedDict, total=False):
    """Options for embedding images.

    Note: 'path' is required at runtime but TypedDict doesn't enforce this.
    """

    path: str  # Path to image file - PNG, JPEG, GIF, BMP (required at runtime)
    scale_width: float  # Scale factor for width (1.0 = original)
    scale_height: float  # Scale factor for height (1.0 = original)
    alt_text: str  # Alternative text for accessibility

class CheckboxOptions(TypedDict, total=False):
    """Options for interactive cell checkboxes.

    Note: 'checked' is required at runtime but TypedDict doesn't enforce this.
    """

    checked: bool  # Initial state: True (checked) or False (unchecked) - required at runtime
    format: ColumnFormat  # Optional cell format (bg_color, font_color, border, etc.)

class TextboxFont(TypedDict, total=False):
    """Font options for textbox text."""

    name: str  # Font family name (e.g. 'Arial', 'Calibri')
    size: float  # Font size in points
    bold: bool
    italic: bool
    underline: bool
    color: str  # '#RRGGBB' or named color

class TextboxOptions(TypedDict, total=False):
    """Options for floating text shapes (textboxes).

    Note: 'text' is required at runtime but TypedDict doesn't enforce this.
    """

    text: str  # Textbox contents (required at runtime)
    width: int  # Width in pixels (default 192)
    height: int  # Height in pixels (default 120)
    x_offset: int  # Horizontal offset within the anchor cell (pixels)
    y_offset: int  # Vertical offset within the anchor cell (pixels)
    font: TextboxFont  # Font properties
    fill_color: str  # Background fill color ('#RRGGBB' or named)
    line_color: str  # Border line color ('#RRGGBB' or named)
    alt_text: str  # Alternative text for accessibility

ChartType = Literal[
    "area", "area_stacked", "area_percent_stacked",
    "stacked_area", "percent_stacked_area",  # aliases
    "bar", "bar_stacked", "bar_percent_stacked",
    "stacked_bar", "percent_stacked_bar",  # aliases
    "column", "col", "column_stacked", "column_percent_stacked",
    "stacked_column", "percent_stacked_column",  # aliases
    "doughnut", "donut",
    "line", "line_stacked", "line_percent_stacked",
    "stacked_line", "percent_stacked_line",  # aliases
    "pie", "radar", "radar_with_markers", "radar_filled",
    "scatter", "scatter_straight", "scatter_straight_with_markers",
    "scatter_smooth", "scatter_smooth_with_markers", "stock",
]

class ChartSeriesOptions(TypedDict, total=False):
    """Options for one chart data series."""

    data_range: str  # Alias for values_range
    values_range: str  # Excel range for series values, e.g. 'Sheet1!$B$2:$B$10'
    values: str  # Alias for values_range
    categories_range: str  # Excel range for categories/X values
    categories: str  # Alias for categories_range
    name: str  # Series name or formula reference
    series_name: str  # Alias for name

class ChartOptions(TypedDict, total=False):
    """Options for native Excel charts.

    Note: 'type' and either 'data_range'/'values_range' or 'series' are required at runtime.
    """

    type: ChartType
    data_range: str  # Alias for values_range
    values_range: str  # Excel range for a single series values
    values: str  # Alias for values_range
    categories_range: str  # Excel range for categories/X values
    categories: str  # Alias for categories_range
    series: list[ChartSeriesOptions]  # Multiple series
    name: str  # Single-series name or formula reference
    series_name: str  # Alias for name
    title: str  # Chart title
    x_axis_name: str  # X/category axis title
    y_axis_name: str  # Y/value axis title
    width: int  # Width in pixels
    height: int  # Height in pixels
    x_offset: int  # Horizontal offset within the anchor cell (pixels)
    y_offset: int  # Vertical offset within the anchor cell (pixels)
    style: int  # Excel chart style id, 1-48
    show_data_table: bool  # Show data table under the chart
    show_legend: bool  # Show chart legend (default True)
    legend_position: Literal["right", "left", "top", "bottom", "top_right"]

SparklineType = Literal["line", "column", "col", "win_loss", "win_lose", "winloss", "winlose"]

class SparklineOptions(TypedDict, total=False):
    """Options for a native Excel sparkline (mini in-cell chart).

    Note: 'range' is required at runtime but TypedDict doesn't enforce this.
    """

    range: str  # Sheet-qualified data range, e.g. 'Sheet1!A2:C2' (1D) or 'Sheet1!A2:C10' (2D, group)
    type: SparklineType  # Sparkline style (default 'line')
    style: int  # Built-in sparkline style id, 1-36
    markers: bool  # Show a marker on every data point
    high_point: bool  # Highlight the highest point
    low_point: bool  # Highlight the lowest point
    first_point: bool  # Highlight the first point
    last_point: bool  # Highlight the last point
    negative_points: bool  # Highlight negative points
    show_axis: bool  # Show a horizontal axis line
    show_hidden_data: bool  # Plot data in hidden rows/columns
    group_max: bool  # Use a common max across a grouped sparkline
    group_min: bool  # Use a common min across a grouped sparkline
    right_to_left: bool  # Plot the data right-to-left
    column_order: bool  # Plot data column-by-column instead of row-by-row
    color: str  # Sparkline series color ('#RRGGBB' or named)
    high_point_color: str  # High-point marker color
    low_point_color: str  # Low-point marker color
    first_point_color: str  # First-point marker color
    last_point_color: str  # Last-point marker color
    negative_points_color: str  # Negative-points marker color
    markers_color: str  # Marker color
    line_weight: float  # Line weight in points (line sparklines)
    custom_max: float  # Custom vertical-axis maximum
    custom_min: float  # Custom vertical-axis minimum
    date_range: str  # Sheet-qualified range supplying X-axis date values, e.g. 'Sheet1!A1:C1'

class CellValueOptions(TypedDict, total=False):
    """Options for a cell write with custom formatting.

    Note: 'value' is required at runtime but TypedDict doesn't enforce this.
    """

    value: str | int | float | bool  # The cell value (required at runtime)
    num_format: str  # Excel number format string, e.g. '@' for text, '0.00' for decimal
    align_horizontal: str  # 'left', 'center', 'right', 'fill', 'justify', 'center_across', 'distributed'
    align_vertical: str  # 'top', 'center', 'bottom', 'justify', 'distributed'
    wrap_text: bool  # Enable text wrapping within cell

class SheetOptions(TypedDict, total=False):
    """Per-sheet options for dfs_to_xlsx. All fields are optional."""

    header: bool
    autofit: bool
    table_style: str | None
    freeze_panes: bool
    column_widths: dict[int | str, int | float] | None  # Keys: int index or '_all'
    row_heights: dict[int, int | float] | None
    table_name: str | None
    header_format: HeaderFormat | None
    column_formats: dict[str, ColumnFormat] | None  # Pattern -> format ('prefix*', '*suffix', '*contains*', exact)
    conditional_formats: dict[str, ConditionalFormat | list[ConditionalFormat]] | None  # Column/pattern -> config
    formula_columns: dict[str, str] | None  # Column name -> Excel formula template with {row} placeholder
    merged_ranges: list[tuple[str, str] | tuple[str, str, HeaderFormat]] | None  # (range, text[, format])
    hyperlinks: list[tuple[str, str] | tuple[str, str, str]] | None  # (cell, url[, display_text])
    comments: dict[str, str | CommentOptions] | None  # Cell ref -> comment text or options
    validations: dict[str, ValidationOptions] | None  # Column name/pattern -> validation options
    rich_text: dict[str, list[tuple[str, RichTextFormat] | str]] | None  # Cell ref -> segments
    images: dict[str, str | ImageOptions] | None  # Cell ref -> image path or options
    checkboxes: dict[str, bool | CheckboxOptions] | None  # Cell ref -> checked state or options
    textboxes: dict[str, str | TextboxOptions] | None  # Cell ref -> text or textbox options
    charts: dict[str, ChartOptions] | None  # Cell ref -> native Excel chart options
    sparklines: dict[str, SparklineOptions] | None  # Location ref -> sparkline options
    cells: dict[str, str | int | float | bool | CellValueOptions] | None  # Cell ref -> value or options

def csv_to_xlsx(
    input_path: PathArg,
    output_path: PathArg,
    sheet_name: str = "Sheet1",
    parallel: bool = False,
    date_order: DateOrder = "auto",
) -> tuple[int, int]:
    """Convert a CSV file to XLSX format with automatic type detection.

    Args:
        input_path: Path to the input CSV file.
        output_path: Path for the output XLSX file.
        sheet_name: Name of the worksheet (default: "Sheet1").
        parallel: Use multi-core parallel processing (default: False).
            Faster for large files (100K+ rows) but uses more memory.
        date_order: Date parsing order for ambiguous dates like "01-02-2024".
            "auto" - ISO first, then European (DMY), then US (MDY).
            "mdy" or "us" - US format: 01-02-2024 = January 2nd.
            "dmy" or "eu" - European format: 01-02-2024 = February 1st.

    Returns:
        Tuple of (rows, columns) written to the Excel file.

    Raises:
        ValueError: If the conversion fails.
    """

def df_to_xlsx(
    df: object,
    output_path: PathArg,
    sheet_name: str = "Sheet1",
    header: bool = True,
    autofit: bool = False,
    table_style: str | None = None,
    freeze_panes: bool = False,
    column_widths: dict[int | str, int | float] | None = None,
    table_name: str | None = None,
    header_format: HeaderFormat | None = None,
    row_heights: dict[int, int | float] | None = None,
    constant_memory: bool = False,
    column_formats: dict[str, ColumnFormat] | None = None,
    conditional_formats: dict[str, ConditionalFormat | list[ConditionalFormat]] | None = None,
    formula_columns: dict[str, str] | None = None,
    merged_ranges: list[tuple[str, str] | tuple[str, str, HeaderFormat]] | None = None,
    hyperlinks: list[tuple[str, str] | tuple[str, str, str]] | None = None,
    comments: dict[str, str | CommentOptions] | None = None,
    validations: dict[str, ValidationOptions] | None = None,
    rich_text: dict[str, list[tuple[str, RichTextFormat] | str]] | None = None,
    images: dict[str, str | ImageOptions] | None = None,
    checkboxes: dict[str, bool | CheckboxOptions] | None = None,
    textboxes: dict[str, str | TextboxOptions] | None = None,
    charts: dict[str, ChartOptions] | None = None,
    defined_names: dict[str, str] | None = None,
    cells: dict[str, str | int | float | bool | CellValueOptions] | None = None,
    sparklines: dict[str, SparklineOptions] | None = None,
) -> tuple[int, int]:
    """Convert a pandas or polars DataFrame to XLSX format.

    Args:
        df: pandas DataFrame or polars DataFrame to export.
        output_path: Path for the output XLSX file.
        sheet_name: Name of the worksheet (default: "Sheet1").
        header: Include column names as header row (default: True).
        autofit: Automatically adjust column widths to fit content (default: False).
        table_style: Apply Excel table formatting (default: None).
            Styles: "Light1"-"Light21", "Medium1"-"Medium28", "Dark1"-"Dark11", "None".
        freeze_panes: Freeze the header row for easier scrolling (default: False).
        column_widths: Dict mapping column index to width. Use '_all' to cap all columns.
        table_name: Custom name for the Excel table (requires table_style).
        header_format: Dict of header cell formatting options.
        row_heights: Dict mapping row index to height in points.
        constant_memory: Use streaming mode for minimal RAM usage (default: False).
            When enabled, emits RuntimeWarning and disables: table_style, freeze_panes,
            row_heights, autofit, column_widths with autofit cap, conditional_formats,
            formula_columns, merged_ranges, hyperlinks, comments, validations, rich_text,
            images, checkboxes, textboxes, charts, sparklines, and cells. Plain column_widths,
            header_format, and column_formats remain supported.
        column_formats: Dict mapping column name patterns to format options.
            Patterns: 'prefix*', '*suffix', '*contains*', or exact match.
            First matching pattern wins (order preserved).
        conditional_formats: Dict mapping column names to conditional format configs.
            Supported types: '2_color_scale', '3_color_scale', 'data_bar', 'icon_set', 'cell'.
            Example: {'score': {'type': '2_color_scale', 'min_color': '#FF0000', 'max_color': '#00FF00'}}
        formula_columns: Dict mapping new column names to Excel formula templates.
            Use {row} placeholder for the current row number (1-based Excel row).
            Example: {'Total': '=A{row}+B{row}', 'Percentage': '=C{row}/D{row}*100'}
        merged_ranges: List of (range, text) or (range, text, format) tuples to merge cells.
            Range uses Excel notation (e.g., 'A1:D1'). Format uses HeaderFormat options.
            Example: [('A1:B1', 'Title'), ('C1:D1', 'Subtitle', {'bold': True})]
        hyperlinks: List of (cell, url) or (cell, url, display_text) tuples for clickable links.
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
        checkboxes: Dict mapping cell refs to interactive checkboxes.
            Simple form: {'A1': True, 'A2': False}
            Dict form: {'A3': {'checked': True, 'format': {'bg_color': '#C6EFCE'}}}
        textboxes: Dict mapping cell refs to floating text shapes.
            Simple form: {'B2': 'Some text'}
            Dict form: {'B2': {'text': 'Note', 'width': 200, 'height': 100,
                        'x_offset': 10, 'y_offset': 5,
                        'font': {'name': 'Arial', 'size': 14, 'bold': True, 'color': '#FF0000'},
                        'fill_color': '#F0F0F0', 'line_color': '#000000',
                        'alt_text': 'Descriptive alt text'}}
        charts: Dict mapping cell refs to native Excel chart configs.
            Example: {'D2': {'type': 'bar', 'data_range': 'Sheet1!$B$2:$B$10',
                      'categories_range': 'Sheet1!$A$2:$A$10', 'title': 'Monthly Activity'}}
        sparklines: Dict mapping a location ref to a sparkline (mini in-cell chart) config.
            A single-cell key (e.g. 'D2') places one sparkline; a range key (e.g. 'D2:D10')
            places a grouped sparkline, one per row of the data range. 'range' is required and
            must be sheet-qualified (e.g. 'Sheet1!A2:C10'), like a chart range.
            Example: {'D2:D10': {'range': 'Sheet1!A2:C10', 'type': 'line', 'markers': True}}
        defined_names: Dict mapping name to Excel reference for workbook-level defined names.
            Example: {'MyRange': '=Sheet1!$A$1:$D$100'}
        cells: Dict mapping cell refs to values for arbitrary cell writes.
            Values can be simple (str, int, float, bool) or dicts with 'value' and optional 'num_format'.
            Cells are written after DataFrame data, so they can overwrite existing values.
            Example: {'B9': 'Label', 'D6': {'value': '934728173849', 'num_format': '@'}}

    Returns:
        Tuple of (rows, columns) written to the Excel file.
    """

def dfs_to_xlsx(
    sheets: list[tuple[object, str] | tuple[object, str, SheetOptions]],
    output_path: PathArg,
    header: bool = True,
    autofit: bool = False,
    table_style: str | None = None,
    freeze_panes: bool = False,
    column_widths: dict[int | str, int | float] | None = None,
    table_name: str | None = None,
    header_format: HeaderFormat | None = None,
    row_heights: dict[int, int | float] | None = None,
    constant_memory: bool = False,
    column_formats: dict[str, ColumnFormat] | None = None,
    conditional_formats: dict[str, ConditionalFormat | list[ConditionalFormat]] | None = None,
    formula_columns: dict[str, str] | None = None,
    merged_ranges: list[tuple[str, str] | tuple[str, str, HeaderFormat]] | None = None,
    hyperlinks: list[tuple[str, str] | tuple[str, str, str]] | None = None,
    comments: dict[str, str | CommentOptions] | None = None,
    validations: dict[str, ValidationOptions] | None = None,
    rich_text: dict[str, list[tuple[str, RichTextFormat] | str]] | None = None,
    images: dict[str, str | ImageOptions] | None = None,
    checkboxes: dict[str, bool | CheckboxOptions] | None = None,
    textboxes: dict[str, str | TextboxOptions] | None = None,
    charts: dict[str, ChartOptions] | None = None,
    defined_names: dict[str, str] | None = None,
    cells: dict[str, str | int | float | bool | CellValueOptions] | None = None,
    sparklines: dict[str, SparklineOptions] | None = None,
) -> list[tuple[int, int]]:
    """Write multiple DataFrames to separate sheets in a single workbook.

    Args:
        sheets: List of (DataFrame, sheet_name) or (DataFrame, sheet_name, options) tuples.
        output_path: Path for the output XLSX file.
        header: Include column names as header row (default: True).
        autofit: Automatically adjust column widths (default: False).
        table_style: Apply Excel table formatting (default: None).
        freeze_panes: Freeze the header row (default: False).
        column_widths: Dict mapping column index to width. Use '_all' to cap all columns.
        table_name: Custom name for Excel tables (requires table_style).
        header_format: Dict of header cell formatting options.
        row_heights: Dict mapping row index to height in points.
        constant_memory: Use streaming mode (default: False).
            When enabled, emits RuntimeWarning and disables: table_style, freeze_panes,
            row_heights, autofit, column_widths with autofit cap, conditional_formats,
            formula_columns, merged_ranges, hyperlinks, comments, validations, rich_text,
            images, checkboxes, textboxes, charts, sparklines, and cells. Plain column_widths,
            header_format, and column_formats remain supported.
        column_formats: Dict mapping column name patterns to format options.
            Patterns: 'prefix*', '*suffix', '*contains*', or exact match.
        conditional_formats: Dict mapping column names to conditional format configs.
            Supported types: '2_color_scale', '3_color_scale', 'data_bar', 'icon_set', 'cell'.
        formula_columns: Dict mapping new column names to Excel formula templates.
            Use {row} placeholder for the current row number (1-based Excel row).
        merged_ranges: List of (range, text) or (range, text, format) tuples to merge cells.
            Range uses Excel notation (e.g., 'A1:D1'). Format uses HeaderFormat options.
        hyperlinks: List of (cell, url) or (cell, url, display_text) tuples for clickable links.
            Cell uses Excel notation (e.g., 'A1'). Display text is optional.
        comments: Dict mapping cell refs to comment text or CommentOptions.
        validations: Dict mapping column name/pattern to data validation config.
        rich_text: Dict mapping cell refs to list of (text, format) tuples or plain strings.
        images: Dict mapping cell refs to image path or ImageOptions.
        checkboxes: Dict mapping cell refs to interactive checkboxes.
            Simple form: {'A1': True}
            Dict form: {'A1': {'checked': True, 'format': {'bg_color': '#C6EFCE'}}}
        textboxes: Dict mapping cell refs to floating text shapes.
            Simple form: {'B2': 'text'}
            Dict form: {'B2': {'text': 'Note', 'width': 200, 'font': {'bold': True}}}
        charts: Dict mapping cell refs to native Excel chart configs.
        sparklines: Dict mapping a location ref to a sparkline (mini in-cell chart) config.
            Range key (e.g. 'D2:D10') makes a grouped sparkline; single cell makes one.
            'range' must be sheet-qualified, e.g. 'Sheet1!A2:C10'.
            Example: {'D2:D10': {'range': 'Sheet1!A2:C10', 'type': 'line', 'markers': True}}
        defined_names: Dict mapping name to Excel reference for workbook-level defined names.
            Example: {'MyRange': '=Sheet1!$A$1:$D$100'}
        cells: Dict mapping cell refs to values for arbitrary cell writes.
            Values can be simple (str, int, float, bool) or dicts with 'value' and optional 'num_format'.
            Example: {'B9': 'Label', 'D6': {'value': '934728173849', 'num_format': '@'}}

    Returns:
        List of (rows, columns) tuples, one per written sheet.
    """

def version() -> str:
    """Return the version of the xlsxturbo library."""

__version__: str
