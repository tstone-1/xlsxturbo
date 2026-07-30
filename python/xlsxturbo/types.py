"""Option shapes for annotating xlsxturbo calls.

Every option that takes a dict has a ``TypedDict`` here, and every option that
takes one of a fixed set of strings has a ``Literal`` alias. Import them
directly -- these are real runtime objects, so no ``TYPE_CHECKING`` guard is
needed::

    from xlsxturbo.types import HeaderFormat

    header: HeaderFormat = {"bold": True, "bg_color": "#DDDDDD"}
    xlsxturbo.df_to_xlsx(df, "out.xlsx", header_format=header)

This module is the authoritative home for these shapes;
``xlsxturbo/xlsxturbo.pyi`` imports them rather than declaring its own copies.
It has no dependencies beyond the standard library, and nothing here imports the
compiled extension, so it is safe to import from anywhere.

.. note::
   Field annotations are strings, because this module uses
   ``from __future__ import annotations`` -- that is what lets fields be written
   as ``bool | str`` while still importing on Python 3.9, where evaluating that
   expression raises ``TypeError``. The consequence is that
   ``typing.get_type_hints()`` on these classes fails on 3.9 and works from 3.10
   onwards. Static type checking is unaffected on every version.
"""

from __future__ import annotations

from os import PathLike
from typing import Literal, TypedDict, Union

# The public surface, authoritative rather than descriptive: `from
# xlsxturbo.types import *` gives exactly these names. Without it the four
# imports above were re-exported too, and `tests/test_types_module.py` had to
# hide them behind a hardcoded exclusion list -- which meant the test validated
# a cleaner namespace than users actually received, and every future typing
# helper would have needed another exclusion. That test now compares against
# this list.
__all__ = [
    "CellValueOptions",
    "ChartOptions",
    "ChartSeriesOptions",
    "ChartType",
    "CheckboxOptions",
    "ColumnFormat",
    "CommentOptions",
    "ConditionalFormat",
    "DateOrder",
    "HeaderFormat",
    "ImageOptions",
    "PathArg",
    "RichTextFormat",
    "SheetOptions",
    "SparklineOptions",
    "SparklineType",
    "TextboxFont",
    "TextboxOptions",
    "ValidationOptions",
    "ValidationType",
]

PathArg = Union[str, PathLike[str]]

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


class _ConditionalRequired(TypedDict):
    """Required field of :class:`ConditionalFormat`; see there for the full shape.

    Split out so the requirement is expressed in the type rather than only
    in prose. A single ``total=False`` TypedDict would let a checker accept
    an empty dict, which the runtime rejects.
    """

    type: str  # Required: '2_color_scale', '3_color_scale', 'data_bar', 'icon_set', 'cell'


class ConditionalFormat(_ConditionalRequired, total=False):
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


class _CommentRequired(TypedDict):
    """Required field of :class:`CommentOptions`; see there for the full shape.

    Split out so the requirement is expressed in the type rather than only
    in prose. A single ``total=False`` TypedDict would let a checker accept
    an empty dict, which the runtime rejects.
    """

    text: str  # The comment text (required)


class CommentOptions(_CommentRequired, total=False):
    """Options for cell comments/notes."""

    author: str  # Author name for the comment


class _ValidationRequired(TypedDict):
    """Required field of :class:`ValidationOptions`; see there for the full shape.

    Split out so the requirement is expressed in the type rather than only
    in prose. A single ``total=False`` TypedDict would let a checker accept
    an empty dict, which the runtime rejects.
    """

    type: ValidationType  # Required: validation type


class ValidationOptions(_ValidationRequired, total=False):
    """Data validation options for a column. 'type' is required.

    Supported types:
    - 'list': Dropdown with specified values
    - 'whole_number': Integer between min and max
    - 'decimal': Decimal number between min and max
    - 'text_length': Text length between min and max

    For 'whole_number', 'min'/'max' are bounded to the i32 range
    (-2147483648..=2147483647); a value outside that range raises ValueError
    naming the field and range, instead of a misleading generic type error.
    """

    values: list[str]  # For 'list' type: dropdown options
    min: int | float  # For number/text_length: minimum value (defaults to type minimum if omitted)
    max: int | float  # For number/text_length: maximum value (defaults to type maximum if omitted)
    input_title: str  # Title for input prompt
    input_message: str  # Message for input prompt
    error_title: str  # Title for error message
    error_message: str  # Message for error message


class RichTextFormat(TypedDict, total=False):
    """Format options for a rich text segment.

    Font-level keys only. A segment is an inline run inside one cell, so
    cell-level keys (borders, alignment, wrap_text) would never render and are
    rejected at runtime; format the cell itself via column_formats or cells.
    """

    bold: bool
    italic: bool
    font_color: str  # '#RRGGBB' or named color
    bg_color: str  # '#RRGGBB' or named color
    font_size: float
    underline: bool


class _ImageRequired(TypedDict):
    """Required field of :class:`ImageOptions`; see there for the full shape.

    Split out so the requirement is expressed in the type rather than only
    in prose. A single ``total=False`` TypedDict would let a checker accept
    an empty dict, which the runtime rejects.
    """

    path: str  # Path to image file - PNG, JPEG, GIF, BMP (required)


class ImageOptions(_ImageRequired, total=False):
    """Options for embedding images."""

    scale_width: float  # Scale factor for width (1.0 = original)
    scale_height: float  # Scale factor for height (1.0 = original)
    alt_text: str  # Alternative text for accessibility


class _CheckboxRequired(TypedDict):
    """Required field of :class:`CheckboxOptions`; see there for the full shape.

    Split out so the requirement is expressed in the type rather than only
    in prose. A single ``total=False`` TypedDict would let a checker accept
    an empty dict, which the runtime rejects.
    """

    checked: bool  # Initial state: True (checked) or False (unchecked) - required at runtime


class CheckboxOptions(_CheckboxRequired, total=False):
    """Options for interactive cell checkboxes."""

    format: ColumnFormat  # Optional cell format (bg_color, font_color, border, etc.)


class TextboxFont(TypedDict, total=False):
    """Font options for textbox text."""

    name: str  # Font family name (e.g. 'Arial', 'Calibri')
    size: float  # Font size in points
    bold: bool
    italic: bool
    underline: bool
    color: str  # '#RRGGBB' or named color


class _TextboxRequired(TypedDict):
    """Required field of :class:`TextboxOptions`; see there for the full shape.

    Split out so the requirement is expressed in the type rather than only
    in prose. A single ``total=False`` TypedDict would let a checker accept
    an empty dict, which the runtime rejects.
    """

    text: str  # Textbox contents (required)


class TextboxOptions(_TextboxRequired, total=False):
    """Options for floating text shapes (textboxes)."""

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
    """Options for one chart data series.

    Note: 'values_range'/'values'/'data_range' and 'categories_range'/'categories'
    must include a sheet name (e.g. 'Sheet1!$B$2:$B$10'); a bare range like
    '$B$2:$B$10' raises ValueError.
    """

    data_range: str  # Alias for values_range
    values_range: str  # Excel range for series values, e.g. 'Sheet1!$B$2:$B$10'
    values: str  # Alias for values_range
    categories_range: str  # Excel range for categories/X values
    categories: str  # Alias for categories_range
    name: str  # Series name or formula reference
    series_name: str  # Alias for name


class _ChartRequired(TypedDict):
    """Required field of :class:`ChartOptions`; see there for the full shape.

    Split out so the requirement is expressed in the type rather than only
    in prose. A single ``total=False`` TypedDict would let a checker accept
    an empty dict, which the runtime rejects.
    """

    type: ChartType


class ChartOptions(_ChartRequired, total=False):
    """Options for native Excel charts.

    Note: 'type' and either 'data_range'/'values_range' or 'series' are required at runtime.
    Note: 'values_range'/'values'/'data_range' and 'categories_range'/'categories'
    must include a sheet name (e.g. 'Sheet1!$B$2:$B$10'); a bare range raises
    ValueError. This also applies to a chart-level 'categories_range'/'categories'
    used as the fallback for series that don't specify their own.
    """

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
    style: int  # Excel chart style id, 1-48 (validated; out-of-range raises)
    show_data_table: bool  # Show data table under the chart
    show_legend: bool  # Show chart legend (default True)
    legend_position: Literal["right", "left", "top", "bottom", "top_right"]

SparklineType = Literal["line", "column", "col", "win_loss", "win_lose", "winloss", "winlose"]


class _SparklineRequired(TypedDict):
    """Required field of :class:`SparklineOptions`; see there for the full shape.

    Split out so the requirement is expressed in the type rather than only
    in prose. A single ``total=False`` TypedDict would let a checker accept
    an empty dict, which the runtime rejects.
    """

    range: str  # Sheet-qualified data range, e.g. 'Sheet1!A2:C2' (1D) or 'Sheet1!A2:C10' (2D, group)


class SparklineOptions(_SparklineRequired, total=False):
    """Options for a native Excel sparkline (mini in-cell chart)."""

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


class _CellValueRequired(TypedDict):
    """Required field of :class:`CellValueOptions`; see there for the full shape.

    Split out so the requirement is expressed in the type rather than only
    in prose. A single ``total=False`` TypedDict would let a checker accept
    an empty dict, which the runtime rejects.
    """

    value: str | int | float | bool  # The cell value (required)


class CellValueOptions(_CellValueRequired, total=False):
    """Options for a cell write with custom formatting."""

    num_format: str  # Excel number format string, e.g. '@' for text, '0.00' for decimal
    align_horizontal: str  # 'left', 'center', 'right', 'fill', 'justify', 'center_across', 'distributed'
    align_vertical: str  # 'top', 'center', 'bottom', 'justify', 'distributed'
    wrap_text: bool  # Enable text wrapping within cell


class SheetOptions(TypedDict, total=False):
    """Per-sheet options for dfs_to_xlsx. All fields are optional.

    For the dict/list-valued options below (column_widths, header_format,
    column_formats, conditional_formats, formula_columns, merged_ranges,
    hyperlinks, comments, validations, rich_text, images, checkboxes,
    textboxes, charts, sparklines, cells): passing an explicitly empty dict/list for a sheet
    disables that global default for that sheet rather than falling back to
    it. Omitting the key entirely (or passing None) still falls back to the
    global default as before.

    A per-sheet `column_widths` dict combined with `autofit=True` (global or
    per-sheet) no longer suppresses autofit: explicit widths win for the
    columns they name, and every other column is still autofitted. An
    `'_all'` entry caps the autofit width for those other columns instead of
    overriding it. An explicitly empty `column_widths: {}` for a sheet has no
    explicit widths to apply, so the sheet is autofitted exactly as if
    `column_widths` had been omitted.
    """

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
