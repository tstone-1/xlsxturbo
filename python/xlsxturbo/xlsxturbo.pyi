"""Type stubs for the xlsxturbo compiled (Rust) extension module.

Declares what the compiled module actually provides: four functions, the version
string, and the exception classes built in ``src/errors.rs``. The option shapes
live in ``xlsxturbo.types`` -- a real runtime module -- and are re-exported here
so that ``from xlsxturbo.xlsxturbo import HeaderFormat`` keeps type-checking for
code written before they moved.
"""

from xlsxturbo.types import (
    CellValueOptions as CellValueOptions,
    ChartOptions as ChartOptions,
    ChartSeriesOptions as ChartSeriesOptions,
    ChartType as ChartType,
    CheckboxOptions as CheckboxOptions,
    ColumnFormat as ColumnFormat,
    CommentOptions as CommentOptions,
    ConditionalFormat as ConditionalFormat,
    DateOrder as DateOrder,
    HeaderFormat as HeaderFormat,
    ImageOptions as ImageOptions,
    PathArg as PathArg,
    RichTextFormat as RichTextFormat,
    SheetOptions as SheetOptions,
    SparklineOptions as SparklineOptions,
    SparklineType as SparklineType,
    TextboxFont as TextboxFont,
    TextboxOptions as TextboxOptions,
    ValidationOptions as ValidationOptions,
    ValidationType as ValidationType,
)

__all__ = [
    "CellValueOptions",
    "ChartOptions",
    "ChartSeriesOptions",
    "ChartType",
    "CheckboxOptions",
    "ColumnFormat",
    "CommentOptions",
    "ConditionalFormat",
    "ConfigurationError",
    "ConfigurationTypeError",
    "DateOrder",
    "FileError",
    "HeaderFormat",
    "ImageOptions",
    "InputDataError",
    "PathArg",
    "RichTextFormat",
    "SheetOptions",
    "SparklineOptions",
    "SparklineType",
    "TextboxFont",
    "TextboxOptions",
    "ValidationOptions",
    "ValidationType",
    "WorkbookValidationError",
    "XlsxTurboError",
    "__version__",
    "csv_to_xlsx",
    "df_to_xlsx",
    "dfs_to_xlsx",
    "version",
]

# --- Exception hierarchy ---------------------------------------------------
#
# Built at module initialisation in `src/errors.rs`, so these are real runtime
# classes rather than stub-only types. Each carries the builtin exception the
# same failure raised before 0.19, which is what keeps `except ValueError` and
# `except TypeError` working; see `docs/errors.md`.

class XlsxTurboError(Exception):
    """Base class for every exception raised by xlsxturbo."""

class ConfigurationError(XlsxTurboError, ValueError):
    """An option or argument has an invalid value."""

class ConfigurationTypeError(XlsxTurboError, TypeError):
    """An option or argument has the wrong type."""

class InputDataError(XlsxTurboError, ValueError):
    """The object passed as data is not a supported DataFrame."""

class FileError(XlsxTurboError, OSError, ValueError):
    """A filesystem read or write failed."""

class WorkbookValidationError(ConfigurationError):
    """The configuration is well-formed, but Excel does not permit it."""

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

    Note:
        String cells preserve surrounding whitespace (e.g. " padded " is
        written back exactly as given); type detection trims a private copy
        of the value to classify it, so a whitespace-padded number like
        " 123 " is still detected and written as a number.
        Dates before 1900-03-01 cannot be represented as a correct Excel
        serial number (Excel's 1900 leap-year bug assumes a phantom
        1900-02-29) and are written as plain text instead; 1900-03-01 is the
        first date written as a real Excel date.

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
            Combined with column_widths: explicit widths win for the columns
            they name; every other column is still autofitted (rather than
            left at Excel's default width). Add an '_all' entry in
            column_widths to cap the autofit width instead of overriding it.
        table_style: Apply Excel table formatting (default: None).
            Styles: "Light1"-"Light21", "Medium1"-"Medium28", "Dark1"-"Dark11", "None".
        freeze_panes: Freeze the header row for easier scrolling (default: False).
        column_widths: Dict mapping column index to width. Use '_all' to cap all columns.
            An integer key must be a non-negative index within Excel's column range
            (0..=16383); a negative key, a key beyond 16383, or a non-integer/non-'_all'
            key raises. A key beyond the DataFrame's column count is applied to that
            column anyway (it is no longer silently ignored). With autofit=True and
            no '_all' key: listed columns get the explicit width, unlisted columns
            are autofitted. With autofit=True and an '_all' key: '_all' caps the
            autofit width for unlisted columns instead of overriding it.
        table_name: Custom name for the Excel table (requires table_style).
            Effective table names must be unique across the workbook after sanitization.
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
            Every pattern must match at least one column or ValueError is raised.
        conditional_formats: Dict mapping column names to conditional format configs.
            Supported types: '2_color_scale', '3_color_scale', 'data_bar', 'icon_set', 'cell'.
            Every name or pattern must match at least one column.
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
            Every name or pattern must match at least one column.
            Types: 'list' (dropdown), 'whole_number', 'decimal', 'text_length'.
            Example: {'Status': {'type': 'list', 'values': ['Open', 'Closed']}}
            'whole_number' min/max are bounded to the i32 range (-2147483648..=2147483647);
            a value outside that range raises ValueError naming the field and range,
            instead of a misleading generic type error.
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
        charts: Dict mapping cell refs to native Excel chart configs. 'values_range'/'values'/
            'data_range' and 'categories_range'/'categories' (including a chart-level
            'categories_range'/'categories' used as a fallback for series without their own)
            must include a sheet name (e.g. 'Sheet1!$B$2:$B$10'); a bare range raises ValueError.
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

    Note:
        date/datetime values before 1900-03-01 cannot be represented as a
        correct Excel serial number (Excel's 1900 leap-year bug assumes a
        phantom 1900-02-29) and are written as plain text instead;
        1900-03-01 is the first date written as a real Excel date.
        Subclasses of datetime.datetime/datetime.date (e.g. pandas Timestamp,
        or a user-defined subclass) in an object-dtype column are written as
        real Excel datetimes/dates, not as their str() representation.

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
            Combined with column_widths: explicit widths win for the columns
            they name; every other column is still autofitted (rather than
            left at Excel's default width). Add an '_all' entry in
            column_widths to cap the autofit width instead of overriding it.
        table_style: Apply Excel table formatting (default: None).
        freeze_panes: Freeze the header row (default: False).
        column_widths: Dict mapping column index to width. Use '_all' to cap all columns.
            An integer key must be a non-negative index within Excel's column range
            (0..=16383); a negative key, a key beyond 16383, or a non-integer/non-'_all'
            key raises. A key beyond the DataFrame's column count is applied to that
            column anyway (it is no longer silently ignored). With autofit=True and
            no '_all' key: listed columns get the explicit width, unlisted columns
            are autofitted. With autofit=True and an '_all' key: '_all' caps the
            autofit width for unlisted columns instead of overriding it.
        table_name: Custom name for Excel tables (requires table_style). Effective
            table names must be unique across the workbook after sanitization.
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
            Every pattern must match at least one column or ValueError is raised.
        conditional_formats: Dict mapping column names to conditional format configs.
            Supported types: '2_color_scale', '3_color_scale', 'data_bar', 'icon_set', 'cell'.
            Every name or pattern must match at least one column.
        formula_columns: Dict mapping new column names to Excel formula templates.
            Use {row} placeholder for the current row number (1-based Excel row).
        merged_ranges: List of (range, text) or (range, text, format) tuples to merge cells.
            Range uses Excel notation (e.g., 'A1:D1'). Format uses HeaderFormat options.
        hyperlinks: List of (cell, url) or (cell, url, display_text) tuples for clickable links.
            Cell uses Excel notation (e.g., 'A1'). Display text is optional.
        comments: Dict mapping cell refs to comment text or CommentOptions.
        validations: Dict mapping column name/pattern to data validation config.
            Every name or pattern must match at least one column.
            'whole_number' min/max are bounded to the i32 range (-2147483648..=2147483647);
            a value outside that range raises ValueError naming the field and range,
            instead of a misleading generic type error.
        rich_text: Dict mapping cell refs to list of (text, format) tuples or plain strings.
        images: Dict mapping cell refs to image path or ImageOptions.
        checkboxes: Dict mapping cell refs to interactive checkboxes.
            Simple form: {'A1': True}
            Dict form: {'A1': {'checked': True, 'format': {'bg_color': '#C6EFCE'}}}
        textboxes: Dict mapping cell refs to floating text shapes.
            Simple form: {'B2': 'text'}
            Dict form: {'B2': {'text': 'Note', 'width': 200, 'font': {'bold': True}}}
        charts: Dict mapping cell refs to native Excel chart configs. 'values_range'/'values'/
            'data_range' and 'categories_range'/'categories' (including a chart-level
            fallback used by series without their own) must include a sheet name
            (e.g. 'Sheet1!$B$2:$B$10'); a bare range raises ValueError.
        sparklines: Dict mapping a location ref to a sparkline (mini in-cell chart) config.
            Range key (e.g. 'D2:D10') makes a grouped sparkline; single cell makes one.
            'range' must be sheet-qualified, e.g. 'Sheet1!A2:C10'.
            Example: {'D2:D10': {'range': 'Sheet1!A2:C10', 'type': 'line', 'markers': True}}
        defined_names: Dict mapping name to Excel reference for workbook-level defined names.
            Example: {'MyRange': '=Sheet1!$A$1:$D$100'}
        cells: Dict mapping cell refs to values for arbitrary cell writes.
            Values can be simple (str, int, float, bool) or dicts with 'value' and optional 'num_format'.
            Example: {'B9': 'Label', 'D6': {'value': '934728173849', 'num_format': '@'}}

    Note:
        date/datetime values before 1900-03-01 cannot be represented as a
        correct Excel serial number (Excel's 1900 leap-year bug assumes a
        phantom 1900-02-29) and are written as plain text instead;
        1900-03-01 is the first date written as a real Excel date.
        Subclasses of datetime.datetime/datetime.date in an object-dtype
        column are written as real Excel datetimes/dates, not as their
        str() representation.

    Returns:
        List of (rows, columns) tuples, one per written sheet.
    """

def version() -> str:
    """Return the version of the xlsxturbo library."""

__version__: str
