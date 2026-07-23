//! xlsxturbo - High-performance Excel writer with automatic type detection
//!
//! This library provides fast DataFrame and CSV to Excel conversion:
//! - Integers and floats → Excel numbers
//! - Booleans (true/false) → Excel booleans
//! - Dates → Excel dates
//! - Datetimes → Excel datetimes
//! - NaN/Inf/None → Empty cells
//! - Everything else → Strings
//!
//! Supports pandas DataFrames, polars DataFrames, and CSV files.

mod apply;
mod convert;
mod extract;
mod parse;
mod types;
mod workbook;
mod write;

// Re-export public API for the CLI binary (main.rs)
pub use convert::{convert_csv_to_xlsx, convert_csv_to_xlsx_parallel};
pub use types::DateOrder;

use convert::{convert_dataframe_to_xlsx, dataframe_row_count, write_configured_sheet};
use extract::{
    extract_cells, extract_charts, extract_checkboxes, extract_column_formats,
    extract_column_widths, extract_comments, extract_conditional_formats, extract_formula_columns,
    extract_header_format, extract_hyperlinks, extract_images, extract_merged_ranges,
    extract_rich_text, extract_sheet_info, extract_sparklines, extract_textboxes,
    extract_validations,
};
use parse::sanitize_table_name;
use types::pytype_name;
use types::ExtractedOptions;
use types::WriteConfig;
use workbook::apply_defined_names;

use pyo3::prelude::*;
use rust_xlsxwriter::Workbook;
use std::collections::HashMap;

fn path_arg_to_string(value: &Bound<'_, PyAny>, param_name: &str) -> PyResult<String> {
    if let Ok(path) = value.extract::<String>() {
        return Ok(path);
    }
    if let Ok(pathlike) = value.call_method0("__fspath__") {
        if let Ok(path) = pathlike.extract::<String>() {
            return Ok(path);
        }
    }
    Err(pyo3::exceptions::PyTypeError::new_err(format!(
        "'{}' must be str or a path-like object returning str (bytes paths are not supported), got {}",
        param_name,
        pytype_name(value)
    )))
}

/// Helper: cast a PyAny to PyDict or raise TypeError with a clear message.
fn require_dict<'py>(
    value: &Bound<'py, PyAny>,
    param_name: &str,
) -> PyResult<Bound<'py, pyo3::types::PyDict>> {
    value.cast::<pyo3::types::PyDict>().cloned().map_err(|_| {
        pyo3::exceptions::PyTypeError::new_err(format!(
            "expected dict for '{}', got {}",
            param_name,
            pytype_name(value)
        ))
    })
}

/// Helper: cast a PyAny to PyList or raise TypeError with a clear message.
fn require_list<'py>(
    value: &Bound<'py, PyAny>,
    param_name: &str,
) -> PyResult<Bound<'py, pyo3::types::PyList>> {
    value.cast::<pyo3::types::PyList>().cloned().map_err(|_| {
        pyo3::exceptions::PyTypeError::new_err(format!(
            "expected list for '{}', got {}",
            param_name,
            pytype_name(value)
        ))
    })
}

struct RawOptions<'a, 'py> {
    column_widths: Option<&'a Bound<'py, PyAny>>,
    header_format: Option<&'a Bound<'py, PyAny>>,
    column_formats: Option<&'a Bound<'py, PyAny>>,
    conditional_formats: Option<&'a Bound<'py, PyAny>>,
    formula_columns: Option<&'a Bound<'py, PyAny>>,
    merged_ranges: Option<&'a Bound<'py, PyAny>>,
    hyperlinks: Option<&'a Bound<'py, PyAny>>,
    comments: Option<&'a Bound<'py, PyAny>>,
    validations: Option<&'a Bound<'py, PyAny>>,
    rich_text: Option<&'a Bound<'py, PyAny>>,
    images: Option<&'a Bound<'py, PyAny>>,
    checkboxes: Option<&'a Bound<'py, PyAny>>,
    textboxes: Option<&'a Bound<'py, PyAny>>,
    charts: Option<&'a Bound<'py, PyAny>>,
    sparklines: Option<&'a Bound<'py, PyAny>>,
    cells: Option<&'a Bound<'py, PyAny>>,
}

/// Extract and validate all optional write parameters from Python into typed Rust structs.
fn extract_options(raw: &RawOptions<'_, '_>) -> PyResult<ExtractedOptions> {
    Ok(ExtractedOptions {
        column_widths: raw
            .column_widths
            .map(|v| require_dict(v, "column_widths").and_then(|d| extract_column_widths(&d)))
            .transpose()?,
        header_format: raw
            .header_format
            .map(|v| require_dict(v, "header_format").and_then(|d| extract_header_format(&d)))
            .transpose()?,
        column_formats: raw
            .column_formats
            .map(|v| require_dict(v, "column_formats").and_then(|d| extract_column_formats(&d)))
            .transpose()?,
        conditional_formats: raw
            .conditional_formats
            .map(|v| {
                require_dict(v, "conditional_formats").and_then(|d| extract_conditional_formats(&d))
            })
            .transpose()?,
        formula_columns: raw
            .formula_columns
            .map(|v| require_dict(v, "formula_columns").and_then(|d| extract_formula_columns(&d)))
            .transpose()?,
        merged_ranges: raw
            .merged_ranges
            .map(|v| require_list(v, "merged_ranges").and_then(|l| extract_merged_ranges(&l)))
            .transpose()?,
        hyperlinks: raw
            .hyperlinks
            .map(|v| require_list(v, "hyperlinks").and_then(|l| extract_hyperlinks(&l)))
            .transpose()?,
        comments: raw
            .comments
            .map(|v| require_dict(v, "comments").and_then(|d| extract_comments(&d)))
            .transpose()?,
        validations: raw
            .validations
            .map(|v| require_dict(v, "validations").and_then(|d| extract_validations(&d)))
            .transpose()?,
        rich_text: raw
            .rich_text
            .map(|v| require_dict(v, "rich_text").and_then(|d| extract_rich_text(&d)))
            .transpose()?,
        images: raw
            .images
            .map(|v| require_dict(v, "images").and_then(|d| extract_images(&d)))
            .transpose()?,
        checkboxes: raw
            .checkboxes
            .map(|v| require_dict(v, "checkboxes").and_then(|d| extract_checkboxes(&d)))
            .transpose()?,
        textboxes: raw
            .textboxes
            .map(|v| require_dict(v, "textboxes").and_then(|d| extract_textboxes(&d)))
            .transpose()?,
        charts: raw
            .charts
            .map(|v| require_dict(v, "charts").and_then(|d| extract_charts(&d)))
            .transpose()?,
        sparklines: raw
            .sparklines
            .map(|v| require_dict(v, "sparklines").and_then(|d| extract_sparklines(&d)))
            .transpose()?,
        cells: raw
            .cells
            .map(|v| require_dict(v, "cells").and_then(|d| extract_cells(&d)))
            .transpose()?,
    })
}

/// Convert a CSV file to XLSX format with automatic type detection.
///
/// Reads a CSV file and converts it to an Excel XLSX file, automatically
/// detecting data types:
/// - Integers and floats become Excel numbers
/// - "true"/"false" (case-insensitive) become Excel booleans
/// - Dates (YYYY-MM-DD, DD-MM-YYYY, MM-DD-YYYY) become Excel dates
/// - Datetimes (ISO 8601) become Excel datetimes
/// - NaN/Inf values become empty cells
/// - Everything else becomes text
///
/// Args:
///     input_path: Path to the input CSV file
///     output_path: Path for the output XLSX file
///     sheet_name: Name of the worksheet (default: "Sheet1")
///     parallel: Use multi-core parallel processing (default: False).
///               Faster for large files (100K+ rows) but uses more memory.
///     date_order: Date parsing order for ambiguous dates like "01-02-2024" (default: "auto").
///                 "auto" - ISO first, then European (DMY), then US (MDY)
///                 "mdy" or "us" - US format: 01-02-2024 = January 2nd
///                 "dmy" or "eu" - European format: 01-02-2024 = February 1st
///
/// Returns:
///     Tuple of (rows, columns) written to the Excel file
///
/// Raises:
///     ValueError: If the conversion fails
///
/// Example:
///     >>> import xlsxturbo
///     >>> rows, cols = xlsxturbo.csv_to_xlsx("data.csv", "output.xlsx")
///     >>> # For US date format (MM-DD-YYYY):
///     >>> rows, cols = xlsxturbo.csv_to_xlsx("data.csv", "out.xlsx", date_order="us")
///     >>> # For large files, use parallel processing:
///     >>> rows, cols = xlsxturbo.csv_to_xlsx("big.csv", "out.xlsx", parallel=True)
#[pyfunction]
#[pyo3(signature = (input_path, output_path, sheet_name = "Sheet1", parallel = false, date_order = "auto"))]
fn csv_to_xlsx(
    py: Python<'_>,
    input_path: &Bound<'_, PyAny>,
    output_path: &Bound<'_, PyAny>,
    sheet_name: &str,
    parallel: bool,
    date_order: &str,
) -> PyResult<(u32, u16)> {
    let input_path = path_arg_to_string(input_path, "input_path")?;
    let output_path = path_arg_to_string(output_path, "output_path")?;
    let sheet_name = sheet_name.to_string();
    let order = DateOrder::parse(date_order).ok_or_else(|| {
        pyo3::exceptions::PyValueError::new_err(format!(
            "Invalid date_order '{}'. Valid values: auto, mdy, us, dmy, eu, european",
            date_order
        ))
    })?;

    // No Python objects are touched below this point, so release the GIL for
    // the (potentially rayon-parallel) pure-Rust conversion work.
    let result = py.detach(|| {
        if parallel {
            convert_csv_to_xlsx_parallel(&input_path, &output_path, &sheet_name, order)
        } else {
            convert_csv_to_xlsx(&input_path, &output_path, &sheet_name, order)
        }
    });
    result.map_err(pyo3::exceptions::PyValueError::new_err)
}

/// Convert a pandas or polars DataFrame to XLSX format.
///
/// This function writes a DataFrame directly to an Excel XLSX file,
/// preserving data types without intermediate CSV conversion.
///
/// Args:
///     df: pandas DataFrame or polars DataFrame to export
///     output_path: Path for the output XLSX file
///     sheet_name: Name of the worksheet (default: "Sheet1")
///     header: Include column names as header row (default: True)
///     autofit: Automatically adjust column widths to fit content (default: False)
///              Combined with column_widths: explicit widths win for the columns
///              they name; every other column is still autofitted (rather than
///              left at Excel's default width). Add an "_all" entry in
///              column_widths to cap the autofit width instead of overriding it.
///     table_style: Apply Excel table formatting with this style name (default: None).
///                  Styles: "Light1"-"Light21", "Medium1"-"Medium28", "Dark1"-"Dark11", "None"
///                  Tables include autofilter dropdowns and banded rows.
///     freeze_panes: Freeze the header row for easier scrolling (default: False)
///     column_widths: Dict mapping column index (0-based) or "_all" to width in characters
///                    (default: None). Example: {0: 20, 1: 15, 3: 30} sets widths for columns
///                    A, B, and D. An integer key must be a non-negative index within Excel's
///                    column range (0..=16383); a negative key, a key beyond 16383, or a
///                    non-integer/non-"_all" key raises. With autofit=True and no "_all" key:
///                    listed columns get the explicit width, unlisted columns are autofitted.
///                    With autofit=True and an "_all" key: "_all" caps the autofit width for
///                    unlisted columns instead of overriding it.
///     row_heights: Dict mapping row index (0-based) to height in points (default: None)
///                  Example: {0: 20, 5: 30} sets heights for specific rows
///     constant_memory: Use constant memory mode for large files (default: False).
///                      Emits a RuntimeWarning and disables: table_style, freeze_panes,
///                      row_heights, autofit, column_widths with autofit cap, conditional_formats,
///                      formula_columns, merged_ranges, hyperlinks, comments, validations,
///                      rich_text, images, checkboxes, textboxes, charts, sparklines, and cells.
///                      Plain column_widths, header_format, and column_formats remain supported.
///     column_formats: Dict mapping column name patterns to format dicts (default: None)
///                     Supports wildcards: "prefix*", "*suffix", "*contains*", or exact match.
///                     Format options: bg_color, font_color, num_format, bold, italic, underline, border.
///                     Example: {"price_*": {"bg_color": "#D6EAF8", "num_format": "$#,##0.00"}}
///     conditional_formats: Dict mapping column names/patterns to conditional format configs (default: None)
///                          Supported types: 2_color_scale, 3_color_scale, data_bar, icon_set, cell
///                          Example: {"score": {"type": "2_color_scale", "min_color": "#FF0000", "max_color": "#00FF00"}}
///     table_name: Custom name for the Excel table (requires table_style; default: auto-generated).
///                 Must be alphanumeric/underscore, max 255 chars.
///     formula_columns: Dict mapping column names to Excel formula templates (default: None).
///                      Use {row} as placeholder for the current row number.
///                      Example: {"Total": "=SUM(A{row}:C{row})"}
///     merged_ranges: List of merge specs: (range, text) or (range, text, format_dict) (default: None).
///                    Example: [("A1:D1", "Title", {"bold": True, "bg_color": "#4F81BD"})]
///     hyperlinks: List of link specs: (cell_ref, url) or (cell_ref, url, display_text) (default: None).
///                 Example: [("A1", "https://example.com", "Click here")]
///     comments: Dict mapping cell refs to note text or config dict (default: None).
///               Example: {"A1": "Note text"} or {"A1": {"text": "Note", "author": "John"}}
///     validations: Dict mapping column names/patterns to validation configs (default: None).
///                  Types: list, whole_number, decimal, text_length
///                  (aliases accepted, e.g. integer/number/length — see README).
///                  Example: {"status": {"type": "list", "values": ["Open", "Closed"]}}
///                  For "whole_number", min/max are bounded to the i32 range
///                  (-2147483648..=2147483647); a value outside that range raises
///                  ValueError naming the field and range.
///     rich_text: Dict mapping cell refs to lists of formatted text segments (default: None).
///                Example: {"A1": [("Bold text", {"bold": True}), (" normal text",)]}
///     images: Dict mapping cell refs to image paths or config dicts (default: None).
///             Example: {"A1": "logo.png"} or {"A1": {"path": "logo.png", "scale_width": 0.5}}
///     checkboxes: Dict mapping cell refs to checkbox state (default: None).
///                 Simple form: {"A1": True, "A2": False}
///                 Dict form with optional cell format: {"A3": {"checked": True, "format": {"bg_color": "#C6EFCE"}}}
///     textboxes: Dict mapping cell refs to floating text shapes (default: None).
///                Simple form: {"B2": "Some text"}
///                Dict form: {"B2": {"text": "Note", "width": 200, "height": 100,
///                            "x_offset": 10, "y_offset": 5,
///                            "font": {"name": "Arial", "size": 14, "bold": True, "color": "#FF0000"},
///                            "fill_color": "#F0F0F0", "line_color": "#000000",
///                            "alt_text": "Descriptive alt text"}}
///     charts: Dict mapping cell refs to native Excel chart configs (default: None).
///             "data_range"/"values_range"/"values" and "categories_range"/"categories"
///             (including a chart-level fallback used by series without their own)
///             must include a sheet name (e.g. "Sheet1!$B$2:$B$10"); a bare range
///             raises ValueError.
///             Example: {"D2": {"type": "bar", "data_range": "Sheet1!$B$2:$B$10",
///                       "categories_range": "Sheet1!$A$2:$A$10", "title": "Monthly Activity"}}
///     sparklines: Dict mapping a location ref to a sparkline (mini in-cell chart) config (default: None).
///                 A single-cell key (e.g. "D2") places one sparkline; a range key (e.g. "D2:D10")
///                 places a grouped sparkline, one per row of the data range.
///                 Required key: "range" (the data to plot, sheet-qualified like a chart range,
///                 e.g. "Sheet1!A2:C10"). Options: type ("line", "column", "win_loss"),
///                 style (1-36), markers, high_point, low_point, first_point, last_point,
///                 negative_points, show_axis, color and the *_point colors, line_weight,
///                 custom_max, custom_min, group_max, group_min, date_range.
///                 Example: {"D2:D10": {"range": "Sheet1!A2:C10", "type": "line", "markers": True}}
///     defined_names: Dict mapping name to Excel reference for workbook-level defined names (default: None).
///                    Example: {"MyRange": "=Sheet1!$A$1:$D$100"}
///     cells: Dict mapping cell refs to values for arbitrary cell writes (default: None).
///            Values can be simple (str, int, float, bool) or dicts with "value" and optional
///            "num_format", "align_horizontal", "align_vertical", and "wrap_text".
///            Cells are written after all DataFrame data, so they can overwrite data cells.
///            Example: {"B9": "Label", "D6": {"value": "934728173849", "num_format": "@"}}
///
/// Returns:
///     Tuple of (rows, columns) written to the Excel file
///
/// Raises:
///     ValueError: If the conversion fails
///
/// Example:
///     >>> import xlsxturbo
///     >>> import pandas as pd
///     >>> df = pd.DataFrame({'name': ['Alice', 'Bob'], 'age': [30, 25]})
///     >>> xlsxturbo.df_to_xlsx(df, "output.xlsx")
///     (3, 2)
///     >>> # With table formatting and auto-width columns:
///     >>> xlsxturbo.df_to_xlsx(df, "styled.xlsx", table_style="Medium9", autofit=True, freeze_panes=True)
///     >>> # With custom column widths and row heights:
///     >>> xlsxturbo.df_to_xlsx(df, "custom.xlsx", column_widths={0: 25, 1: 10}, row_heights={0: 20})
///     >>> # For very large files, use constant_memory mode:
///     >>> xlsxturbo.df_to_xlsx(large_df, "big.xlsx", constant_memory=True)
#[pyfunction]
#[pyo3(signature = (
    df,
    output_path,
    sheet_name = "Sheet1",
    header = true,
    autofit = false,
    table_style = None,
    freeze_panes = false,
    column_widths = None,
    table_name = None,
    header_format = None,
    row_heights = None,
    constant_memory = false,
    column_formats = None,
    conditional_formats = None,
    formula_columns = None,
    merged_ranges = None,
    hyperlinks = None,
    comments = None,
    validations = None,
    rich_text = None,
    images = None,
    checkboxes = None,
    textboxes = None,
    charts = None,
    defined_names = None,
    cells = None,
    sparklines = None,
))]
#[allow(clippy::too_many_arguments)]
fn df_to_xlsx<'py>(
    py: Python<'py>,
    df: &Bound<'py, PyAny>,
    output_path: &Bound<'py, PyAny>,
    sheet_name: &str,
    header: bool,
    autofit: bool,
    table_style: Option<&str>,
    freeze_panes: bool,
    column_widths: Option<&Bound<'py, PyAny>>,
    table_name: Option<String>,
    header_format: Option<&Bound<'py, PyAny>>,
    row_heights: Option<HashMap<u32, f64>>,
    constant_memory: bool,
    column_formats: Option<&Bound<'py, PyAny>>,
    conditional_formats: Option<&Bound<'py, PyAny>>,
    formula_columns: Option<&Bound<'py, PyAny>>,
    merged_ranges: Option<&Bound<'py, PyAny>>,
    hyperlinks: Option<&Bound<'py, PyAny>>,
    comments: Option<&Bound<'py, PyAny>>,
    validations: Option<&Bound<'py, PyAny>>,
    rich_text: Option<&Bound<'py, PyAny>>,
    images: Option<&Bound<'py, PyAny>>,
    checkboxes: Option<&Bound<'py, PyAny>>,
    textboxes: Option<&Bound<'py, PyAny>>,
    charts: Option<&Bound<'py, PyAny>>,
    defined_names: Option<HashMap<String, String>>,
    cells: Option<&Bound<'py, PyAny>>,
    sparklines: Option<&Bound<'py, PyAny>>,
) -> PyResult<(u32, u16)> {
    let output_path = path_arg_to_string(output_path, "output_path")?;
    let opts = extract_options(&RawOptions {
        column_widths,
        header_format,
        column_formats,
        conditional_formats,
        formula_columns,
        merged_ranges,
        hyperlinks,
        comments,
        validations,
        rich_text,
        images,
        checkboxes,
        textboxes,
        charts,
        sparklines,
        cells,
    })?;

    convert_dataframe_to_xlsx(
        py,
        df,
        &output_path,
        sheet_name,
        header,
        autofit,
        table_style,
        freeze_panes,
        table_name.as_deref(),
        row_heights.as_ref(),
        constant_memory,
        &opts,
        defined_names.as_ref(),
    )
    .map_err(pyo3::exceptions::PyValueError::new_err)
}

/// Get the version of the xlsxturbo library
#[pyfunction]
fn version() -> &'static str {
    env!("CARGO_PKG_VERSION")
}

/// Write multiple DataFrames to separate sheets in a single workbook.
///
/// This is a convenience function that writes multiple DataFrames to
/// separate sheets in one workbook, which is more efficient than
/// calling df_to_xlsx multiple times.
///
/// Args:
///     sheets: List of tuples. Each tuple can be:
///             - (DataFrame, sheet_name) - uses global defaults
///             - (DataFrame, sheet_name, options_dict) - per-sheet overrides
///             Options dict keys: header, autofit, table_style, freeze_panes,
///             column_widths, row_heights, table_name, header_format, column_formats,
///             conditional_formats, formula_columns, merged_ranges, hyperlinks,
///             comments, validations, rich_text, images, checkboxes, textboxes, charts,
///             sparklines, cells
///     output_path: Path for the output XLSX file
///     header: Include column names as header row (default: True)
///     autofit: Automatically adjust column widths to fit content (default: False)
///              Combined with column_widths: explicit widths win for the columns
///              they name; every other column is still autofitted (rather than
///              left at Excel's default width). Add an "_all" entry in
///              column_widths to cap the autofit width instead of overriding it.
///     table_style: Apply Excel table formatting with this style name (default: None).
///                  Styles: "Light1"-"Light21", "Medium1"-"Medium28", "Dark1"-"Dark11", "None"
///                  Tables include autofilter dropdowns and banded rows.
///     freeze_panes: Freeze the header row for easier scrolling (default: False)
///     column_widths: Dict mapping column index or "_all" to width in characters (default: None)
///                    Example: {0: 20, "_all": 50} sets col A to 20, caps others at 50. An
///                    integer key must be a non-negative index within Excel's column range
///                    (0..=16383); a negative key, a key beyond 16383, or a
///                    non-integer/non-"_all" key raises. With autofit=True and no "_all"
///                    key: listed columns get the explicit width, unlisted columns are
///                    autofitted. With autofit=True and an "_all" key: "_all" caps the
///                    autofit width for unlisted columns instead of overriding it.
///     table_name: Name for Excel table (requires table_style; default: auto-generated).
///         Effective names must be unique across the workbook after sanitization.
///     header_format: Dict with header cell formatting options (default: None)
///                    Example: {"bold": True, "bg_color": "#4F81BD", "font_color": "white"}
///     row_heights: Dict mapping row index (0-based) to height in points (default: None)
///     constant_memory: Use constant memory mode for large files (default: False).
///                      Emits a RuntimeWarning and disables: table_style, freeze_panes,
///                      row_heights, autofit, column_widths with autofit cap, conditional_formats,
///                      formula_columns, merged_ranges, hyperlinks, comments, validations,
///                      rich_text, images, checkboxes, textboxes, charts, sparklines, and cells.
///                      Plain column_widths, header_format, and column_formats remain supported.
///     column_formats: Dict mapping column name patterns to format dicts (default: None)
///                     Supports wildcards: "prefix*", "*suffix", "*contains*", or exact match.
///                     Format options: bg_color, font_color, num_format, bold, italic, underline, border.
///                     Example: {"price_*": {"bg_color": "#D6EAF8", "num_format": "$#,##0.00"}}
///     conditional_formats: Dict mapping column names to conditional format configs (default: None)
///                          Supported types: 2_color_scale, 3_color_scale, data_bar, icon_set, cell
///                          Example: {"score": {"type": "2_color_scale", "min_color": "#FF0000", "max_color": "#00FF00"}}
///     formula_columns: Dict mapping column names to Excel formula templates (default: None).
///                      Use {row} as placeholder for the current row number.
///     merged_ranges: List of merge specs: (range, text) or (range, text, format_dict) (default: None).
///     hyperlinks: List of link specs: (cell_ref, url) or (cell_ref, url, display_text) (default: None).
///     comments: Dict mapping cell refs to note text or config dict (default: None).
///     validations: Dict mapping column names/patterns to validation configs (default: None).
///                  Types: list, whole_number, decimal, text_length
///                  (aliases accepted, e.g. integer/number/length — see README).
///                  For "whole_number", min/max are bounded to the i32 range
///                  (-2147483648..=2147483647); a value outside that range raises
///                  ValueError naming the field and range.
///     rich_text: Dict mapping cell refs to lists of formatted text segments (default: None).
///     images: Dict mapping cell refs to image paths or config dicts (default: None).
///     checkboxes: Dict mapping cell refs to checkbox state (bool) or config dict (default: None).
///                 Example: {"A1": True} or {"A1": {"checked": True, "format": {"bg_color": "#C6EFCE"}}}
///     textboxes: Dict mapping cell refs to floating text shapes (default: None).
///                Example: {"B2": "text"} or {"B2": {"text": "Note", "width": 200, "font": {"bold": True}}}
///     charts: Dict mapping cell refs to native Excel chart configs (default: None).
///             "data_range"/"values_range"/"values" and "categories_range"/"categories"
///             (including a chart-level fallback used by series without their own)
///             must include a sheet name (e.g. "Sheet1!$B$2:$B$10"); a bare range
///             raises ValueError.
///     sparklines: Dict mapping a location ref to a sparkline (mini in-cell chart) config (default: None).
///                 Range key (e.g. "D2:D10") makes a grouped sparkline; single cell makes one.
///                 "range" must be sheet-qualified, e.g. "Sheet1!A2:C10".
///                 Example: {"D2:D10": {"range": "Sheet1!A2:C10", "type": "line", "markers": True}}
///     defined_names: Dict mapping name to Excel reference for workbook-level defined names (default: None).
///                    Example: {"MyRange": "=Sheet1!$A$1:$D$100"}
///     cells: Dict mapping cell refs to values for arbitrary cell writes (default: None).
///            Values can be simple (str, int, float, bool) or dicts with "value" and optional
///            "num_format", "align_horizontal", "align_vertical", and "wrap_text".
///            Example: {"B9": "Label", "D6": {"value": "934728173849", "num_format": "@"}}
///
/// Returns:
///     List of (rows, columns) tuples for each sheet
///
/// Raises:
///     ValueError: If the conversion fails
///
/// Example:
///     >>> import xlsxturbo
///     >>> import pandas as pd
///     >>> df1 = pd.DataFrame({'a': [1, 2]})
///     >>> df2 = pd.DataFrame({'b': [3, 4]})
///     >>> xlsxturbo.dfs_to_xlsx([(df1, "Sheet1"), (df2, "Sheet2")], "out.xlsx")
///     >>> # With styling applied to all sheets:
///     >>> xlsxturbo.dfs_to_xlsx([(df1, "Sales"), (df2, "Regions")], "report.xlsx",
///     ...                       table_style="Medium9", autofit=True, freeze_panes=True)
///     >>> # With per-sheet options (header=False for one sheet):
///     >>> xlsxturbo.dfs_to_xlsx([
///     ...     (df1, "Data", {"header": True, "table_style": "Medium2"}),
///     ...     (df2, "Instructions", {"header": False})
///     ... ], "report.xlsx", autofit=True)
#[pyfunction]
#[pyo3(signature = (
    sheets,
    output_path,
    header = true,
    autofit = false,
    table_style = None,
    freeze_panes = false,
    column_widths = None,
    table_name = None,
    header_format = None,
    row_heights = None,
    constant_memory = false,
    column_formats = None,
    conditional_formats = None,
    formula_columns = None,
    merged_ranges = None,
    hyperlinks = None,
    comments = None,
    validations = None,
    rich_text = None,
    images = None,
    checkboxes = None,
    textboxes = None,
    charts = None,
    defined_names = None,
    cells = None,
    sparklines = None,
))]
#[allow(clippy::too_many_arguments)]
fn dfs_to_xlsx<'py>(
    py: Python<'py>,
    sheets: Vec<Bound<'py, PyAny>>,
    output_path: &Bound<'py, PyAny>,
    header: bool,
    autofit: bool,
    table_style: Option<&str>,
    freeze_panes: bool,
    column_widths: Option<&Bound<'py, PyAny>>,
    table_name: Option<String>,
    header_format: Option<&Bound<'py, PyAny>>,
    row_heights: Option<HashMap<u32, f64>>,
    constant_memory: bool,
    column_formats: Option<&Bound<'py, PyAny>>,
    conditional_formats: Option<&Bound<'py, PyAny>>,
    formula_columns: Option<&Bound<'py, PyAny>>,
    merged_ranges: Option<&Bound<'py, PyAny>>,
    hyperlinks: Option<&Bound<'py, PyAny>>,
    comments: Option<&Bound<'py, PyAny>>,
    validations: Option<&Bound<'py, PyAny>>,
    rich_text: Option<&Bound<'py, PyAny>>,
    images: Option<&Bound<'py, PyAny>>,
    checkboxes: Option<&Bound<'py, PyAny>>,
    textboxes: Option<&Bound<'py, PyAny>>,
    charts: Option<&Bound<'py, PyAny>>,
    defined_names: Option<HashMap<String, String>>,
    cells: Option<&Bound<'py, PyAny>>,
    sparklines: Option<&Bound<'py, PyAny>>,
) -> PyResult<Vec<(u32, u16)>> {
    let output_path = path_arg_to_string(output_path, "output_path")?;
    if sheets.is_empty() {
        return Err(pyo3::exceptions::PyValueError::new_err(
            "dfs_to_xlsx requires at least one sheet, got an empty list",
        ));
    }
    let mut workbook = Workbook::new();
    let mut stats = Vec::new();
    let mut table_names: HashMap<String, String> = HashMap::new();

    let opts = extract_options(&RawOptions {
        column_widths,
        header_format,
        column_formats,
        conditional_formats,
        formula_columns,
        merged_ranges,
        hyperlinks,
        comments,
        validations,
        rich_text,
        images,
        checkboxes,
        textboxes,
        charts,
        sparklines,
        cells,
    })?;

    for sheet_tuple in sheets {
        let (df, sheet_name, sheet_config) = extract_sheet_info(&sheet_tuple)?;

        // Merge per-sheet scalar options with global defaults
        let effective_header = sheet_config.header.unwrap_or(header);
        let effective_autofit = sheet_config.autofit.unwrap_or(autofit);
        let effective_table_style: Option<String> = match &sheet_config.table_style {
            Some(style_opt) => style_opt.clone(),
            None => table_style.map(|s| s.to_string()),
        };
        let effective_freeze_panes = sheet_config.freeze_panes.unwrap_or(freeze_panes);
        let effective_table_name: Option<String> = sheet_config
            .table_name
            .as_ref()
            .cloned()
            .or_else(|| table_name.clone());
        let effective_row_heights: Option<&HashMap<u32, f64>> =
            sheet_config.row_heights.as_ref().or(row_heights.as_ref());

        // A table is only actually created when there's at least one data row
        // (see the `row_count > 0` gate in `apply_worksheet_features`), so an
        // empty DataFrame never claims a table name here either — otherwise
        // two empty sheets sharing a table name would false-positive as a
        // conflict.
        if !constant_memory && effective_header && effective_table_style.is_some() {
            let row_count = dataframe_row_count(&df).map_err(|e| {
                pyo3::exceptions::PyValueError::new_err(format!("sheet '{}': {}", sheet_name, e))
            })?;
            if row_count > 0 {
                if let Some(name) = effective_table_name.as_deref() {
                    let sanitized = sanitize_table_name(name);
                    let key = sanitized.to_ascii_lowercase();
                    if let Some(previous_sheet) = table_names.insert(key, sheet_name.clone()) {
                        return Err(pyo3::exceptions::PyValueError::new_err(format!(
                            "Duplicate table name '{}' for sheets '{}' and '{}'. Excel table names must be unique within a workbook",
                            sanitized, previous_sheet, sheet_name
                        )));
                    }
                }
            }
        }

        // Merge per-sheet complex options with global defaults (references, no cloning needed)
        let effective_opts = sheet_config.merge_with(&opts);

        let sheet_config_write = WriteConfig {
            include_header: effective_header,
            autofit: effective_autofit,
            table_style: effective_table_style.as_deref(),
            freeze_panes: effective_freeze_panes,
            table_name: effective_table_name.as_deref(),
            row_heights: effective_row_heights,
            constant_memory,
        };

        let result = write_configured_sheet(
            py,
            &mut workbook,
            &df,
            &sheet_name,
            &sheet_config_write,
            effective_opts,
        )
        .map_err(|e| {
            pyo3::exceptions::PyValueError::new_err(format!("sheet '{}': {}", sheet_name, e))
        })?;

        stats.push(result);
    }

    apply_defined_names(&mut workbook, defined_names.as_ref())
        .map_err(pyo3::exceptions::PyValueError::new_err)?;

    // Save workbook
    workbook.save(&output_path).map_err(|e| {
        pyo3::exceptions::PyValueError::new_err(format!(
            "Failed to save workbook to '{}': {}",
            output_path, e
        ))
    })?;

    Ok(stats)
}

/// xlsxturbo - High-performance Excel writer
///
/// A Rust-powered library for converting DataFrames and CSV files to Excel XLSX format.
/// Substantially faster than pandas + openpyxl; see the README for machine-labeled
/// benchmark tables (the durable, reproducible source for performance numbers).
///
/// Features:
/// - Direct DataFrame support (pandas and polars)
/// - Automatic type detection (numbers, booleans, dates, datetimes)
/// - Proper Excel formatting for dates and times
/// - Handles NaN/Inf/None gracefully
/// - Memory-efficient for large files
///
/// Example:
///     >>> import xlsxturbo
///     >>> import pandas as pd
///     >>> df = pd.DataFrame({'a': [1, 2], 'b': [3.14, 2.71]})
///     >>> xlsxturbo.df_to_xlsx(df, "output.xlsx")
///     (3, 2)
#[pymodule]
fn xlsxturbo(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(csv_to_xlsx, m)?)?;
    m.add_function(wrap_pyfunction!(df_to_xlsx, m)?)?;
    m.add_function(wrap_pyfunction!(dfs_to_xlsx, m)?)?;
    m.add_function(wrap_pyfunction!(version, m)?)?;
    m.add("__version__", env!("CARGO_PKG_VERSION"))?;
    Ok(())
}
