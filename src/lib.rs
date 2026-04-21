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

// Re-export public API for the CLI binary (main.rs)
pub use convert::{convert_csv_to_xlsx, convert_csv_to_xlsx_parallel};
pub use types::DateOrder;

use convert::{convert_dataframe_to_xlsx, write_sheet_data};
use extract::{
    extract_cells, extract_checkboxes, extract_column_formats, extract_column_widths,
    extract_comments, extract_conditional_formats, extract_formula_columns, extract_header_format,
    extract_hyperlinks, extract_images, extract_merged_ranges, extract_rich_text,
    extract_sheet_info, extract_textboxes, extract_validations,
};
use types::pytype_name;
use types::ExtractedOptions;
use types::WriteConfig;

use pyo3::prelude::*;
use rust_xlsxwriter::Workbook;
use std::collections::HashMap;

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

/// Extract and validate all optional write parameters from Python into typed Rust structs.
#[allow(clippy::too_many_arguments)]
fn extract_options<'py>(
    column_widths: Option<&Bound<'py, PyAny>>,
    header_format: Option<&Bound<'py, PyAny>>,
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
    cells: Option<&Bound<'py, PyAny>>,
) -> PyResult<ExtractedOptions> {
    Ok(ExtractedOptions {
        column_widths: column_widths
            .map(|v| require_dict(v, "column_widths").and_then(|d| extract_column_widths(&d)))
            .transpose()?,
        header_format: header_format
            .map(|v| require_dict(v, "header_format").and_then(|d| extract_header_format(&d)))
            .transpose()?,
        column_formats: column_formats
            .map(|v| require_dict(v, "column_formats").and_then(|d| extract_column_formats(&d)))
            .transpose()?,
        conditional_formats: conditional_formats
            .map(|v| {
                require_dict(v, "conditional_formats").and_then(|d| extract_conditional_formats(&d))
            })
            .transpose()?,
        formula_columns: formula_columns
            .map(|v| require_dict(v, "formula_columns").and_then(|d| extract_formula_columns(&d)))
            .transpose()?,
        merged_ranges: merged_ranges
            .map(|v| require_list(v, "merged_ranges").and_then(|l| extract_merged_ranges(&l)))
            .transpose()?,
        hyperlinks: hyperlinks
            .map(|v| require_list(v, "hyperlinks").and_then(|l| extract_hyperlinks(&l)))
            .transpose()?,
        comments: comments
            .map(|v| require_dict(v, "comments").and_then(|d| extract_comments(&d)))
            .transpose()?,
        validations: validations
            .map(|v| require_dict(v, "validations").and_then(|d| extract_validations(&d)))
            .transpose()?,
        rich_text: rich_text
            .map(|v| require_dict(v, "rich_text").and_then(|d| extract_rich_text(&d)))
            .transpose()?,
        images: images
            .map(|v| require_dict(v, "images").and_then(|d| extract_images(&d)))
            .transpose()?,
        checkboxes: checkboxes
            .map(|v| require_dict(v, "checkboxes").and_then(|d| extract_checkboxes(&d)))
            .transpose()?,
        textboxes: textboxes
            .map(|v| require_dict(v, "textboxes").and_then(|d| extract_textboxes(&d)))
            .transpose()?,
        cells: cells
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
    input_path: &str,
    output_path: &str,
    sheet_name: &str,
    parallel: bool,
    date_order: &str,
) -> PyResult<(u32, u16)> {
    let order = DateOrder::parse(date_order).ok_or_else(|| {
        pyo3::exceptions::PyValueError::new_err(format!(
            "Invalid date_order '{}'. Valid values: auto, mdy, us, dmy, eu",
            date_order
        ))
    })?;

    let result = if parallel {
        convert_csv_to_xlsx_parallel(input_path, output_path, sheet_name, order)
    } else {
        convert_csv_to_xlsx(input_path, output_path, sheet_name, order)
    };
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
///     table_style: Apply Excel table formatting with this style name (default: None).
///                  Styles: "Light1"-"Light21", "Medium1"-"Medium28", "Dark1"-"Dark11", "None"
///                  Tables include autofilter dropdowns and banded rows.
///     freeze_panes: Freeze the header row for easier scrolling (default: False)
///     column_widths: Dict mapping column index (0-based) to width in characters (default: None)
///                    Example: {0: 20, 1: 15, 3: 30} sets widths for columns A, B, and D
///     row_heights: Dict mapping row index (0-based) to height in points (default: None)
///                  Example: {0: 20, 5: 30} sets heights for specific rows
///     constant_memory: Use constant memory mode for large files (default: False).
///                      Reduces memory usage but disables table_style, freeze_panes,
///                      row_heights, and autofit features.
///     column_formats: Dict mapping column name patterns to format dicts (default: None)
///                     Supports wildcards: "prefix*", "*suffix", "*contains*", or exact match.
///                     Format options: bg_color, font_color, num_format, bold, italic, underline, border.
///                     Example: {"price_*": {"bg_color": "#D6EAF8", "num_format": "$#,##0.00"}}
///     conditional_formats: Dict mapping column names/patterns to conditional format configs (default: None)
///                          Supported types: 2_color_scale, 3_color_scale, data_bar, icon_set
///                          Example: {"score": {"type": "2_color_scale", "min_color": "#FF0000", "max_color": "#00FF00"}}
///     table_name: Custom name for the Excel table (default: auto-generated).
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
///                  Types: list, whole_number, decimal, text_length.
///                  Example: {"status": {"type": "list", "values": ["Open", "Closed"]}}
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
///     defined_names: Dict mapping name to Excel reference for workbook-level defined names (default: None).
///                    Example: {"MyRange": "=Sheet1!$A$1:$D$100"}
///     cells: Dict mapping cell refs to values for arbitrary cell writes (default: None).
///            Values can be simple (str, int, float, bool) or dicts with "value" and optional "num_format".
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
    defined_names = None,
    cells = None,
))]
#[allow(clippy::too_many_arguments)]
fn df_to_xlsx<'py>(
    py: Python<'py>,
    df: &Bound<'py, PyAny>,
    output_path: &str,
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
    defined_names: Option<HashMap<String, String>>,
    cells: Option<&Bound<'py, PyAny>>,
) -> PyResult<(u32, u16)> {
    let opts = extract_options(
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
        cells,
    )?;

    convert_dataframe_to_xlsx(
        py,
        df,
        output_path,
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
///             comments, validations, rich_text, images, checkboxes, textboxes, cells
///     output_path: Path for the output XLSX file
///     header: Include column names as header row (default: True)
///     autofit: Automatically adjust column widths to fit content (default: False)
///     table_style: Apply Excel table formatting with this style name (default: None).
///                  Styles: "Light1"-"Light21", "Medium1"-"Medium28", "Dark1"-"Dark11", "None"
///                  Tables include autofilter dropdowns and banded rows.
///     freeze_panes: Freeze the header row for easier scrolling (default: False)
///     column_widths: Dict mapping column index or "_all" to width in characters (default: None)
///                    Example: {0: 20, "_all": 50} sets col A to 20, caps others at 50
///     table_name: Name for Excel table (default: auto-generated)
///     header_format: Dict with header cell formatting options (default: None)
///                    Example: {"bold": True, "bg_color": "#4F81BD", "font_color": "white"}
///     row_heights: Dict mapping row index (0-based) to height in points (default: None)
///     constant_memory: Use constant memory mode for large files (default: False).
///     column_formats: Dict mapping column name patterns to format dicts (default: None)
///                     Supports wildcards: "prefix*", "*suffix", "*contains*", or exact match.
///                     Format options: bg_color, font_color, num_format, bold, italic, underline.
///                     Example: {"price_*": {"bg_color": "#D6EAF8", "num_format": "$#,##0.00"}}
///     conditional_formats: Dict mapping column names to conditional format configs (default: None)
///                          Supported types: 2_color_scale, 3_color_scale, data_bar, icon_set
///                          Example: {"score": {"type": "2_color_scale", "min_color": "#FF0000", "max_color": "#00FF00"}}
///     formula_columns: Dict mapping column names to Excel formula templates (default: None).
///                      Use {row} as placeholder for the current row number.
///     merged_ranges: List of merge specs: (range, text) or (range, text, format_dict) (default: None).
///     hyperlinks: List of link specs: (cell_ref, url) or (cell_ref, url, display_text) (default: None).
///     comments: Dict mapping cell refs to note text or config dict (default: None).
///     validations: Dict mapping column names/patterns to validation configs (default: None).
///                  Types: list, whole_number, decimal, text_length.
///     rich_text: Dict mapping cell refs to lists of formatted text segments (default: None).
///     images: Dict mapping cell refs to image paths or config dicts (default: None).
///     checkboxes: Dict mapping cell refs to checkbox state (bool) or config dict (default: None).
///                 Example: {"A1": True} or {"A1": {"checked": True, "format": {"bg_color": "#C6EFCE"}}}
///     textboxes: Dict mapping cell refs to floating text shapes (default: None).
///                Example: {"B2": "text"} or {"B2": {"text": "Note", "width": 200, "font": {"bold": True}}}
///     defined_names: Dict mapping name to Excel reference for workbook-level defined names (default: None).
///                    Example: {"MyRange": "=Sheet1!$A$1:$D$100"}
///     cells: Dict mapping cell refs to values for arbitrary cell writes (default: None).
///            Values can be simple (str, int, float, bool) or dicts with "value" and optional "num_format".
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
    defined_names = None,
    cells = None,
))]
#[allow(clippy::too_many_arguments)]
fn dfs_to_xlsx<'py>(
    py: Python<'py>,
    sheets: Vec<Bound<'py, PyAny>>,
    output_path: &str,
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
    defined_names: Option<HashMap<String, String>>,
    cells: Option<&Bound<'py, PyAny>>,
) -> PyResult<Vec<(u32, u16)>> {
    let mut workbook = Workbook::new();
    let mut stats = Vec::new();

    let opts = extract_options(
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
        cells,
    )?;

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

        // Merge per-sheet complex options with global defaults (references, no cloning needed)
        let effective_opts = sheet_config.merge_with(&opts);

        let worksheet = if constant_memory {
            workbook.add_worksheet_with_constant_memory()
        } else {
            workbook.add_worksheet()
        };
        worksheet.set_name(&sheet_name).map_err(|e| {
            pyo3::exceptions::PyValueError::new_err(format!(
                "Failed to set sheet name '{}': {}",
                sheet_name, e
            ))
        })?;

        let sheet_config_write = WriteConfig {
            include_header: effective_header,
            autofit: effective_autofit,
            table_style: effective_table_style.as_deref(),
            freeze_panes: effective_freeze_panes,
            table_name: effective_table_name.as_deref(),
            row_heights: effective_row_heights,
            constant_memory,
        };

        let result = write_sheet_data(py, worksheet, &df, &sheet_config_write, effective_opts)
            .map_err(pyo3::exceptions::PyValueError::new_err)?;

        stats.push(result);
    }

    // Apply defined names (workbook-level)
    if let Some(ref names) = defined_names {
        for (name, reference) in names {
            workbook.define_name(name, reference).map_err(|e| {
                pyo3::exceptions::PyValueError::new_err(format!(
                    "Failed to define name '{}': {}",
                    name, e
                ))
            })?;
        }
    }

    // Save workbook
    workbook
        .save(output_path)
        .map_err(|e| pyo3::exceptions::PyValueError::new_err(format!("Failed to save: {}", e)))?;

    Ok(stats)
}

/// xlsxturbo - High-performance Excel writer
///
/// A Rust-powered library for converting DataFrames and CSV files to Excel XLSX format.
/// ~6x faster than pandas + openpyxl (benchmark: 100K rows x 50 columns).
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
