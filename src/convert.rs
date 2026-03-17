//! Core conversion functions for CSV and DataFrame to XLSX

use crate::features::{
    apply_cells, apply_column_widths, apply_column_widths_with_autofit_cap, apply_comments,
    apply_conditional_formats, apply_formula_columns, apply_hyperlinks, apply_images,
    apply_merged_ranges, apply_rich_text, apply_validations,
};
use crate::parse::{
    build_column_formats, naive_date_to_excel, naive_datetime_to_excel, parse_header_format,
    parse_table_style, parse_value, sanitize_table_name,
};
use crate::types::{
    extract_columns, is_polars_dataframe, CellValue, DateOrder, EffectiveOpts, ExtractedOptions,
    WriteConfig,
};
use csv::ReaderBuilder;
use pyo3::prelude::*;
use pyo3::types::{PyBool, PyFloat, PyInt, PyString};
use rayon::prelude::*;
use rust_xlsxwriter::{Format, Table, Worksheet, XlsxError};
use std::collections::HashMap;
use std::fs::File;
use std::io::BufReader;

/// Maximum safe integer for lossless f64 representation (2^53).
/// Integers beyond this range lose precision when cast to f64.
const MAX_SAFE_INT: i64 = 1 << 53;

/// Excel number format strings (shared with features::apply_cells)
pub(crate) const DATE_NUM_FORMAT: &str = "yyyy-mm-dd";
pub(crate) const DATETIME_NUM_FORMAT: &str = "yyyy-mm-dd hh:mm:ss";

/// Write a string to a cell, applying column format if provided
fn write_str(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    val: impl Into<String>,
    fmt: Option<&Format>,
) -> Result<(), String> {
    let s = val.into();
    if let Some(f) = fmt {
        worksheet.write_string_with_format(row, col, &s, f)
    } else {
        worksheet.write_string(row, col, &s)
    }
    .map(|_| ())
    .map_err(|e| e.to_string())
}

/// Write a number to a cell, applying column format if provided
fn write_num(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    val: f64,
    fmt: Option<&Format>,
) -> Result<(), String> {
    if let Some(f) = fmt {
        worksheet.write_number_with_format(row, col, val, f)
    } else {
        worksheet.write_number(row, col, val)
    }
    .map(|_| ())
    .map_err(|e| e.to_string())
}

/// Write a boolean to a cell, applying column format if provided
fn write_bool(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    val: bool,
    fmt: Option<&Format>,
) -> Result<(), String> {
    if let Some(f) = fmt {
        worksheet.write_boolean_with_format(row, col, val, f)
    } else {
        worksheet.write_boolean(row, col, val)
    }
    .map(|_| ())
    .map_err(|e| e.to_string())
}

/// Write an integer, falling back to string for values beyond f64 precision (>2^53)
fn write_int(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    val: i64,
    fmt: Option<&Format>,
) -> Result<(), String> {
    if val.abs() > MAX_SAFE_INT {
        write_str(worksheet, row, col, val.to_string(), fmt)
    } else {
        write_num(worksheet, row, col, val as f64, fmt)
    }
}

/// Write a float, treating NaN/Inf as empty
fn write_float(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    val: f64,
    fmt: Option<&Format>,
) -> Result<(), String> {
    if val.is_nan() || val.is_infinite() {
        write_str(worksheet, row, col, "", fmt)
    } else {
        write_num(worksheet, row, col, val, fmt)
    }
}

/// Write a cell value to the worksheet with appropriate formatting
pub(crate) fn write_cell(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    value: CellValue,
    date_format: &Format,
    datetime_format: &Format,
) -> Result<(), XlsxError> {
    match value {
        CellValue::Empty => {
            // Write empty string rather than leaving cell blank so Excel table
            // formatting renders consistently. Trade-off: COUNTA will count these.
            worksheet.write_string(row, col, "")?;
        }
        CellValue::Integer(v) => {
            if v.abs() > MAX_SAFE_INT {
                worksheet.write_string(row, col, v.to_string())?;
            } else {
                worksheet.write_number(row, col, v as f64)?;
            }
        }
        CellValue::Float(v) => {
            worksheet.write_number(row, col, v)?;
        }
        CellValue::Boolean(v) => {
            worksheet.write_boolean(row, col, v)?;
        }
        CellValue::Date(v) => {
            worksheet.write_number_with_format(row, col, v, date_format)?;
        }
        CellValue::DateTime(v) => {
            worksheet.write_number_with_format(row, col, v, datetime_format)?;
        }
        CellValue::String(v) => {
            worksheet.write_string(row, col, &v)?;
        }
    }
    Ok(())
}

/// Convert a CSV file to XLSX format with automatic type detection.
///
/// # Arguments
/// * `input_path` - Path to the input CSV file
/// * `output_path` - Path for the output XLSX file
/// * `sheet_name` - Name of the worksheet (default: "Sheet1")
/// * `date_order` - Date parsing order for ambiguous dates (default: Auto)
///
/// # Returns
/// * `Ok((rows, cols))` - Number of rows and columns written
/// * `Err(message)` - Error description if conversion fails
pub fn convert_csv_to_xlsx(
    input_path: &str,
    output_path: &str,
    sheet_name: &str,
    date_order: DateOrder,
) -> Result<(u32, u16), String> {
    // Open CSV file
    let file = File::open(input_path).map_err(|e| format!("Failed to open input file: {}", e))?;
    let reader = BufReader::with_capacity(1024 * 1024, file);
    let mut csv_reader = ReaderBuilder::new()
        .has_headers(false)
        .flexible(true)
        .from_reader(reader);

    // Create workbook and worksheet
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .set_name(sheet_name)
        .map_err(|e| format!("Failed to set sheet name: {}", e))?;

    // Create formats for dates and datetimes
    let date_format = Format::new().set_num_format(DATE_NUM_FORMAT);
    let datetime_format = Format::new().set_num_format(DATETIME_NUM_FORMAT);

    let mut row_count: u32 = 0;
    let mut col_count: u16 = 0;

    // Process records
    for result in csv_reader.records() {
        let record = result.map_err(|e| format!("CSV parse error at row {}: {}", row_count, e))?;
        let num_cols = u16::try_from(record.len())
            .map_err(|_| format!("Column count {} exceeds u16 limit", record.len()))?;
        if num_cols > col_count {
            col_count = num_cols;
        }

        for (col_idx, value) in record.iter().enumerate() {
            let cell_value = parse_value(value, date_order);
            let col = u16::try_from(col_idx)
                .map_err(|_| format!("Column index {} exceeds u16 limit", col_idx))?;
            write_cell(
                worksheet,
                row_count,
                col,
                cell_value,
                &date_format,
                &datetime_format,
            )
            .map_err(|e| format!("Write error at ({}, {}): {}", row_count, col_idx, e))?;
        }

        row_count = row_count
            .checked_add(1)
            .ok_or("Row count exceeds u32 limit")?;
    }

    // Save workbook
    workbook
        .save(output_path)
        .map_err(|e| format!("Failed to save workbook: {}", e))?;

    Ok((row_count, col_count))
}

/// Convert a CSV file to XLSX format using parallel processing.
///
/// This version reads all records into memory, parses them in parallel,
/// then writes sequentially. Best for large files with complex type detection.
pub fn convert_csv_to_xlsx_parallel(
    input_path: &str,
    output_path: &str,
    sheet_name: &str,
    date_order: DateOrder,
) -> Result<(u32, u16), String> {
    // Open CSV file
    let file = File::open(input_path).map_err(|e| format!("Failed to open input file: {}", e))?;
    let reader = BufReader::with_capacity(1024 * 1024, file);
    let mut csv_reader = ReaderBuilder::new()
        .has_headers(false)
        .flexible(true)
        .from_reader(reader);

    // Read all records into memory
    let records: Vec<Vec<String>> = csv_reader
        .records()
        .enumerate()
        .map(|(row_idx, result)| {
            result
                .map(|record| record.iter().map(|s| s.to_string()).collect())
                .map_err(|e| format!("CSV parse error at row {}: {}", row_idx, e))
        })
        .collect::<Result<Vec<_>, _>>()?;

    let row_count = u32::try_from(records.len())
        .map_err(|_| format!("Row count {} exceeds u32 limit", records.len()))?;
    let max_cols = records.iter().map(|r| r.len()).max().unwrap_or(0);
    let col_count = u16::try_from(max_cols)
        .map_err(|_| format!("Column count {} exceeds u16 limit", max_cols))?;

    // Parse all values in parallel
    let parsed_rows: Vec<Vec<CellValue>> = records
        .par_iter()
        .map(|row| {
            row.iter()
                .map(|value| parse_value(value, date_order))
                .collect()
        })
        .collect();

    // Create workbook and worksheet
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .set_name(sheet_name)
        .map_err(|e| format!("Failed to set sheet name: {}", e))?;

    // Create formats for dates and datetimes
    let date_format = Format::new().set_num_format(DATE_NUM_FORMAT);
    let datetime_format = Format::new().set_num_format(DATETIME_NUM_FORMAT);

    // Write parsed values sequentially
    for (row_idx, row) in parsed_rows.into_iter().enumerate() {
        let row_u32 = u32::try_from(row_idx)
            .map_err(|_| format!("Row index {} exceeds u32 limit", row_idx))?;
        for (col_idx, cell_value) in row.into_iter().enumerate() {
            let col_u16 = u16::try_from(col_idx)
                .map_err(|_| format!("Column index {} exceeds u16 limit", col_idx))?;
            write_cell(
                worksheet,
                row_u32,
                col_u16,
                cell_value,
                &date_format,
                &datetime_format,
            )
            .map_err(|e| format!("Write error at ({}, {}): {}", row_idx, col_idx, e))?;
        }
    }

    // Save workbook
    workbook
        .save(output_path)
        .map_err(|e| format!("Failed to save workbook: {}", e))?;

    Ok((row_count, col_count))
}

// ============================================================================
// DataFrame support
// ============================================================================

/// Write a Python value to the worksheet with optional column format
pub(crate) fn write_py_value_with_format(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    value: &Bound<'_, PyAny>,
    date_format: &Format,
    datetime_format: &Format,
    column_format: Option<&Format>,
) -> Result<(), String> {
    // Check for None first
    if value.is_none() {
        return write_str(worksheet, row, col, "", column_format);
    }

    // Check for pandas NA/NaT
    let type_name = value
        .get_type()
        .name()
        .map_err(|e| format!("Failed to get type name: {}", e))?
        .to_string();
    if type_name == "NAType" || type_name == "NaTType" {
        return write_str(worksheet, row, col, "", column_format);
    }

    // Try boolean first (before int, since bool is subclass of int in Python)
    if let Ok(b) = value.cast::<PyBool>() {
        return write_bool(worksheet, row, col, b.is_true(), column_format);
    }

    // Try datetime (before date, since datetime is subclass of date)
    if type_name == "datetime" || type_name == "Timestamp" {
        let year: i32 = value
            .getattr("year")
            .and_then(|v| v.extract())
            .map_err(|e| format!("Failed to extract datetime year: {}", e))?;
        let month: u32 = value
            .getattr("month")
            .and_then(|v| v.extract())
            .map_err(|e| format!("Failed to extract datetime month: {}", e))?;
        let day: u32 = value
            .getattr("day")
            .and_then(|v| v.extract())
            .map_err(|e| format!("Failed to extract datetime day: {}", e))?;
        let hour: u32 = value
            .getattr("hour")
            .and_then(|v| v.extract())
            .map_err(|e| format!("Failed to extract datetime hour: {}", e))?;
        let minute: u32 = value
            .getattr("minute")
            .and_then(|v| v.extract())
            .map_err(|e| format!("Failed to extract datetime minute: {}", e))?;
        let second: u32 = value
            .getattr("second")
            .and_then(|v| v.extract())
            .map_err(|e| format!("Failed to extract datetime second: {}", e))?;

        let date = chrono::NaiveDate::from_ymd_opt(year, month, day).ok_or_else(|| {
            format!(
                "Invalid datetime date: year={}, month={}, day={}",
                year, month, day
            )
        })?;
        let time = chrono::NaiveTime::from_hms_opt(hour, minute, second).ok_or_else(|| {
            format!(
                "Invalid datetime time: hour={}, minute={}, second={}",
                hour, minute, second
            )
        })?;
        let dt = chrono::NaiveDateTime::new(date, time);
        let excel_dt = naive_datetime_to_excel(dt);
        let fmt = column_format.unwrap_or(datetime_format);
        return write_num(worksheet, row, col, excel_dt, Some(fmt));
    }

    // Try date
    if type_name == "date" {
        let year: i32 = value
            .getattr("year")
            .and_then(|v| v.extract())
            .map_err(|e| format!("Failed to extract date year: {}", e))?;
        let month: u32 = value
            .getattr("month")
            .and_then(|v| v.extract())
            .map_err(|e| format!("Failed to extract date month: {}", e))?;
        let day: u32 = value
            .getattr("day")
            .and_then(|v| v.extract())
            .map_err(|e| format!("Failed to extract date day: {}", e))?;

        let date = chrono::NaiveDate::from_ymd_opt(year, month, day)
            .ok_or_else(|| format!("Invalid date: year={}, month={}, day={}", year, month, day))?;
        let excel_date = naive_date_to_excel(date);
        let fmt = column_format.unwrap_or(date_format);
        return write_num(worksheet, row, col, excel_date, Some(fmt));
    }

    // Try integer
    if let Ok(i) = value.cast::<PyInt>() {
        if let Ok(val) = i.extract::<i64>() {
            return write_int(worksheet, row, col, val, column_format);
        }
    }

    // Try float
    if let Ok(f) = value.cast::<PyFloat>() {
        if let Ok(val) = f.extract::<f64>() {
            return write_float(worksheet, row, col, val, column_format);
        }
    }

    // Try to extract as i64 first (covers numpy int types, before f64 to avoid precision loss)
    if let Ok(val) = value.extract::<i64>() {
        return write_int(worksheet, row, col, val, column_format);
    }

    // Try to extract as f64 (covers numpy float types)
    if let Ok(val) = value.extract::<f64>() {
        return write_float(worksheet, row, col, val, column_format);
    }

    // Try string
    if let Ok(s) = value.cast::<PyString>() {
        return write_str(worksheet, row, col, s.to_string(), column_format);
    }

    // Fallback: convert to string
    let s = value
        .str()
        .map_err(|e| format!("Failed to convert value to string: {}", e))?
        .to_string();
    write_str(worksheet, row, col, s, column_format)
}

/// Write DataFrame data and apply all features to a worksheet.
///
/// This is the shared per-sheet write logic used by both `convert_dataframe_to_xlsx`
/// (single-sheet) and `dfs_to_xlsx` (multi-sheet). The caller is responsible for
/// creating the workbook/worksheet and saving.
pub(crate) fn write_sheet_data(
    py: Python<'_>,
    worksheet: &mut Worksheet,
    df: &Bound<'_, PyAny>,
    config: &WriteConfig<'_>,
    opts: EffectiveOpts<'_>,
) -> Result<(u32, u16), String> {
    // Warn if constant_memory is enabled with incompatible options
    if config.constant_memory {
        let mut disabled: Vec<&str> = Vec::new();
        if config.table_style.is_some() {
            disabled.push("table_style");
        }
        if config.freeze_panes {
            disabled.push("freeze_panes");
        }
        if config.autofit {
            disabled.push("autofit");
        }
        if config.row_heights.is_some() {
            disabled.push("row_heights");
        }
        if opts.formula_columns.is_some() {
            disabled.push("formula_columns");
        }
        if opts.conditional_formats.is_some() {
            disabled.push("conditional_formats");
        }
        if opts.merged_ranges.is_some() {
            disabled.push("merged_ranges");
        }
        if opts.hyperlinks.is_some() {
            disabled.push("hyperlinks");
        }
        if opts.comments.is_some() {
            disabled.push("comments");
        }
        if opts.validations.is_some() {
            disabled.push("validations");
        }
        if opts.rich_text.is_some() {
            disabled.push("rich_text");
        }
        if opts.images.is_some() {
            disabled.push("images");
        }
        if opts.cells.is_some() {
            disabled.push("cells");
        }
        if !disabled.is_empty() {
            let warnings = py
                .import("warnings")
                .map_err(|e| format!("Failed to import warnings: {}", e))?;
            let msg = format!(
                "constant_memory=True disables these features (they will be silently skipped): {}",
                disabled.join(", ")
            );
            warnings
                .call_method1("warn", (msg,))
                .map_err(|e| format!("Failed to emit warning: {}", e))?;
        }
    }

    // Create formats
    let date_format = Format::new().set_num_format(DATE_NUM_FORMAT);
    let datetime_format = Format::new().set_num_format(DATETIME_NUM_FORMAT);

    // Parse header format if provided
    let header_fmt = if let Some(fmt_dict) = opts.header_format {
        Some(parse_header_format(py, fmt_dict)?)
    } else {
        None
    };

    let mut row_idx: u32 = 0;

    // Get column names
    let is_polars = is_polars_dataframe(df)?;
    let columns: Vec<String> = extract_columns(df, is_polars)?;

    let col_count = u16::try_from(columns.len())
        .map_err(|_| format!("Column count {} exceeds u16 limit", columns.len()))?;

    // Build column formats if provided
    let col_formats: Vec<Option<Format>> = if let Some(cf) = opts.column_formats {
        build_column_formats(py, &columns, cf)?
    } else {
        vec![None; columns.len()]
    };

    // Track max content lengths for autofit+cap (only when both are active)
    let track_widths = config.autofit && opts.column_widths.is_some_and(|w| w.contains_key("_all"));
    let mut max_lens = vec![0usize; columns.len()];

    // Write header if requested
    if config.include_header {
        for (col_idx, col_name) in columns.iter().enumerate() {
            let col = col_idx as u16; // safe: col_count already validated via u16::try_from
            if track_widths {
                max_lens[col_idx] = col_name.len();
            }
            if let Some(ref fmt) = header_fmt {
                worksheet
                    .write_string_with_format(row_idx, col, col_name, fmt)
                    .map_err(|e| e.to_string())?;
            } else {
                worksheet
                    .write_string(row_idx, col, col_name)
                    .map_err(|e| e.to_string())?;
            }
        }
        row_idx = 1;
    }

    // Get row count
    let row_count: usize = if df.hasattr("shape").unwrap_or(false) {
        let shape = df
            .getattr("shape")
            .map_err(|e| format!("Failed to get DataFrame shape: {}", e))?;
        let shape_tuple: (usize, usize) = shape
            .extract()
            .map_err(|e| format!("Failed to extract DataFrame shape: {}", e))?;
        shape_tuple.0
    } else {
        df.call_method0("__len__")
            .map_err(|e| format!("Failed to get DataFrame length: {}", e))?
            .extract()
            .map_err(|e| format!("Failed to extract DataFrame length: {}", e))?
    };

    if is_polars {
        // Polars: iterate using rows()
        let rows = df
            .call_method0("iter_rows")
            .map_err(|e| format!("Failed to iterate polars rows: {}", e))?;
        let iter = rows
            .try_iter()
            .map_err(|e| format!("Failed to create polars row iterator: {}", e))?;
        for row_result in iter {
            let row = row_result.map_err(|e| format!("Failed to read polars row: {}", e))?;
            let row_iter = row
                .try_iter()
                .map_err(|e| format!("Failed to iterate polars row values: {}", e))?;
            let row_tuple: Vec<Bound<'_, PyAny>> = row_iter
                .collect::<Result<Vec<_>, _>>()
                .map_err(|e| format!("Failed to collect polars row values: {}", e))?;

            for (col_idx, value) in row_tuple.iter().enumerate() {
                let col = col_idx as u16; // safe: col_count already validated via u16::try_from
                if track_widths {
                    let len = value.str().map(|s| s.to_string_lossy().len()).unwrap_or(0);
                    if len > max_lens[col_idx] {
                        max_lens[col_idx] = len;
                    }
                }
                write_py_value_with_format(
                    worksheet,
                    row_idx,
                    col,
                    value,
                    &date_format,
                    &datetime_format,
                    col_formats.get(col_idx).and_then(|f| f.as_ref()),
                )?;
            }
            row_idx = row_idx
                .checked_add(1)
                .ok_or("Row count exceeds u32 limit")?;
        }
    } else {
        // Pandas: use .values for faster access
        let values = df
            .getattr("values")
            .map_err(|e| format!("Failed to access DataFrame.values: {}", e))?;

        for i in 0..row_count {
            let row = values
                .get_item(i)
                .map_err(|e| format!("Failed to get row {}: {}", i, e))?;

            #[allow(clippy::needless_range_loop)]
            for col_idx in 0..columns.len() {
                let value = row
                    .get_item(col_idx)
                    .map_err(|e| format!("Failed to get value at ({}, {}): {}", i, col_idx, e))?;

                if track_widths {
                    let len = value.str().map(|s| s.to_string_lossy().len()).unwrap_or(0);
                    if len > max_lens[col_idx] {
                        max_lens[col_idx] = len;
                    }
                }

                let col = col_idx as u16; // safe: col_count already validated via u16::try_from
                write_py_value_with_format(
                    worksheet,
                    row_idx,
                    col,
                    &value,
                    &date_format,
                    &datetime_format,
                    col_formats.get(col_idx).and_then(|f| f.as_ref()),
                )?;
            }
            row_idx = row_idx
                .checked_add(1)
                .ok_or("Row count exceeds u32 limit")?;
        }
    }

    // Convert tracked content lengths to approximate Excel column widths
    let content_widths: Vec<f64> = if track_widths {
        max_lens
            .iter()
            .map(|&len| (len as f64 + 1.0).max(8.43))
            .collect()
    } else {
        Vec::new()
    };

    // Apply all worksheet features (table, formulas, formatting, etc.)
    let total_col_count = apply_worksheet_features(
        py,
        worksheet,
        &columns,
        col_count,
        row_idx,
        row_count,
        config,
        header_fmt.as_ref(),
        &opts,
        &content_widths,
    )?;

    Ok((row_idx, total_col_count))
}

/// Apply all worksheet features after data has been written.
///
/// Handles: table formatting, formula columns, conditional formats, freeze panes,
/// column widths/autofit, row heights, merged ranges, hyperlinks, comments,
/// validations, rich text, and images. All features except column widths are
/// skipped in constant_memory mode.
#[allow(clippy::too_many_arguments)]
fn apply_worksheet_features(
    py: Python<'_>,
    worksheet: &mut Worksheet,
    columns: &[String],
    col_count: u16,
    last_row_idx: u32,
    row_count: usize,
    config: &WriteConfig<'_>,
    header_fmt: Option<&Format>,
    opts: &EffectiveOpts<'_>,
    content_widths: &[f64],
) -> Result<u16, String> {
    // In constant_memory mode, only column widths (without autofit) are supported
    if config.constant_memory {
        if let Some(widths) = opts.column_widths {
            apply_column_widths(worksheet, col_count, widths)?;
        }
        return Ok(col_count);
    }

    // Add Excel Table if requested (requires header + at least one data row)
    if let Some(style_name) = config.table_style {
        if row_count > 0 && config.include_header {
            let style = parse_table_style(style_name)?;
            let mut table = Table::new().set_style(style);

            if let Some(name) = config.table_name {
                let sanitized = sanitize_table_name(name);
                table = table.set_name(&sanitized);
            }

            let last_row = last_row_idx.saturating_sub(1);
            let last_col = col_count.saturating_sub(1);

            worksheet
                .add_table(0, 0, last_row, last_col, &table)
                .map_err(|e| format!("Failed to add table: {}", e))?;
        }
    }

    let data_row_start = if config.include_header { 1u32 } else { 0u32 };
    let data_row_end = last_row_idx.saturating_sub(1);
    let has_data_rows = row_count > 0 && data_row_end >= data_row_start;

    // Apply formula columns
    let mut total_col_count = col_count;
    if let Some(formulas) = opts.formula_columns {
        if !formulas.is_empty() && has_data_rows {
            let formula_cols_added = apply_formula_columns(
                worksheet,
                formulas,
                col_count,
                data_row_start,
                data_row_end,
                config.include_header,
                header_fmt,
            )?;
            total_col_count = col_count
                .checked_add(formula_cols_added)
                .ok_or("Total column count exceeds u16 limit")?;
        }
    }

    // Apply conditional formats
    if let Some(cond_fmts) = opts.conditional_formats {
        if has_data_rows {
            apply_conditional_formats(
                py,
                worksheet,
                columns,
                data_row_start,
                data_row_end,
                cond_fmts,
            )?;
        }
    }

    // Freeze panes (freeze header row)
    if config.freeze_panes && config.include_header {
        worksheet
            .set_freeze_panes(1, 0)
            .map_err(|e| format!("Failed to freeze panes: {}", e))?;
    }

    // Apply custom column widths and/or autofit
    if let Some(widths) = opts.column_widths {
        if config.autofit && widths.contains_key("_all") {
            apply_column_widths_with_autofit_cap(worksheet, col_count, widths, content_widths)?;
        } else {
            apply_column_widths(worksheet, col_count, widths)?;
        }
    } else if config.autofit {
        worksheet.autofit();
    }

    // Apply custom row heights
    if let Some(heights) = config.row_heights {
        for (&row_idx_h, &height) in heights.iter() {
            worksheet
                .set_row_height(row_idx_h, height)
                .map_err(|e| format!("Failed to set row height: {}", e))?;
        }
    }

    // Apply merged ranges
    if let Some(ranges) = opts.merged_ranges {
        if !ranges.is_empty() {
            apply_merged_ranges(py, worksheet, ranges)?;
        }
    }

    // Apply hyperlinks
    if let Some(links) = opts.hyperlinks {
        if !links.is_empty() {
            apply_hyperlinks(worksheet, links)?;
        }
    }

    // Apply comments/notes
    if let Some(cmts) = opts.comments {
        if !cmts.is_empty() {
            apply_comments(worksheet, cmts)?;
        }
    }

    // Apply data validations
    if let Some(vals) = opts.validations {
        if has_data_rows {
            apply_validations(py, worksheet, columns, data_row_start, data_row_end, vals)?;
        }
    }

    // Apply rich text
    if let Some(rt) = opts.rich_text {
        if !rt.is_empty() {
            apply_rich_text(py, worksheet, rt)?;
        }
    }

    // Apply images
    if let Some(imgs) = opts.images {
        if !imgs.is_empty() {
            apply_images(py, worksheet, imgs)?;
        }
    }

    // Apply cells (arbitrary cell writes, after all DataFrame data)
    if let Some(cells) = opts.cells {
        if !cells.is_empty() {
            apply_cells(py, worksheet, cells)?;
        }
    }

    Ok(total_col_count)
}

/// Convert a DataFrame (pandas or polars) to XLSX format
#[allow(clippy::too_many_arguments)]
pub(crate) fn convert_dataframe_to_xlsx(
    py: Python<'_>,
    df: &Bound<'_, PyAny>,
    output_path: &str,
    sheet_name: &str,
    include_header: bool,
    autofit: bool,
    table_style: Option<&str>,
    freeze_panes: bool,
    table_name: Option<&str>,
    row_heights: Option<&HashMap<u32, f64>>,
    constant_memory: bool,
    opts: &ExtractedOptions,
    defined_names: Option<&HashMap<String, String>>,
) -> Result<(u32, u16), String> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = if constant_memory {
        workbook.add_worksheet_with_constant_memory()
    } else {
        workbook.add_worksheet()
    };
    worksheet
        .set_name(sheet_name)
        .map_err(|e| format!("Failed to set sheet name: {}", e))?;

    let config = WriteConfig {
        include_header,
        autofit,
        table_style,
        freeze_panes,
        table_name,
        row_heights,
        constant_memory,
    };

    let result = write_sheet_data(py, worksheet, df, &config, opts.as_effective())?;

    // Apply defined names (workbook-level)
    if let Some(names) = defined_names {
        for (name, reference) in names {
            workbook
                .define_name(name, reference)
                .map_err(|e| format!("Failed to define name '{}': {}", name, e))?;
        }
    }

    workbook
        .save(output_path)
        .map_err(|e| format!("Failed to save workbook: {}", e))?;

    Ok(result)
}
