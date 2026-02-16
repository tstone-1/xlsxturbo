//! Core conversion functions for CSV and DataFrame to XLSX

use crate::features::{
    apply_column_widths, apply_column_widths_with_autofit_cap, apply_comments,
    apply_conditional_formats, apply_formula_columns, apply_hyperlinks, apply_images,
    apply_merged_ranges, apply_rich_text, apply_validations,
};
use crate::parse::{
    build_column_formats, naive_date_to_excel, naive_datetime_to_excel, parse_header_format,
    parse_table_style, parse_value, sanitize_table_name,
};
use crate::types::{extract_columns, is_polars_dataframe, CellValue, DateOrder, ExtractedOptions};
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
    let date_format = Format::new().set_num_format("yyyy-mm-dd");
    let datetime_format = Format::new().set_num_format("yyyy-mm-dd hh:mm:ss");

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

        row_count += 1;
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
    let date_format = Format::new().set_num_format("yyyy-mm-dd");
    let datetime_format = Format::new().set_num_format("yyyy-mm-dd hh:mm:ss");

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
        if let Some(fmt) = column_format {
            worksheet
                .write_string_with_format(row, col, "", fmt)
                .map_err(|e| e.to_string())?;
        } else {
            worksheet
                .write_string(row, col, "")
                .map_err(|e| e.to_string())?;
        }
        return Ok(());
    }

    // Check for pandas NA/NaT
    let type_name = value
        .get_type()
        .name()
        .map_err(|e| e.to_string())?
        .to_string();
    if type_name == "NAType" || type_name == "NaTType" {
        if let Some(fmt) = column_format {
            worksheet
                .write_string_with_format(row, col, "", fmt)
                .map_err(|e| e.to_string())?;
        } else {
            worksheet
                .write_string(row, col, "")
                .map_err(|e| e.to_string())?;
        }
        return Ok(());
    }

    // Try boolean first (before int, since bool is subclass of int in Python)
    if let Ok(b) = value.cast::<PyBool>() {
        worksheet
            .write_boolean(row, col, b.is_true())
            .map_err(|e| e.to_string())?;
        return Ok(());
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

        if let Some(date) = chrono::NaiveDate::from_ymd_opt(year, month, day) {
            if let Some(time) = chrono::NaiveTime::from_hms_opt(hour, minute, second) {
                let dt = chrono::NaiveDateTime::new(date, time);
                let excel_dt = naive_datetime_to_excel(dt);
                // For datetime, use column format if provided, otherwise datetime_format
                let fmt = column_format.unwrap_or(datetime_format);
                worksheet
                    .write_number_with_format(row, col, excel_dt, fmt)
                    .map_err(|e| e.to_string())?;
                return Ok(());
            }
        }
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

        if let Some(date) = chrono::NaiveDate::from_ymd_opt(year, month, day) {
            let excel_date = naive_date_to_excel(date);
            // For date, use column format if provided, otherwise date_format
            let fmt = column_format.unwrap_or(date_format);
            worksheet
                .write_number_with_format(row, col, excel_date, fmt)
                .map_err(|e| e.to_string())?;
            return Ok(());
        }
    }

    // Try integer
    if let Ok(i) = value.cast::<PyInt>() {
        if let Ok(val) = i.extract::<i64>() {
            if val.abs() > MAX_SAFE_INT {
                // Write as string to avoid precision loss for large integers
                if let Some(fmt) = column_format {
                    worksheet
                        .write_string_with_format(row, col, val.to_string(), fmt)
                        .map_err(|e| e.to_string())?;
                } else {
                    worksheet
                        .write_string(row, col, val.to_string())
                        .map_err(|e| e.to_string())?;
                }
            } else if let Some(fmt) = column_format {
                worksheet
                    .write_number_with_format(row, col, val as f64, fmt)
                    .map_err(|e| e.to_string())?;
            } else {
                worksheet
                    .write_number(row, col, val as f64)
                    .map_err(|e| e.to_string())?;
            }
            return Ok(());
        }
    }

    // Try float
    if let Ok(f) = value.cast::<PyFloat>() {
        if let Ok(val) = f.extract::<f64>() {
            if val.is_nan() || val.is_infinite() {
                if let Some(fmt) = column_format {
                    worksheet
                        .write_string_with_format(row, col, "", fmt)
                        .map_err(|e| e.to_string())?;
                } else {
                    worksheet
                        .write_string(row, col, "")
                        .map_err(|e| e.to_string())?;
                }
            } else if let Some(fmt) = column_format {
                worksheet
                    .write_number_with_format(row, col, val, fmt)
                    .map_err(|e| e.to_string())?;
            } else {
                worksheet
                    .write_number(row, col, val)
                    .map_err(|e| e.to_string())?;
            }
            return Ok(());
        }
    }

    // Try to extract as i64 first (covers numpy int types, before f64 to avoid precision loss)
    if let Ok(val) = value.extract::<i64>() {
        if val.abs() > MAX_SAFE_INT {
            if let Some(fmt) = column_format {
                worksheet
                    .write_string_with_format(row, col, val.to_string(), fmt)
                    .map_err(|e| e.to_string())?;
            } else {
                worksheet
                    .write_string(row, col, val.to_string())
                    .map_err(|e| e.to_string())?;
            }
        } else if let Some(fmt) = column_format {
            worksheet
                .write_number_with_format(row, col, val as f64, fmt)
                .map_err(|e| e.to_string())?;
        } else {
            worksheet
                .write_number(row, col, val as f64)
                .map_err(|e| e.to_string())?;
        }
        return Ok(());
    }

    // Try to extract as f64 (covers numpy float types)
    if let Ok(val) = value.extract::<f64>() {
        if val.is_nan() || val.is_infinite() {
            if let Some(fmt) = column_format {
                worksheet
                    .write_string_with_format(row, col, "", fmt)
                    .map_err(|e| e.to_string())?;
            } else {
                worksheet
                    .write_string(row, col, "")
                    .map_err(|e| e.to_string())?;
            }
        } else if let Some(fmt) = column_format {
            worksheet
                .write_number_with_format(row, col, val, fmt)
                .map_err(|e| e.to_string())?;
        } else {
            worksheet
                .write_number(row, col, val)
                .map_err(|e| e.to_string())?;
        }
        return Ok(());
    }

    // Try to extract as bool
    if let Ok(val) = value.extract::<bool>() {
        worksheet
            .write_boolean(row, col, val)
            .map_err(|e| e.to_string())?;
        return Ok(());
    }

    // Try string
    if let Ok(s) = value.cast::<PyString>() {
        if let Some(fmt) = column_format {
            worksheet
                .write_string_with_format(row, col, s.to_string(), fmt)
                .map_err(|e| e.to_string())?;
        } else {
            worksheet
                .write_string(row, col, s.to_string())
                .map_err(|e| e.to_string())?;
        }
        return Ok(());
    }

    // Fallback: convert to string
    let s = value.str().map_err(|e| e.to_string())?.to_string();
    if let Some(fmt) = column_format {
        worksheet
            .write_string_with_format(row, col, &s, fmt)
            .map_err(|e| e.to_string())?;
    } else {
        worksheet
            .write_string(row, col, &s)
            .map_err(|e| e.to_string())?;
    }

    Ok(())
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
) -> Result<(u32, u16), String> {
    // Create workbook and worksheet
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = if constant_memory {
        workbook.add_worksheet_with_constant_memory()
    } else {
        workbook.add_worksheet()
    };
    worksheet
        .set_name(sheet_name)
        .map_err(|e| format!("Failed to set sheet name: {}", e))?;

    // Create formats
    let date_format = Format::new().set_num_format("yyyy-mm-dd");
    let datetime_format = Format::new().set_num_format("yyyy-mm-dd hh:mm:ss");

    // Parse header format if provided
    let header_fmt = if let Some(ref fmt_dict) = opts.header_format {
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
    let col_formats: Vec<Option<Format>> = if let Some(ref cf) = opts.column_formats {
        build_column_formats(py, &columns, cf)?
    } else {
        vec![None; columns.len()]
    };

    // Write header if requested (and not using table, since table handles headers)
    if include_header && table_style.is_none() {
        for (col_idx, col_name) in columns.iter().enumerate() {
            let col = col_idx as u16; // safe: col_count already validated via u16::try_from
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
        row_idx += 1;
    }

    // If using table with header, write header in row 0
    let data_start_row = if table_style.is_some() && include_header {
        for (col_idx, col_name) in columns.iter().enumerate() {
            let col = col_idx as u16; // safe: col_count already validated via u16::try_from
            if let Some(ref fmt) = header_fmt {
                worksheet
                    .write_string_with_format(0, col, col_name, fmt)
                    .map_err(|e| e.to_string())?;
            } else {
                worksheet
                    .write_string(0, col, col_name)
                    .map_err(|e| e.to_string())?;
            }
        }
        row_idx = 1;
        0u32
    } else {
        row_idx.saturating_sub(1)
    };

    // Get row count
    let row_count: usize = if df.hasattr("shape").unwrap_or(false) {
        let shape = df
            .getattr("shape")
            .map_err(|e: pyo3::PyErr| e.to_string())?;
        let shape_tuple: (usize, usize) =
            shape.extract().map_err(|e: pyo3::PyErr| e.to_string())?;
        shape_tuple.0
    } else {
        df.call_method0("__len__")
            .map_err(|e: pyo3::PyErr| e.to_string())?
            .extract()
            .map_err(|e: pyo3::PyErr| e.to_string())?
    };

    if is_polars {
        // Polars: iterate using rows()
        let rows = df.call_method0("iter_rows").map_err(|e| e.to_string())?;
        let iter = rows.try_iter().map_err(|e| e.to_string())?;
        for row_result in iter {
            let row = row_result.map_err(|e| e.to_string())?;
            let row_iter = row.try_iter().map_err(|e| e.to_string())?;
            let row_tuple: Vec<Bound<'_, PyAny>> = row_iter
                .collect::<Result<Vec<_>, _>>()
                .map_err(|e: PyErr| e.to_string())?;

            for (col_idx, value) in row_tuple.iter().enumerate() {
                let col = col_idx as u16; // safe: col_count already validated via u16::try_from
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
            row_idx += 1;
        }
    } else {
        // Pandas: use .values for faster access
        let values = df.getattr("values").map_err(|e| e.to_string())?;

        for i in 0..row_count {
            let row = values
                .get_item(i)
                .map_err(|e| format!("Failed to get row {}: {}", i, e))?;

            for col_idx in 0..columns.len() {
                let value = row
                    .get_item(col_idx)
                    .map_err(|e| format!("Failed to get value at ({}, {}): {}", i, col_idx, e))?;

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
            row_idx += 1;
        }
    }

    // Add Excel Table if requested (not supported in constant_memory mode)
    // Tables require at least one data row, so skip if DataFrame is empty
    if let Some(style_name) = table_style {
        if !constant_memory && row_count > 0 {
            let style = parse_table_style(style_name)?;
            let mut table = Table::new().set_style(style);

            // Apply table name if provided
            if let Some(name) = table_name {
                let sanitized = sanitize_table_name(name);
                table = table.set_name(&sanitized);
            }

            let last_row = row_idx.saturating_sub(1);
            let last_col = col_count.saturating_sub(1);

            if last_row >= data_start_row {
                worksheet
                    .add_table(data_start_row, 0, last_row, last_col, &table)
                    .map_err(|e| format!("Failed to add table: {}", e))?;
            }
        }
    }

    // Apply formula columns (append calculated columns after data)
    // Formula columns are added after the original data columns
    let mut total_col_count = col_count;
    if let Some(ref formulas) = opts.formula_columns {
        if !formulas.is_empty() && row_count > 0 {
            let data_row_start = if include_header { 1u32 } else { 0u32 };
            let data_row_end = row_idx.saturating_sub(1);
            if data_row_end >= data_row_start {
                let formula_cols_added = apply_formula_columns(
                    worksheet,
                    formulas,
                    col_count, // Start after original data columns
                    data_row_start,
                    data_row_end,
                    header_fmt.as_ref(),
                )?;
                total_col_count += formula_cols_added;
            }
        }
    }

    // Apply conditional formats (not supported in constant_memory mode)
    if let Some(ref cond_fmts) = opts.conditional_formats {
        if !constant_memory && row_count > 0 {
            let data_row_start = if include_header { 1 } else { 0 };
            let data_row_end = row_idx.saturating_sub(1);
            if data_row_end >= data_row_start {
                apply_conditional_formats(
                    py,
                    worksheet,
                    &columns,
                    data_row_start,
                    data_row_end,
                    cond_fmts,
                )?;
            }
        }
    }

    // Freeze panes (freeze header row) - not supported in constant_memory mode
    if freeze_panes && include_header && !constant_memory {
        worksheet
            .set_freeze_panes(1, 0)
            .map_err(|e| format!("Failed to freeze panes: {}", e))?;
    }

    // Apply custom column widths and/or autofit
    if let Some(ref widths) = opts.column_widths {
        if autofit && widths.contains_key("_all") && !constant_memory {
            // Autofit first, then apply cap from '_all' and specific widths
            apply_column_widths_with_autofit_cap(worksheet, col_count, widths, constant_memory)?;
        } else {
            // Just apply the specified widths
            apply_column_widths(worksheet, col_count, widths)?;
        }
    } else if autofit && !constant_memory {
        // Just autofit, no width constraints
        worksheet.autofit();
    }

    // Apply custom row heights if specified (not supported in constant_memory mode)
    if let Some(heights) = row_heights {
        if !constant_memory {
            for (&row_idx_h, &height) in heights.iter() {
                worksheet
                    .set_row_height(row_idx_h, height)
                    .map_err(|e| format!("Failed to set row height: {}", e))?;
            }
        }
    }

    // Apply merged ranges (not supported in constant_memory mode)
    if let Some(ref ranges) = opts.merged_ranges {
        if !constant_memory && !ranges.is_empty() {
            apply_merged_ranges(py, worksheet, ranges)?;
        }
    }

    // Apply hyperlinks (not supported in constant_memory mode)
    if let Some(ref links) = opts.hyperlinks {
        if !constant_memory && !links.is_empty() {
            apply_hyperlinks(worksheet, links)?;
        }
    }

    // Apply comments/notes (not supported in constant_memory mode)
    if let Some(ref cmts) = opts.comments {
        if !constant_memory && !cmts.is_empty() {
            apply_comments(worksheet, cmts)?;
        }
    }

    // Apply data validations (not supported in constant_memory mode)
    if let Some(ref vals) = opts.validations {
        if !constant_memory && row_count > 0 {
            let data_row_start = if include_header { 1 } else { 0 };
            let data_row_end = row_idx.saturating_sub(1);
            if data_row_end >= data_row_start {
                apply_validations(py, worksheet, &columns, data_row_start, data_row_end, vals)?;
            }
        }
    }

    // Apply rich text (not supported in constant_memory mode)
    if let Some(ref rt) = opts.rich_text {
        if !constant_memory && !rt.is_empty() {
            apply_rich_text(py, worksheet, rt)?;
        }
    }

    // Apply images (not supported in constant_memory mode)
    if let Some(ref imgs) = opts.images {
        if !constant_memory && !imgs.is_empty() {
            apply_images(py, worksheet, imgs)?;
        }
    }

    // Save workbook
    workbook
        .save(output_path)
        .map_err(|e| format!("Failed to save workbook: {}", e))?;

    Ok((row_idx, total_col_count))
}
