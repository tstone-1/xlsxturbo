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

use chrono::Timelike;
use csv::ReaderBuilder;
use pyo3::prelude::*;
use pyo3::types::{PyBool, PyFloat, PyInt, PyString};
use rayon::prelude::*;
use rust_xlsxwriter::{Format, Table, TableStyle, Workbook, Worksheet, XlsxError};
use std::fs::File;
use std::io::BufReader;

/// Date formats we recognize
const DATE_PATTERNS: &[&str] = &[
    "%Y-%m-%d", // 2024-01-15
    "%Y/%m/%d", // 2024/01/15
    "%d-%m-%Y", // 15-01-2024
    "%d/%m/%Y", // 15/01/2024
    "%m-%d-%Y", // 01-15-2024
    "%m/%d/%Y", // 01/15/2024
];

/// Datetime formats we recognize
const DATETIME_PATTERNS: &[&str] = &[
    "%Y-%m-%dT%H:%M:%S",    // ISO 8601
    "%Y-%m-%d %H:%M:%S",    // Common format
    "%Y-%m-%dT%H:%M:%S%.f", // ISO 8601 with fractional seconds
    "%Y-%m-%d %H:%M:%S%.f", // With fractional seconds
];

/// Represents the detected type of a cell value
#[derive(Debug, Clone)]
enum CellValue {
    Empty,
    Integer(i64),
    Float(f64),
    Boolean(bool),
    Date(f64),     // Excel serial date
    DateTime(f64), // Excel serial datetime
    String(String),
}

/// Parse a table style string to TableStyle enum.
/// Supports: "Light1"-"Light21", "Medium1"-"Medium28", "Dark1"-"Dark11", "None"
fn parse_table_style(style: &str) -> TableStyle {
    match style {
        "None" => TableStyle::None,
        "Light1" => TableStyle::Light1,
        "Light2" => TableStyle::Light2,
        "Light3" => TableStyle::Light3,
        "Light4" => TableStyle::Light4,
        "Light5" => TableStyle::Light5,
        "Light6" => TableStyle::Light6,
        "Light7" => TableStyle::Light7,
        "Light8" => TableStyle::Light8,
        "Light9" => TableStyle::Light9,
        "Light10" => TableStyle::Light10,
        "Light11" => TableStyle::Light11,
        "Light12" => TableStyle::Light12,
        "Light13" => TableStyle::Light13,
        "Light14" => TableStyle::Light14,
        "Light15" => TableStyle::Light15,
        "Light16" => TableStyle::Light16,
        "Light17" => TableStyle::Light17,
        "Light18" => TableStyle::Light18,
        "Light19" => TableStyle::Light19,
        "Light20" => TableStyle::Light20,
        "Light21" => TableStyle::Light21,
        "Medium1" => TableStyle::Medium1,
        "Medium2" => TableStyle::Medium2,
        "Medium3" => TableStyle::Medium3,
        "Medium4" => TableStyle::Medium4,
        "Medium5" => TableStyle::Medium5,
        "Medium6" => TableStyle::Medium6,
        "Medium7" => TableStyle::Medium7,
        "Medium8" => TableStyle::Medium8,
        "Medium9" => TableStyle::Medium9,
        "Medium10" => TableStyle::Medium10,
        "Medium11" => TableStyle::Medium11,
        "Medium12" => TableStyle::Medium12,
        "Medium13" => TableStyle::Medium13,
        "Medium14" => TableStyle::Medium14,
        "Medium15" => TableStyle::Medium15,
        "Medium16" => TableStyle::Medium16,
        "Medium17" => TableStyle::Medium17,
        "Medium18" => TableStyle::Medium18,
        "Medium19" => TableStyle::Medium19,
        "Medium20" => TableStyle::Medium20,
        "Medium21" => TableStyle::Medium21,
        "Medium22" => TableStyle::Medium22,
        "Medium23" => TableStyle::Medium23,
        "Medium24" => TableStyle::Medium24,
        "Medium25" => TableStyle::Medium25,
        "Medium26" => TableStyle::Medium26,
        "Medium27" => TableStyle::Medium27,
        "Medium28" => TableStyle::Medium28,
        "Dark1" => TableStyle::Dark1,
        "Dark2" => TableStyle::Dark2,
        "Dark3" => TableStyle::Dark3,
        "Dark4" => TableStyle::Dark4,
        "Dark5" => TableStyle::Dark5,
        "Dark6" => TableStyle::Dark6,
        "Dark7" => TableStyle::Dark7,
        "Dark8" => TableStyle::Dark8,
        "Dark9" => TableStyle::Dark9,
        "Dark10" => TableStyle::Dark10,
        "Dark11" => TableStyle::Dark11,
        _ => TableStyle::Medium9, // Default Excel table style
    }
}

/// Parse a string value and detect its type
fn parse_value(value: &str) -> CellValue {
    let trimmed = value.trim();

    if trimmed.is_empty() {
        return CellValue::Empty;
    }

    // Try integer
    if let Ok(int_val) = trimmed.parse::<i64>() {
        return CellValue::Integer(int_val);
    }

    // Try float
    if let Ok(float_val) = trimmed.parse::<f64>() {
        if float_val.is_nan() || float_val.is_infinite() {
            return CellValue::Empty;
        }
        return CellValue::Float(float_val);
    }

    // Try boolean
    if trimmed.eq_ignore_ascii_case("true") {
        return CellValue::Boolean(true);
    }
    if trimmed.eq_ignore_ascii_case("false") {
        return CellValue::Boolean(false);
    }

    // Try datetime (before date, as datetime patterns are more specific)
    for pattern in DATETIME_PATTERNS {
        if let Ok(dt) = chrono::NaiveDateTime::parse_from_str(trimmed, pattern) {
            let excel_date = naive_datetime_to_excel(dt);
            return CellValue::DateTime(excel_date);
        }
    }

    // Try date
    for pattern in DATE_PATTERNS {
        if let Ok(date) = chrono::NaiveDate::parse_from_str(trimmed, pattern) {
            let excel_date = naive_date_to_excel(date);
            return CellValue::Date(excel_date);
        }
    }

    // Default to string
    CellValue::String(trimmed.to_string())
}

/// Convert NaiveDate to Excel serial date number
fn naive_date_to_excel(date: chrono::NaiveDate) -> f64 {
    // Excel epoch is December 30, 1899 (accounting for the 1900 leap year bug)
    let excel_epoch = chrono::NaiveDate::from_ymd_opt(1899, 12, 30).unwrap();
    let duration = date.signed_duration_since(excel_epoch);
    duration.num_days() as f64
}

/// Convert NaiveDateTime to Excel serial datetime number
fn naive_datetime_to_excel(dt: chrono::NaiveDateTime) -> f64 {
    let date_part = naive_date_to_excel(dt.date());
    let time = dt.time();
    let time_fraction = (time.num_seconds_from_midnight() as f64) / 86400.0;
    date_part + time_fraction
}

/// Write a cell value to the worksheet with appropriate formatting
fn write_cell(
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
            worksheet.write_number(row, col, v as f64)?;
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
///
/// # Returns
/// * `Ok((rows, cols))` - Number of rows and columns written
/// * `Err(message)` - Error description if conversion fails
pub fn convert_csv_to_xlsx(
    input_path: &str,
    output_path: &str,
    sheet_name: &str,
) -> Result<(u32, u16), String> {
    // Open CSV file
    let file = File::open(input_path).map_err(|e| format!("Failed to open input file: {}", e))?;
    let reader = BufReader::with_capacity(1024 * 1024, file);
    let mut csv_reader = ReaderBuilder::new()
        .has_headers(false)
        .flexible(true)
        .from_reader(reader);

    // Create workbook and worksheet
    let mut workbook = Workbook::new();
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
        let num_cols = record.len() as u16;
        if num_cols > col_count {
            col_count = num_cols;
        }

        for (col_idx, value) in record.iter().enumerate() {
            let cell_value = parse_value(value);
            write_cell(
                worksheet,
                row_count,
                col_idx as u16,
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

    let row_count = records.len() as u32;
    let col_count = records.iter().map(|r| r.len()).max().unwrap_or(0) as u16;

    // Parse all values in parallel
    let parsed_rows: Vec<Vec<CellValue>> = records
        .par_iter()
        .map(|row| row.iter().map(|value| parse_value(value)).collect())
        .collect();

    // Create workbook and worksheet
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .set_name(sheet_name)
        .map_err(|e| format!("Failed to set sheet name: {}", e))?;

    // Create formats for dates and datetimes
    let date_format = Format::new().set_num_format("yyyy-mm-dd");
    let datetime_format = Format::new().set_num_format("yyyy-mm-dd hh:mm:ss");

    // Write parsed values sequentially
    for (row_idx, row) in parsed_rows.into_iter().enumerate() {
        for (col_idx, cell_value) in row.into_iter().enumerate() {
            write_cell(
                worksheet,
                row_idx as u32,
                col_idx as u16,
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

/// Write a Python value to the worksheet, detecting type automatically
fn write_py_value(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    value: &Bound<'_, PyAny>,
    date_format: &Format,
    datetime_format: &Format,
) -> Result<(), String> {
    // Check for None first
    if value.is_none() {
        worksheet
            .write_string(row, col, "")
            .map_err(|e| e.to_string())?;
        return Ok(());
    }

    // Check for pandas NA/NaT
    let type_name = value
        .get_type()
        .name()
        .map_err(|e| e.to_string())?
        .to_string();
    if type_name == "NAType" || type_name == "NaTType" {
        worksheet
            .write_string(row, col, "")
            .map_err(|e| e.to_string())?;
        return Ok(());
    }

    // Try boolean first (before int, since bool is subclass of int in Python)
    if let Ok(b) = value.downcast::<PyBool>() {
        worksheet
            .write_boolean(row, col, b.is_true())
            .map_err(|e| e.to_string())?;
        return Ok(());
    }

    // Try datetime (before date, since datetime is subclass of date)
    // Check by type name since PyDateTime is not available in abi3 mode
    if type_name == "datetime" || type_name == "Timestamp" {
        // pandas Timestamp or datetime.datetime
        let year: i32 = value
            .getattr("year")
            .ok()
            .and_then(|v| v.extract().ok())
            .unwrap_or(1900);
        let month: u32 = value
            .getattr("month")
            .ok()
            .and_then(|v| v.extract().ok())
            .unwrap_or(1);
        let day: u32 = value
            .getattr("day")
            .ok()
            .and_then(|v| v.extract().ok())
            .unwrap_or(1);
        let hour: u32 = value
            .getattr("hour")
            .ok()
            .and_then(|v| v.extract().ok())
            .unwrap_or(0);
        let minute: u32 = value
            .getattr("minute")
            .ok()
            .and_then(|v| v.extract().ok())
            .unwrap_or(0);
        let second: u32 = value
            .getattr("second")
            .ok()
            .and_then(|v| v.extract().ok())
            .unwrap_or(0);

        if let Some(date) = chrono::NaiveDate::from_ymd_opt(year, month, day) {
            if let Some(time) = chrono::NaiveTime::from_hms_opt(hour, minute, second) {
                let dt = chrono::NaiveDateTime::new(date, time);
                let excel_dt = naive_datetime_to_excel(dt);
                worksheet
                    .write_number_with_format(row, col, excel_dt, datetime_format)
                    .map_err(|e| e.to_string())?;
                return Ok(());
            }
        }
    }

    // Try date
    if type_name == "date" {
        let year: i32 = value
            .getattr("year")
            .ok()
            .and_then(|v| v.extract().ok())
            .unwrap_or(1900);
        let month: u32 = value
            .getattr("month")
            .ok()
            .and_then(|v| v.extract().ok())
            .unwrap_or(1);
        let day: u32 = value
            .getattr("day")
            .ok()
            .and_then(|v| v.extract().ok())
            .unwrap_or(1);

        if let Some(date) = chrono::NaiveDate::from_ymd_opt(year, month, day) {
            let excel_date = naive_date_to_excel(date);
            worksheet
                .write_number_with_format(row, col, excel_date, date_format)
                .map_err(|e| e.to_string())?;
            return Ok(());
        }
    }

    // Try integer
    if let Ok(i) = value.downcast::<PyInt>() {
        if let Ok(val) = i.extract::<i64>() {
            worksheet
                .write_number(row, col, val as f64)
                .map_err(|e| e.to_string())?;
            return Ok(());
        }
    }

    // Try float
    if let Ok(f) = value.downcast::<PyFloat>() {
        if let Ok(val) = f.extract::<f64>() {
            if val.is_nan() || val.is_infinite() {
                worksheet
                    .write_string(row, col, "")
                    .map_err(|e| e.to_string())?;
            } else {
                worksheet
                    .write_number(row, col, val)
                    .map_err(|e| e.to_string())?;
            }
            return Ok(());
        }
    }

    // Try to extract as f64 (covers numpy types)
    if let Ok(val) = value.extract::<f64>() {
        if val.is_nan() || val.is_infinite() {
            worksheet
                .write_string(row, col, "")
                .map_err(|e| e.to_string())?;
        } else {
            worksheet
                .write_number(row, col, val)
                .map_err(|e| e.to_string())?;
        }
        return Ok(());
    }

    // Try to extract as i64 (covers numpy int types)
    if let Ok(val) = value.extract::<i64>() {
        worksheet
            .write_number(row, col, val as f64)
            .map_err(|e| e.to_string())?;
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
    if let Ok(s) = value.downcast::<PyString>() {
        worksheet
            .write_string(row, col, s.to_string())
            .map_err(|e| e.to_string())?;
        return Ok(());
    }

    // Fallback: convert to string
    let s = value.str().map_err(|e| e.to_string())?.to_string();
    worksheet
        .write_string(row, col, &s)
        .map_err(|e| e.to_string())?;

    Ok(())
}

/// Convert a DataFrame (pandas or polars) to XLSX format
#[allow(clippy::too_many_arguments)]
fn convert_dataframe_to_xlsx(
    _py: Python<'_>,
    df: &Bound<'_, PyAny>,
    output_path: &str,
    sheet_name: &str,
    include_header: bool,
    autofit: bool,
    table_style: Option<&str>,
    freeze_panes: bool,
) -> Result<(u32, u16), String> {
    // Create workbook and worksheet
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .set_name(sheet_name)
        .map_err(|e| format!("Failed to set sheet name: {}", e))?;

    // Create formats
    let date_format = Format::new().set_num_format("yyyy-mm-dd");
    let datetime_format = Format::new().set_num_format("yyyy-mm-dd hh:mm:ss");

    let mut row_idx: u32 = 0;

    // Get column names - check polars first since it also has .columns
    let columns: Vec<String> =
        if df.hasattr("schema").unwrap_or(false) && !df.hasattr("iloc").unwrap_or(false) {
            // polars DataFrame (has schema but no iloc)
            let cols = df.getattr("columns").map_err(|e| e.to_string())?;
            cols.extract().map_err(|e| e.to_string())?
        } else if df.hasattr("columns").unwrap_or(false) {
            // pandas DataFrame
            let cols = df.getattr("columns").map_err(|e| e.to_string())?;
            let col_list = cols.call_method0("tolist").map_err(|e| e.to_string())?;
            col_list.extract().map_err(|e| e.to_string())?
        } else {
            return Err("Unsupported DataFrame type".to_string());
        };

    let col_count = columns.len() as u16;

    // Write header if requested (and not using table, since table handles headers)
    if include_header && table_style.is_none() {
        for (col_idx, col_name) in columns.iter().enumerate() {
            worksheet
                .write_string(row_idx, col_idx as u16, col_name)
                .map_err(|e| e.to_string())?;
        }
        row_idx += 1;
    }

    // If using table with header, write header in row 0
    let data_start_row = if table_style.is_some() && include_header {
        for (col_idx, col_name) in columns.iter().enumerate() {
            worksheet
                .write_string(0, col_idx as u16, col_name)
                .map_err(|e| e.to_string())?;
        }
        row_idx = 1;
        0u32
    } else {
        row_idx.saturating_sub(1)
    };

    // Get row count
    let row_count: usize = if df.hasattr("shape").unwrap_or(false) {
        let shape = df.getattr("shape").map_err(|e| e.to_string())?;
        let shape_tuple: (usize, usize) = shape.extract().map_err(|e| e.to_string())?;
        shape_tuple.0
    } else {
        df.call_method0("__len__")
            .map_err(|e| e.to_string())?
            .extract()
            .map_err(|e| e.to_string())?
    };

    // Check if it's a polars DataFrame
    let is_polars = df.hasattr("schema").unwrap_or(false) && !df.hasattr("iloc").unwrap_or(false);

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
                write_py_value(
                    worksheet,
                    row_idx,
                    col_idx as u16,
                    value,
                    &date_format,
                    &datetime_format,
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

                write_py_value(
                    worksheet,
                    row_idx,
                    col_idx as u16,
                    &value,
                    &date_format,
                    &datetime_format,
                )?;
            }
            row_idx += 1;
        }
    }

    // Add Excel Table if requested
    if let Some(style_name) = table_style {
        let style = parse_table_style(style_name);
        let table = Table::new().set_style(style);

        let last_row = row_idx.saturating_sub(1);
        let last_col = col_count.saturating_sub(1);

        if last_row >= data_start_row {
            worksheet
                .add_table(data_start_row, 0, last_row, last_col, &table)
                .map_err(|e| format!("Failed to add table: {}", e))?;
        }
    }

    // Freeze panes (freeze header row)
    if freeze_panes && include_header {
        worksheet
            .set_freeze_panes(1, 0)
            .map_err(|e| format!("Failed to freeze panes: {}", e))?;
    }

    // Autofit columns
    if autofit {
        worksheet.autofit();
    }

    // Save workbook
    workbook
        .save(output_path)
        .map_err(|e| format!("Failed to save workbook: {}", e))?;

    Ok((row_idx, col_count))
}

// ============================================================================
// Python bindings
// ============================================================================

/// Convert a CSV file to XLSX format with automatic type detection.
///
/// This function reads a CSV file and writes it to an Excel XLSX file,
/// automatically detecting and converting data types:
/// - Numbers (integers and floats) become Excel numbers
/// - "true"/"false" become Excel booleans
/// - Dates (YYYY-MM-DD, etc.) become Excel dates with formatting
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
///     >>> # For large files, use parallel processing:
///     >>> rows, cols = xlsxturbo.csv_to_xlsx("big.csv", "out.xlsx", parallel=True)
#[pyfunction]
#[pyo3(signature = (input_path, output_path, sheet_name = "Sheet1", parallel = false))]
fn csv_to_xlsx(
    input_path: &str,
    output_path: &str,
    sheet_name: &str,
    parallel: bool,
) -> PyResult<(u32, u16)> {
    let result = if parallel {
        convert_csv_to_xlsx_parallel(input_path, output_path, sheet_name)
    } else {
        convert_csv_to_xlsx(input_path, output_path, sheet_name)
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
#[pyfunction]
#[pyo3(signature = (df, output_path, sheet_name = "Sheet1", header = true, autofit = false, table_style = None, freeze_panes = false))]
#[allow(clippy::too_many_arguments)]
fn df_to_xlsx(
    py: Python<'_>,
    df: &Bound<'_, PyAny>,
    output_path: &str,
    sheet_name: &str,
    header: bool,
    autofit: bool,
    table_style: Option<&str>,
    freeze_panes: bool,
) -> PyResult<(u32, u16)> {
    convert_dataframe_to_xlsx(
        py,
        df,
        output_path,
        sheet_name,
        header,
        autofit,
        table_style,
        freeze_panes,
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
///     sheets: List of (DataFrame, sheet_name) tuples
///     output_path: Path for the output XLSX file
///     header: Include column names as header row (default: True)
///     autofit: Automatically adjust column widths to fit content (default: False)
///     table_style: Apply Excel table formatting with this style name (default: None).
///                  Styles: "Light1"-"Light21", "Medium1"-"Medium28", "Dark1"-"Dark11", "None"
///                  Tables include autofilter dropdowns and banded rows.
///     freeze_panes: Freeze the header row for easier scrolling (default: False)
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
#[pyfunction]
#[pyo3(signature = (sheets, output_path, header = true, autofit = false, table_style = None, freeze_panes = false))]
fn dfs_to_xlsx(
    _py: Python<'_>,
    sheets: Vec<(Bound<'_, PyAny>, String)>,
    output_path: &str,
    header: bool,
    autofit: bool,
    table_style: Option<&str>,
    freeze_panes: bool,
) -> PyResult<Vec<(u32, u16)>> {
    let mut workbook = Workbook::new();
    let mut stats = Vec::new();

    // Create formats
    let date_format = Format::new().set_num_format("yyyy-mm-dd");
    let datetime_format = Format::new().set_num_format("yyyy-mm-dd hh:mm:ss");

    for (df, sheet_name) in sheets {
        let worksheet = workbook.add_worksheet();
        worksheet.set_name(&sheet_name).map_err(|e| {
            pyo3::exceptions::PyValueError::new_err(format!(
                "Failed to set sheet name '{}': {}",
                sheet_name, e
            ))
        })?;

        let mut row_idx: u32 = 0;

        // Get column names - check polars first
        let columns: Vec<String> =
            if df.hasattr("schema").unwrap_or(false) && !df.hasattr("iloc").unwrap_or(false) {
                let cols = df
                    .getattr("columns")
                    .map_err(|e| pyo3::exceptions::PyValueError::new_err(e.to_string()))?;
                cols.extract()
                    .map_err(|e| pyo3::exceptions::PyValueError::new_err(e.to_string()))?
            } else if df.hasattr("columns").unwrap_or(false) {
                let cols = df
                    .getattr("columns")
                    .map_err(|e| pyo3::exceptions::PyValueError::new_err(e.to_string()))?;
                let col_list = cols
                    .call_method0("tolist")
                    .map_err(|e| pyo3::exceptions::PyValueError::new_err(e.to_string()))?;
                col_list
                    .extract()
                    .map_err(|e| pyo3::exceptions::PyValueError::new_err(e.to_string()))?
            } else {
                return Err(pyo3::exceptions::PyValueError::new_err(
                    "Unsupported DataFrame type",
                ));
            };

        let col_count = columns.len() as u16;

        // Write header if requested
        if header {
            for (col_idx, col_name) in columns.iter().enumerate() {
                worksheet
                    .write_string(row_idx, col_idx as u16, col_name)
                    .map_err(|e| pyo3::exceptions::PyValueError::new_err(e.to_string()))?;
            }
            row_idx += 1;
        }

        // Get row count and check if polars
        let row_count: usize = if df.hasattr("shape").unwrap_or(false) {
            let shape = df
                .getattr("shape")
                .map_err(|e| pyo3::exceptions::PyValueError::new_err(e.to_string()))?;
            let shape_tuple: (usize, usize) = shape
                .extract()
                .map_err(|e| pyo3::exceptions::PyValueError::new_err(e.to_string()))?;
            shape_tuple.0
        } else {
            df.call_method0("__len__")
                .map_err(|e| pyo3::exceptions::PyValueError::new_err(e.to_string()))?
                .extract()
                .map_err(|e| pyo3::exceptions::PyValueError::new_err(e.to_string()))?
        };

        let is_polars =
            df.hasattr("schema").unwrap_or(false) && !df.hasattr("iloc").unwrap_or(false);

        // Write data rows
        if is_polars {
            let rows = df
                .call_method0("iter_rows")
                .map_err(|e| pyo3::exceptions::PyValueError::new_err(e.to_string()))?;
            let iter = rows
                .try_iter()
                .map_err(|e| pyo3::exceptions::PyValueError::new_err(e.to_string()))?;
            for row_result in iter {
                let row = row_result
                    .map_err(|e| pyo3::exceptions::PyValueError::new_err(e.to_string()))?;
                let row_iter = row
                    .try_iter()
                    .map_err(|e| pyo3::exceptions::PyValueError::new_err(e.to_string()))?;
                let row_tuple: Vec<Bound<'_, PyAny>> = row_iter
                    .collect::<Result<Vec<_>, _>>()
                    .map_err(|e: PyErr| pyo3::exceptions::PyValueError::new_err(e.to_string()))?;

                for (col_idx, value) in row_tuple.iter().enumerate() {
                    write_py_value(
                        worksheet,
                        row_idx,
                        col_idx as u16,
                        value,
                        &date_format,
                        &datetime_format,
                    )
                    .map_err(pyo3::exceptions::PyValueError::new_err)?;
                }
                row_idx += 1;
            }
        } else {
            let values = df
                .getattr("values")
                .map_err(|e| pyo3::exceptions::PyValueError::new_err(e.to_string()))?;
            for i in 0..row_count {
                let row = values.get_item(i).map_err(|e| {
                    pyo3::exceptions::PyValueError::new_err(format!(
                        "Failed to get row {}: {}",
                        i, e
                    ))
                })?;

                for col_idx in 0..columns.len() {
                    let value = row.get_item(col_idx).map_err(|e| {
                        pyo3::exceptions::PyValueError::new_err(format!(
                            "Failed to get value at ({}, {}): {}",
                            i, col_idx, e
                        ))
                    })?;

                    write_py_value(
                        worksheet,
                        row_idx,
                        col_idx as u16,
                        &value,
                        &date_format,
                        &datetime_format,
                    )
                    .map_err(pyo3::exceptions::PyValueError::new_err)?;
                }
                row_idx += 1;
            }
        }

        // Add Excel Table if requested
        if let Some(style_name) = table_style {
            let style = parse_table_style(style_name);
            let table = Table::new().set_style(style);

            let data_start_row = 0u32;
            let last_row = row_idx.saturating_sub(1);
            let last_col = col_count.saturating_sub(1);

            if last_row >= data_start_row {
                worksheet
                    .add_table(data_start_row, 0, last_row, last_col, &table)
                    .map_err(|e| {
                        pyo3::exceptions::PyValueError::new_err(format!(
                            "Failed to add table: {}",
                            e
                        ))
                    })?;
            }
        }

        // Freeze panes (freeze header row)
        if freeze_panes && header {
            worksheet.set_freeze_panes(1, 0).map_err(|e| {
                pyo3::exceptions::PyValueError::new_err(format!("Failed to freeze panes: {}", e))
            })?;
        }

        // Autofit columns
        if autofit {
            worksheet.autofit();
        }

        stats.push((row_idx, col_count));
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
/// Up to 25x faster than pure Python solutions.
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_integer() {
        assert!(matches!(parse_value("123"), CellValue::Integer(123)));
        assert!(matches!(parse_value("-456"), CellValue::Integer(-456)));
    }

    #[test]
    fn test_parse_float() {
        if let CellValue::Float(v) = parse_value("3.14") {
            assert!((v - 3.14).abs() < 0.001);
        } else {
            panic!("Expected float");
        }
    }

    #[test]
    fn test_parse_boolean() {
        assert!(matches!(parse_value("true"), CellValue::Boolean(true)));
        assert!(matches!(parse_value("TRUE"), CellValue::Boolean(true)));
        assert!(matches!(parse_value("false"), CellValue::Boolean(false)));
        assert!(matches!(parse_value("False"), CellValue::Boolean(false)));
    }

    #[test]
    fn test_parse_empty() {
        assert!(matches!(parse_value(""), CellValue::Empty));
        assert!(matches!(parse_value("   "), CellValue::Empty));
        assert!(matches!(parse_value("NaN"), CellValue::Empty));
    }

    #[test]
    fn test_parse_date() {
        assert!(matches!(parse_value("2024-01-15"), CellValue::Date(_)));
        assert!(matches!(parse_value("2024/01/15"), CellValue::Date(_)));
    }

    #[test]
    fn test_parse_datetime() {
        assert!(matches!(
            parse_value("2024-01-15T10:30:00"),
            CellValue::DateTime(_)
        ));
        assert!(matches!(
            parse_value("2024-01-15 10:30:00"),
            CellValue::DateTime(_)
        ));
    }

    #[test]
    fn test_parse_string() {
        assert!(matches!(parse_value("hello"), CellValue::String(_)));
    }
}
