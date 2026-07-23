//! Shared worksheet cell writers.

use crate::parse::{naive_date_to_excel, naive_datetime_to_excel};
use crate::types::CellValue;
use pyo3::prelude::*;
use pyo3::types::{PyBool, PyDate, PyDateTime, PyFloat, PyInt, PyString};
use rust_xlsxwriter::{Format, Worksheet, XlsxError};

/// Maximum safe integer for lossless f64 representation (2^53).
/// Integers beyond this range lose precision when cast to f64.
const MAX_SAFE_INT: i64 = 1 << 53;
const MAX_SAFE_INT_U64: u64 = MAX_SAFE_INT as u64;

/// Whether an integer of the given magnitude fits in an f64 without precision
/// loss. Single source of the overflow-to-string policy shared by every integer
/// write path (`write_int`, `write_uint`, and the `CellValue::Integer` arm).
fn int_fits_f64(magnitude: u64) -> bool {
    magnitude <= MAX_SAFE_INT_U64
}

/// Excel number format strings (shared with apply::apply_cells)
pub(crate) const DATE_NUM_FORMAT: &str = "yyyy-mm-dd";
pub(crate) const DATETIME_NUM_FORMAT: &str = "yyyy-mm-dd hh:mm:ss";

/// Write a string to a cell, applying column format if provided.
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
    .map_err(|e| format!("Failed to write text at row {}, col {}: {}", row, col, e))
}

/// Write a number to a cell, applying column format if provided.
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
    .map_err(|e| format!("Failed to write number at row {}, col {}: {}", row, col, e))
}

/// Write a boolean to a cell, applying column format if provided.
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
    .map_err(|e| format!("Failed to write boolean at row {}, col {}: {}", row, col, e))
}

/// Write an integer, falling back to string for values beyond f64 precision.
fn write_int(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    val: i64,
    fmt: Option<&Format>,
) -> Result<(), String> {
    if int_fits_f64(val.unsigned_abs()) {
        write_num(worksheet, row, col, val as f64, fmt)
    } else {
        write_str(worksheet, row, col, val.to_string(), fmt)
    }
}

fn write_uint(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    val: u64,
    fmt: Option<&Format>,
) -> Result<(), String> {
    if int_fits_f64(val) {
        write_num(worksheet, row, col, val as f64, fmt)
    } else {
        write_str(worksheet, row, col, val.to_string(), fmt)
    }
}

/// Write a float, treating NaN/Inf as empty.
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

/// Write a cell value to the worksheet with appropriate formatting.
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
            if int_fits_f64(v.unsigned_abs()) {
                worksheet.write_number(row, col, v as f64)?;
            } else {
                worksheet.write_string(row, col, v.to_string())?;
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

/// Write a Python value to the worksheet with optional column format.
pub(crate) fn write_py_value_with_format(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    value: &Bound<'_, PyAny>,
    date_format: &Format,
    datetime_format: &Format,
    column_format: Option<&Format>,
) -> Result<(), String> {
    // Check for None first.
    if value.is_none() {
        return write_str(worksheet, row, col, "", column_format);
    }

    // Boolean first (before int, since bool is subclass of int in Python).
    if let Ok(b) = value.cast::<PyBool>() {
        return write_bool(worksheet, row, col, b.is_true(), column_format);
    }

    if let Ok(i) = value.cast::<PyInt>() {
        if let Ok(val) = i.extract::<i64>() {
            return write_int(worksheet, row, col, val, column_format);
        }
        let s = i
            .str()
            .map_err(|e| format!("Failed to convert Python int to string: {}", e))?
            .to_string();
        return write_str(worksheet, row, col, s, column_format);
    }

    if let Ok(f) = value.cast::<PyFloat>() {
        if let Ok(val) = f.extract::<f64>() {
            return write_float(worksheet, row, col, val, column_format);
        }
    }

    if let Ok(s) = value.cast::<PyString>() {
        return write_str(worksheet, row, col, s.to_string(), column_format);
    }

    let type_name = value
        .get_type()
        .name()
        .map_err(|e| format!("Failed to get type name: {}", e))?
        .to_string();

    if type_name == "NAType" || type_name == "NaTType" {
        return write_str(worksheet, row, col, "", column_format);
    }

    // numpy scalar bool. Checked after the `PyBool` cast above (which only
    // matches real Python bool) and before the generic int/float fallbacks
    // below, since `np.bool_`/`np.bool` satisfies `__index__` and would
    // otherwise be silently written as the number 0/1. The type name is
    // "bool_" on numpy 1.x and "bool" on numpy 2.x.
    if type_name == "bool_" || type_name == "bool" {
        if let Ok(val) = value.extract::<bool>() {
            return write_bool(worksheet, row, col, val, column_format);
        }
    }

    if type_name == "datetime64" {
        let us_since_epoch: i64 = value
            .call_method1("astype", ("datetime64[us]",))
            .and_then(|v| v.call_method1("astype", ("int64",)))
            .and_then(|v| v.call_method0("item"))
            .and_then(|v| v.extract())
            .map_err(|e| format!("Failed to convert numpy datetime64 scalar: {}", e))?;
        if us_since_epoch == i64::MIN {
            return write_str(worksheet, row, col, "", column_format);
        }

        let seconds = us_since_epoch.div_euclid(1_000_000);
        let nanosecond = (us_since_epoch.rem_euclid(1_000_000) as u32)
            .checked_mul(1_000)
            .ok_or("numpy datetime64 microsecond remainder exceeds nanosecond range")?;
        let dt = chrono::DateTime::from_timestamp(seconds, nanosecond)
            .ok_or_else(|| format!("Invalid numpy datetime64 timestamp: {}", us_since_epoch))?
            .naive_utc();
        let excel_dt = naive_datetime_to_excel(dt);
        // Dates before 1900-03-01 (serial 61) can't be represented correctly
        // due to Excel's 1900 leap-year bug; fall back to string.
        if excel_dt < 61.0 {
            let s = value
                .str()
                .map_err(|e| format!("Failed to convert numpy datetime64 to string: {}", e))?
                .to_string();
            return write_str(worksheet, row, col, s, column_format);
        }
        let fmt = column_format.unwrap_or(datetime_format);
        return write_num(worksheet, row, col, excel_dt, Some(fmt));
    }

    // Datetime before date, since datetime is subclass of date. Use a typed
    // isinstance-style check (via `cast::<PyDateTime>`) so subclasses such as
    // pendulum.DateTime or freezegun's FakeDatetime are caught here too,
    // instead of falling through to the generic str() path below. The
    // type-name comparisons are kept as a defensive fast path/fallback
    // (pandas Timestamp is itself a datetime subclass, so the typed check
    // already covers it, but the string check costs nothing extra here).
    if value.cast::<PyDateTime>().is_ok() || type_name == "datetime" || type_name == "Timestamp" {
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
        let microsecond: u32 = value
            .getattr("microsecond")
            .and_then(|v| v.extract())
            .unwrap_or(0);
        let nanosecond_remainder: u32 = value
            .getattr("nanosecond")
            .and_then(|v| v.extract())
            .unwrap_or(0);
        let nanosecond = microsecond
            .checked_mul(1_000)
            .and_then(|v| v.checked_add(nanosecond_remainder))
            .ok_or("Datetime fractional seconds exceed nanosecond range")?;

        let date = chrono::NaiveDate::from_ymd_opt(year, month, day).ok_or_else(|| {
            format!(
                "Invalid datetime date: year={}, month={}, day={}",
                year, month, day
            )
        })?;
        let time = chrono::NaiveTime::from_hms_nano_opt(hour, minute, second, nanosecond)
            .ok_or_else(|| {
                format!(
                    "Invalid datetime time: hour={}, minute={}, second={}, nanosecond={}",
                    hour, minute, second, nanosecond
                )
            })?;
        let dt = chrono::NaiveDateTime::new(date, time);
        let excel_dt = naive_datetime_to_excel(dt);
        // Dates before 1900-03-01 (serial 61) can't be represented correctly
        // due to Excel's 1900 leap-year bug; fall back to string.
        if excel_dt < 61.0 {
            let s = value
                .str()
                .map_err(|e| format!("Failed to convert datetime to string: {}", e))?
                .to_string();
            return write_str(worksheet, row, col, s, column_format);
        }
        let fmt = column_format.unwrap_or(datetime_format);
        return write_num(worksheet, row, col, excel_dt, Some(fmt));
    }

    // Typed check first, same rationale as the datetime branch above; this
    // branch is only reached for plain dates because the datetime branch
    // (checked first) already returns for any datetime/subclass instance,
    // even though `PyDate` would also match those (datetime is-a date).
    if value.cast::<PyDate>().is_ok() || type_name == "date" {
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
        // Dates before 1900-03-01 (serial 61) can't be represented correctly
        // due to Excel's 1900 leap-year bug; fall back to string.
        if excel_date < 61.0 {
            let s = value
                .str()
                .map_err(|e| format!("Failed to convert date to string: {}", e))?
                .to_string();
            return write_str(worksheet, row, col, s, column_format);
        }
        let fmt = column_format.unwrap_or(date_format);
        return write_num(worksheet, row, col, excel_date, Some(fmt));
    }

    // numpy scalar int (before f64 to avoid precision loss).
    if let Ok(val) = value.extract::<i64>() {
        return write_int(worksheet, row, col, val, column_format);
    }

    if let Ok(val) = value.extract::<u64>() {
        return write_uint(worksheet, row, col, val, column_format);
    }

    // numpy scalar float.
    if let Ok(val) = value.extract::<f64>() {
        return write_float(worksheet, row, col, val, column_format);
    }

    let s = value
        .str()
        .map_err(|e| format!("Failed to convert value to string: {}", e))?
        .to_string();
    write_str(worksheet, row, col, s, column_format)
}
