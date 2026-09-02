use crate::types::{CellValue, DateOrder, DATETIME_PATTERNS};
use chrono::Timelike;

/// Parse a string value and detect its type
pub(crate) fn parse_value(value: &str, date_order: DateOrder) -> CellValue {
    let trimmed = value.trim();

    if trimmed.is_empty() {
        return CellValue::Empty;
    }

    // Try integer
    if let Ok(int_val) = trimmed.parse::<i64>() {
        return CellValue::Integer(int_val);
    }

    // An integer too large for `i64` must not reach the `f64` branch below.
    // `f64` parses any run of digits and silently rounds it, which contradicts
    // the guarantee the DataFrame path keeps (`write.rs::int_fits_f64`) and the
    // documentation states: integers above 2^53 are written as text. Measured
    // before this branch existed, a CSV cell `12345678901234567890` reached the
    // workbook as 12345678901234567000, and a digit string long enough to parse
    // as infinity reached it as an empty cell.
    //
    // There is no numeric arm to take here. Anything that fails `i64` and is
    // still an integer exceeds `i64::MAX`, hence 2^53, so the shared
    // overflow-to-string policy resolves to "text" for every value on this path.
    //
    // Both screens are behind a length gate so that ordinary cells -- every
    // float, every word, every date -- do not pay for them. 19 is exact rather
    // than generous: `i64::MAX` is 9223372036854775807, 19 digits, so the
    // smallest integer literal that can fail the parse above is
    // 9223372036854775808 (19 bytes). A sign or a leading zero only makes the
    // literal longer, and the negative side needs 20 bytes (`i64::MIN` is
    // -9223372036854775808, so the first failing negative is
    // -9223372036854775809). Anything shorter than 19 bytes that reaches here
    // is therefore not an integer literal at all, and both screens would
    // decline it.
    if trimmed.len() >= 19 {
        if let Ok(uint_val) = trimmed.parse::<u64>() {
            // Still inside a Rust integer type, so round-trip it: the written
            // text is canonical the same way the `i64` arm turns "+42" into 42
            // and "007" into 7.
            return CellValue::String(uint_val.to_string());
        }
        if is_ascii_integer_literal(trimmed) {
            // Beyond `u64` (or below `i64::MIN`) there is no integer type left
            // to normalise against, so the caller's digits go through as given
            // -- including any sign and any leading zeros.
            return CellValue::String(trimmed.to_string());
        }
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
            let excel_dt = naive_datetime_to_excel(dt);
            // Dates before 1900-03-01 (serial 61) can't be represented
            // correctly: Excel's serial numbering assumes a phantom
            // 1900-02-29 that never existed, so `naive_date_to_excel`'s
            // epoch-based formula is one day ahead of the real Excel
            // serial for any date in Jan/Feb 1900. Fall back to string
            // rather than write a date that renders one day late.
            if excel_dt < 61.0 {
                return CellValue::String(value.to_string());
            }
            return CellValue::DateTime(excel_dt);
        }
    }

    // Try date with locale-aware ordering
    for pattern in date_order.patterns() {
        if let Ok(date) = chrono::NaiveDate::parse_from_str(trimmed, pattern) {
            let excel_date = naive_date_to_excel(date);
            // See the comment on the datetime branch above: dates before
            // 1900-03-01 (serial 61) can't be represented correctly because
            // of Excel's 1900 leap-year bug, so fall back to string.
            if excel_date < 61.0 {
                return CellValue::String(value.to_string());
            }
            return CellValue::Date(excel_date);
        }
    }

    // Default to string. Use the original (untrimmed) value so leading/
    // trailing whitespace on genuine string cells is preserved; `trimmed`
    // is only used above for type detection.
    CellValue::String(value.to_string())
}

/// Whether `text` is an ASCII integer literal, `[+-]?[0-9]+`.
///
/// Deliberately narrower than what `f64::from_str` accepts: no exponent, no
/// decimal point, no non-ASCII digits. That keeps "1e3" and "1.0" on the float
/// branch and leaves a non-ASCII digit string such as "٣٤" falling through to
/// the string default, both exactly as before this screen existed.
fn is_ascii_integer_literal(text: &str) -> bool {
    let digits = text.strip_prefix(['+', '-']).unwrap_or(text);
    !digits.is_empty() && digits.bytes().all(|b| b.is_ascii_digit())
}

/// Convert NaiveDate to Excel serial date number.
///
/// The epoch used here (1899-12-30) only produces the correct Excel serial
/// for dates on or after 1900-03-01 (serial 61). Excel's serial numbering
/// assumes a phantom leap day, 1900-02-29, that never actually existed, so
/// for any date before 1900-03-01 this formula returns a value one day
/// ahead of what Excel would show. Callers that accept arbitrary dates
/// must reject results below serial 61 rather than write them as dates.
pub(crate) fn naive_date_to_excel(date: chrono::NaiveDate) -> f64 {
    // Excel epoch is December 30, 1899 (accounting for the 1900 leap year bug)
    // SAFETY: constant date literal, always valid
    let excel_epoch =
        chrono::NaiveDate::from_ymd_opt(1899, 12, 30).expect("Excel epoch date is always valid");
    let duration = date.signed_duration_since(excel_epoch);
    duration.num_days() as f64
}

/// Convert NaiveDateTime to Excel serial datetime number
pub(crate) fn naive_datetime_to_excel(dt: chrono::NaiveDateTime) -> f64 {
    let date_part = naive_date_to_excel(dt.date());
    let time = dt.time();
    let seconds = time.num_seconds_from_midnight() as f64;
    let fractional_seconds = time.nanosecond() as f64 / 1_000_000_000.0;
    let time_fraction = (seconds + fractional_seconds) / 86400.0;
    date_part + time_fraction
}
