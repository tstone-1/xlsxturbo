//! Parsing and utility functions

use crate::types::{CellValue, DateOrder, DATETIME_PATTERNS};
use chrono::Timelike;
use indexmap::IndexMap;
use pyo3::prelude::*;
use pyo3::Py;
use rust_xlsxwriter::{ConditionalFormatIconType, Format, TableStyle};
use std::collections::HashMap;

/// Parse a table style string into a `TableStyle` enum value
pub(crate) fn parse_table_style(style: &str) -> Result<TableStyle, String> {
    match style {
        "None" => Ok(TableStyle::None),
        "Light1" => Ok(TableStyle::Light1),
        "Light2" => Ok(TableStyle::Light2),
        "Light3" => Ok(TableStyle::Light3),
        "Light4" => Ok(TableStyle::Light4),
        "Light5" => Ok(TableStyle::Light5),
        "Light6" => Ok(TableStyle::Light6),
        "Light7" => Ok(TableStyle::Light7),
        "Light8" => Ok(TableStyle::Light8),
        "Light9" => Ok(TableStyle::Light9),
        "Light10" => Ok(TableStyle::Light10),
        "Light11" => Ok(TableStyle::Light11),
        "Light12" => Ok(TableStyle::Light12),
        "Light13" => Ok(TableStyle::Light13),
        "Light14" => Ok(TableStyle::Light14),
        "Light15" => Ok(TableStyle::Light15),
        "Light16" => Ok(TableStyle::Light16),
        "Light17" => Ok(TableStyle::Light17),
        "Light18" => Ok(TableStyle::Light18),
        "Light19" => Ok(TableStyle::Light19),
        "Light20" => Ok(TableStyle::Light20),
        "Light21" => Ok(TableStyle::Light21),
        "Medium1" => Ok(TableStyle::Medium1),
        "Medium2" => Ok(TableStyle::Medium2),
        "Medium3" => Ok(TableStyle::Medium3),
        "Medium4" => Ok(TableStyle::Medium4),
        "Medium5" => Ok(TableStyle::Medium5),
        "Medium6" => Ok(TableStyle::Medium6),
        "Medium7" => Ok(TableStyle::Medium7),
        "Medium8" => Ok(TableStyle::Medium8),
        "Medium9" => Ok(TableStyle::Medium9),
        "Medium10" => Ok(TableStyle::Medium10),
        "Medium11" => Ok(TableStyle::Medium11),
        "Medium12" => Ok(TableStyle::Medium12),
        "Medium13" => Ok(TableStyle::Medium13),
        "Medium14" => Ok(TableStyle::Medium14),
        "Medium15" => Ok(TableStyle::Medium15),
        "Medium16" => Ok(TableStyle::Medium16),
        "Medium17" => Ok(TableStyle::Medium17),
        "Medium18" => Ok(TableStyle::Medium18),
        "Medium19" => Ok(TableStyle::Medium19),
        "Medium20" => Ok(TableStyle::Medium20),
        "Medium21" => Ok(TableStyle::Medium21),
        "Medium22" => Ok(TableStyle::Medium22),
        "Medium23" => Ok(TableStyle::Medium23),
        "Medium24" => Ok(TableStyle::Medium24),
        "Medium25" => Ok(TableStyle::Medium25),
        "Medium26" => Ok(TableStyle::Medium26),
        "Medium27" => Ok(TableStyle::Medium27),
        "Medium28" => Ok(TableStyle::Medium28),
        "Dark1" => Ok(TableStyle::Dark1),
        "Dark2" => Ok(TableStyle::Dark2),
        "Dark3" => Ok(TableStyle::Dark3),
        "Dark4" => Ok(TableStyle::Dark4),
        "Dark5" => Ok(TableStyle::Dark5),
        "Dark6" => Ok(TableStyle::Dark6),
        "Dark7" => Ok(TableStyle::Dark7),
        "Dark8" => Ok(TableStyle::Dark8),
        "Dark9" => Ok(TableStyle::Dark9),
        "Dark10" => Ok(TableStyle::Dark10),
        "Dark11" => Ok(TableStyle::Dark11),
        _ => Err(format!(
            "Unknown table_style '{}'. Valid styles: Light1-Light21, Medium1-Medium28, Dark1-Dark11, None",
            style
        )),
    }
}

/// Parse a cell reference like "A1" into (row, col) - 0-based
pub(crate) fn parse_cell_ref(cell_ref: &str) -> Result<(u32, u16), String> {
    let cell_ref = cell_ref.trim().to_uppercase();
    if cell_ref.is_empty() {
        return Err("Empty cell reference".to_string());
    }

    // Find where letters end and numbers begin
    let col_end = cell_ref
        .chars()
        .take_while(|c| c.is_ascii_alphabetic())
        .count();
    if col_end == 0 {
        return Err(format!(
            "Invalid cell reference '{}': no column letters",
            cell_ref
        ));
    }

    let col_str = &cell_ref[..col_end];
    let row_str = &cell_ref[col_end..];

    if row_str.is_empty() {
        return Err(format!(
            "Invalid cell reference '{}': no row number",
            cell_ref
        ));
    }

    // Convert column letters to 0-based index (A=0, B=1, ..., Z=25, AA=26, etc.)
    let col: u16 = col_str
        .chars()
        .fold(0u16, |acc, c| acc * 26 + (c as u16 - 'A' as u16 + 1))
        .saturating_sub(1);

    // Parse row number (Excel rows are 1-based, so must be >= 1)
    let row_1based: u32 = row_str
        .parse::<u32>()
        .map_err(|_| format!("Invalid row number in cell reference '{}'", cell_ref))?;

    if row_1based == 0 {
        return Err(format!(
            "Invalid cell reference '{}': row number must be >= 1 (Excel rows are 1-based)",
            cell_ref
        ));
    }

    // Convert to 0-based index
    let row = row_1based - 1;

    Ok((row, col))
}

/// Parse a cell range like "A1:D1" into (first_row, first_col, last_row, last_col) - 0-based
pub(crate) fn parse_cell_range(range_str: &str) -> Result<(u32, u16, u32, u16), String> {
    let parts: Vec<&str> = range_str.split(':').collect();
    if parts.len() != 2 {
        return Err(format!(
            "Invalid cell range '{}': expected format 'A1:B2'",
            range_str
        ));
    }

    let (first_row, first_col) = parse_cell_ref(parts[0])?;
    let (last_row, last_col) = parse_cell_ref(parts[1])?;

    Ok((first_row, first_col, last_row, last_col))
}

/// Parse icon type string into `ConditionalFormatIconType`
pub(crate) fn parse_icon_type(icon_type: &str) -> Result<ConditionalFormatIconType, String> {
    match icon_type.to_lowercase().as_str() {
        "3_arrows" | "3arrows" => Ok(ConditionalFormatIconType::ThreeArrows),
        "3_arrows_gray" | "3arrowsgray" => Ok(ConditionalFormatIconType::ThreeArrowsGray),
        "3_flags" | "3flags" => Ok(ConditionalFormatIconType::ThreeFlags),
        "3_traffic_lights" | "3trafficlights" | "traffic_lights" => {
            Ok(ConditionalFormatIconType::ThreeTrafficLights)
        }
        "3_traffic_lights_rimmed" | "3trafficlightsrimmed" => {
            Ok(ConditionalFormatIconType::ThreeTrafficLightsWithRim)
        }
        "3_signs" | "3signs" => Ok(ConditionalFormatIconType::ThreeSigns),
        "3_symbols" | "3symbols" => Ok(ConditionalFormatIconType::ThreeSymbolsCircled),
        "3_symbols_uncircled" | "3symbolsuncircled" => {
            Ok(ConditionalFormatIconType::ThreeSymbols)
        }
        "4_arrows" | "4arrows" => Ok(ConditionalFormatIconType::FourArrows),
        "4_arrows_gray" | "4arrowsgray" => Ok(ConditionalFormatIconType::FourArrowsGray),
        "4_rating" | "4rating" => Ok(ConditionalFormatIconType::FourHistograms),
        "4_traffic_lights" | "4trafficlights" => {
            Ok(ConditionalFormatIconType::FourTrafficLights)
        }
        "5_arrows" | "5arrows" => Ok(ConditionalFormatIconType::FiveArrows),
        "5_arrows_gray" | "5arrowsgray" => Ok(ConditionalFormatIconType::FiveArrowsGray),
        "5_rating" | "5rating" => Ok(ConditionalFormatIconType::FiveHistograms),
        "5_quarters" | "5quarters" => Ok(ConditionalFormatIconType::FiveQuadrants),
        _ => Err(format!(
            "Unknown icon_type '{}'. Valid types: 3_arrows, 3_arrows_gray, 3_flags, 3_traffic_lights, 3_traffic_lights_rimmed, 3_signs, 3_symbols, 3_symbols_uncircled, 4_arrows, 4_arrows_gray, 4_rating, 4_traffic_lights, 5_arrows, 5_arrows_gray, 5_quarters, 5_rating",
            icon_type
        )),
    }
}

/// Sanitize a string for use as an Excel table name
pub(crate) fn sanitize_table_name(name: &str) -> String {
    let mut sanitized: String = name
        .chars()
        .map(|c| {
            if c.is_alphanumeric() || c == '_' {
                c
            } else {
                '_'
            }
        })
        .collect();

    // Must start with letter or underscore
    if sanitized.chars().next().is_none_or(|c| c.is_ascii_digit()) {
        sanitized = format!("_{}", sanitized);
    }

    // Max 255 chars
    sanitized.truncate(255);
    sanitized
}

/// Parse color string (hex #RRGGBB or named color) to u32
pub(crate) fn parse_color(color_str: &str) -> Result<u32, String> {
    let color = color_str.trim();
    if let Some(hex) = color.strip_prefix('#') {
        if hex.len() != 6 {
            return Err(format!(
                "Invalid hex color '{}': expected 6 characters after #, got {}",
                color,
                hex.len()
            ));
        }
        u32::from_str_radix(hex, 16).map_err(|_| format!("Invalid hex color: {}", color))
    } else {
        match color.to_lowercase().as_str() {
            "white" => Ok(0xFFFFFF),
            "black" => Ok(0x000000),
            "red" => Ok(0xFF0000),
            "green" => Ok(0x00FF00),
            "blue" => Ok(0x0000FF),
            "yellow" => Ok(0xFFFF00),
            "cyan" => Ok(0x00FFFF),
            "magenta" => Ok(0xFF00FF),
            "gray" | "grey" => Ok(0x808080),
            "silver" => Ok(0xC0C0C0),
            "orange" => Ok(0xFFA500),
            "purple" => Ok(0x800080),
            "navy" => Ok(0x000080),
            "teal" => Ok(0x008080),
            "maroon" => Ok(0x800000),
            _ => Err(format!("Unknown color: {}", color)),
        }
    }
}

/// Parse header format dictionary into rust_xlsxwriter Format
/// Delegates to parse_format_dict without column-specific options
pub(crate) fn parse_header_format(
    py: Python<'_>,
    fmt_dict: &HashMap<String, Py<PyAny>>,
) -> Result<Format, String> {
    parse_format_dict(py, fmt_dict, false)
}

/// Shared format parser for both header and column formats.
/// When `include_column_options` is true, also handles num_format and border.
fn parse_format_dict(
    py: Python<'_>,
    fmt_dict: &HashMap<String, Py<PyAny>>,
    include_column_options: bool,
) -> Result<Format, String> {
    let mut format = Format::new();

    if let Some(bold_obj) = fmt_dict.get("bold") {
        let bold: bool = bold_obj.bind(py).extract().unwrap_or(false);
        if bold {
            format = format.set_bold();
        }
    }

    if let Some(italic_obj) = fmt_dict.get("italic") {
        let italic: bool = italic_obj.bind(py).extract().unwrap_or(false);
        if italic {
            format = format.set_italic();
        }
    }

    if let Some(bg_obj) = fmt_dict.get("bg_color") {
        if let Ok(color_str) = bg_obj.bind(py).extract::<String>() {
            let color = parse_color(&color_str)?;
            format = format.set_background_color(color);
        }
    }

    if let Some(font_obj) = fmt_dict.get("font_color") {
        if let Ok(color_str) = font_obj.bind(py).extract::<String>() {
            let color = parse_color(&color_str)?;
            format = format.set_font_color(color);
        }
    }

    if let Some(size_obj) = fmt_dict.get("font_size") {
        if let Ok(size) = size_obj.bind(py).extract::<f64>() {
            format = format.set_font_size(size);
        }
    }

    if let Some(underline_obj) = fmt_dict.get("underline") {
        let underline: bool = underline_obj.bind(py).extract().unwrap_or(false);
        if underline {
            format = format.set_underline(rust_xlsxwriter::FormatUnderline::Single);
        }
    }

    if include_column_options {
        if let Some(num_fmt_obj) = fmt_dict.get("num_format") {
            if let Ok(num_fmt_str) = num_fmt_obj.bind(py).extract::<String>() {
                format = format.set_num_format(&num_fmt_str);
            }
        }

        if let Some(border_obj) = fmt_dict.get("border") {
            let border: bool = border_obj.bind(py).extract().unwrap_or(false);
            if border {
                format = format.set_border(rust_xlsxwriter::FormatBorder::Thin);
            }
        }
    }

    Ok(format)
}

/// Check if a column name matches a wildcard pattern.
/// Supports: "prefix*", "*suffix", "*contains*", or exact match
pub(crate) fn matches_pattern(column_name: &str, pattern: &str) -> bool {
    let starts_with_star = pattern.starts_with('*');
    let ends_with_star = pattern.ends_with('*');

    match (starts_with_star, ends_with_star) {
        (true, true) => {
            // *contains* - match substring; lone "*" matches everything
            if pattern.len() <= 2 {
                return true;
            }
            let inner = &pattern[1..pattern.len() - 1];
            column_name.contains(inner)
        }
        (true, false) => {
            // *suffix - match ending
            let suffix = &pattern[1..];
            column_name.ends_with(suffix)
        }
        (false, true) => {
            // prefix* - match beginning
            let prefix = &pattern[..pattern.len() - 1];
            column_name.starts_with(prefix)
        }
        (false, false) => {
            // Exact match
            column_name == pattern
        }
    }
}

/// Parse column format dictionary into rust_xlsxwriter Format
/// Delegates to parse_format_dict with column-specific options enabled
pub(crate) fn parse_column_format(
    py: Python<'_>,
    fmt_dict: &HashMap<String, Py<PyAny>>,
) -> Result<Format, String> {
    parse_format_dict(py, fmt_dict, true)
}

/// Build a vector of column formats, one for each column.
/// Returns None for columns with no matching pattern.
/// Uses IndexMap to preserve pattern order - first matching pattern wins.
pub(crate) fn build_column_formats(
    py: Python<'_>,
    columns: &[String],
    column_formats: &IndexMap<String, HashMap<String, Py<PyAny>>>,
) -> Result<Vec<Option<Format>>, String> {
    let mut formats = Vec::with_capacity(columns.len());

    for col_name in columns {
        // Find the first matching pattern (order preserved by IndexMap)
        let mut matched_format: Option<Format> = None;
        for (pattern, fmt_dict) in column_formats {
            if matches_pattern(col_name, pattern) {
                matched_format = Some(parse_column_format(py, fmt_dict)?);
                break;
            }
        }
        formats.push(matched_format);
    }

    Ok(formats)
}

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

    // Try date with locale-aware ordering
    for pattern in date_order.patterns() {
        if let Ok(date) = chrono::NaiveDate::parse_from_str(trimmed, pattern) {
            let excel_date = naive_date_to_excel(date);
            return CellValue::Date(excel_date);
        }
    }

    // Default to string
    CellValue::String(trimmed.to_string())
}

/// Convert NaiveDate to Excel serial date number
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
    let time_fraction = (time.num_seconds_from_midnight() as f64) / 86400.0;
    date_part + time_fraction
}
