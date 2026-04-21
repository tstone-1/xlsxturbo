//! Parsing and utility functions

use crate::types::{pytype_name, CellValue, DateOrder, DATETIME_PATTERNS};
use chrono::Timelike;
use indexmap::IndexMap;
use pyo3::prelude::*;
use pyo3::Py;
use rust_xlsxwriter::{
    Color, ConditionalFormatIconType, Format, FormatAlign, FormatBorder, TableStyle,
};
use std::collections::HashMap;

/// Generate a table style lookup match from a list of (string, variant) pairs.
macro_rules! table_style_match {
    ($style:expr, $( $name:literal => $variant:ident ),+ $(,)?) => {
        match $style {
            $( $name => Ok(TableStyle::$variant), )+
            _ => Err(format!(
                "Unknown table_style '{}'. Valid styles: Light1-Light21, Medium1-Medium28, Dark1-Dark11, None",
                $style
            )),
        }
    };
}

/// Parse a table style string into a `TableStyle` enum value.
/// Synced with rust_xlsxwriter 0.94 TableStyle variants.
pub(crate) fn parse_table_style(style: &str) -> Result<TableStyle, String> {
    table_style_match!(style,
        "None" => None,
        "Light1" => Light1, "Light2" => Light2, "Light3" => Light3, "Light4" => Light4,
        "Light5" => Light5, "Light6" => Light6, "Light7" => Light7, "Light8" => Light8,
        "Light9" => Light9, "Light10" => Light10, "Light11" => Light11, "Light12" => Light12,
        "Light13" => Light13, "Light14" => Light14, "Light15" => Light15, "Light16" => Light16,
        "Light17" => Light17, "Light18" => Light18, "Light19" => Light19, "Light20" => Light20,
        "Light21" => Light21,
        "Medium1" => Medium1, "Medium2" => Medium2, "Medium3" => Medium3, "Medium4" => Medium4,
        "Medium5" => Medium5, "Medium6" => Medium6, "Medium7" => Medium7, "Medium8" => Medium8,
        "Medium9" => Medium9, "Medium10" => Medium10, "Medium11" => Medium11, "Medium12" => Medium12,
        "Medium13" => Medium13, "Medium14" => Medium14, "Medium15" => Medium15, "Medium16" => Medium16,
        "Medium17" => Medium17, "Medium18" => Medium18, "Medium19" => Medium19, "Medium20" => Medium20,
        "Medium21" => Medium21, "Medium22" => Medium22, "Medium23" => Medium23, "Medium24" => Medium24,
        "Medium25" => Medium25, "Medium26" => Medium26, "Medium27" => Medium27, "Medium28" => Medium28,
        "Dark1" => Dark1, "Dark2" => Dark2, "Dark3" => Dark3, "Dark4" => Dark4,
        "Dark5" => Dark5, "Dark6" => Dark6, "Dark7" => Dark7, "Dark8" => Dark8,
        "Dark9" => Dark9, "Dark10" => Dark10, "Dark11" => Dark11,
    )
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
    // Use checked arithmetic to detect overflow on adversarial input
    let col_u32: u32 = col_str
        .chars()
        .try_fold(0u32, |acc, c| {
            acc.checked_mul(26)?.checked_add(c as u32 - 'A' as u32 + 1)
        })
        .ok_or_else(|| format!("Column '{}' is too large", col_str))?
        .saturating_sub(1);
    if col_u32 > 16383 {
        return Err(format!(
            "Column '{}' exceeds Excel's maximum column (XFD = 16384)",
            col_str
        ));
    }
    let col = col_u32 as u16;

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

/// Parse border style string into `FormatBorder` enum value.
pub(crate) fn parse_border_style(style: &str) -> Result<FormatBorder, String> {
    match style.to_lowercase().as_str() {
        "thin" => Ok(FormatBorder::Thin),
        "medium" => Ok(FormatBorder::Medium),
        "thick" => Ok(FormatBorder::Thick),
        "dashed" => Ok(FormatBorder::Dashed),
        "dotted" => Ok(FormatBorder::Dotted),
        "double" => Ok(FormatBorder::Double),
        "hair" => Ok(FormatBorder::Hair),
        "medium_dashed" | "mediumdashed" => Ok(FormatBorder::MediumDashed),
        "dash_dot" | "dashdot" => Ok(FormatBorder::DashDot),
        "medium_dash_dot" | "mediumdashdot" => Ok(FormatBorder::MediumDashDot),
        "dash_dot_dot" | "dashdotdot" => Ok(FormatBorder::DashDotDot),
        "medium_dash_dot_dot" | "mediumdashdotdot" => Ok(FormatBorder::MediumDashDotDot),
        "slant_dash_dot" | "slantdashdot" => Ok(FormatBorder::SlantDashDot),
        _ => Err(format!(
            "Unknown border style '{}'. Valid styles: thin, medium, thick, dashed, dotted, \
             double, hair, medium_dashed, dash_dot, medium_dash_dot, dash_dot_dot, \
             medium_dash_dot_dot, slant_dash_dot",
            style
        )),
    }
}

/// Parse horizontal alignment string into `FormatAlign` enum value.
pub(crate) fn parse_horizontal_alignment(align: &str) -> Result<FormatAlign, String> {
    match align.to_lowercase().as_str() {
        "left" => Ok(FormatAlign::Left),
        "center" => Ok(FormatAlign::Center),
        "right" => Ok(FormatAlign::Right),
        "fill" => Ok(FormatAlign::Fill),
        "justify" => Ok(FormatAlign::Justify),
        "center_across" => Ok(FormatAlign::CenterAcross),
        "distributed" => Ok(FormatAlign::Distributed),
        _ => Err(format!(
            "Unknown horizontal alignment '{}'. Valid values: left, center, right, \
             fill, justify, center_across, distributed",
            align
        )),
    }
}

/// Parse vertical alignment string into `FormatAlign` enum value.
pub(crate) fn parse_vertical_alignment(align: &str) -> Result<FormatAlign, String> {
    match align.to_lowercase().as_str() {
        "top" => Ok(FormatAlign::Top),
        "center" => Ok(FormatAlign::VerticalCenter),
        "bottom" => Ok(FormatAlign::Bottom),
        "justify" => Ok(FormatAlign::VerticalJustify),
        "distributed" => Ok(FormatAlign::VerticalDistributed),
        _ => Err(format!(
            "Unknown vertical alignment '{}'. Valid values: top, center, bottom, \
             justify, distributed",
            align
        )),
    }
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

/// Parse color string into a rust_xlsxwriter `Color` enum.
/// Wraps `parse_color` — used by features whose setters take `impl Into<Color>`
/// rather than a raw `u32` (shapes, charts, sparklines).
pub(crate) fn parse_color_enum(color_str: &str) -> Result<Color, String> {
    parse_color(color_str).map(Color::RGB)
}

/// Parse header format dictionary into rust_xlsxwriter Format
/// Delegates to parse_format_dict without column-specific options
pub(crate) fn parse_header_format(
    py: Python<'_>,
    fmt_dict: &HashMap<String, Py<PyAny>>,
) -> Result<Format, String> {
    parse_format_dict(py, fmt_dict, false)
}

/// Parse a rich text segment format dictionary into rust_xlsxwriter Format.
/// Rich text segments only carry font-level formatting (bold, italic, color,
/// size, underline, bg_color). Borders/alignment/wrap/num_format are meaningless
/// for an inline text run, so we reuse the no-column-options parser.
pub(crate) fn parse_rich_text_format(
    py: Python<'_>,
    fmt_dict: &HashMap<String, Py<PyAny>>,
) -> Result<Format, String> {
    parse_format_dict(py, fmt_dict, false)
}

/// Keys accepted by `parse_format_dict` regardless of context.
const FORMAT_KEYS_BASE: &[&str] = &[
    "bold",
    "italic",
    "underline",
    "bg_color",
    "font_color",
    "font_size",
    "border",
    "border_left",
    "border_right",
    "border_top",
    "border_bottom",
    "border_color",
    "align_horizontal",
    "align_vertical",
    "wrap_text",
];

/// Keys accepted only when `include_column_options` is true.
const FORMAT_KEYS_COLUMN: &[&str] = &["num_format"];

/// Extract a bool field. None values are treated as unset. Wrong types error.
fn get_bool_field(
    py: Python<'_>,
    fmt_dict: &HashMap<String, Py<PyAny>>,
    key: &str,
) -> Result<Option<bool>, String> {
    let Some(obj) = fmt_dict.get(key) else {
        return Ok(None);
    };
    let bound = obj.bind(py);
    if bound.is_none() {
        return Ok(None);
    }
    bound.extract::<bool>().map(Some).map_err(|_| {
        format!(
            "format option '{}' must be a bool, got {}",
            key,
            pytype_name(bound)
        )
    })
}

/// Extract a string field. None values are treated as unset. Wrong types error.
fn get_string_field(
    py: Python<'_>,
    fmt_dict: &HashMap<String, Py<PyAny>>,
    key: &str,
) -> Result<Option<String>, String> {
    let Some(obj) = fmt_dict.get(key) else {
        return Ok(None);
    };
    let bound = obj.bind(py);
    if bound.is_none() {
        return Ok(None);
    }
    bound.extract::<String>().map(Some).map_err(|_| {
        format!(
            "format option '{}' must be a string, got {}",
            key,
            pytype_name(bound)
        )
    })
}

/// Extract an f64 field. None values are treated as unset. Wrong types error.
fn get_f64_field(
    py: Python<'_>,
    fmt_dict: &HashMap<String, Py<PyAny>>,
    key: &str,
) -> Result<Option<f64>, String> {
    let Some(obj) = fmt_dict.get(key) else {
        return Ok(None);
    };
    let bound = obj.bind(py);
    if bound.is_none() {
        return Ok(None);
    }
    bound.extract::<f64>().map(Some).map_err(|_| {
        format!(
            "format option '{}' must be a number, got {}",
            key,
            pytype_name(bound)
        )
    })
}

/// Extract a border field accepting bool (True=thin) or a style name string.
/// None, missing, or `false` return Ok(None). Unknown types error.
fn get_border_field(
    py: Python<'_>,
    fmt_dict: &HashMap<String, Py<PyAny>>,
    key: &str,
) -> Result<Option<FormatBorder>, String> {
    let Some(obj) = fmt_dict.get(key) else {
        return Ok(None);
    };
    let bound = obj.bind(py);
    if bound.is_none() {
        return Ok(None);
    }
    if let Ok(style_str) = bound.extract::<String>() {
        return Ok(Some(parse_border_style(&style_str)?));
    }
    if let Ok(flag) = bound.extract::<bool>() {
        return Ok(if flag { Some(FormatBorder::Thin) } else { None });
    }
    Err(format!(
        "format option '{}' must be a bool or a style name string, got {}",
        key,
        pytype_name(bound)
    ))
}

/// Shared format parser for header, column, and rich-text formats.
/// When `include_column_options` is true, also handles `num_format`.
/// Unknown keys produce a clear error listing the valid options.
fn parse_format_dict(
    py: Python<'_>,
    fmt_dict: &HashMap<String, Py<PyAny>>,
    include_column_options: bool,
) -> Result<Format, String> {
    // Reject unknown keys so typos (e.g. 'color' vs 'font_color') surface
    // immediately rather than silently producing unformatted output.
    for key in fmt_dict.keys() {
        let known = FORMAT_KEYS_BASE.contains(&key.as_str())
            || (include_column_options && FORMAT_KEYS_COLUMN.contains(&key.as_str()));
        if !known {
            let mut valid: Vec<&str> = FORMAT_KEYS_BASE.to_vec();
            if include_column_options {
                valid.extend_from_slice(FORMAT_KEYS_COLUMN);
            }
            return Err(format!(
                "Unknown format option '{}'. Valid options: {}",
                key,
                valid.join(", ")
            ));
        }
    }

    let mut format = Format::new();

    if get_bool_field(py, fmt_dict, "bold")?.unwrap_or(false) {
        format = format.set_bold();
    }

    if get_bool_field(py, fmt_dict, "italic")?.unwrap_or(false) {
        format = format.set_italic();
    }

    if get_bool_field(py, fmt_dict, "underline")?.unwrap_or(false) {
        format = format.set_underline(rust_xlsxwriter::FormatUnderline::Single);
    }

    if let Some(color_str) = get_string_field(py, fmt_dict, "bg_color")? {
        format = format.set_background_color(parse_color(&color_str)?);
    }

    if let Some(color_str) = get_string_field(py, fmt_dict, "font_color")? {
        format = format.set_font_color(parse_color(&color_str)?);
    }

    if let Some(size) = get_f64_field(py, fmt_dict, "font_size")? {
        format = format.set_font_size(size);
    }

    if include_column_options {
        if let Some(num_fmt_str) = get_string_field(py, fmt_dict, "num_format")? {
            format = format.set_num_format(&num_fmt_str);
        }
    }

    // Borders: `border` applies to all four sides; per-side keys override.
    if let Some(style) = get_border_field(py, fmt_dict, "border")? {
        format = format.set_border(style);
    }

    for (key, setter) in [
        (
            "border_left",
            Format::set_border_left as fn(Format, FormatBorder) -> Format,
        ),
        (
            "border_right",
            Format::set_border_right as fn(Format, FormatBorder) -> Format,
        ),
        (
            "border_top",
            Format::set_border_top as fn(Format, FormatBorder) -> Format,
        ),
        (
            "border_bottom",
            Format::set_border_bottom as fn(Format, FormatBorder) -> Format,
        ),
    ] {
        if let Some(style) = get_border_field(py, fmt_dict, key)? {
            format = setter(format, style);
        }
    }

    if let Some(color_str) = get_string_field(py, fmt_dict, "border_color")? {
        format = format.set_border_color(parse_color(&color_str)?);
    }

    if let Some(align_str) = get_string_field(py, fmt_dict, "align_horizontal")? {
        format = format.set_align(parse_horizontal_alignment(&align_str)?);
    }

    if let Some(align_str) = get_string_field(py, fmt_dict, "align_vertical")? {
        format = format.set_align(parse_vertical_alignment(&align_str)?);
    }

    if get_bool_field(py, fmt_dict, "wrap_text")?.unwrap_or(false) {
        format = format.set_text_wrap();
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
            let excel_dt = naive_datetime_to_excel(dt);
            // Excel doesn't support dates before 1900-01-01;
            // negative/zero serial numbers render as ##### in Excel
            if excel_dt <= 0.0 {
                return CellValue::String(trimmed.to_string());
            }
            return CellValue::DateTime(excel_dt);
        }
    }

    // Try date with locale-aware ordering
    for pattern in date_order.patterns() {
        if let Ok(date) = chrono::NaiveDate::parse_from_str(trimmed, pattern) {
            let excel_date = naive_date_to_excel(date);
            // Excel doesn't support dates before 1900-01-01 (serial 1);
            // negative/zero serial numbers render as ##### in Excel
            if excel_date <= 0.0 {
                return CellValue::String(trimmed.to_string());
            }
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

#[cfg(test)]
mod tests {
    use super::{
        matches_pattern, naive_date_to_excel, parse_border_style, parse_cell_range, parse_cell_ref,
        parse_color, parse_horizontal_alignment, parse_table_style, parse_value,
        parse_vertical_alignment, sanitize_table_name,
    };
    use crate::types::{CellValue, DateOrder};

    #[test]
    fn test_parse_integer() {
        assert!(matches!(
            parse_value("123", DateOrder::Auto),
            CellValue::Integer(123)
        ));
        assert!(matches!(
            parse_value("-456", DateOrder::Auto),
            CellValue::Integer(-456)
        ));
    }

    #[test]
    fn test_parse_float() {
        let value = parse_value("3.14", DateOrder::Auto);
        assert!(
            matches!(value, CellValue::Float(_)),
            "Expected CellValue::Float, got {:?}",
            value
        );
        if let CellValue::Float(v) = value {
            assert!((v - 3.14).abs() < 0.001);
        }
    }

    #[test]
    fn test_parse_boolean() {
        assert!(matches!(
            parse_value("true", DateOrder::Auto),
            CellValue::Boolean(true)
        ));
        assert!(matches!(
            parse_value("TRUE", DateOrder::Auto),
            CellValue::Boolean(true)
        ));
        assert!(matches!(
            parse_value("false", DateOrder::Auto),
            CellValue::Boolean(false)
        ));
        assert!(matches!(
            parse_value("False", DateOrder::Auto),
            CellValue::Boolean(false)
        ));
    }

    #[test]
    fn test_parse_empty() {
        assert!(matches!(parse_value("", DateOrder::Auto), CellValue::Empty));
        assert!(matches!(
            parse_value("   ", DateOrder::Auto),
            CellValue::Empty
        ));
        assert!(matches!(
            parse_value("NaN", DateOrder::Auto),
            CellValue::Empty
        ));
    }

    #[test]
    fn test_parse_date() {
        assert!(matches!(
            parse_value("2024-01-15", DateOrder::Auto),
            CellValue::Date(_)
        ));
        assert!(matches!(
            parse_value("2024/01/15", DateOrder::Auto),
            CellValue::Date(_)
        ));
    }

    #[test]
    fn test_parse_datetime() {
        assert!(matches!(
            parse_value("2024-01-15T10:30:00", DateOrder::Auto),
            CellValue::DateTime(_)
        ));
        assert!(matches!(
            parse_value("2024-01-15 10:30:00", DateOrder::Auto),
            CellValue::DateTime(_)
        ));
    }

    #[test]
    fn test_parse_string() {
        assert!(matches!(
            parse_value("hello", DateOrder::Auto),
            CellValue::String(_)
        ));
    }

    #[test]
    fn test_matches_pattern_exact() {
        assert!(matches_pattern("column_name", "column_name"));
        assert!(!matches_pattern("column_name", "other"));
    }

    #[test]
    fn test_matches_pattern_prefix() {
        assert!(matches_pattern("price_usd", "price_*"));
        assert!(matches_pattern("price_", "price_*"));
        assert!(!matches_pattern("cost_usd", "price_*"));
    }

    #[test]
    fn test_matches_pattern_suffix() {
        assert!(matches_pattern("col_weight", "*_weight"));
        assert!(matches_pattern("_weight", "*_weight"));
        assert!(!matches_pattern("col_height", "*_weight"));
    }

    #[test]
    fn test_matches_pattern_contains() {
        assert!(matches_pattern("leadframe_difference", "*difference*"));
        assert!(matches_pattern("difference", "*difference*"));
        assert!(matches_pattern("my_difference_col", "*difference*"));
        assert!(!matches_pattern("other_column", "*difference*"));
    }

    #[test]
    fn test_matches_pattern_wildcard() {
        // Single "*" matches everything
        assert!(matches_pattern("anything", "*"));
        assert!(matches_pattern("", "*"));
        // Double "**" also matches everything
        assert!(matches_pattern("anything", "**"));
        assert!(matches_pattern("", "**"));
    }

    // --- parse_cell_ref tests ---

    #[test]
    fn test_parse_cell_ref_basic() {
        assert_eq!(parse_cell_ref("A1").unwrap(), (0, 0));
        assert_eq!(parse_cell_ref("B2").unwrap(), (1, 1));
        assert_eq!(parse_cell_ref("Z1").unwrap(), (0, 25));
        assert_eq!(parse_cell_ref("AA1").unwrap(), (0, 26));
        assert_eq!(parse_cell_ref("AZ1").unwrap(), (0, 51));
    }

    #[test]
    fn test_parse_cell_ref_case_insensitive() {
        assert_eq!(parse_cell_ref("a1").unwrap(), (0, 0));
        assert_eq!(parse_cell_ref("aa1").unwrap(), (0, 26));
    }

    #[test]
    fn test_parse_cell_ref_max_column() {
        // XFD = 16384th column = index 16383
        assert_eq!(parse_cell_ref("XFD1").unwrap(), (0, 16383));
    }

    #[test]
    fn test_parse_cell_ref_overflow_column() {
        assert!(parse_cell_ref("ZZZZ1").is_err());
    }

    #[test]
    fn test_parse_cell_ref_exceeds_excel_max() {
        // XFE = 16385th column, exceeds Excel max
        assert!(parse_cell_ref("XFE1").is_err());
    }

    #[test]
    fn test_parse_cell_ref_row_zero() {
        assert!(parse_cell_ref("A0").is_err());
    }

    #[test]
    fn test_parse_cell_ref_empty() {
        assert!(parse_cell_ref("").is_err());
    }

    #[test]
    fn test_parse_cell_ref_no_row() {
        assert!(parse_cell_ref("A").is_err());
    }

    #[test]
    fn test_parse_cell_ref_no_column() {
        assert!(parse_cell_ref("1").is_err());
    }

    // --- parse_cell_range tests ---

    #[test]
    fn test_parse_cell_range_basic() {
        assert_eq!(parse_cell_range("A1:B2").unwrap(), (0, 0, 1, 1));
        assert_eq!(parse_cell_range("A1:D1").unwrap(), (0, 0, 0, 3));
    }

    #[test]
    fn test_parse_cell_range_invalid_format() {
        assert!(parse_cell_range("A1").is_err()); // no colon
        assert!(parse_cell_range("A1:B2:C3").is_err()); // too many colons
    }

    // --- parse_color tests ---

    #[test]
    fn test_parse_color_hex() {
        assert_eq!(parse_color("#FF0000").unwrap(), 0xFF0000);
        assert_eq!(parse_color("#000000").unwrap(), 0x000000);
        assert_eq!(parse_color("#FFFFFF").unwrap(), 0xFFFFFF);
        assert_eq!(parse_color("#4F81BD").unwrap(), 0x4F81BD);
    }

    #[test]
    fn test_parse_color_named() {
        assert_eq!(parse_color("red").unwrap(), 0xFF0000);
        assert_eq!(parse_color("Red").unwrap(), 0xFF0000);
        assert_eq!(parse_color("WHITE").unwrap(), 0xFFFFFF);
        assert_eq!(parse_color("gray").unwrap(), 0x808080);
        assert_eq!(parse_color("grey").unwrap(), 0x808080);
    }

    #[test]
    fn test_parse_color_invalid() {
        assert!(parse_color("#FFF").is_err()); // too short
        assert!(parse_color("#GGGGGG").is_err()); // invalid hex
        assert!(parse_color("chartreuse").is_err()); // unsupported name
    }

    #[test]
    fn test_parse_color_whitespace() {
        assert_eq!(parse_color("  #FF0000  ").unwrap(), 0xFF0000);
        assert_eq!(parse_color("  red  ").unwrap(), 0xFF0000);
    }

    // --- sanitize_table_name tests ---

    #[test]
    fn test_sanitize_table_name_valid() {
        assert_eq!(sanitize_table_name("MyTable"), "MyTable");
        assert_eq!(sanitize_table_name("_table1"), "_table1");
    }

    #[test]
    fn test_sanitize_table_name_special_chars() {
        assert_eq!(sanitize_table_name("My Table!"), "My_Table_");
        assert_eq!(sanitize_table_name("data-2024"), "data_2024");
    }

    #[test]
    fn test_sanitize_table_name_starts_with_digit() {
        assert_eq!(sanitize_table_name("123Data"), "_123Data");
    }

    #[test]
    fn test_sanitize_table_name_truncation() {
        let long_name = "a".repeat(300);
        let sanitized = sanitize_table_name(&long_name);
        assert_eq!(sanitized.len(), 255);
    }

    #[test]
    fn test_sanitize_table_name_empty() {
        assert_eq!(sanitize_table_name(""), "_");
    }

    // --- parse_table_style tests ---

    #[test]
    fn test_parse_table_style_valid() {
        assert!(parse_table_style("None").is_ok());
        assert!(parse_table_style("Light1").is_ok());
        assert!(parse_table_style("Medium14").is_ok());
        assert!(parse_table_style("Dark11").is_ok());
    }

    #[test]
    fn test_parse_table_style_invalid() {
        assert!(parse_table_style("light1").is_err()); // case-sensitive
        assert!(parse_table_style("Medium29").is_err()); // out of range
        assert!(parse_table_style("Dark12").is_err()); // out of range
        assert!(parse_table_style("").is_err());
    }

    // --- naive_date_to_excel tests ---

    #[test]
    fn test_naive_date_to_excel_epoch() {
        // Excel epoch is 1899-12-30, so 1900-01-01 = day 2
        let date = chrono::NaiveDate::from_ymd_opt(1900, 1, 1).unwrap();
        assert_eq!(naive_date_to_excel(date), 2.0);
    }

    #[test]
    fn test_naive_date_to_excel_known_date() {
        // 2024-01-15 is a known Excel serial date
        let date = chrono::NaiveDate::from_ymd_opt(2024, 1, 15).unwrap();
        assert_eq!(naive_date_to_excel(date), 45306.0);
    }

    // --- DateOrder tests ---

    #[test]
    fn test_date_order_parse() {
        assert_eq!(DateOrder::parse("auto"), Some(DateOrder::Auto));
        assert_eq!(DateOrder::parse("mdy"), Some(DateOrder::MDY));
        assert_eq!(DateOrder::parse("us"), Some(DateOrder::MDY));
        assert_eq!(DateOrder::parse("dmy"), Some(DateOrder::DMY));
        assert_eq!(DateOrder::parse("eu"), Some(DateOrder::DMY));
        assert_eq!(DateOrder::parse("european"), Some(DateOrder::DMY));
        assert_eq!(DateOrder::parse("AUTO"), Some(DateOrder::Auto));
        assert_eq!(DateOrder::parse("invalid"), None);
        assert_eq!(DateOrder::parse(""), None);
    }

    // --- parse_border_style tests ---

    #[test]
    fn test_parse_border_style_valid() {
        use rust_xlsxwriter::FormatBorder;
        assert_eq!(parse_border_style("thin").unwrap(), FormatBorder::Thin);
        assert_eq!(parse_border_style("medium").unwrap(), FormatBorder::Medium);
        assert_eq!(parse_border_style("thick").unwrap(), FormatBorder::Thick);
        assert_eq!(parse_border_style("dashed").unwrap(), FormatBorder::Dashed);
        assert_eq!(parse_border_style("dotted").unwrap(), FormatBorder::Dotted);
        assert_eq!(parse_border_style("double").unwrap(), FormatBorder::Double);
        assert_eq!(parse_border_style("hair").unwrap(), FormatBorder::Hair);
    }

    #[test]
    fn test_parse_border_style_case_insensitive() {
        use rust_xlsxwriter::FormatBorder;
        assert_eq!(parse_border_style("THIN").unwrap(), FormatBorder::Thin);
        assert_eq!(parse_border_style("Thick").unwrap(), FormatBorder::Thick);
        assert_eq!(parse_border_style("Medium").unwrap(), FormatBorder::Medium);
    }

    #[test]
    fn test_parse_border_style_aliases() {
        use rust_xlsxwriter::FormatBorder;
        assert_eq!(
            parse_border_style("medium_dashed").unwrap(),
            FormatBorder::MediumDashed
        );
        assert_eq!(
            parse_border_style("mediumdashed").unwrap(),
            FormatBorder::MediumDashed
        );
        assert_eq!(
            parse_border_style("dash_dot").unwrap(),
            FormatBorder::DashDot
        );
        assert_eq!(
            parse_border_style("dashdot").unwrap(),
            FormatBorder::DashDot
        );
        assert_eq!(
            parse_border_style("slant_dash_dot").unwrap(),
            FormatBorder::SlantDashDot
        );
        assert_eq!(
            parse_border_style("slantdashdot").unwrap(),
            FormatBorder::SlantDashDot
        );
    }

    #[test]
    fn test_parse_border_style_invalid() {
        assert!(parse_border_style("").is_err());
        assert!(parse_border_style("bold").is_err());
        assert!(parse_border_style("heavy").is_err());
    }

    // --- parse_horizontal_alignment tests ---

    #[test]
    fn test_parse_horizontal_alignment_valid() {
        use rust_xlsxwriter::FormatAlign;
        assert_eq!(
            parse_horizontal_alignment("left").unwrap(),
            FormatAlign::Left
        );
        assert_eq!(
            parse_horizontal_alignment("center").unwrap(),
            FormatAlign::Center
        );
        assert_eq!(
            parse_horizontal_alignment("right").unwrap(),
            FormatAlign::Right
        );
        assert_eq!(
            parse_horizontal_alignment("fill").unwrap(),
            FormatAlign::Fill
        );
        assert_eq!(
            parse_horizontal_alignment("justify").unwrap(),
            FormatAlign::Justify
        );
        assert_eq!(
            parse_horizontal_alignment("CENTER").unwrap(),
            FormatAlign::Center
        );
    }

    #[test]
    fn test_parse_horizontal_alignment_invalid() {
        assert!(parse_horizontal_alignment("").is_err());
        assert!(parse_horizontal_alignment("top").is_err());
        assert!(parse_horizontal_alignment("middle").is_err());
    }

    // --- parse_vertical_alignment tests ---

    #[test]
    fn test_parse_vertical_alignment_valid() {
        use rust_xlsxwriter::FormatAlign;
        assert_eq!(parse_vertical_alignment("top").unwrap(), FormatAlign::Top);
        assert_eq!(
            parse_vertical_alignment("center").unwrap(),
            FormatAlign::VerticalCenter
        );
        assert_eq!(
            parse_vertical_alignment("bottom").unwrap(),
            FormatAlign::Bottom
        );
        assert_eq!(
            parse_vertical_alignment("justify").unwrap(),
            FormatAlign::VerticalJustify
        );
        assert_eq!(parse_vertical_alignment("TOP").unwrap(), FormatAlign::Top);
    }

    #[test]
    fn test_parse_vertical_alignment_invalid() {
        assert!(parse_vertical_alignment("").is_err());
        assert!(parse_vertical_alignment("left").is_err());
        assert!(parse_vertical_alignment("right").is_err());
        assert!(parse_vertical_alignment("general").is_err());
    }

    // --- naive_datetime_to_excel tests ---

    #[test]
    fn test_naive_datetime_to_excel_noon() {
        let dt = chrono::NaiveDate::from_ymd_opt(2024, 1, 15)
            .unwrap()
            .and_hms_opt(12, 0, 0)
            .unwrap();
        let result = super::naive_datetime_to_excel(dt);
        // 2024-01-15 = 45306.0, noon = 0.5
        assert!((result - 45306.5).abs() < 0.001);
    }

    #[test]
    fn test_naive_datetime_to_excel_midnight() {
        let dt = chrono::NaiveDate::from_ymd_opt(2024, 1, 15)
            .unwrap()
            .and_hms_opt(0, 0, 0)
            .unwrap();
        let result = super::naive_datetime_to_excel(dt);
        assert!((result - 45306.0).abs() < 0.001);
    }

    #[test]
    fn test_naive_datetime_to_excel_end_of_day() {
        let dt = chrono::NaiveDate::from_ymd_opt(2024, 1, 15)
            .unwrap()
            .and_hms_opt(23, 59, 59)
            .unwrap();
        let result = super::naive_datetime_to_excel(dt);
        assert!((result - 45307.0).abs() < 0.001); // just under next day
    }

    // --- parse_icon_type tests ---

    #[test]
    fn test_parse_icon_type_valid() {
        assert!(super::parse_icon_type("3_arrows").is_ok());
        assert!(super::parse_icon_type("3arrows").is_ok());
        assert!(super::parse_icon_type("3_flags").is_ok());
        assert!(super::parse_icon_type("3_traffic_lights").is_ok());
        assert!(super::parse_icon_type("4_arrows").is_ok());
        assert!(super::parse_icon_type("5_quarters").is_ok());
        assert!(super::parse_icon_type("5_rating").is_ok());
    }

    #[test]
    fn test_parse_icon_type_case_insensitive() {
        assert!(super::parse_icon_type("3_ARROWS").is_ok());
        assert!(super::parse_icon_type("5_Quarters").is_ok());
    }

    #[test]
    fn test_parse_icon_type_invalid() {
        assert!(super::parse_icon_type("").is_err());
        assert!(super::parse_icon_type("6_arrows").is_err());
        assert!(super::parse_icon_type("invalid").is_err());
    }

    // --- naive_date_to_excel pre-epoch guard tests ---

    #[test]
    fn test_naive_date_to_excel_pre_epoch() {
        // Dates before 1900-01-01 should be treated as strings, not invalid serial numbers
        let result = super::parse_value("1899-01-01", crate::types::DateOrder::Auto);
        assert!(matches!(result, crate::types::CellValue::String(_)));
    }
}
