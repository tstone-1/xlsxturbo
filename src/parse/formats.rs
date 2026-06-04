use crate::types::{extract_opt, pytype_name};
use indexmap::IndexMap;
use pyo3::{prelude::*, Py};
use rust_xlsxwriter::{ConditionalFormatIconType, Format, FormatAlign, FormatBorder};
use std::collections::HashMap;

use super::colors::parse_color;
use super::patterns::matches_pattern;

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
    extract_opt(py, fmt_dict.get(key), |bound| {
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
    extract_opt(py, fmt_dict.get(key), |bound| {
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
    extract_opt(py, fmt_dict.get(key), |bound| {
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
