//! Excel application functions for applying features to worksheets

use crate::convert::{write_py_value_with_format, DATETIME_NUM_FORMAT, DATE_NUM_FORMAT};
use crate::extract::pydict_to_hashmap;
use crate::parse::{
    matches_pattern, parse_cell_range, parse_cell_ref, parse_color, parse_column_format,
    parse_header_format, parse_horizontal_alignment, parse_icon_type, parse_rich_text_format,
    parse_vertical_alignment,
};
use crate::types::{
    CellWrite, CheckboxConfig, Comment, ConditionalFormatConfigs, Hyperlink, ImageConfig,
    MergedRange, RichTextSegment, ValidationConfig,
};
use indexmap::IndexMap;
use pyo3::prelude::*;
use rust_xlsxwriter::{
    ConditionalFormat2ColorScale, ConditionalFormat3ColorScale, ConditionalFormatBlank,
    ConditionalFormatCell, ConditionalFormatCellRule, ConditionalFormatDataBar,
    ConditionalFormatDataBarDirection, ConditionalFormatIconSet, ConditionalFormatText,
    ConditionalFormatTextRule, DataValidation, DataValidationErrorStyle, Format, Image, Note,
    Worksheet,
};
use std::collections::HashMap;

/// Apply column widths to worksheet, supporting '_all' global cap
pub(crate) fn apply_column_widths(
    worksheet: &mut Worksheet,
    col_count: u16,
    widths: &HashMap<String, f64>,
) -> Result<(), String> {
    let global_width = widths.get("_all").copied();

    for col_idx in 0..col_count {
        let col_key = col_idx.to_string();
        // Specific column overrides '_all'
        if let Some(width) = widths.get(&col_key) {
            worksheet
                .set_column_width(col_idx, *width)
                .map_err(|e| format!("Failed to set column width: {}", e))?;
        } else if let Some(width) = global_width {
            worksheet
                .set_column_width(col_idx, width)
                .map_err(|e| format!("Failed to set column width: {}", e))?;
        }
    }
    Ok(())
}

/// Apply column widths with autofit and cap: autofit each column to content, then cap at '_all'.
/// Uses pre-computed content widths to apply min(autofit, cap) per column.
///
/// Caller must ensure this is NOT called in constant_memory mode (autofit is unsupported).
pub(crate) fn apply_column_widths_with_autofit_cap(
    worksheet: &mut Worksheet,
    col_count: u16,
    widths: &HashMap<String, f64>,
    content_widths: &[f64],
) -> Result<(), String> {
    let global_cap = widths.get("_all").copied().unwrap_or(f64::MAX);

    for col_idx in 0..col_count {
        let col_key = col_idx.to_string();
        if let Some(width) = widths.get(&col_key) {
            // Specific width overrides autofit and cap
            worksheet
                .set_column_width(col_idx, *width)
                .map_err(|e| format!("Failed to set column width: {}", e))?;
        } else {
            // Autofit capped at '_all'
            let autofit_width = content_widths
                .get(col_idx as usize)
                .copied()
                .unwrap_or(8.43); // Excel default
            let capped = autofit_width.min(global_cap);
            worksheet
                .set_column_width(col_idx, capped)
                .map_err(|e| format!("Failed to set column width: {}", e))?;
        }
    }
    Ok(())
}

/// Apply formula columns to worksheet
/// Formula templates can use {row} which is replaced with the actual row number (1-based)
pub(crate) fn apply_formula_columns(
    worksheet: &mut Worksheet,
    formula_columns: &IndexMap<String, String>,
    start_col: u16,
    data_start_row: u32,
    data_end_row: u32,
    include_header: bool,
    header_format: Option<&Format>,
) -> Result<u16, String> {
    let mut col_offset = 0u16;

    for (col_name, formula_template) in formula_columns {
        let col_idx = start_col
            .checked_add(col_offset)
            .ok_or("Formula column index exceeds u16 limit")?;

        // Write header for formula column (only when headers are enabled)
        if include_header {
            if let Some(fmt) = header_format {
                worksheet
                    .write_string_with_format(0, col_idx, col_name, fmt)
                    .map_err(|e| format!("Failed to write formula column header: {}", e))?;
            } else {
                worksheet
                    .write_string(0, col_idx, col_name)
                    .map_err(|e| format!("Failed to write formula column header: {}", e))?;
            }
        }

        // Write formula for each data row
        for row in data_start_row..=data_end_row {
            // Replace {row} with actual row number (Excel is 1-based)
            let excel_row = row + 1; // Convert 0-based to 1-based
            let formula = formula_template.replace("{row}", &excel_row.to_string());

            worksheet
                .write_formula(row, col_idx, formula.as_str())
                .map_err(|e| format!("Failed to write formula at row {}: {}", row, e))?;
        }

        col_offset = col_offset
            .checked_add(1)
            .ok_or("Formula column count exceeds u16 limit")?;
    }

    Ok(col_offset)
}

/// Apply merged ranges to worksheet
pub(crate) fn apply_merged_ranges(
    py: Python<'_>,
    worksheet: &mut Worksheet,
    merged_ranges: &[MergedRange],
) -> Result<(), String> {
    for (range_str, text, format_dict) in merged_ranges {
        let (first_row, first_col, last_row, last_col) = parse_cell_range(range_str)?;

        // Build format if provided
        let format = if let Some(fmt_dict) = format_dict {
            let parsed = parse_header_format(py, fmt_dict)?;
            Some(parsed)
        } else {
            None
        };

        // Apply merge with or without format
        if let Some(ref fmt) = format {
            worksheet
                .merge_range(first_row, first_col, last_row, last_col, text, fmt)
                .map_err(|e| format!("Failed to merge range '{}': {}", range_str, e))?;
        } else {
            // Create default center-aligned format for merged cells
            let default_fmt = Format::new().set_align(rust_xlsxwriter::FormatAlign::Center);
            worksheet
                .merge_range(first_row, first_col, last_row, last_col, text, &default_fmt)
                .map_err(|e| format!("Failed to merge range '{}': {}", range_str, e))?;
        }
    }

    Ok(())
}

/// Apply hyperlinks to worksheet
pub(crate) fn apply_hyperlinks(
    worksheet: &mut Worksheet,
    hyperlinks: &[Hyperlink],
) -> Result<(), String> {
    for (cell_ref, url, display_text) in hyperlinks {
        let (row, col) = parse_cell_ref(cell_ref)?;

        if let Some(text) = display_text {
            worksheet
                .write_url_with_text(row, col, url.as_str(), text.as_str())
                .map_err(|e| format!("Failed to write hyperlink at '{}': {}", cell_ref, e))?;
        } else {
            worksheet
                .write_url(row, col, url.as_str())
                .map_err(|e| format!("Failed to write hyperlink at '{}': {}", cell_ref, e))?;
        }
    }

    Ok(())
}

/// Apply comments/notes to worksheet
pub(crate) fn apply_comments(
    worksheet: &mut Worksheet,
    comments: &HashMap<String, Comment>,
) -> Result<(), String> {
    for (cell_ref, (text, author)) in comments {
        let (row, col) = parse_cell_ref(cell_ref)?;

        let mut note = Note::new(text);
        if let Some(auth) = author {
            note = note.set_author(auth);
        }

        worksheet
            .insert_note(row, col, &note)
            .map_err(|e| format!("Failed to insert note at '{}': {}", cell_ref, e))?;
    }

    Ok(())
}

/// Extract an optional string field from a validation config dict.
/// Returns Ok(None) for missing or Python-None values. Wrong types produce an error.
fn validation_string_field(
    py: Python<'_>,
    config: &ValidationConfig,
    col_pattern: &str,
    key: &str,
) -> Result<Option<String>, String> {
    let Some(obj) = config.get(key) else {
        return Ok(None);
    };
    let bound = obj.bind(py);
    if bound.is_none() {
        return Ok(None);
    }
    bound
        .extract::<String>()
        .map(Some)
        .map_err(|_| format!("validations['{}']: '{}' must be a string", col_pattern, key))
}

/// Apply data validations to worksheet
pub(crate) fn apply_validations(
    py: Python<'_>,
    worksheet: &mut Worksheet,
    columns: &[String],
    data_start_row: u32,
    data_end_row: u32,
    validations: &IndexMap<String, ValidationConfig>,
) -> Result<(), String> {
    for (col_pattern, config) in validations {
        // Find matching columns
        let col_indices: Vec<u16> = columns
            .iter()
            .enumerate()
            .filter(|(_, name)| matches_pattern(name, col_pattern))
            .map(|(idx, _)| idx as u16) // safe: col_count already validated via u16::try_from
            .collect();

        if col_indices.is_empty() {
            continue;
        }

        // Get validation type
        let val_type: String = config
            .get("type")
            .ok_or_else(|| format!("validations['{}']: missing 'type' key", col_pattern))?
            .bind(py)
            .extract()
            .map_err(|e| format!("validations['{}']: invalid 'type': {}", col_pattern, e))?;

        for col_idx in col_indices {
            let validation = match val_type.to_lowercase().as_str() {
                "list" => {
                    // List validation: dropdown with values
                    let values: Vec<String> = config
                        .get("values")
                        .ok_or_else(|| {
                            format!(
                                "validations['{}']: list type requires 'values'",
                                col_pattern
                            )
                        })?
                        .bind(py)
                        .extract()
                        .map_err(|e| {
                            format!("validations['{}']: invalid 'values': {}", col_pattern, e)
                        })?;

                    // Check Excel's 255 character limit for list validation
                    let total_chars: usize = values.iter().map(|s| s.len()).sum::<usize>()
                        + values.len().saturating_sub(1); // commas between items
                    if total_chars > 255 {
                        return Err(format!(
                            "validations['{}']: list values exceed Excel's 255 character limit ({} chars). \
                             Use fewer or shorter values.",
                            col_pattern, total_chars
                        ));
                    }

                    let values_refs: Vec<&str> = values.iter().map(|s| s.as_str()).collect();
                    DataValidation::new()
                        .allow_list_strings(&values_refs)
                        .map_err(|e| format!("Failed to create list validation: {}", e))?
                }
                "whole_number" | "whole" | "integer" => {
                    // Whole number validation with between rule
                    let min: i32 = config
                        .get("min")
                        .and_then(|v| v.bind(py).extract().ok())
                        .unwrap_or(i32::MIN);
                    let max: i32 = config
                        .get("max")
                        .and_then(|v| v.bind(py).extract().ok())
                        .unwrap_or(i32::MAX);
                    DataValidation::new()
                        .allow_whole_number(rust_xlsxwriter::DataValidationRule::Between(min, max))
                }
                "decimal" | "number" => {
                    // Decimal validation with between rule
                    let min: f64 = config
                        .get("min")
                        .and_then(|v| v.bind(py).extract().ok())
                        .unwrap_or(f64::MIN);
                    let max: f64 = config
                        .get("max")
                        .and_then(|v| v.bind(py).extract().ok())
                        .unwrap_or(f64::MAX);
                    DataValidation::new().allow_decimal_number(
                        rust_xlsxwriter::DataValidationRule::Between(min, max),
                    )
                }
                "text_length" | "textlength" | "length" => {
                    // Text length validation with between rule
                    let min: u32 = config
                        .get("min")
                        .and_then(|v| v.bind(py).extract().ok())
                        .unwrap_or(0);
                    let max: u32 = config
                        .get("max")
                        .and_then(|v| v.bind(py).extract().ok())
                        .unwrap_or(u32::MAX);
                    DataValidation::new()
                        .allow_text_length(rust_xlsxwriter::DataValidationRule::Between(min, max))
                }
                _ => {
                    return Err(format!(
                        "Unknown validation type '{}'. Valid types: list, whole_number, decimal, text_length",
                        val_type
                    ));
                }
            };

            // Add optional input message
            let validation = if let Some(msg) =
                validation_string_field(py, config, col_pattern, "input_message")?
            {
                let title = validation_string_field(py, config, col_pattern, "input_title")?
                    .unwrap_or_default();
                validation
                    .set_input_title(&title)
                    .map_err(|e| format!("Failed to set input title: {}", e))?
                    .set_input_message(&msg)
                    .map_err(|e| format!("Failed to set input message: {}", e))?
            } else {
                validation
            };

            // Add optional error message
            let validation = if let Some(msg) =
                validation_string_field(py, config, col_pattern, "error_message")?
            {
                let title = validation_string_field(py, config, col_pattern, "error_title")?
                    .unwrap_or_default();
                validation
                    .set_error_title(&title)
                    .map_err(|e| format!("Failed to set error title: {}", e))?
                    .set_error_message(&msg)
                    .map_err(|e| format!("Failed to set error message: {}", e))?
                    .set_error_style(DataValidationErrorStyle::Stop)
            } else {
                validation
            };

            worksheet
                .add_data_validation(data_start_row, col_idx, data_end_row, col_idx, &validation)
                .map_err(|e| format!("Failed to add validation: {}", e))?;
        }
    }

    Ok(())
}

/// Apply rich text to worksheet
pub(crate) fn apply_rich_text(
    py: Python<'_>,
    worksheet: &mut Worksheet,
    rich_text: &HashMap<String, Vec<RichTextSegment>>,
) -> Result<(), String> {
    for (cell_ref, segments) in rich_text {
        let (row, col) = parse_cell_ref(cell_ref)?;

        // Build formats and strings separately
        let mut formats: Vec<Format> = Vec::new();
        let mut texts: Vec<String> = Vec::new();

        for (text, format_dict) in segments {
            if let Some(fmt_dict) = format_dict {
                let format = parse_rich_text_format(py, fmt_dict)?;
                formats.push(format);
            } else {
                formats.push(Format::new());
            }
            texts.push(text.clone());
        }

        // Create the segments as tuples of (&Format, &str)
        let rich_segments: Vec<(&Format, &str)> = formats
            .iter()
            .zip(texts.iter())
            .map(|(f, t)| (f, t.as_str()))
            .collect();

        if !rich_segments.is_empty() {
            worksheet
                .write_rich_string(row, col, &rich_segments)
                .map_err(|e| format!("Failed to write rich text at '{}': {}", cell_ref, e))?;
        }
    }

    Ok(())
}

/// Extract an optional f64 image option. Wrong types produce an error.
fn image_f64_field(
    py: Python<'_>,
    opts: &HashMap<String, Py<PyAny>>,
    cell_ref: &str,
    key: &str,
) -> Result<Option<f64>, String> {
    let Some(obj) = opts.get(key) else {
        return Ok(None);
    };
    let bound = obj.bind(py);
    if bound.is_none() {
        return Ok(None);
    }
    bound
        .extract::<f64>()
        .map(Some)
        .map_err(|_| format!("images['{}']: '{}' must be a number", cell_ref, key))
}

/// Extract an optional string image option. Wrong types produce an error.
fn image_string_field(
    py: Python<'_>,
    opts: &HashMap<String, Py<PyAny>>,
    cell_ref: &str,
    key: &str,
) -> Result<Option<String>, String> {
    let Some(obj) = opts.get(key) else {
        return Ok(None);
    };
    let bound = obj.bind(py);
    if bound.is_none() {
        return Ok(None);
    }
    bound
        .extract::<String>()
        .map(Some)
        .map_err(|_| format!("images['{}']: '{}' must be a string", cell_ref, key))
}

/// Apply images to worksheet
pub(crate) fn apply_images(
    py: Python<'_>,
    worksheet: &mut Worksheet,
    images: &HashMap<String, ImageConfig>,
) -> Result<(), String> {
    for (cell_ref, (path, options)) in images {
        let (row, col) = parse_cell_ref(cell_ref)?;

        let mut image =
            Image::new(path).map_err(|e| format!("Failed to load image '{}': {}", path, e))?;

        // Apply options if provided
        if let Some(opts) = options {
            const IMAGE_KEYS: &[&str] = &["path", "scale_width", "scale_height", "alt_text"];
            for key in opts.keys() {
                if !IMAGE_KEYS.contains(&key.as_str()) {
                    return Err(format!(
                        "images['{}']: unknown option '{}'. Valid: scale_width, scale_height, alt_text",
                        cell_ref, key
                    ));
                }
            }
            if let Some(scale) = image_f64_field(py, opts, cell_ref, "scale_width")? {
                image = image.set_scale_width(scale);
            }
            if let Some(scale) = image_f64_field(py, opts, cell_ref, "scale_height")? {
                image = image.set_scale_height(scale);
            }
            if let Some(alt) = image_string_field(py, opts, cell_ref, "alt_text")? {
                image = image.set_alt_text(&alt);
            }
        }

        worksheet
            .insert_image(row, col, &image)
            .map_err(|e| format!("Failed to insert image at '{}': {}", cell_ref, e))?;
    }

    Ok(())
}

/// Apply checkboxes to worksheet
pub(crate) fn apply_checkboxes(
    py: Python<'_>,
    worksheet: &mut Worksheet,
    checkboxes: &HashMap<String, CheckboxConfig>,
) -> Result<(), String> {
    for (cell_ref, (checked, format_dict)) in checkboxes {
        let (row, col) = parse_cell_ref(cell_ref)?;

        if let Some(fmt_dict) = format_dict {
            let fmt = parse_column_format(py, fmt_dict)
                .map_err(|e| format!("checkboxes['{}']: {}", cell_ref, e))?;
            worksheet
                .insert_checkbox_with_format(row, col, *checked, &fmt)
                .map_err(|e| format!("checkboxes['{}']: {}", cell_ref, e))?;
        } else {
            worksheet
                .insert_checkbox(row, col, *checked)
                .map_err(|e| format!("checkboxes['{}']: {}", cell_ref, e))?;
        }
    }

    Ok(())
}

/// Apply arbitrary cell writes to a worksheet
pub(crate) fn apply_cells(
    py: Python<'_>,
    worksheet: &mut Worksheet,
    cells: &[CellWrite],
) -> Result<(), String> {
    let date_format = Format::new().set_num_format(DATE_NUM_FORMAT);
    let datetime_format = Format::new().set_num_format(DATETIME_NUM_FORMAT);

    for cell in cells {
        let value = cell.value.bind(py);
        let has_formatting = cell.num_format.is_some()
            || cell.align_horizontal.is_some()
            || cell.align_vertical.is_some()
            || cell.wrap_text;
        let fmt = if has_formatting {
            let mut f = Format::new();
            if let Some(nf) = &cell.num_format {
                f = f.set_num_format(nf);
            }
            if let Some(ah) = &cell.align_horizontal {
                f = f.set_align(parse_horizontal_alignment(ah)?);
            }
            if let Some(av) = &cell.align_vertical {
                f = f.set_align(parse_vertical_alignment(av)?);
            }
            if cell.wrap_text {
                f = f.set_text_wrap();
            }
            Some(f)
        } else {
            None
        };
        write_py_value_with_format(
            worksheet,
            cell.row,
            cell.col,
            value,
            &date_format,
            &datetime_format,
            fmt.as_ref(),
        )?;
    }
    Ok(())
}

/// Add a cell-rule conditional format (Blank/Text/Cell) with an optional `format` applied first.
macro_rules! add_cell_cf {
    ($ws:expr, $r0:expr, $c:expr, $r1:expr, $cf:expr, $fmt:expr, $what:literal) => {{
        let __cf = match $fmt {
            Some(f) => $cf.set_format(f),
            None => $cf,
        };
        $ws.add_conditional_format($r0, $c, $r1, $c, &__cf)
            .map_err(|e| format!(concat!("Failed to add ", $what, ": {}"), e))?;
    }};
}

/// Add a visualization-style conditional format (color scale, data bar, icon set).
/// These types don't accept a user Format — their appearance IS the format.
macro_rules! add_viz_cf {
    ($ws:expr, $r0:expr, $c:expr, $r1:expr, $cf:expr, $what:literal) => {{
        $ws.add_conditional_format($r0, $c, $r1, $c, &$cf)
            .map_err(|e| format!(concat!("Failed to add ", $what, ": {}"), e))?;
    }};
}

/// Extract an optional color field from a conditional format config.
/// None values are treated as unset. Wrong types error with context.
fn cf_optional_color(
    py: Python<'_>,
    config: &HashMap<String, Py<PyAny>>,
    col_pattern: &str,
    key: &str,
) -> Result<Option<u32>, String> {
    let Some(obj) = config.get(key) else {
        return Ok(None);
    };
    let bound = obj.bind(py);
    if bound.is_none() {
        return Ok(None);
    }
    let color_str = bound.extract::<String>().map_err(|_| {
        format!(
            "conditional_formats['{}']: '{}' must be a color string",
            col_pattern, key
        )
    })?;
    parse_color(&color_str).map(Some)
}

/// Extract an optional bool field from a conditional format config.
fn cf_optional_bool(
    py: Python<'_>,
    config: &HashMap<String, Py<PyAny>>,
    col_pattern: &str,
    key: &str,
) -> Result<Option<bool>, String> {
    let Some(obj) = config.get(key) else {
        return Ok(None);
    };
    let bound = obj.bind(py);
    if bound.is_none() {
        return Ok(None);
    }
    bound.extract::<bool>().map(Some).map_err(|_| {
        format!(
            "conditional_formats['{}']: '{}' must be a bool",
            col_pattern, key
        )
    })
}

/// Extract a required string field.
fn cf_required_string(
    py: Python<'_>,
    config: &HashMap<String, Py<PyAny>>,
    col_pattern: &str,
    key: &str,
) -> Result<String, String> {
    config
        .get(key)
        .ok_or_else(|| {
            format!(
                "conditional_formats['{}']: missing '{}' key",
                col_pattern, key
            )
        })?
        .bind(py)
        .extract::<String>()
        .map_err(|e| {
            format!(
                "conditional_formats['{}']: invalid '{}': {}",
                col_pattern, key, e
            )
        })
}

/// Extract a required f64 field.
fn cf_required_f64(
    py: Python<'_>,
    config: &HashMap<String, Py<PyAny>>,
    col_pattern: &str,
    key: &str,
) -> Result<f64, String> {
    config
        .get(key)
        .ok_or_else(|| {
            format!(
                "conditional_formats['{}']: missing '{}' key",
                col_pattern, key
            )
        })?
        .bind(py)
        .extract::<f64>()
        .map_err(|e| {
            format!(
                "conditional_formats['{}']: invalid '{}': {}",
                col_pattern, key, e
            )
        })
}

/// Parse the optional `format` dict on a cell-rule conditional format config.
fn parse_cf_format(
    py: Python<'_>,
    config: &HashMap<String, Py<PyAny>>,
    col_pattern: &str,
) -> Result<Option<Format>, String> {
    let Some(obj) = config.get("format") else {
        return Ok(None);
    };
    let bound = obj.bind(py);
    if bound.is_none() {
        return Ok(None);
    }
    let fmt_dict = bound.cast::<pyo3::types::PyDict>().map_err(|_| {
        format!(
            "conditional_formats['{}']: 'format' must be a dict",
            col_pattern
        )
    })?;
    let map = pydict_to_hashmap(fmt_dict)
        .map_err(|e| format!("conditional_formats['{}']: {}", col_pattern, e))?;
    parse_column_format(py, &map).map(Some)
}

/// Apply a 2-color-scale conditional format.
fn apply_2_color_scale(
    py: Python<'_>,
    worksheet: &mut Worksheet,
    col_pattern: &str,
    config: &HashMap<String, Py<PyAny>>,
    col_idx: u16,
    data_start_row: u32,
    data_end_row: u32,
) -> Result<(), String> {
    let mut cf = ConditionalFormat2ColorScale::new();
    if let Some(c) = cf_optional_color(py, config, col_pattern, "min_color")? {
        cf = cf.set_minimum_color(c);
    }
    if let Some(c) = cf_optional_color(py, config, col_pattern, "max_color")? {
        cf = cf.set_maximum_color(c);
    }
    add_viz_cf!(
        worksheet,
        data_start_row,
        col_idx,
        data_end_row,
        cf,
        "2_color_scale"
    );
    Ok(())
}

/// Apply a 3-color-scale conditional format.
fn apply_3_color_scale(
    py: Python<'_>,
    worksheet: &mut Worksheet,
    col_pattern: &str,
    config: &HashMap<String, Py<PyAny>>,
    col_idx: u16,
    data_start_row: u32,
    data_end_row: u32,
) -> Result<(), String> {
    let mut cf = ConditionalFormat3ColorScale::new();
    if let Some(c) = cf_optional_color(py, config, col_pattern, "min_color")? {
        cf = cf.set_minimum_color(c);
    }
    if let Some(c) = cf_optional_color(py, config, col_pattern, "mid_color")? {
        cf = cf.set_midpoint_color(c);
    }
    if let Some(c) = cf_optional_color(py, config, col_pattern, "max_color")? {
        cf = cf.set_maximum_color(c);
    }
    add_viz_cf!(
        worksheet,
        data_start_row,
        col_idx,
        data_end_row,
        cf,
        "3_color_scale"
    );
    Ok(())
}

/// Apply a data-bar conditional format.
fn apply_data_bar(
    py: Python<'_>,
    worksheet: &mut Worksheet,
    col_pattern: &str,
    config: &HashMap<String, Py<PyAny>>,
    col_idx: u16,
    data_start_row: u32,
    data_end_row: u32,
) -> Result<(), String> {
    let mut cf = ConditionalFormatDataBar::new();
    if let Some(c) = cf_optional_color(py, config, col_pattern, "bar_color")? {
        cf = cf.set_fill_color(c);
    }
    if let Some(c) = cf_optional_color(py, config, col_pattern, "border_color")? {
        cf = cf.set_border_color(c);
    }
    if cf_optional_bool(py, config, col_pattern, "solid")?.unwrap_or(false) {
        cf = cf.set_solid_fill(true);
    }
    if let Some(obj) = config.get("direction") {
        let bound = obj.bind(py);
        if !bound.is_none() {
            let s = bound.extract::<String>().map_err(|_| {
                format!(
                    "conditional_formats['{}']: 'direction' must be a string",
                    col_pattern
                )
            })?;
            let dir = match s.to_lowercase().as_str() {
                "left_to_right" | "ltr" => ConditionalFormatDataBarDirection::LeftToRight,
                "right_to_left" | "rtl" => ConditionalFormatDataBarDirection::RightToLeft,
                "context" | "" => ConditionalFormatDataBarDirection::Context,
                _ => {
                    return Err(format!(
                        "Unknown direction '{}'. Valid: left_to_right, right_to_left, context",
                        s
                    ));
                }
            };
            cf = cf.set_direction(dir);
        }
    }
    add_viz_cf!(
        worksheet,
        data_start_row,
        col_idx,
        data_end_row,
        cf,
        "data_bar"
    );
    Ok(())
}

/// Apply an icon-set conditional format.
fn apply_icon_set(
    py: Python<'_>,
    worksheet: &mut Worksheet,
    col_pattern: &str,
    config: &HashMap<String, Py<PyAny>>,
    col_idx: u16,
    data_start_row: u32,
    data_end_row: u32,
) -> Result<(), String> {
    let mut cf = ConditionalFormatIconSet::new();
    if let Some(obj) = config.get("icon_type") {
        let bound = obj.bind(py);
        if !bound.is_none() {
            let s = bound.extract::<String>().map_err(|_| {
                format!(
                    "conditional_formats['{}']: 'icon_type' must be a string",
                    col_pattern
                )
            })?;
            cf = cf.set_icon_type(parse_icon_type(&s)?);
        }
    }
    if cf_optional_bool(py, config, col_pattern, "reverse")?.unwrap_or(false) {
        cf = cf.reverse_icons(true);
    }
    if cf_optional_bool(py, config, col_pattern, "icons_only")?.unwrap_or(false) {
        cf = cf.show_icons_only(true);
    }
    add_viz_cf!(
        worksheet,
        data_start_row,
        col_idx,
        data_end_row,
        cf,
        "icon_set"
    );
    Ok(())
}

/// Apply a cell-rule conditional format, dispatching by criteria family.
fn apply_cell_conditional(
    py: Python<'_>,
    worksheet: &mut Worksheet,
    col_pattern: &str,
    config: &HashMap<String, Py<PyAny>>,
    col_idx: u16,
    data_start_row: u32,
    data_end_row: u32,
) -> Result<(), String> {
    let criteria: String = config
        .get("criteria")
        .ok_or_else(|| {
            format!(
                "conditional_formats['{}']: 'cell' type requires 'criteria' key",
                col_pattern
            )
        })?
        .bind(py)
        .extract()
        .map_err(|e| {
            format!(
                "conditional_formats['{}']: invalid 'criteria': {}",
                col_pattern, e
            )
        })?;

    let fmt = parse_cf_format(py, config, col_pattern)?;
    let criteria_lower = criteria.to_lowercase();

    // Blank / no-blank
    match criteria_lower.as_str() {
        "blanks" | "blank" => {
            add_cell_cf!(
                worksheet,
                data_start_row,
                col_idx,
                data_end_row,
                ConditionalFormatBlank::new(),
                fmt,
                "blanks format"
            );
            return Ok(());
        }
        "no_blanks" | "no blanks" | "not_blank" | "not blank" => {
            add_cell_cf!(
                worksheet,
                data_start_row,
                col_idx,
                data_end_row,
                ConditionalFormatBlank::new().invert(),
                fmt,
                "no_blanks format"
            );
            return Ok(());
        }
        _ => {}
    }

    // Text rules (require string 'value')
    if let Some(rule) = match criteria_lower.as_str() {
        "containing" | "contains" | "text_contains" => Some(ConditionalFormatTextRule::Contains(
            cf_required_string(py, config, col_pattern, "value")?,
        )),
        "not_containing" | "not containing" | "does_not_contain" | "does not contain" => {
            Some(ConditionalFormatTextRule::DoesNotContain(
                cf_required_string(py, config, col_pattern, "value")?,
            ))
        }
        "begins_with" | "begins with" | "starts_with" | "starts with" => {
            Some(ConditionalFormatTextRule::BeginsWith(cf_required_string(
                py,
                config,
                col_pattern,
                "value",
            )?))
        }
        "ends_with" | "ends with" => Some(ConditionalFormatTextRule::EndsWith(cf_required_string(
            py,
            config,
            col_pattern,
            "value",
        )?)),
        _ => None,
    } {
        add_cell_cf!(
            worksheet,
            data_start_row,
            col_idx,
            data_end_row,
            ConditionalFormatText::new().set_rule(rule),
            fmt,
            "text rule format"
        );
        return Ok(());
    }

    // Range rules (require numeric min_value/max_value)
    if let Some(rule) = match criteria_lower.as_str() {
        "between" => Some(ConditionalFormatCellRule::Between(
            cf_required_f64(py, config, col_pattern, "min_value")?,
            cf_required_f64(py, config, col_pattern, "max_value")?,
        )),
        "not_between" | "not between" => Some(ConditionalFormatCellRule::NotBetween(
            cf_required_f64(py, config, col_pattern, "min_value")?,
            cf_required_f64(py, config, col_pattern, "max_value")?,
        )),
        _ => None,
    } {
        add_cell_cf!(
            worksheet,
            data_start_row,
            col_idx,
            data_end_row,
            ConditionalFormatCell::new().set_rule(rule),
            fmt,
            "range rule format"
        );
        return Ok(());
    }

    // Single-value comparison rules. Preserve numeric vs string in the Excel rule
    // so "ERROR" matches as text and 100 matches as number.
    let value_obj = config.get("value").ok_or_else(|| {
        format!(
            "conditional_formats['{}']: missing 'value' key",
            col_pattern
        )
    })?;
    let bound = value_obj.bind(py);

    macro_rules! make_rule {
        ($variant:ident) => {
            if let Ok(v) = bound.extract::<f64>() {
                ConditionalFormatCell::new().set_rule(ConditionalFormatCellRule::$variant(v))
            } else if let Ok(s) = bound.extract::<String>() {
                ConditionalFormatCell::new().set_rule(ConditionalFormatCellRule::$variant(s))
            } else {
                return Err(format!(
                    "conditional_formats['{}']: 'value' must be a string or number",
                    col_pattern
                ));
            }
        };
    }

    let cf = match criteria_lower.as_str() {
        "equal_to" | "equal to" | "==" | "eq" => make_rule!(EqualTo),
        "not_equal_to" | "not equal to" | "!=" | "ne" => make_rule!(NotEqualTo),
        "greater_than" | "greater than" | ">" | "gt" => make_rule!(GreaterThan),
        "less_than" | "less than" | "<" | "lt" => make_rule!(LessThan),
        "greater_than_or_equal_to" | "greater than or equal to" | ">=" | "gte" => {
            make_rule!(GreaterThanOrEqualTo)
        }
        "less_than_or_equal_to" | "less than or equal to" | "<=" | "lte" => {
            make_rule!(LessThanOrEqualTo)
        }
        _ => {
            return Err(format!(
                "Unknown criteria '{}'. Valid: equal_to, not_equal_to, \
                 greater_than, less_than, greater_than_or_equal_to, \
                 less_than_or_equal_to, between, not_between, \
                 containing, not_containing, begins_with, ends_with, \
                 blanks, no_blanks",
                criteria
            ));
        }
    };

    add_cell_cf!(
        worksheet,
        data_start_row,
        col_idx,
        data_end_row,
        cf,
        fmt,
        "cell format"
    );
    Ok(())
}

/// Apply a single conditional format config to a column range.
/// Dispatches by `type` to a family-specific helper.
fn apply_single_conditional_format(
    py: Python<'_>,
    worksheet: &mut Worksheet,
    col_pattern: &str,
    config: &HashMap<String, Py<PyAny>>,
    col_idx: u16,
    data_start_row: u32,
    data_end_row: u32,
) -> Result<(), String> {
    let format_type: String = config
        .get("type")
        .ok_or_else(|| format!("conditional_formats['{}']: missing 'type' key", col_pattern))?
        .bind(py)
        .extract()
        .map_err(|e| {
            format!(
                "conditional_formats['{}']: invalid 'type': {}",
                col_pattern, e
            )
        })?;

    match format_type.to_lowercase().as_str() {
        "2_color_scale" | "2colorscale" | "two_color_scale" => apply_2_color_scale(
            py,
            worksheet,
            col_pattern,
            config,
            col_idx,
            data_start_row,
            data_end_row,
        ),
        "3_color_scale" | "3colorscale" | "three_color_scale" => apply_3_color_scale(
            py,
            worksheet,
            col_pattern,
            config,
            col_idx,
            data_start_row,
            data_end_row,
        ),
        "data_bar" | "databar" => apply_data_bar(
            py,
            worksheet,
            col_pattern,
            config,
            col_idx,
            data_start_row,
            data_end_row,
        ),
        "icon_set" | "iconset" => apply_icon_set(
            py,
            worksheet,
            col_pattern,
            config,
            col_idx,
            data_start_row,
            data_end_row,
        ),
        "cell" => apply_cell_conditional(
            py,
            worksheet,
            col_pattern,
            config,
            col_idx,
            data_start_row,
            data_end_row,
        ),
        _ => Err(format!(
            "Unknown conditional format type '{}'. Valid types: \
             2_color_scale, 3_color_scale, data_bar, icon_set, cell",
            format_type
        )),
    }
}

/// Apply conditional formats to a worksheet
/// Supports: 2_color_scale, 3_color_scale, data_bar, icon_set, cell
/// Uses IndexMap to preserve pattern order (first match wins for overlapping patterns)
pub(crate) fn apply_conditional_formats(
    py: Python<'_>,
    worksheet: &mut Worksheet,
    columns: &[String],
    data_start_row: u32,
    data_end_row: u32,
    cond_formats: &ConditionalFormatConfigs,
) -> Result<(), String> {
    for (col_pattern, configs) in cond_formats {
        let col_indices: Vec<u16> = columns
            .iter()
            .enumerate()
            .filter(|(_, name)| matches_pattern(name, col_pattern))
            .map(|(idx, _)| idx as u16) // safe: col_count already validated via u16::try_from
            .collect();

        if col_indices.is_empty() {
            continue;
        }

        for config in configs {
            for &col_idx in &col_indices {
                apply_single_conditional_format(
                    py,
                    worksheet,
                    col_pattern,
                    config,
                    col_idx,
                    data_start_row,
                    data_end_row,
                )?;
            }
        }
    }

    Ok(())
}
