//! Conditional formatting application helpers.

use crate::extract::pydict_to_hashmap;
use crate::parse::{matches_pattern, parse_color, parse_column_format, parse_icon_type};
use crate::types::ConditionalFormatConfigs;
use pyo3::prelude::*;
use rust_xlsxwriter::{
    ConditionalFormat2ColorScale, ConditionalFormat3ColorScale, ConditionalFormatBlank,
    ConditionalFormatCell, ConditionalFormatCellRule, ConditionalFormatDataBar,
    ConditionalFormatDataBarDirection, ConditionalFormatIconSet, ConditionalFormatText,
    ConditionalFormatTextRule, Format, Worksheet,
};
use std::collections::HashMap;

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
