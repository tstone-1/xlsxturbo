//! Conditional formatting application helpers.

use crate::parse::{matches_pattern, parse_color, parse_column_format, parse_icon_type};
use crate::types::{ConditionalFormatConfigs, OptionMap};
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
fn cf_optional_color(view: &OptionMap<'_, '_>, key: &str) -> Result<Option<u32>, String> {
    let Some(color_str) = view.string(key)? else {
        return Ok(None);
    };
    parse_color(&color_str).map(Some)
}

/// Parse the optional `format` dict on a cell-rule conditional format config.
fn parse_cf_format(view: &OptionMap<'_, '_>) -> Result<Option<Format>, String> {
    match view.dict("format")? {
        Some(map) => Ok(Some(parse_column_format(view.py(), &map, view.context())?)),
        None => Ok(None),
    }
}

/// Apply a 2-color-scale conditional format.
fn apply_2_color_scale(
    view: &OptionMap<'_, '_>,
    worksheet: &mut Worksheet,
    col_idx: u16,
    data_start_row: u32,
    data_end_row: u32,
) -> Result<(), String> {
    let mut cf = ConditionalFormat2ColorScale::new();
    if let Some(c) = cf_optional_color(view, "min_color")? {
        cf = cf.set_minimum_color(c);
    }
    if let Some(c) = cf_optional_color(view, "max_color")? {
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
    view: &OptionMap<'_, '_>,
    worksheet: &mut Worksheet,
    col_idx: u16,
    data_start_row: u32,
    data_end_row: u32,
) -> Result<(), String> {
    let mut cf = ConditionalFormat3ColorScale::new();
    if let Some(c) = cf_optional_color(view, "min_color")? {
        cf = cf.set_minimum_color(c);
    }
    if let Some(c) = cf_optional_color(view, "mid_color")? {
        cf = cf.set_midpoint_color(c);
    }
    if let Some(c) = cf_optional_color(view, "max_color")? {
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
    view: &OptionMap<'_, '_>,
    worksheet: &mut Worksheet,
    col_idx: u16,
    data_start_row: u32,
    data_end_row: u32,
) -> Result<(), String> {
    let mut cf = ConditionalFormatDataBar::new();
    if let Some(c) = cf_optional_color(view, "bar_color")? {
        cf = cf.set_fill_color(c);
    }
    if let Some(c) = cf_optional_color(view, "border_color")? {
        cf = cf.set_border_color(c);
    }
    if view.bool("solid")?.unwrap_or(false) {
        cf = cf.set_solid_fill(true);
    }
    if let Some(s) = view.string("direction")? {
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
    view: &OptionMap<'_, '_>,
    worksheet: &mut Worksheet,
    col_idx: u16,
    data_start_row: u32,
    data_end_row: u32,
) -> Result<(), String> {
    let mut cf = ConditionalFormatIconSet::new();
    if let Some(s) = view.string("icon_type")? {
        cf = cf.set_icon_type(parse_icon_type(&s)?);
    }
    if view.bool("reverse")?.unwrap_or(false) {
        cf = cf.reverse_icons(true);
    }
    if view.bool("icons_only")?.unwrap_or(false) {
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
    view: &OptionMap<'_, '_>,
    worksheet: &mut Worksheet,
    col_pattern: &str,
    col_idx: u16,
    data_start_row: u32,
    data_end_row: u32,
) -> Result<(), String> {
    let criteria: String = view
        .get("criteria")
        .ok_or_else(|| {
            format!(
                "conditional_formats['{}']: 'cell' type requires 'criteria' key",
                col_pattern
            )
        })?
        .bind(view.py())
        .extract()
        .map_err(|e| {
            format!(
                "conditional_formats['{}']: invalid 'criteria': {}",
                col_pattern, e
            )
        })?;

    let fmt = parse_cf_format(view)?;
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
            view.required_string("value")?,
        )),
        "not_containing" | "not containing" | "does_not_contain" | "does not contain" => Some(
            ConditionalFormatTextRule::DoesNotContain(view.required_string("value")?),
        ),
        "begins_with" | "begins with" | "starts_with" | "starts with" => Some(
            ConditionalFormatTextRule::BeginsWith(view.required_string("value")?),
        ),
        "ends_with" | "ends with" => Some(ConditionalFormatTextRule::EndsWith(
            view.required_string("value")?,
        )),
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
            view.required_f64("min_value")?,
            view.required_f64("max_value")?,
        )),
        "not_between" | "not between" => Some(ConditionalFormatCellRule::NotBetween(
            view.required_f64("min_value")?,
            view.required_f64("max_value")?,
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
    let value_obj = view.get("value").ok_or_else(|| {
        format!(
            "conditional_formats['{}']: missing 'value' key",
            col_pattern
        )
    })?;
    let bound = value_obj.bind(view.py());

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
/// Dispatches by `type` to a family-specific helper. Unlike `validations`,
/// conditional format configs are dispatched to one of five type-specific
/// families with disjoint key sets, so the valid-keys list — and thus the
/// unknown-key error — is per-type; the resolved format type is passed to
/// `OptionMap::reject_unknown_for` as its qualifier.
fn apply_single_conditional_format(
    py: Python<'_>,
    worksheet: &mut Worksheet,
    col_pattern: &str,
    config: &HashMap<String, Py<PyAny>>,
    col_idx: u16,
    data_start_row: u32,
    data_end_row: u32,
) -> Result<(), String> {
    let view = OptionMap::new(
        py,
        config,
        format!("conditional_formats['{}']", col_pattern),
    );

    let format_type: String = view
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
        "2_color_scale" | "2colorscale" | "two_color_scale" => {
            view.reject_unknown_for(&format_type, &["type", "min_color", "max_color"])?;
            apply_2_color_scale(&view, worksheet, col_idx, data_start_row, data_end_row)
        }
        "3_color_scale" | "3colorscale" | "three_color_scale" => {
            view.reject_unknown_for(
                &format_type,
                &["type", "min_color", "mid_color", "max_color"],
            )?;
            apply_3_color_scale(&view, worksheet, col_idx, data_start_row, data_end_row)
        }
        "data_bar" | "databar" => {
            view.reject_unknown_for(
                &format_type,
                &["type", "bar_color", "border_color", "solid", "direction"],
            )?;
            apply_data_bar(&view, worksheet, col_idx, data_start_row, data_end_row)
        }
        "icon_set" | "iconset" => {
            view.reject_unknown_for(
                &format_type,
                &["type", "icon_type", "reverse", "icons_only"],
            )?;
            apply_icon_set(&view, worksheet, col_idx, data_start_row, data_end_row)
        }
        "cell" => {
            view.reject_unknown_for(
                &format_type,
                &[
                    "type",
                    "criteria",
                    "format",
                    "value",
                    "min_value",
                    "max_value",
                ],
            )?;
            apply_cell_conditional(
                &view,
                worksheet,
                col_pattern,
                col_idx,
                data_start_row,
                data_end_row,
            )
        }
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
            return Err(format!(
                "conditional_formats['{}']: pattern matched no columns",
                col_pattern
            ));
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

// Unknown-key rejection policy (including the qualifier/"Valid for <type>"
// phrasing) is unit-tested once, without a Python interpreter, in
// `types::reject_unknown_keys_tests` — the single source of truth this
// module's `OptionMap::reject_unknown_for` calls delegate to.
