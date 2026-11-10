//! Conditional formatting application helpers.

use crate::parse::{
    matches_pattern, parse_color, parse_column_format, parse_icon_type, WithOptionContext,
};
use crate::types::{pytype_name, ConditionalFormatConfigs, OptionMap};
use pyo3::prelude::*;
use rust_xlsxwriter::{
    ConditionalFormat, ConditionalFormat2ColorScale, ConditionalFormat3ColorScale,
    ConditionalFormatBlank, ConditionalFormatCell, ConditionalFormatCellRule,
    ConditionalFormatDataBar, ConditionalFormatDataBarDirection, ConditionalFormatIconSet,
    ConditionalFormatText, ConditionalFormatTextRule, ConditionalFormatType, Format, Worksheet,
};
use std::collections::HashMap;

/// One conditional format, built and checked but not yet on a worksheet.
///
/// Building is kept apart from writing so that a frame without data rows still
/// has its configs checked: before, both sat behind one "has data rows" test,
/// and a typo in a scheduled export passed on every empty day and failed on
/// the first day with data. An enum rather than a boxed trait object, because
/// `Worksheet::add_conditional_format` takes a sized type.
enum BuiltFormat {
    TwoColorScale(ConditionalFormat2ColorScale),
    ThreeColorScale(ConditionalFormat3ColorScale),
    DataBar(ConditionalFormatDataBar),
    IconSet(ConditionalFormatIconSet),
    Blank(ConditionalFormatBlank, &'static str),
    Text(ConditionalFormatText),
    Range(ConditionalFormatCell),
    Cell(ConditionalFormatCell),
}

impl BuiltFormat {
    /// The name used in the "Failed to add ..." message.
    fn what(&self) -> &'static str {
        match self {
            Self::TwoColorScale(_) => "2_color_scale",
            Self::ThreeColorScale(_) => "3_color_scale",
            Self::DataBar(_) => "data_bar",
            Self::IconSet(_) => "icon_set",
            Self::Blank(_, what) => what,
            Self::Text(_) => "text rule format",
            Self::Range(_) => "range rule format",
            Self::Cell(_) => "cell format",
        }
    }

    /// Run the crate's own check, which `add_conditional_format` would run
    /// at write time, so that a frame with no rows refuses the same configs.
    fn validate(&self) -> Result<(), String> {
        let result = match self {
            Self::TwoColorScale(cf) => cf.validate(),
            Self::ThreeColorScale(cf) => cf.validate(),
            Self::DataBar(cf) => cf.validate(),
            Self::IconSet(cf) => cf.validate(),
            Self::Blank(cf, _) => cf.validate(),
            Self::Text(cf) => cf.validate(),
            Self::Range(cf) | Self::Cell(cf) => cf.validate(),
        };
        result.map_err(|e| format!("Failed to add {}: {}", self.what(), e))
    }

    /// Add the format to one column's data rows.
    fn add(
        &self,
        worksheet: &mut Worksheet,
        col_idx: u16,
        data_start_row: u32,
        data_end_row: u32,
    ) -> Result<(), String> {
        let (r0, r1, c) = (data_start_row, data_end_row, col_idx);
        let result = match self {
            Self::TwoColorScale(cf) => worksheet.add_conditional_format(r0, c, r1, c, cf),
            Self::ThreeColorScale(cf) => worksheet.add_conditional_format(r0, c, r1, c, cf),
            Self::DataBar(cf) => worksheet.add_conditional_format(r0, c, r1, c, cf),
            Self::IconSet(cf) => worksheet.add_conditional_format(r0, c, r1, c, cf),
            Self::Blank(cf, _) => worksheet.add_conditional_format(r0, c, r1, c, cf),
            Self::Text(cf) => worksheet.add_conditional_format(r0, c, r1, c, cf),
            Self::Range(cf) | Self::Cell(cf) => worksheet.add_conditional_format(r0, c, r1, c, cf),
        };
        result
            .map(|_| ())
            .map_err(|e| format!("Failed to add {}: {}", self.what(), e))
    }
}

/// Apply the optional user `format` of a cell-rule conditional format.
/// Visualization types (scales, bars, icons) take none: their look is the format.
macro_rules! with_format {
    ($cf:expr, $fmt:expr) => {
        match $fmt {
            Some(f) => $cf.set_format(f),
            None => $cf,
        }
    };
}

/// Extract an optional color field from a conditional format config.
/// None values are treated as unset. Wrong types error with context.
fn cf_optional_color(view: &OptionMap<'_, '_>, key: &str) -> Result<Option<u32>, String> {
    let Some(color_str) = view.string(key)? else {
        return Ok(None);
    };
    parse_color(&color_str)
        .in_field(view.context(), key)
        .map(Some)
}

/// Parse the optional `format` dict on a cell-rule conditional format config.
fn parse_cf_format(view: &OptionMap<'_, '_>) -> Result<Option<Format>, String> {
    match view.dict("format")? {
        Some(map) => {
            // Differential styles cannot carry these base-cell properties. An
            // explicit None means "absent", as it does for every other key.
            for key in [
                "font_name",
                "quote_prefix",
                "font_size",
                "align_horizontal",
                "align_vertical",
                "wrap_text",
            ] {
                if map.get(key).is_some_and(|v| !v.bind(view.py()).is_none()) {
                    return Err(format!(
                        "{}: '{}' is not supported in conditional formats; use column_formats instead",
                        view.context(), key
                    ));
                }
            }
            Ok(Some(parse_column_format(view.py(), &map, view.context())?))
        }
        None => Ok(None),
    }
}

/// Read the optional fixed scale of a colour scale or data bar: `min_value`
/// and `max_value`, finite numbers mapped to `ConditionalFormatType::Number`.
///
/// Without them each column scales from its own lowest to its own highest
/// value, so the bars of two columns cannot be compared (issue #67).
///
/// A reversed pair is refused: rust_xlsxwriter writes the two `cfvo` values as
/// given, so `min_value=10, max_value=0` reached the file unchanged, and it is
/// almost always a swapped pair. Equal bounds stay accepted.
fn cf_scale_bounds(view: &OptionMap<'_, '_>) -> Result<(Option<f64>, Option<f64>), String> {
    let min_value = view.finite_f64("min_value")?;
    let max_value = view.finite_f64("max_value")?;
    if let (Some(lo), Some(hi)) = (min_value, max_value) {
        if lo > hi {
            return Err(format!(
                "{}: 'min_value' ({}) must not be greater than 'max_value' ({})",
                view.context(),
                lo,
                hi
            ));
        }
    }
    Ok((min_value, max_value))
}

/// Build a 2-color-scale conditional format.
fn build_2_color_scale(view: &OptionMap<'_, '_>) -> Result<BuiltFormat, String> {
    let mut cf = ConditionalFormat2ColorScale::new();
    let (min_value, max_value) = cf_scale_bounds(view)?;
    if let Some(v) = min_value {
        cf = cf.set_minimum(ConditionalFormatType::Number, v);
    }
    if let Some(v) = max_value {
        cf = cf.set_maximum(ConditionalFormatType::Number, v);
    }
    if let Some(c) = cf_optional_color(view, "min_color")? {
        cf = cf.set_minimum_color(c);
    }
    if let Some(c) = cf_optional_color(view, "max_color")? {
        cf = cf.set_maximum_color(c);
    }
    Ok(BuiltFormat::TwoColorScale(cf))
}

/// Build a 3-color-scale conditional format.
fn build_3_color_scale(view: &OptionMap<'_, '_>) -> Result<BuiltFormat, String> {
    let mut cf = ConditionalFormat3ColorScale::new();
    let (min_value, max_value) = cf_scale_bounds(view)?;
    if let Some(v) = min_value {
        cf = cf.set_minimum(ConditionalFormatType::Number, v);
    }
    if let Some(v) = max_value {
        cf = cf.set_maximum(ConditionalFormatType::Number, v);
    }
    if let Some(c) = cf_optional_color(view, "min_color")? {
        cf = cf.set_minimum_color(c);
    }
    if let Some(c) = cf_optional_color(view, "mid_color")? {
        cf = cf.set_midpoint_color(c);
    }
    if let Some(c) = cf_optional_color(view, "max_color")? {
        cf = cf.set_maximum_color(c);
    }
    Ok(BuiltFormat::ThreeColorScale(cf))
}

/// Build a data-bar conditional format.
fn build_data_bar(view: &OptionMap<'_, '_>) -> Result<BuiltFormat, String> {
    let mut cf = ConditionalFormatDataBar::new();
    let (min_value, max_value) = cf_scale_bounds(view)?;
    if let Some(v) = min_value {
        cf = cf.set_minimum(ConditionalFormatType::Number, v);
    }
    if let Some(v) = max_value {
        cf = cf.set_maximum(ConditionalFormatType::Number, v);
    }
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
                    "{}: 'direction': Unknown direction '{}'. \
                     Valid: left_to_right, right_to_left, context",
                    view.context(),
                    s
                ));
            }
        };
        cf = cf.set_direction(dir);
    }
    Ok(BuiltFormat::DataBar(cf))
}

/// Build an icon-set conditional format.
fn build_icon_set(view: &OptionMap<'_, '_>) -> Result<BuiltFormat, String> {
    let mut cf = ConditionalFormatIconSet::new();
    if let Some(s) = view.string("icon_type")? {
        cf = cf.set_icon_type(parse_icon_type(&s).in_field(view.context(), "icon_type")?);
    }
    if view.bool("reverse")?.unwrap_or(false) {
        cf = cf.reverse_icons(true);
    }
    if view.bool("icons_only")?.unwrap_or(false) {
        cf = cf.show_icons_only(true);
    }
    Ok(BuiltFormat::IconSet(cf))
}

/// Build a cell-rule conditional format, dispatching by criteria family.
fn build_cell_conditional(
    view: &OptionMap<'_, '_>,
    col_pattern: &str,
) -> Result<BuiltFormat, String> {
    // `string()` rather than a hand-rolled `extract()`: a raw `PyErr` flattened
    // into this layer's `String` channel renders as "<class>: <message>", so the
    // caller reads a `TypeError:` prefix inside a `ConfigurationError`. The
    // missing-key wording stays bespoke -- it names the format type that
    // requires the key, which the generic "missing 'criteria' key" would drop.
    let criteria: String = view.string("criteria")?.ok_or_else(|| {
        format!(
            "conditional_formats['{}']: 'cell' type requires 'criteria' key",
            col_pattern
        )
    })?;

    let fmt = parse_cf_format(view)?;
    let criteria_lower = criteria.to_lowercase();

    // Blank / no-blank
    match criteria_lower.as_str() {
        "blanks" | "blank" => {
            return Ok(BuiltFormat::Blank(
                with_format!(ConditionalFormatBlank::new(), fmt),
                "blanks format",
            ));
        }
        "no_blanks" | "no blanks" | "not_blank" | "not blank" => {
            return Ok(BuiltFormat::Blank(
                with_format!(ConditionalFormatBlank::new().invert(), fmt),
                "no_blanks format",
            ));
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
        return Ok(BuiltFormat::Text(with_format!(
            ConditionalFormatText::new().set_rule(rule),
            fmt
        )));
    }

    // Range rules (require numeric min_value/max_value)
    if let Some(rule) = match criteria_lower.as_str() {
        "between" => Some(ConditionalFormatCellRule::Between(
            view.required_finite_f64("min_value")?,
            view.required_finite_f64("max_value")?,
        )),
        "not_between" | "not between" => Some(ConditionalFormatCellRule::NotBetween(
            view.required_finite_f64("min_value")?,
            view.required_finite_f64("max_value")?,
        )),
        _ => None,
    } {
        return Ok(BuiltFormat::Range(with_format!(
            ConditionalFormatCell::new().set_rule(rule),
            fmt
        )));
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
    // A Python bool extracts as f64, so `value=True` used to become the rule
    // `= 1`. Refused, as `OptionMap`'s numeric accessors refuse it for
    // `min_value`/`max_value` on the same dict.
    let value_type_error = || {
        format!(
            "{}: 'value' must be a string or number, got {}",
            view.context(),
            pytype_name(bound)
        )
    };
    if bound.is_instance_of::<pyo3::types::PyBool>() {
        return Err(value_type_error());
    }

    macro_rules! make_rule {
        ($variant:ident) => {
            if let Ok(v) = bound.extract::<f64>() {
                if !v.is_finite() {
                    return Err(format!(
                        "{}: 'value' must be a finite number, got {}",
                        view.context(),
                        v
                    ));
                }
                ConditionalFormatCell::new().set_rule(ConditionalFormatCellRule::$variant(v))
            } else if let Ok(s) = bound.extract::<String>() {
                ConditionalFormatCell::new().set_rule(ConditionalFormatCellRule::$variant(s))
            } else {
                return Err(value_type_error());
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
                "{}: 'criteria': Unknown criteria '{}'. Valid: equal_to, not_equal_to, \
                 greater_than, less_than, greater_than_or_equal_to, \
                 less_than_or_equal_to, between, not_between, \
                 containing, not_containing, begins_with, ends_with, \
                 blanks, no_blanks",
                view.context(),
                criteria
            ));
        }
    };

    Ok(BuiltFormat::Cell(with_format!(cf, fmt)))
}

/// Build and check a single conditional format config.
/// Dispatches by `type` to a family-specific helper. Unlike `validations`,
/// conditional format configs are dispatched to one of five type-specific
/// families with disjoint key sets, so the valid-keys list — and thus the
/// unknown-key error — is per-type; the resolved format type is passed to
/// `OptionMap::reject_unknown_for` as its qualifier.
fn build_single_conditional_format(
    py: Python<'_>,
    col_pattern: &str,
    config: &HashMap<String, Py<PyAny>>,
) -> Result<BuiltFormat, String> {
    let view = OptionMap::new(
        py,
        config,
        format!("conditional_formats['{}']", col_pattern),
    );

    // Through the `OptionMap` accessor, not a hand-rolled `extract()`: a raw
    // `PyErr` flattened into this layer's `String` channel carries its own class
    // name into the message of whatever the boundary raises. `required_string`
    // produces the same missing-key wording this site used before.
    let format_type: String = view.required_string("type")?;

    let built = match format_type.to_lowercase().as_str() {
        "2_color_scale" | "2colorscale" | "two_color_scale" => {
            view.reject_unknown_for(
                &format_type,
                &["type", "min_color", "max_color", "min_value", "max_value"],
            )?;
            build_2_color_scale(&view)
        }
        "3_color_scale" | "3colorscale" | "three_color_scale" => {
            view.reject_unknown_for(
                &format_type,
                &[
                    "type",
                    "min_color",
                    "mid_color",
                    "max_color",
                    "min_value",
                    "max_value",
                ],
            )?;
            build_3_color_scale(&view)
        }
        "data_bar" | "databar" => {
            view.reject_unknown_for(
                &format_type,
                &[
                    "type",
                    "bar_color",
                    "border_color",
                    "solid",
                    "direction",
                    "min_value",
                    "max_value",
                ],
            )?;
            build_data_bar(&view)
        }
        "icon_set" | "iconset" => {
            view.reject_unknown_for(
                &format_type,
                &["type", "icon_type", "reverse", "icons_only"],
            )?;
            build_icon_set(&view)
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
            build_cell_conditional(&view, col_pattern)
        }
        _ => Err(format!(
            "{}: 'type': Unknown conditional format type '{}'. Valid types: \
             2_color_scale, 3_color_scale, data_bar, icon_set, cell",
            view.context(),
            format_type
        )),
    }?;
    built.validate()?;
    Ok(built)
}

/// Apply conditional formats to a worksheet
/// Supports: 2_color_scale, 3_color_scale, data_bar, icon_set, cell
/// Uses IndexMap to preserve pattern order (first match wins for overlapping patterns)
///
/// `data_rows` is the inclusive `(first, last)` range of data rows, or `None`
/// for a frame without any: every config is then still matched against the
/// columns and checked, and only the write is skipped.
pub(crate) fn apply_conditional_formats(
    py: Python<'_>,
    worksheet: &mut Worksheet,
    columns: &[String],
    data_rows: Option<(u32, u32)>,
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

        // Built once per config, as validations are: the format is the same
        // for every column the pattern matches, only the range differs. The
        // loop order (config, then column) is the order rules are written in,
        // and so their priority.
        for config in configs {
            let built = build_single_conditional_format(py, col_pattern, config)?;
            let Some((data_start_row, data_end_row)) = data_rows else {
                continue;
            };
            for &col_idx in &col_indices {
                built.add(worksheet, col_idx, data_start_row, data_end_row)?;
            }
        }
    }

    Ok(())
}

// Unknown-key rejection policy (including the qualifier/"Valid for <type>"
// phrasing) is unit-tested once, without a Python interpreter, in
// `types::reject_unknown_keys_tests` — the single source of truth this
// module's `OptionMap::reject_unknown_for` calls delegate to.
