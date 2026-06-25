//! Native Excel sparkline application helpers.
//!
//! Sparklines are mini in-cell charts. A location key that is a single cell
//! (e.g. `"D2"`) places one sparkline via [`Worksheet::add_sparkline`]; a range
//! key (e.g. `"D2:D10"`) places a grouped sparkline — one per row of the data
//! range — via [`Worksheet::add_sparkline_group`].

use crate::parse::{parse_cell_range, parse_cell_ref, parse_color_enum};
use crate::types::{extract_opt, SparklineConfig};
use pyo3::prelude::*;
use rust_xlsxwriter::{Sparkline, SparklineType, Worksheet};
use std::collections::HashMap;

const SPARKLINE_KEYS: &[&str] = &[
    "range",
    "type",
    "style",
    "markers",
    "high_point",
    "low_point",
    "first_point",
    "last_point",
    "negative_points",
    "show_axis",
    "show_hidden_data",
    "group_max",
    "group_min",
    "right_to_left",
    "column_order",
    "color",
    "high_point_color",
    "low_point_color",
    "first_point_color",
    "last_point_color",
    "negative_points_color",
    "markers_color",
    "line_weight",
    "custom_max",
    "custom_min",
    "date_range",
];

fn parse_sparkline_type(sparkline_type: &str) -> Result<SparklineType, String> {
    match sparkline_type.to_lowercase().as_str() {
        "line" => Ok(SparklineType::Line),
        "column" | "col" => Ok(SparklineType::Column),
        "win_loss" | "win_lose" | "winloss" | "winlose" => Ok(SparklineType::WinLose),
        _ => Err(format!(
            "Unknown sparkline type '{}'. Valid: line, column, win_loss",
            sparkline_type
        )),
    }
}

fn spark_string(
    py: Python<'_>,
    opts: &HashMap<String, Py<PyAny>>,
    loc: &str,
    key: &str,
) -> Result<Option<String>, String> {
    extract_opt(py, opts.get(key), |_| {
        format!("sparklines['{}']: '{}' must be a string", loc, key)
    })
}

fn spark_bool(
    py: Python<'_>,
    opts: &HashMap<String, Py<PyAny>>,
    loc: &str,
    key: &str,
) -> Result<Option<bool>, String> {
    extract_opt(py, opts.get(key), |_| {
        format!("sparklines['{}']: '{}' must be a bool", loc, key)
    })
}

fn spark_u8(
    py: Python<'_>,
    opts: &HashMap<String, Py<PyAny>>,
    loc: &str,
    key: &str,
) -> Result<Option<u8>, String> {
    extract_opt(py, opts.get(key), |_| {
        format!("sparklines['{}']: '{}' must be an integer 0-255", loc, key)
    })
}

fn spark_f64(
    py: Python<'_>,
    opts: &HashMap<String, Py<PyAny>>,
    loc: &str,
    key: &str,
) -> Result<Option<f64>, String> {
    extract_opt(py, opts.get(key), |_| {
        format!("sparklines['{}']: '{}' must be a number", loc, key)
    })
}

fn build_sparkline(
    py: Python<'_>,
    loc: &str,
    config: &HashMap<String, Py<PyAny>>,
) -> Result<Sparkline, String> {
    for key in config.keys() {
        if !SPARKLINE_KEYS.contains(&key.as_str()) {
            return Err(format!(
                "sparklines['{}']: unknown option '{}'. Valid: {}",
                loc,
                key,
                SPARKLINE_KEYS.join(", ")
            ));
        }
    }

    let range = spark_string(py, config, loc, "range")?
        .ok_or_else(|| format!("sparklines['{}']: missing required 'range' key", loc))?;
    // rust_xlsxwriter requires a sheet-qualified range string (like charts);
    // a bare range silently yields an empty data range, so reject it early.
    if !range.contains('!') {
        return Err(format!(
            "sparklines['{}']: 'range' must include a sheet name, e.g. 'Sheet1!{}'",
            loc, range
        ));
    }
    let mut sparkline = Sparkline::new().set_range(range.as_str());

    if let Some(sparkline_type) = spark_string(py, config, loc, "type")? {
        let parsed = parse_sparkline_type(&sparkline_type)
            .map_err(|e| format!("sparklines['{}']: {}", loc, e))?;
        sparkline = sparkline.set_type(parsed);
    }
    if let Some(style) = spark_u8(py, config, loc, "style")? {
        if !(1..=36).contains(&style) {
            return Err(format!(
                "sparklines['{}']: 'style' must be in the range 1-36, got {}",
                loc, style
            ));
        }
        sparkline = sparkline.set_style(style);
    }

    // Boolean display toggles. Local macro keeps the builder reassignment DRY
    // without obscuring which key maps to which rust_xlsxwriter setter.
    macro_rules! apply_bool {
        ($key:literal, $method:ident) => {
            if let Some(enable) = spark_bool(py, config, loc, $key)? {
                sparkline = sparkline.$method(enable);
            }
        };
    }
    apply_bool!("markers", show_markers);
    apply_bool!("high_point", show_high_point);
    apply_bool!("low_point", show_low_point);
    apply_bool!("first_point", show_first_point);
    apply_bool!("last_point", show_last_point);
    apply_bool!("negative_points", show_negative_points);
    apply_bool!("show_axis", show_axis);
    apply_bool!("show_hidden_data", show_hidden_data);
    apply_bool!("group_max", set_group_max);
    apply_bool!("group_min", set_group_min);
    apply_bool!("right_to_left", set_right_to_left);
    apply_bool!("column_order", set_column_order);

    // Color setters all take `impl Into<Color>`.
    macro_rules! apply_color {
        ($key:literal, $method:ident) => {
            if let Some(color_str) = spark_string(py, config, loc, $key)? {
                let color = parse_color_enum(&color_str)
                    .map_err(|e| format!("sparklines['{}']: '{}': {}", loc, $key, e))?;
                sparkline = sparkline.$method(color);
            }
        };
    }
    apply_color!("color", set_sparkline_color);
    apply_color!("high_point_color", set_high_point_color);
    apply_color!("low_point_color", set_low_point_color);
    apply_color!("first_point_color", set_first_point_color);
    apply_color!("last_point_color", set_last_point_color);
    apply_color!("negative_points_color", set_negative_points_color);
    apply_color!("markers_color", set_markers_color);

    if let Some(weight) = spark_f64(py, config, loc, "line_weight")? {
        sparkline = sparkline.set_line_weight(weight);
    }
    if let Some(max) = spark_f64(py, config, loc, "custom_max")? {
        sparkline = sparkline.set_custom_max(max);
    }
    if let Some(min) = spark_f64(py, config, loc, "custom_min")? {
        sparkline = sparkline.set_custom_min(min);
    }
    if let Some(date_range) = spark_string(py, config, loc, "date_range")? {
        if !date_range.contains('!') {
            return Err(format!(
                "sparklines['{}']: 'date_range' must include a sheet name, e.g. 'Sheet1!{}'",
                loc, date_range
            ));
        }
        sparkline = sparkline.set_date_range(date_range.as_str());
    }

    Ok(sparkline)
}

/// Apply native Excel sparklines to a worksheet.
pub(crate) fn apply_sparklines(
    py: Python<'_>,
    worksheet: &mut Worksheet,
    sparklines: &HashMap<String, SparklineConfig>,
) -> Result<(), String> {
    for (loc, config) in sparklines {
        let sparkline = build_sparkline(py, loc, config)?;

        if loc.contains(':') {
            let (first_row, first_col, last_row, last_col) = parse_cell_range(loc)?;
            // A grouped sparkline location must be a single row or single column
            // (one sparkline per cell). A 2D block is ambiguous, so reject it
            // rather than let rust_xlsxwriter place sparklines unexpectedly.
            if first_row != last_row && first_col != last_col {
                return Err(format!(
                    "sparklines['{}']: a grouped location must be a single row or column, not a 2D block",
                    loc
                ));
            }
            worksheet
                .add_sparkline_group(first_row, first_col, last_row, last_col, &sparkline)
                .map_err(|e| format!("sparklines['{}']: {}", loc, e))?;
        } else {
            let (row, col) = parse_cell_ref(loc)?;
            worksheet
                .add_sparkline(row, col, &sparkline)
                .map_err(|e| format!("sparklines['{}']: {}", loc, e))?;
        }
    }

    Ok(())
}
