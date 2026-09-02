//! Native Excel chart application helpers.

use crate::parse::{parse_cell_ref, WithOptionContext};
use crate::types::{pydict_to_hashmap_str, ChartConfig, OptionMap};
use indexmap::IndexMap;
use pyo3::prelude::*;
use rust_xlsxwriter::{Chart, ChartDataTable, ChartLegendPosition, ChartType, Worksheet};

fn parse_chart_type(chart_type: &str) -> Result<ChartType, String> {
    match chart_type.to_lowercase().as_str() {
        "area" => Ok(ChartType::Area),
        "area_stacked" | "stacked_area" => Ok(ChartType::AreaStacked),
        "area_percent_stacked" | "percent_stacked_area" => Ok(ChartType::AreaPercentStacked),
        "bar" => Ok(ChartType::Bar),
        "bar_stacked" | "stacked_bar" => Ok(ChartType::BarStacked),
        "bar_percent_stacked" | "percent_stacked_bar" => Ok(ChartType::BarPercentStacked),
        "column" | "col" => Ok(ChartType::Column),
        "column_stacked" | "stacked_column" => Ok(ChartType::ColumnStacked),
        "column_percent_stacked" | "percent_stacked_column" => {
            Ok(ChartType::ColumnPercentStacked)
        }
        "doughnut" | "donut" => Ok(ChartType::Doughnut),
        "line" => Ok(ChartType::Line),
        "line_stacked" | "stacked_line" => Ok(ChartType::LineStacked),
        "line_percent_stacked" | "percent_stacked_line" => Ok(ChartType::LinePercentStacked),
        "pie" => Ok(ChartType::Pie),
        "radar" => Ok(ChartType::Radar),
        "radar_with_markers" => Ok(ChartType::RadarWithMarkers),
        "radar_filled" => Ok(ChartType::RadarFilled),
        "scatter" => Ok(ChartType::Scatter),
        "scatter_straight" => Ok(ChartType::ScatterStraight),
        "scatter_straight_with_markers" => Ok(ChartType::ScatterStraightWithMarkers),
        "scatter_smooth" => Ok(ChartType::ScatterSmooth),
        "scatter_smooth_with_markers" => Ok(ChartType::ScatterSmoothWithMarkers),
        "stock" => Ok(ChartType::Stock),
        _ => Err(format!(
            "Unknown chart type '{}'. Valid: area, bar, column, doughnut, line, pie, radar, scatter, stock and stacked variants",
            chart_type
        )),
    }
}

fn parse_legend_position(position: &str) -> Result<ChartLegendPosition, String> {
    match position.to_lowercase().as_str() {
        "right" => Ok(ChartLegendPosition::Right),
        "left" => Ok(ChartLegendPosition::Left),
        "top" => Ok(ChartLegendPosition::Top),
        "bottom" => Ok(ChartLegendPosition::Bottom),
        "top_right" | "topright" => Ok(ChartLegendPosition::TopRight),
        _ => Err(format!(
            "Unknown legend position '{}'. Valid: right, left, top, bottom, top_right",
            position
        )),
    }
}

/// Look up the first of `keys` that is present in `config`, returning the
/// matching key name alongside its string value so callers can report which
/// key a validation error applies to.
///
/// Every key in `keys` is probed — not just up to the first present one —
/// so a malformed alias later in priority order (e.g. `values` holding an
/// int while `values_range` is already a valid string) still raises its type
/// error instead of being silently skipped because an earlier alias won.
fn first_present_string_field(
    view: &OptionMap<'_, '_>,
    keys: &[&'static str],
) -> Result<Option<(&'static str, String)>, String> {
    let mut found: Option<(&'static str, String)> = None;
    for &key in keys {
        let value = view.string(key)?;
        if found.is_none() {
            if let Some(v) = value {
                found = Some((key, v));
            }
        }
    }
    Ok(found)
}

/// rust_xlsxwriter's `ChartRange::new_from_string` leaves the sheet name
/// empty when `value` has no `!`. For a values range that surfaces later as a
/// misleading validate() error; for a categories range it is silently
/// ignored and the chart falls back to default 1..N axis labels. Reject both
/// cases up front with a clear, actionable error (mirrors the sparklines
/// range guard in src/apply/sparklines.rs).
fn require_sheet_qualified(cell_ref: &str, key: &str, value: &str) -> Result<(), String> {
    if value.contains('!') {
        Ok(())
    } else {
        Err(format!(
            "charts['{}']: '{}' must include a sheet name, e.g. 'Sheet1!{}'",
            cell_ref, key, value
        ))
    }
}

fn add_chart_series(
    chart: &mut Chart,
    cell_ref: &str,
    view: &OptionMap<'_, '_>,
    default_categories: Option<&str>,
) -> Result<(), String> {
    let (values_key, values) =
        first_present_string_field(view, &["values_range", "values", "data_range"])?.ok_or_else(
            || {
                format!(
                    "charts['{}']: chart series requires 'values_range', 'values', or 'data_range'",
                    cell_ref
                )
            },
        )?;
    require_sheet_qualified(cell_ref, values_key, &values)?;

    // Categories are optional; only validate sheet-qualification when a
    // categories range is actually present (explicit or via the chart-level
    // default, which was already validated when it was computed).
    let categories = match first_present_string_field(view, &["categories_range", "categories"])? {
        Some((key, v)) => {
            require_sheet_qualified(cell_ref, key, &v)?;
            Some(v)
        }
        None => default_categories.map(str::to_string),
    };

    let name = view.string("name")?.or(view.string("series_name")?);

    let series = chart.add_series().set_values(values.as_str());
    if let Some(cat) = categories {
        series.set_categories(cat.as_str());
    }
    if let Some(series_name) = name {
        series.set_name(series_name.as_str());
    }

    Ok(())
}

/// Apply native Excel charts to worksheet.
pub(crate) fn apply_charts(
    py: Python<'_>,
    worksheet: &mut Worksheet,
    charts: &IndexMap<String, ChartConfig>,
) -> Result<(), String> {
    const CHART_KEYS: &[&str] = &[
        "type",
        "data_range",
        "values_range",
        "values",
        "categories_range",
        "categories",
        "series",
        "series_name",
        "name",
        "title",
        "x_axis_name",
        "y_axis_name",
        "width",
        "height",
        "x_offset",
        "y_offset",
        "style",
        "show_data_table",
        "show_legend",
        "legend_position",
    ];

    const SERIES_KEYS: &[&str] = &[
        "values_range",
        "values",
        "data_range",
        "categories_range",
        "categories",
        "name",
        "series_name",
    ];

    // Chart-level keys that only ever describe a single series. With `series`
    // present each item carries its own, and the chart-level copy was read by
    // nothing -- silently, which is the one thing this file does nowhere else.
    //
    // `categories_range`/`categories` are deliberately absent: those *are* read
    // at chart level alongside `series`, as the documented fallback for items
    // that do not carry their own. So is everything that is not a series
    // concept (title, axes, size, style, legend).
    const CHART_LEVEL_SERIES_KEYS: &[&str] = &[
        "data_range",
        "values_range",
        "values",
        "name",
        "series_name",
    ];

    for (cell_ref, config) in charts {
        let context = format!("charts['{}']", cell_ref);
        let (row, col) = parse_cell_ref(cell_ref).in_option(&context)?;
        let view = OptionMap::new(py, config, context.clone());
        view.reject_unknown(CHART_KEYS)?;

        let chart_type = view
            .string("type")?
            .ok_or_else(|| format!("charts['{}']: missing 'type' key", cell_ref))?;
        let mut chart = Chart::new(parse_chart_type(&chart_type).in_field(&context, "type")?);

        if config
            .get("series")
            .is_some_and(|series_obj| !series_obj.bind(py).is_none())
        {
            // Reject the contradiction before parsing anything: a chart-level
            // range or name beside `series` states one thing and gets another.
            let conflicting: Vec<String> = CHART_LEVEL_SERIES_KEYS
                .iter()
                .filter(|key| {
                    config
                        .get(**key)
                        .is_some_and(|value| !value.bind(py).is_none())
                })
                .map(|key| format!("'{}'", key))
                .collect();
            if !conflicting.is_empty() {
                return Err(format!(
                    "charts['{}']: {} cannot be combined with 'series'; a series-level \
                     range or name belongs inside the series list",
                    cell_ref,
                    conflicting.join(", ")
                ));
            }

            let series_obj = config.get("series").expect("checked above");
            let series_list = series_obj
                .bind(py)
                .cast::<pyo3::types::PyList>()
                .map_err(|_| format!("charts['{}']: 'series' must be a list", cell_ref))?;
            if series_list.is_empty() {
                return Err(format!(
                    "charts['{}']: 'series' must not be empty",
                    cell_ref
                ));
            }

            // Chart-level categories fallback used by any series item that
            // doesn't specify its own; computed once (not per series item)
            // and validated up front since it is reused unchanged.
            let default_categories =
                match first_present_string_field(&view, &["categories_range", "categories"])? {
                    Some((key, v)) => {
                        require_sheet_qualified(cell_ref, key, &v)?;
                        Some(v)
                    }
                    None => None,
                };

            for (idx, item) in series_list.iter().enumerate() {
                let series_dict = item.cast::<pyo3::types::PyDict>().map_err(|_| {
                    format!("charts['{}']: series item {} must be a dict", cell_ref, idx)
                })?;
                let series_config = pydict_to_hashmap_str(
                    series_dict,
                    &format!("charts['{}']: series item {}", cell_ref, idx),
                )?;
                let series_view = OptionMap::new(
                    py,
                    &series_config,
                    format!("charts['{}']: series item {}", cell_ref, idx),
                );
                series_view.reject_unknown(SERIES_KEYS)?;
                add_chart_series(
                    &mut chart,
                    cell_ref,
                    &series_view,
                    default_categories.as_deref(),
                )?;
            }
        } else {
            add_chart_series(&mut chart, cell_ref, &view, None)?;
        }

        if let Some(title) = view.string("title")? {
            chart.title().set_name(title.as_str());
        }
        if let Some(name) = view.string("x_axis_name")? {
            chart.x_axis().set_name(name.as_str());
        }
        if let Some(name) = view.string("y_axis_name")? {
            chart.y_axis().set_name(name.as_str());
        }
        if let Some(width) = view.u32("width")? {
            chart.set_width(width);
        }
        if let Some(height) = view.u32("height")? {
            chart.set_height(height);
        }
        // Read as i64 and range-check rather than letting `u8` decide: a u8
        // silently accepts 0 and 200, which Excel discards (rust_xlsxwriter
        // only complains on stderr, where Python can't see it), and rejects
        // 300 with a "range 0-255" message that has nothing to do with Excel's
        // limit. Same guard sparklines have on their own `style`.
        if let Some(style) = view.i64("style")? {
            if !(1..=48).contains(&style) {
                return Err(format!(
                    "charts['{}']: 'style' must be in the range 1-48, got {}",
                    cell_ref, style
                ));
            }
            chart.set_style(style as u8);
        }
        if view.bool("show_data_table")?.unwrap_or(false) {
            chart.set_data_table(&ChartDataTable::default());
        }
        if !view.bool("show_legend")?.unwrap_or(true) {
            chart.legend().set_hidden();
        }
        if let Some(position) = view.string("legend_position")? {
            let parsed = parse_legend_position(&position)
                .map_err(|e| format!("charts['{}']: {}", cell_ref, e))?;
            chart.legend().set_position(parsed);
        }

        let x_offset = view.u32("x_offset")?.unwrap_or(0);
        let y_offset = view.u32("y_offset")?.unwrap_or(0);
        worksheet
            .insert_chart_with_offset(row, col, &chart, x_offset, y_offset)
            .map_err(|e| format!("charts['{}']: {}", cell_ref, e))?;
    }

    Ok(())
}
