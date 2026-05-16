//! Image, checkbox, and textbox application helpers.

use crate::extract::pydict_to_hashmap;
use crate::parse::{parse_cell_ref, parse_color_enum, parse_column_format};
use crate::types::{ChartConfig, CheckboxConfig, ImageConfig, TextboxConfig};
use pyo3::prelude::*;
use rust_xlsxwriter::{
    Chart, ChartDataTable, ChartLegendPosition, ChartType, Color, Image, Shape, ShapeFont,
    ShapeFormat, ShapeLine, ShapeSolidFill, Worksheet,
};
use std::collections::HashMap;

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
    for (cell_ref, config) in images {
        let (row, col) = parse_cell_ref(cell_ref)?;

        let mut image = Image::new(&config.path)
            .map_err(|e| format!("Failed to load image '{}': {}", config.path, e))?;

        // Apply options if provided
        if let Some(opts) = &config.options {
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
    for (cell_ref, config) in checkboxes {
        let (row, col) = parse_cell_ref(cell_ref)?;

        if let Some(fmt_dict) = &config.format {
            let fmt = parse_column_format(py, fmt_dict)
                .map_err(|e| format!("checkboxes['{}']: {}", cell_ref, e))?;
            worksheet
                .insert_checkbox_with_format(row, col, config.checked, &fmt)
                .map_err(|e| format!("checkboxes['{}']: {}", cell_ref, e))?;
        } else {
            worksheet
                .insert_checkbox(row, col, config.checked)
                .map_err(|e| format!("checkboxes['{}']: {}", cell_ref, e))?;
        }
    }

    Ok(())
}

/// Extract an optional u32 option from a textbox/shape options dict.
fn textbox_u32_field(
    py: Python<'_>,
    opts: &HashMap<String, Py<PyAny>>,
    cell_ref: &str,
    key: &str,
) -> Result<Option<u32>, String> {
    let Some(obj) = opts.get(key) else {
        return Ok(None);
    };
    let bound = obj.bind(py);
    if bound.is_none() {
        return Ok(None);
    }
    bound.extract::<u32>().map(Some).map_err(|_| {
        format!(
            "textboxes['{}']: '{}' must be a non-negative integer",
            cell_ref, key
        )
    })
}

/// Extract an optional string option from a textbox options dict.
fn textbox_string_field(
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
        .map_err(|_| format!("textboxes['{}']: '{}' must be a string", cell_ref, key))
}

/// Extract an optional color option from a textbox options dict.
fn textbox_color_field(
    py: Python<'_>,
    opts: &HashMap<String, Py<PyAny>>,
    cell_ref: &str,
    key: &str,
) -> Result<Option<Color>, String> {
    let Some(s) = textbox_string_field(py, opts, cell_ref, key)? else {
        return Ok(None);
    };
    parse_color_enum(&s)
        .map(Some)
        .map_err(|e| format!("textboxes['{}']: '{}': {}", cell_ref, key, e))
}

/// Build a `ShapeFont` from a font option dict, validating keys.
fn build_shape_font(
    py: Python<'_>,
    cell_ref: &str,
    font_dict: &HashMap<String, Py<PyAny>>,
) -> Result<ShapeFont, String> {
    const FONT_KEYS: &[&str] = &["name", "size", "bold", "italic", "underline", "color"];
    for key in font_dict.keys() {
        if !FONT_KEYS.contains(&key.as_str()) {
            return Err(format!(
                "textboxes['{}']: unknown font option '{}'. Valid: {}",
                cell_ref,
                key,
                FONT_KEYS.join(", ")
            ));
        }
    }

    let mut font = ShapeFont::new();
    if let Some(name) = font_dict.get("name") {
        let s = name
            .bind(py)
            .extract::<String>()
            .map_err(|_| format!("textboxes['{}']: font.name must be a string", cell_ref))?;
        font = font.set_name(&s);
    }
    if let Some(size) = font_dict.get("size") {
        let n = size
            .bind(py)
            .extract::<f64>()
            .map_err(|_| format!("textboxes['{}']: font.size must be a number", cell_ref))?;
        font = font.set_size(n);
    }
    if let Some(b) = font_dict.get("bold") {
        if b.bind(py)
            .extract::<bool>()
            .map_err(|_| format!("textboxes['{}']: font.bold must be a bool", cell_ref))?
        {
            font = font.set_bold();
        }
    }
    if let Some(i) = font_dict.get("italic") {
        if i.bind(py)
            .extract::<bool>()
            .map_err(|_| format!("textboxes['{}']: font.italic must be a bool", cell_ref))?
        {
            font = font.set_italic();
        }
    }
    if let Some(u) = font_dict.get("underline") {
        if u.bind(py)
            .extract::<bool>()
            .map_err(|_| format!("textboxes['{}']: font.underline must be a bool", cell_ref))?
        {
            font = font.set_underline();
        }
    }
    if let Some(c) = font_dict.get("color") {
        let s = c
            .bind(py)
            .extract::<String>()
            .map_err(|_| format!("textboxes['{}']: font.color must be a string", cell_ref))?;
        let color = parse_color_enum(&s)
            .map_err(|e| format!("textboxes['{}']: font.color: {}", cell_ref, e))?;
        font = font.set_color(color);
    }
    Ok(font)
}

/// Apply textboxes (floating text shapes) to worksheet
pub(crate) fn apply_textboxes(
    py: Python<'_>,
    worksheet: &mut Worksheet,
    textboxes: &HashMap<String, TextboxConfig>,
) -> Result<(), String> {
    const TEXTBOX_KEYS: &[&str] = &[
        "text",
        "width",
        "height",
        "x_offset",
        "y_offset",
        "font",
        "fill_color",
        "line_color",
        "alt_text",
    ];

    for (cell_ref, config) in textboxes {
        let (row, col) = parse_cell_ref(cell_ref)?;

        let mut shape = Shape::textbox().set_text(config.text.as_str());

        let (x_offset, y_offset) = if let Some(opts) = &config.options {
            for key in opts.keys() {
                if !TEXTBOX_KEYS.contains(&key.as_str()) {
                    return Err(format!(
                        "textboxes['{}']: unknown option '{}'. Valid: {}",
                        cell_ref,
                        key,
                        TEXTBOX_KEYS.join(", ")
                    ));
                }
            }

            if let Some(w) = textbox_u32_field(py, opts, cell_ref, "width")? {
                shape = shape.set_width(w);
            }
            if let Some(h) = textbox_u32_field(py, opts, cell_ref, "height")? {
                shape = shape.set_height(h);
            }
            if let Some(alt) = textbox_string_field(py, opts, cell_ref, "alt_text")? {
                shape = shape.set_alt_text(&alt);
            }

            if let Some(font_obj) = opts.get("font") {
                let bound = font_obj.bind(py);
                if !bound.is_none() {
                    let font_dict = bound
                        .cast::<pyo3::types::PyDict>()
                        .map_err(|_| format!("textboxes['{}']: 'font' must be a dict", cell_ref))?;
                    let font_map = pydict_to_hashmap(font_dict)
                        .map_err(|e| format!("textboxes['{}']: {}", cell_ref, e))?;
                    let font = build_shape_font(py, cell_ref, &font_map)?;
                    shape = shape.set_font(&font);
                }
            }

            let fill = textbox_color_field(py, opts, cell_ref, "fill_color")?;
            let line = textbox_color_field(py, opts, cell_ref, "line_color")?;
            if fill.is_some() || line.is_some() {
                let mut format = ShapeFormat::new();
                if let Some(c) = fill {
                    format = format.set_solid_fill(&ShapeSolidFill::new().set_color(c));
                }
                if let Some(c) = line {
                    format = format.set_line(&ShapeLine::new().set_color(c));
                }
                shape = shape.set_format(&format);
            }

            let x = textbox_u32_field(py, opts, cell_ref, "x_offset")?.unwrap_or(0);
            let y = textbox_u32_field(py, opts, cell_ref, "y_offset")?.unwrap_or(0);
            (x, y)
        } else {
            (0, 0)
        };

        // insert_shape itself delegates to insert_shape_with_offset(row, col, shape, 0, 0),
        // so one call handles both offset and no-offset paths.
        worksheet
            .insert_shape_with_offset(row, col, &shape, x_offset, y_offset)
            .map_err(|e| format!("textboxes['{}']: {}", cell_ref, e))?;
    }

    Ok(())
}

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

fn chart_string_field(
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
        .map_err(|_| format!("charts['{}']: '{}' must be a string", cell_ref, key))
}

fn chart_u32_field(
    py: Python<'_>,
    opts: &HashMap<String, Py<PyAny>>,
    cell_ref: &str,
    key: &str,
) -> Result<Option<u32>, String> {
    let Some(obj) = opts.get(key) else {
        return Ok(None);
    };
    let bound = obj.bind(py);
    if bound.is_none() {
        return Ok(None);
    }
    bound.extract::<u32>().map(Some).map_err(|_| {
        format!(
            "charts['{}']: '{}' must be a non-negative integer",
            cell_ref, key
        )
    })
}

fn chart_u8_field(
    py: Python<'_>,
    opts: &HashMap<String, Py<PyAny>>,
    cell_ref: &str,
    key: &str,
) -> Result<Option<u8>, String> {
    let Some(obj) = opts.get(key) else {
        return Ok(None);
    };
    let bound = obj.bind(py);
    if bound.is_none() {
        return Ok(None);
    }
    bound
        .extract::<u8>()
        .map(Some)
        .map_err(|_| format!("charts['{}']: '{}' must be an integer 0-255", cell_ref, key))
}

fn chart_bool_field(
    py: Python<'_>,
    opts: &HashMap<String, Py<PyAny>>,
    cell_ref: &str,
    key: &str,
) -> Result<Option<bool>, String> {
    let Some(obj) = opts.get(key) else {
        return Ok(None);
    };
    let bound = obj.bind(py);
    if bound.is_none() {
        return Ok(None);
    }
    bound
        .extract::<bool>()
        .map(Some)
        .map_err(|_| format!("charts['{}']: '{}' must be a bool", cell_ref, key))
}

fn add_chart_series(
    py: Python<'_>,
    chart: &mut Chart,
    cell_ref: &str,
    config: &HashMap<String, Py<PyAny>>,
    default_categories: Option<&str>,
) -> Result<(), String> {
    let values = chart_string_field(py, config, cell_ref, "values_range")?
        .or(chart_string_field(py, config, cell_ref, "values")?)
        .or(chart_string_field(py, config, cell_ref, "data_range")?)
        .ok_or_else(|| {
            format!(
                "charts['{}']: chart series requires 'values_range', 'values', or 'data_range'",
                cell_ref
            )
        })?;

    let categories = chart_string_field(py, config, cell_ref, "categories_range")?
        .or(chart_string_field(py, config, cell_ref, "categories")?)
        .or_else(|| default_categories.map(str::to_string));
    let name = chart_string_field(py, config, cell_ref, "name")?.or(chart_string_field(
        py,
        config,
        cell_ref,
        "series_name",
    )?);

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
    charts: &HashMap<String, ChartConfig>,
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

    for (cell_ref, config) in charts {
        let (row, col) = parse_cell_ref(cell_ref)?;
        for key in config.keys() {
            if !CHART_KEYS.contains(&key.as_str()) {
                return Err(format!(
                    "charts['{}']: unknown option '{}'. Valid: {}",
                    cell_ref,
                    key,
                    CHART_KEYS.join(", ")
                ));
            }
        }

        let chart_type = chart_string_field(py, config, cell_ref, "type")?
            .ok_or_else(|| format!("charts['{}']: missing 'type' key", cell_ref))?;
        let mut chart = Chart::new(parse_chart_type(&chart_type)?);

        if config
            .get("series")
            .is_some_and(|series_obj| !series_obj.bind(py).is_none())
        {
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
            for (idx, item) in series_list.iter().enumerate() {
                let series_dict = item.cast::<pyo3::types::PyDict>().map_err(|_| {
                    format!("charts['{}']: series item {} must be a dict", cell_ref, idx)
                })?;
                let series_config = pydict_to_hashmap(series_dict)
                    .map_err(|e| format!("charts['{}']: {}", cell_ref, e))?;
                let default_categories =
                    chart_string_field(py, config, cell_ref, "categories_range")?
                        .or(chart_string_field(py, config, cell_ref, "categories")?);
                add_chart_series(
                    py,
                    &mut chart,
                    cell_ref,
                    &series_config,
                    default_categories.as_deref(),
                )?;
            }
        } else {
            add_chart_series(py, &mut chart, cell_ref, config, None)?;
        }

        if let Some(title) = chart_string_field(py, config, cell_ref, "title")? {
            chart.title().set_name(title.as_str());
        }
        if let Some(name) = chart_string_field(py, config, cell_ref, "x_axis_name")? {
            chart.x_axis().set_name(name.as_str());
        }
        if let Some(name) = chart_string_field(py, config, cell_ref, "y_axis_name")? {
            chart.y_axis().set_name(name.as_str());
        }
        if let Some(width) = chart_u32_field(py, config, cell_ref, "width")? {
            chart.set_width(width);
        }
        if let Some(height) = chart_u32_field(py, config, cell_ref, "height")? {
            chart.set_height(height);
        }
        if let Some(style) = chart_u8_field(py, config, cell_ref, "style")? {
            chart.set_style(style);
        }
        if chart_bool_field(py, config, cell_ref, "show_data_table")?.unwrap_or(false) {
            chart.set_data_table(&ChartDataTable::default());
        }
        if !chart_bool_field(py, config, cell_ref, "show_legend")?.unwrap_or(true) {
            chart.legend().set_hidden();
        }
        if let Some(position) = chart_string_field(py, config, cell_ref, "legend_position")? {
            let parsed = parse_legend_position(&position)
                .map_err(|e| format!("charts['{}']: {}", cell_ref, e))?;
            chart.legend().set_position(parsed);
        }

        let x_offset = chart_u32_field(py, config, cell_ref, "x_offset")?.unwrap_or(0);
        let y_offset = chart_u32_field(py, config, cell_ref, "y_offset")?.unwrap_or(0);
        worksheet
            .insert_chart_with_offset(row, col, &chart, x_offset, y_offset)
            .map_err(|e| format!("charts['{}']: {}", cell_ref, e))?;
    }

    Ok(())
}
