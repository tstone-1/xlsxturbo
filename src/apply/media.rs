//! Image, checkbox, and textbox application helpers.

use crate::parse::{parse_cell_ref, parse_color_enum, parse_column_format};
use crate::types::{extract_opt, pydict_to_hashmap, CheckboxConfig, ImageConfig, TextboxConfig};
use pyo3::prelude::*;
use rust_xlsxwriter::{
    Color, Image, Shape, ShapeFont, ShapeFormat, ShapeLine, ShapeSolidFill, Worksheet,
};
use std::collections::HashMap;

/// Extract an optional f64 image option. Wrong types produce an error.
fn image_f64_field(
    py: Python<'_>,
    opts: &HashMap<String, Py<PyAny>>,
    cell_ref: &str,
    key: &str,
) -> Result<Option<f64>, String> {
    extract_opt(py, opts.get(key), |_| {
        format!("images['{}']: '{}' must be a number", cell_ref, key)
    })
}

/// Extract an optional string image option. Wrong types produce an error.
fn image_string_field(
    py: Python<'_>,
    opts: &HashMap<String, Py<PyAny>>,
    cell_ref: &str,
    key: &str,
) -> Result<Option<String>, String> {
    extract_opt(py, opts.get(key), |_| {
        format!("images['{}']: '{}' must be a string", cell_ref, key)
    })
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
    extract_opt(py, opts.get(key), |_| {
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
    extract_opt(py, opts.get(key), |_| {
        format!("textboxes['{}']: '{}' must be a string", cell_ref, key)
    })
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
