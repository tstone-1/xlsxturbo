//! Image, checkbox, and textbox application helpers.

use crate::parse::{parse_cell_ref, parse_color_enum, parse_column_format};
use crate::types::{pydict_to_hashmap, CheckboxConfig, ImageConfig, OptionMap, TextboxConfig};
use indexmap::IndexMap;
use pyo3::prelude::*;
use rust_xlsxwriter::{Image, Shape, ShapeFont, ShapeFormat, ShapeLine, ShapeSolidFill, Worksheet};
use std::collections::HashMap;

/// Apply images to worksheet
pub(crate) fn apply_images(
    py: Python<'_>,
    worksheet: &mut Worksheet,
    images: &IndexMap<String, ImageConfig>,
) -> Result<(), String> {
    const IMAGE_KEYS: &[&str] = &["scale_width", "scale_height", "alt_text"];

    for (cell_ref, config) in images {
        let (row, col) = parse_cell_ref(cell_ref)?;

        let mut image = Image::new(&config.path)
            .map_err(|e| format!("Failed to load image '{}': {}", config.path, e))?;

        // Apply options if provided
        if let Some(opts) = &config.options {
            let view = OptionMap::new(py, opts, format!("images['{}']", cell_ref));
            view.reject_unknown(IMAGE_KEYS)?;
            if let Some(scale) = view.f64("scale_width")? {
                image = image.set_scale_width(scale);
            }
            if let Some(scale) = view.f64("scale_height")? {
                image = image.set_scale_height(scale);
            }
            if let Some(alt) = view.string("alt_text")? {
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
    checkboxes: &IndexMap<String, CheckboxConfig>,
) -> Result<(), String> {
    for (cell_ref, config) in checkboxes {
        let (row, col) = parse_cell_ref(cell_ref)?;

        if let Some(fmt_dict) = &config.format {
            let fmt = parse_column_format(py, fmt_dict, &format!("checkboxes['{}']", cell_ref))?;
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

/// Build a `ShapeFont` from a font option dict, validating keys.
fn build_shape_font(
    py: Python<'_>,
    cell_ref: &str,
    font_dict: &HashMap<String, Py<PyAny>>,
) -> Result<ShapeFont, String> {
    const FONT_KEYS: &[&str] = &["name", "size", "bold", "italic", "underline", "color"];
    let view = OptionMap::new(py, font_dict, format!("textboxes['{}']: font", cell_ref));
    view.reject_unknown(FONT_KEYS)?;

    let mut font = ShapeFont::new();
    if let Some(name) = view.string("name")? {
        font = font.set_name(&name);
    }
    if let Some(size) = view.f64("size")? {
        font = font.set_size(size);
    }
    if view.bool("bold")?.unwrap_or(false) {
        font = font.set_bold();
    }
    if view.bool("italic")?.unwrap_or(false) {
        font = font.set_italic();
    }
    if view.bool("underline")?.unwrap_or(false) {
        font = font.set_underline();
    }
    if let Some(s) = view.string("color")? {
        let color = parse_color_enum(&s)
            .map_err(|e| format!("textboxes['{}']: font: color: {}", cell_ref, e))?;
        font = font.set_color(color);
    }
    Ok(font)
}

/// Apply textboxes (floating text shapes) to worksheet
pub(crate) fn apply_textboxes(
    py: Python<'_>,
    worksheet: &mut Worksheet,
    textboxes: &IndexMap<String, TextboxConfig>,
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
            let view = OptionMap::new(py, opts, format!("textboxes['{}']", cell_ref));
            view.reject_unknown(TEXTBOX_KEYS)?;

            if let Some(w) = view.u32("width")? {
                shape = shape.set_width(w);
            }
            if let Some(h) = view.u32("height")? {
                shape = shape.set_height(h);
            }
            if let Some(alt) = view.string("alt_text")? {
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

            let fill = view
                .string("fill_color")?
                .map(|s| {
                    parse_color_enum(&s)
                        .map_err(|e| format!("textboxes['{}']: 'fill_color': {}", cell_ref, e))
                })
                .transpose()?;
            let line = view
                .string("line_color")?
                .map(|s| {
                    parse_color_enum(&s)
                        .map_err(|e| format!("textboxes['{}']: 'line_color': {}", cell_ref, e))
                })
                .transpose()?;
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

            let x = view.u32("x_offset")?.unwrap_or(0);
            let y = view.u32("y_offset")?.unwrap_or(0);
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
