//! Feature extraction and application functions

use crate::parse::{
    matches_pattern, parse_cell_range, parse_cell_ref, parse_color, parse_header_format,
    parse_icon_type,
};
use crate::types::*;
use indexmap::IndexMap;
use pyo3::prelude::*;
use rust_xlsxwriter::{
    ConditionalFormat2ColorScale, ConditionalFormat3ColorScale, ConditionalFormatDataBar,
    ConditionalFormatDataBarDirection, ConditionalFormatIconSet, DataValidation,
    DataValidationErrorStyle, Format, Image, Note, Worksheet,
};
use std::collections::HashMap;

/// Extract sheet info from a Python tuple (supports both 2-tuple and 3-tuple formats)
/// 2-tuple: (df, sheet_name)
/// 3-tuple: (df, sheet_name, options_dict)
pub(crate) fn extract_sheet_info<'py>(
    sheet_tuple: &Bound<'py, PyAny>,
) -> PyResult<(Bound<'py, PyAny>, String, SheetConfig)> {
    let len: usize = sheet_tuple.len()?;

    if len < 2 {
        return Err(pyo3::exceptions::PyValueError::new_err(
            "Sheet tuple must have at least 2 elements: (df, sheet_name)",
        ));
    }

    let df = sheet_tuple.get_item(0)?;
    let sheet_name: String = sheet_tuple.get_item(1)?.extract()?;

    let config = if len >= 3 {
        let opts = sheet_tuple.get_item(2)?;
        let mut config = SheetConfig::default();

        // Extract optional fields from the dict
        if let Ok(val) = opts.get_item("header") {
            if !val.is_none() {
                config.header = Some(val.extract()?);
            }
        }
        if let Ok(val) = opts.get_item("autofit") {
            if !val.is_none() {
                config.autofit = Some(val.extract()?);
            }
        }
        if let Ok(val) = opts.get_item("table_style") {
            // Handle both None and string values
            if val.is_none() {
                config.table_style = Some(None); // Explicitly no style
            } else {
                config.table_style = Some(Some(val.extract()?));
            }
        }
        if let Ok(val) = opts.get_item("freeze_panes") {
            if !val.is_none() {
                config.freeze_panes = Some(val.extract()?);
            }
        }
        if let Ok(val) = opts.get_item("column_widths") {
            if !val.is_none() {
                // Support both integer keys {0: 20} and string keys {"_all": 50}
                let mut widths: HashMap<String, f64> = HashMap::new();
                if let Ok(dict) = val.cast::<pyo3::types::PyDict>() {
                    for (k, v) in dict.iter() {
                        let key_str = if let Ok(i) = k.extract::<i64>() {
                            i.to_string()
                        } else {
                            k.extract::<String>()?
                        };
                        widths.insert(key_str, v.extract()?);
                    }
                }
                if !widths.is_empty() {
                    config.column_widths = Some(widths);
                }
            }
        }
        if let Ok(val) = opts.get_item("row_heights") {
            if !val.is_none() {
                config.row_heights = Some(val.extract()?);
            }
        }
        if let Ok(val) = opts.get_item("table_name") {
            if !val.is_none() {
                config.table_name = Some(val.extract()?);
            }
        }
        if let Ok(val) = opts.get_item("header_format") {
            if !val.is_none() {
                let mut fmt: HashMap<String, Py<PyAny>> = HashMap::new();
                if let Ok(dict) = val.cast::<pyo3::types::PyDict>() {
                    for (k, v) in dict.iter() {
                        fmt.insert(k.extract()?, v.unbind());
                    }
                }
                if !fmt.is_empty() {
                    config.header_format = Some(fmt);
                }
            }
        }
        if let Ok(val) = opts.get_item("column_formats") {
            if !val.is_none() {
                if let Ok(outer_dict) = val.cast::<pyo3::types::PyDict>() {
                    let mut col_fmts: IndexMap<String, HashMap<String, Py<PyAny>>> =
                        IndexMap::new();
                    for (pattern, fmt_dict) in outer_dict.iter() {
                        let pattern_str: String = pattern.extract()?;
                        if let Ok(inner_dict) = fmt_dict.cast::<pyo3::types::PyDict>() {
                            let mut fmt: HashMap<String, Py<PyAny>> = HashMap::new();
                            for (k, v) in inner_dict.iter() {
                                fmt.insert(k.extract()?, v.unbind());
                            }
                            col_fmts.insert(pattern_str, fmt);
                        }
                    }
                    if !col_fmts.is_empty() {
                        config.column_formats = Some(col_fmts);
                    }
                }
            }
        }
        if let Ok(val) = opts.get_item("conditional_formats") {
            if !val.is_none() {
                if let Ok(outer_dict) = val.cast::<pyo3::types::PyDict>() {
                    let mut cond_fmts: IndexMap<String, HashMap<String, Py<PyAny>>> =
                        IndexMap::new();
                    for (col_name, fmt_dict) in outer_dict.iter() {
                        let col_str: String = col_name.extract()?;
                        if let Ok(inner_dict) = fmt_dict.cast::<pyo3::types::PyDict>() {
                            let mut fmt: HashMap<String, Py<PyAny>> = HashMap::new();
                            for (k, v) in inner_dict.iter() {
                                fmt.insert(k.extract()?, v.unbind());
                            }
                            cond_fmts.insert(col_str, fmt);
                        }
                    }
                    if !cond_fmts.is_empty() {
                        config.conditional_formats = Some(cond_fmts);
                    }
                }
            }
        }
        if let Ok(val) = opts.get_item("formula_columns") {
            if !val.is_none() {
                if let Ok(dict) = val.cast::<pyo3::types::PyDict>() {
                    let mut formulas: IndexMap<String, String> = IndexMap::new();
                    for (col_name, formula) in dict.iter() {
                        let col_str: String = col_name.extract()?;
                        let formula_str: String = formula.extract()?;
                        formulas.insert(col_str, formula_str);
                    }
                    if !formulas.is_empty() {
                        config.formula_columns = Some(formulas);
                    }
                }
            }
        }
        if let Ok(val) = opts.get_item("merged_ranges") {
            if !val.is_none() {
                if let Ok(list) = val.cast::<pyo3::types::PyList>() {
                    let extracted = extract_merged_ranges(list)?;
                    if !extracted.is_empty() {
                        config.merged_ranges = Some(extracted);
                    }
                }
            }
        }
        if let Ok(val) = opts.get_item("hyperlinks") {
            if !val.is_none() {
                if let Ok(list) = val.cast::<pyo3::types::PyList>() {
                    let extracted = extract_hyperlinks(list)?;
                    if !extracted.is_empty() {
                        config.hyperlinks = Some(extracted);
                    }
                }
            }
        }
        if let Ok(val) = opts.get_item("comments") {
            if !val.is_none() {
                if let Ok(dict) = val.cast::<pyo3::types::PyDict>() {
                    let extracted = extract_comments(dict)?;
                    if !extracted.is_empty() {
                        config.comments = Some(extracted);
                    }
                }
            }
        }
        if let Ok(val) = opts.get_item("validations") {
            if !val.is_none() {
                if let Ok(dict) = val.cast::<pyo3::types::PyDict>() {
                    let extracted = extract_validations(dict)?;
                    if !extracted.is_empty() {
                        config.validations = Some(extracted);
                    }
                }
            }
        }
        if let Ok(val) = opts.get_item("rich_text") {
            if !val.is_none() {
                if let Ok(dict) = val.cast::<pyo3::types::PyDict>() {
                    let extracted = extract_rich_text(dict)?;
                    if !extracted.is_empty() {
                        config.rich_text = Some(extracted);
                    }
                }
            }
        }
        if let Ok(val) = opts.get_item("images") {
            if !val.is_none() {
                if let Ok(dict) = val.cast::<pyo3::types::PyDict>() {
                    let extracted = extract_images(dict)?;
                    if !extracted.is_empty() {
                        config.images = Some(extracted);
                    }
                }
            }
        }

        config
    } else {
        SheetConfig::default()
    };

    Ok((df, sheet_name, config))
}

/// Extract column_widths from Python dict, supporting both integer and string keys
pub(crate) fn extract_column_widths(
    py_dict: &Bound<'_, pyo3::types::PyDict>,
) -> PyResult<HashMap<String, f64>> {
    let mut widths: HashMap<String, f64> = HashMap::new();
    for (k, v) in py_dict.iter() {
        let key_str = if let Ok(i) = k.extract::<i64>() {
            i.to_string()
        } else {
            k.extract::<String>()?
        };
        widths.insert(key_str, v.extract()?);
    }
    Ok(widths)
}

/// Extract header_format from Python dict
pub(crate) fn extract_header_format(
    py_dict: &Bound<'_, pyo3::types::PyDict>,
) -> PyResult<HashMap<String, Py<PyAny>>> {
    let mut fmt: HashMap<String, Py<PyAny>> = HashMap::new();
    for (k, v) in py_dict.iter() {
        fmt.insert(k.extract()?, v.unbind());
    }
    Ok(fmt)
}

/// Extract column_formats from Python dict (pattern -> format dict)
/// Uses IndexMap to preserve insertion order from Python dict
pub(crate) fn extract_column_formats(
    py_dict: &Bound<'_, pyo3::types::PyDict>,
) -> PyResult<IndexMap<String, HashMap<String, Py<PyAny>>>> {
    let mut col_fmts: IndexMap<String, HashMap<String, Py<PyAny>>> = IndexMap::new();
    for (pattern, fmt_dict) in py_dict.iter() {
        let pattern_str: String = pattern.extract()?;
        if let Ok(inner_dict) = fmt_dict.cast::<pyo3::types::PyDict>() {
            let mut fmt: HashMap<String, Py<PyAny>> = HashMap::new();
            for (k, v) in inner_dict.iter() {
                fmt.insert(k.extract()?, v.unbind());
            }
            col_fmts.insert(pattern_str, fmt);
        }
    }
    Ok(col_fmts)
}

/// Extract conditional_formats from Python dict (column/pattern -> config dict)
/// Uses IndexMap to preserve insertion order for pattern matching (first match wins)
pub(crate) fn extract_conditional_formats(
    py_dict: &Bound<'_, pyo3::types::PyDict>,
) -> PyResult<IndexMap<String, HashMap<String, Py<PyAny>>>> {
    let mut cond_fmts: IndexMap<String, HashMap<String, Py<PyAny>>> = IndexMap::new();
    for (col_name, fmt_dict) in py_dict.iter() {
        let col_str: String = col_name.extract()?;
        if let Ok(inner_dict) = fmt_dict.cast::<pyo3::types::PyDict>() {
            let mut fmt: HashMap<String, Py<PyAny>> = HashMap::new();
            for (k, v) in inner_dict.iter() {
                fmt.insert(k.extract()?, v.unbind());
            }
            cond_fmts.insert(col_str, fmt);
        }
    }
    Ok(cond_fmts)
}

/// Extract formula_columns from Python dict (column name -> formula template)
/// Uses IndexMap to preserve column order
pub(crate) fn extract_formula_columns(
    py_dict: &Bound<'_, pyo3::types::PyDict>,
) -> PyResult<IndexMap<String, String>> {
    let mut formulas: IndexMap<String, String> = IndexMap::new();
    for (col_name, formula) in py_dict.iter() {
        let col_str: String = col_name.extract()?;
        let formula_str: String = formula.extract()?;
        formulas.insert(col_str, formula_str);
    }
    Ok(formulas)
}

/// Extract merged_ranges from Python list of tuples
/// Each tuple: (range_str, text) or (range_str, text, format_dict)
pub(crate) fn extract_merged_ranges(
    py_list: &Bound<'_, pyo3::types::PyList>,
) -> PyResult<Vec<MergedRange>> {
    let mut ranges = Vec::new();

    for item in py_list.iter() {
        let tuple_len = item.len()?;
        if tuple_len < 2 {
            return Err(pyo3::exceptions::PyValueError::new_err(
                "merged_ranges tuple must have at least 2 elements: (range, text)",
            ));
        }

        let range_str: String = item.get_item(0)?.extract()?;
        let text: String = item.get_item(1)?.extract()?;

        let format_dict = if tuple_len >= 3 {
            let fmt_item = item.get_item(2)?;
            if !fmt_item.is_none() {
                if let Ok(dict) = fmt_item.cast::<pyo3::types::PyDict>() {
                    let mut fmt_map: HashMap<String, Py<PyAny>> = HashMap::new();
                    for (k, v) in dict.iter() {
                        let key: String = k.extract()?;
                        fmt_map.insert(key, v.unbind());
                    }
                    Some(fmt_map)
                } else {
                    None
                }
            } else {
                None
            }
        } else {
            None
        };

        ranges.push((range_str, text, format_dict));
    }

    Ok(ranges)
}

/// Extract hyperlinks from Python list of tuples
/// Each tuple: (cell_ref, url) or (cell_ref, url, display_text)
pub(crate) fn extract_hyperlinks(
    py_list: &Bound<'_, pyo3::types::PyList>,
) -> PyResult<Vec<Hyperlink>> {
    let mut links = Vec::new();

    for item in py_list.iter() {
        let tuple_len = item.len()?;
        if tuple_len < 2 {
            return Err(pyo3::exceptions::PyValueError::new_err(
                "hyperlinks tuple must have at least 2 elements: (cell_ref, url)",
            ));
        }

        let cell_ref: String = item.get_item(0)?.extract()?;
        let url: String = item.get_item(1)?.extract()?;

        let display_text = if tuple_len >= 3 {
            let text_item = item.get_item(2)?;
            if !text_item.is_none() {
                Some(text_item.extract()?)
            } else {
                None
            }
        } else {
            None
        };

        links.push((cell_ref, url, display_text));
    }

    Ok(links)
}

/// Extract comments from Python dict
/// Supports: {'A1': 'text'} or {'A1': {'text': 'note', 'author': 'John'}}
pub(crate) fn extract_comments(
    py_dict: &Bound<'_, pyo3::types::PyDict>,
) -> PyResult<HashMap<String, Comment>> {
    let mut comments: HashMap<String, Comment> = HashMap::new();

    for (cell_ref, value) in py_dict.iter() {
        let cell_str: String = cell_ref.extract()?;

        // Check if value is a dict or simple string
        if let Ok(inner_dict) = value.cast::<pyo3::types::PyDict>() {
            // Dict format: {'text': '...', 'author': '...'}
            let text: String = inner_dict
                .get_item("text")?
                .ok_or_else(|| {
                    pyo3::exceptions::PyValueError::new_err(format!(
                        "Comment at '{}' missing 'text' key",
                        cell_str
                    ))
                })?
                .extract()?;
            let author: Option<String> = if let Ok(Some(a)) = inner_dict.get_item("author") {
                if !a.is_none() {
                    Some(a.extract()?)
                } else {
                    None
                }
            } else {
                None
            };
            comments.insert(cell_str, (text, author));
        } else {
            // Simple string format
            let text: String = value.extract()?;
            comments.insert(cell_str, (text, None));
        }
    }

    Ok(comments)
}

/// Extract validations from Python dict (column name/pattern -> validation config)
pub(crate) fn extract_validations(
    py_dict: &Bound<'_, pyo3::types::PyDict>,
) -> PyResult<IndexMap<String, ValidationConfig>> {
    let mut validations: IndexMap<String, ValidationConfig> = IndexMap::new();
    for (col_name, config) in py_dict.iter() {
        let col_str: String = col_name.extract()?;
        if let Ok(inner_dict) = config.cast::<pyo3::types::PyDict>() {
            let mut cfg: ValidationConfig = HashMap::new();
            for (k, v) in inner_dict.iter() {
                cfg.insert(k.extract()?, v.unbind());
            }
            validations.insert(col_str, cfg);
        } else {
            return Err(pyo3::exceptions::PyTypeError::new_err(format!(
                "validations['{}']: expected dict, got {}",
                col_str,
                config.get_type().name()?
            )));
        }
    }
    Ok(validations)
}

/// Extract rich_text from Python dict (cell_ref -> list of segments)
pub(crate) fn extract_rich_text(
    py_dict: &Bound<'_, pyo3::types::PyDict>,
) -> PyResult<HashMap<String, Vec<RichTextSegment>>> {
    let mut rich_text: HashMap<String, Vec<RichTextSegment>> = HashMap::new();

    for (cell_ref, segments_list) in py_dict.iter() {
        let cell_str: String = cell_ref.extract()?;
        let mut segments: Vec<RichTextSegment> = Vec::new();

        if let Ok(list) = segments_list.cast::<pyo3::types::PyList>() {
            for (idx, item) in list.iter().enumerate() {
                // Check if item is a tuple (text, format_dict) or just a string
                if let Ok(tuple) = item.cast::<pyo3::types::PyTuple>() {
                    let text: String = tuple.get_item(0)?.extract()?;
                    let format_dict = if tuple.len() >= 2 {
                        let fmt_item = tuple.get_item(1)?;
                        if let Ok(dict) = fmt_item.cast::<pyo3::types::PyDict>() {
                            let mut fmt: HashMap<String, Py<PyAny>> = HashMap::new();
                            for (k, v) in dict.iter() {
                                fmt.insert(k.extract()?, v.unbind());
                            }
                            Some(fmt)
                        } else {
                            None
                        }
                    } else {
                        None
                    };
                    segments.push((text, format_dict));
                } else if let Ok(text) = item.extract::<String>() {
                    // Plain string segment
                    segments.push((text, None));
                } else {
                    return Err(pyo3::exceptions::PyTypeError::new_err(format!(
                        "rich_text['{}']: segment {} must be a string or tuple (text, format_dict), got {}",
                        cell_str,
                        idx,
                        item.get_type().name()?
                    )));
                }
            }
        } else {
            return Err(pyo3::exceptions::PyTypeError::new_err(format!(
                "rich_text['{}']: expected list of segments, got {}",
                cell_str,
                segments_list.get_type().name()?
            )));
        }

        if !segments.is_empty() {
            rich_text.insert(cell_str, segments);
        }
    }

    Ok(rich_text)
}

/// Extract images from Python dict (cell_ref -> path or config dict)
pub(crate) fn extract_images(
    py_dict: &Bound<'_, pyo3::types::PyDict>,
) -> PyResult<HashMap<String, ImageConfig>> {
    let mut images: HashMap<String, ImageConfig> = HashMap::new();

    for (cell_ref, value) in py_dict.iter() {
        let cell_str: String = cell_ref.extract()?;

        // Check if value is a dict or simple string (path)
        if let Ok(inner_dict) = value.cast::<pyo3::types::PyDict>() {
            // Dict format: {'path': '...', 'scale_width': 0.5, ...}
            let path: String = inner_dict
                .get_item("path")?
                .ok_or_else(|| {
                    pyo3::exceptions::PyValueError::new_err(format!(
                        "Image at '{}' missing 'path' key",
                        cell_str
                    ))
                })?
                .extract()?;
            let mut options: HashMap<String, Py<PyAny>> = HashMap::new();
            for (k, v) in inner_dict.iter() {
                let key: String = k.extract()?;
                if key != "path" {
                    options.insert(key, v.unbind());
                }
            }
            images.insert(cell_str, (path, Some(options)));
        } else {
            // Simple string format (just path)
            let path: String = value.extract()?;
            images.insert(cell_str, (path, None));
        }
    }

    Ok(images)
}

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

/// Apply column widths with autofit cap: autofit first, then cap columns at '_all' width
pub(crate) fn apply_column_widths_with_autofit_cap(
    worksheet: &mut Worksheet,
    col_count: u16,
    widths: &HashMap<String, f64>,
    constant_memory: bool,
) -> Result<(), String> {
    // First autofit
    if !constant_memory {
        worksheet.autofit();
    }

    // Then apply specific widths and cap at '_all' if specified
    let global_cap = widths.get("_all").copied();

    for col_idx in 0..col_count {
        let col_key = col_idx.to_string();
        if let Some(width) = widths.get(&col_key) {
            // Specific width overrides autofit completely
            worksheet
                .set_column_width(col_idx, *width)
                .map_err(|e| format!("Failed to set column width: {}", e))?;
        } else if let Some(cap) = global_cap {
            // '_all' acts as a cap - only set if current width exceeds cap
            // Since we can't read current width, just set the cap
            worksheet
                .set_column_width(col_idx, cap)
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
    header_format: Option<&Format>,
) -> Result<u16, String> {
    let mut col_offset = 0u16;

    for (col_name, formula_template) in formula_columns {
        let col_idx = start_col + col_offset;

        // Write header for formula column
        if let Some(fmt) = header_format {
            worksheet
                .write_string_with_format(0, col_idx, col_name, fmt)
                .map_err(|e| format!("Failed to write formula column header: {}", e))?;
        } else {
            worksheet
                .write_string(0, col_idx, col_name)
                .map_err(|e| format!("Failed to write formula column header: {}", e))?;
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

        col_offset += 1;
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
            .map(|(idx, _)| idx as u16)
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
            let validation = if let Some(msg_obj) = config.get("input_message") {
                if let Ok(msg) = msg_obj.bind(py).extract::<String>() {
                    let title = config
                        .get("input_title")
                        .and_then(|t| t.bind(py).extract::<String>().ok())
                        .unwrap_or_default();
                    validation
                        .set_input_title(&title)
                        .map_err(|e| format!("Failed to set input title: {}", e))?
                        .set_input_message(&msg)
                        .map_err(|e| format!("Failed to set input message: {}", e))?
                } else {
                    validation
                }
            } else {
                validation
            };

            // Add optional error message
            let validation = if let Some(msg_obj) = config.get("error_message") {
                if let Ok(msg) = msg_obj.bind(py).extract::<String>() {
                    let title = config
                        .get("error_title")
                        .and_then(|t| t.bind(py).extract::<String>().ok())
                        .unwrap_or_default();
                    validation
                        .set_error_title(&title)
                        .map_err(|e| format!("Failed to set error title: {}", e))?
                        .set_error_message(&msg)
                        .map_err(|e| format!("Failed to set error message: {}", e))?
                        .set_error_style(DataValidationErrorStyle::Stop)
                } else {
                    validation
                }
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
    use crate::parse::parse_column_format;

    for (cell_ref, segments) in rich_text {
        let (row, col) = parse_cell_ref(cell_ref)?;

        // Build formats and strings separately
        let mut formats: Vec<Format> = Vec::new();
        let mut texts: Vec<String> = Vec::new();

        for (text, format_dict) in segments {
            if let Some(fmt_dict) = format_dict {
                let format = parse_column_format(py, fmt_dict)?;
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
            if let Some(scale_obj) = opts.get("scale_width") {
                if let Ok(scale) = scale_obj.bind(py).extract::<f64>() {
                    image = image.set_scale_width(scale);
                }
            }
            if let Some(scale_obj) = opts.get("scale_height") {
                if let Ok(scale) = scale_obj.bind(py).extract::<f64>() {
                    image = image.set_scale_height(scale);
                }
            }
            if let Some(alt_obj) = opts.get("alt_text") {
                if let Ok(alt) = alt_obj.bind(py).extract::<String>() {
                    image = image.set_alt_text(&alt);
                }
            }
        }

        worksheet
            .insert_image(row, col, &image)
            .map_err(|e| format!("Failed to insert image at '{}': {}", cell_ref, e))?;
    }

    Ok(())
}

/// Apply conditional formats to a worksheet
/// Supports: 2_color_scale, 3_color_scale, data_bar, icon_set
/// Uses IndexMap to preserve pattern order (first match wins for overlapping patterns)
pub(crate) fn apply_conditional_formats(
    py: Python<'_>,
    worksheet: &mut Worksheet,
    columns: &[String],
    data_start_row: u32,
    data_end_row: u32,
    cond_formats: &IndexMap<String, HashMap<String, Py<PyAny>>>,
) -> Result<(), String> {
    for (col_pattern, config) in cond_formats {
        // Find column index by name (supports exact match or pattern)
        let col_indices: Vec<u16> = columns
            .iter()
            .enumerate()
            .filter(|(_, name)| matches_pattern(name, col_pattern))
            .map(|(idx, _)| idx as u16)
            .collect();

        if col_indices.is_empty() {
            continue; // Skip if no matching columns
        }

        // Get the format type
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

        for col_idx in col_indices {
            match format_type.to_lowercase().as_str() {
                "2_color_scale" | "2colorscale" | "two_color_scale" => {
                    let mut cf = ConditionalFormat2ColorScale::new();

                    // Parse min_color
                    if let Some(min_color_obj) = config.get("min_color") {
                        if let Ok(color_str) = min_color_obj.bind(py).extract::<String>() {
                            let color = parse_color(&color_str)?;
                            cf = cf.set_minimum_color(color);
                        }
                    }

                    // Parse max_color
                    if let Some(max_color_obj) = config.get("max_color") {
                        if let Ok(color_str) = max_color_obj.bind(py).extract::<String>() {
                            let color = parse_color(&color_str)?;
                            cf = cf.set_maximum_color(color);
                        }
                    }

                    worksheet
                        .add_conditional_format(data_start_row, col_idx, data_end_row, col_idx, &cf)
                        .map_err(|e| format!("Failed to add 2_color_scale: {}", e))?;
                }

                "3_color_scale" | "3colorscale" | "three_color_scale" => {
                    let mut cf = ConditionalFormat3ColorScale::new();

                    // Parse min_color
                    if let Some(min_color_obj) = config.get("min_color") {
                        if let Ok(color_str) = min_color_obj.bind(py).extract::<String>() {
                            let color = parse_color(&color_str)?;
                            cf = cf.set_minimum_color(color);
                        }
                    }

                    // Parse mid_color
                    if let Some(mid_color_obj) = config.get("mid_color") {
                        if let Ok(color_str) = mid_color_obj.bind(py).extract::<String>() {
                            let color = parse_color(&color_str)?;
                            cf = cf.set_midpoint_color(color);
                        }
                    }

                    // Parse max_color
                    if let Some(max_color_obj) = config.get("max_color") {
                        if let Ok(color_str) = max_color_obj.bind(py).extract::<String>() {
                            let color = parse_color(&color_str)?;
                            cf = cf.set_maximum_color(color);
                        }
                    }

                    worksheet
                        .add_conditional_format(data_start_row, col_idx, data_end_row, col_idx, &cf)
                        .map_err(|e| format!("Failed to add 3_color_scale: {}", e))?;
                }

                "data_bar" | "databar" => {
                    let mut cf = ConditionalFormatDataBar::new();

                    // Parse bar_color (fill color)
                    if let Some(color_obj) = config.get("bar_color") {
                        if let Ok(color_str) = color_obj.bind(py).extract::<String>() {
                            let color = parse_color(&color_str)?;
                            cf = cf.set_fill_color(color);
                        }
                    }

                    // Parse border_color
                    if let Some(color_obj) = config.get("border_color") {
                        if let Ok(color_str) = color_obj.bind(py).extract::<String>() {
                            let color = parse_color(&color_str)?;
                            cf = cf.set_border_color(color);
                        }
                    }

                    // Parse solid (vs gradient)
                    if let Some(solid_obj) = config.get("solid") {
                        if let Ok(solid) = solid_obj.bind(py).extract::<bool>() {
                            if solid {
                                cf = cf.set_solid_fill(true);
                            }
                        }
                    }

                    // Parse direction
                    if let Some(dir_obj) = config.get("direction") {
                        if let Ok(dir_str) = dir_obj.bind(py).extract::<String>() {
                            let direction = match dir_str.to_lowercase().as_str() {
                                "left_to_right" | "ltr" => {
                                    ConditionalFormatDataBarDirection::LeftToRight
                                }
                                "right_to_left" | "rtl" => {
                                    ConditionalFormatDataBarDirection::RightToLeft
                                }
                                "context" | "" => ConditionalFormatDataBarDirection::Context,
                                _ => {
                                    return Err(format!(
                                        "Unknown direction '{}'. Valid values: left_to_right, right_to_left, context",
                                        dir_str
                                    ));
                                }
                            };
                            cf = cf.set_direction(direction);
                        }
                    }

                    worksheet
                        .add_conditional_format(data_start_row, col_idx, data_end_row, col_idx, &cf)
                        .map_err(|e| format!("Failed to add data_bar: {}", e))?;
                }

                "icon_set" | "iconset" => {
                    let mut cf = ConditionalFormatIconSet::new();

                    // Parse icon_type
                    if let Some(icon_obj) = config.get("icon_type") {
                        if let Ok(icon_str) = icon_obj.bind(py).extract::<String>() {
                            let icon_type = parse_icon_type(&icon_str)?;
                            cf = cf.set_icon_type(icon_type);
                        }
                    }

                    // Parse reverse
                    if let Some(rev_obj) = config.get("reverse") {
                        if let Ok(reverse) = rev_obj.bind(py).extract::<bool>() {
                            if reverse {
                                cf = cf.reverse_icons(true);
                            }
                        }
                    }

                    // Parse icons_only (hide numbers, show only icons)
                    if let Some(icons_only_obj) = config.get("icons_only") {
                        if let Ok(icons_only) = icons_only_obj.bind(py).extract::<bool>() {
                            if icons_only {
                                cf = cf.show_icons_only(true);
                            }
                        }
                    }

                    worksheet
                        .add_conditional_format(data_start_row, col_idx, data_end_row, col_idx, &cf)
                        .map_err(|e| format!("Failed to add icon_set: {}", e))?;
                }

                _ => {
                    return Err(format!(
                        "Unknown conditional format type '{}'. Valid types: 2_color_scale, 3_color_scale, data_bar, icon_set",
                        format_type
                    ));
                }
            }
        }
    }

    Ok(())
}
